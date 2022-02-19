Public Enum DeflateStrategy
    [Default] = 0
    Filtered = 1
    HuffmanOnly = 2
End Enum

Public Class DeflaterEngine
    Inherits DeflaterConstants

    Const TooFar As Integer = 4096

    Private ins_h As Integer
    Private head As Short()
    Private prev As Short()
    Private matchStart As Integer
    Private matchLen As Integer
    Private prevAvailable As Boolean
    Private blockStart As Integer
    Private strstart As Integer
    Private lookahead As Integer
    Private window As Byte()
    Private m_strategy As DeflateStrategy
    Private p_max_chain As Integer
    Private p_max_lazy As Integer
    Private niceLength As Integer
    Private goodLength As Integer
    Private compressionFunction As Integer
    Private inputBuf As Byte()
    Private m_totalIn As Long
    Private inputOff As Integer
    Private inputEnd As Integer
    Private mp_oPending As DeflaterPending
    Private huffman As DeflaterHuffman
    Private m_adler As Adler32

    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function

    Public Sub New(pending As DeflaterPending)
        mp_oPending = pending
        huffman = New DeflaterHuffman(mp_oPending)
        m_adler = New Adler32()
        window = New Byte(2 * WSIZE - 1) {}
        head = New Short(HASH_SIZE - 1) {}
        prev = New Short(WSIZE - 1) {}
        blockStart = InlineAssignHelper(strstart, 1)
    End Sub

    Public Function Deflate(flush As Boolean, finish As Boolean) As Boolean
        Dim progress As Boolean
        Do
            FillWindow()
            Dim canFlush As Boolean = flush AndAlso (inputOff = inputEnd)

            Select Case compressionFunction
                Case DEFLATE_STORED
                    progress = DeflateStored(canFlush, finish)
                    Exit Select
                Case DEFLATE_FAST
                    progress = DeflateFast(canFlush, finish)
                    Exit Select
                Case DEFLATE_SLOW
                    progress = DeflateSlow(canFlush, finish)
                    Exit Select
                Case Else
                    Throw New InvalidOperationException("unknown compressionFunction")
            End Select
        Loop While mp_oPending.IsFlushed AndAlso progress
        Return progress
    End Function


    Public Sub SetInput(buffer As Byte(), offset As Integer, count As Integer)
        If buffer Is Nothing Then
            Throw New ArgumentNullException("buffer")
        End If

        If offset < 0 Then
            Throw New ArgumentOutOfRangeException("offset")
        End If

        If count < 0 Then
            Throw New ArgumentOutOfRangeException("count")
        End If

        If inputOff < inputEnd Then
            Throw New InvalidOperationException("Old input was not completely processed")
        End If

        Dim [end] As Integer = offset + count


        If (offset > [end]) OrElse ([end] > buffer.Length) Then
            Throw New ArgumentOutOfRangeException("count")
        End If

        inputBuf = buffer
        inputOff = offset
        inputEnd = [end]
    End Sub


    Public Function NeedsInput() As Boolean
        Return (inputEnd = inputOff)
    End Function


    Public Sub SetDictionary(buffer As Byte(), offset As Integer, length As Integer)
        m_adler.Update(buffer, offset, length)
        If length < MIN_MATCH Then
            Return
        End If

        If length > MAX_DIST Then
            offset += length - MAX_DIST
            length = MAX_DIST
        End If

        System.Array.Copy(buffer, offset, window, strstart, length)

        UpdateHash()
        length -= 1
        'while (--length > 0) 
        While PreD(length) > 0
            InsertString()
            strstart += 1
        End While
        strstart += 2
        blockStart = strstart
    End Sub


    Public Sub Reset()
        huffman.Reset()
        m_adler.Reset()
        blockStart = InlineAssignHelper(strstart, 1)
        lookahead = 0
        m_totalIn = 0
        prevAvailable = False
        matchLen = MIN_MATCH - 1

        For i As Integer = 0 To HASH_SIZE - 1
            head(i) = 0
        Next

        For i As Integer = 0 To WSIZE - 1
            prev(i) = 0
        Next
    End Sub


    Public Sub ResetAdler()
        m_adler.Reset()
    End Sub


    Public ReadOnly Property Adler() As Integer
        Get
            Return ForceToInt32(m_adler.Value)
        End Get
    End Property

    Public ReadOnly Property TotalIn() As Long
        Get
            Return m_totalIn
        End Get
    End Property

    Public Property Strategy() As DeflateStrategy
        Get
            Return m_strategy
        End Get
        Set(value As DeflateStrategy)
            m_strategy = value
        End Set
    End Property

    Public Sub SetLevel(level As Integer)
        If (level < 0) OrElse (level > 9) Then
            Throw New ArgumentOutOfRangeException("level")
        End If

        goodLength = DeflaterConstants.GOOD_LENGTH(level)
        p_max_lazy = DeflaterConstants.MAX_LAZY(level)
        niceLength = DeflaterConstants.NICE_LENGTH(level)
        p_max_chain = DeflaterConstants.MAX_CHAIN(level)

        If DeflaterConstants.COMPR_FUNC(level) <> compressionFunction Then

            Select Case compressionFunction
                Case DEFLATE_STORED
                    If strstart > blockStart Then
                        huffman.FlushStoredBlock(window, blockStart, strstart - blockStart, False)
                        blockStart = strstart
                    End If
                    UpdateHash()
                    Exit Select

                Case DEFLATE_FAST
                    If strstart > blockStart Then
                        huffman.FlushBlock(window, blockStart, strstart - blockStart, False)
                        blockStart = strstart
                    End If
                    Exit Select

                Case DEFLATE_SLOW
                    If prevAvailable Then
                        huffman.TallyLit(window(strstart - 1) And &HFF)
                    End If
                    If strstart > blockStart Then
                        huffman.FlushBlock(window, blockStart, strstart - blockStart, False)
                        blockStart = strstart
                    End If
                    prevAvailable = False
                    matchLen = MIN_MATCH - 1
                    Exit Select
            End Select
            compressionFunction = COMPR_FUNC(level)
        End If
    End Sub


    Public Sub FillWindow()
        If strstart >= WSIZE + MAX_DIST Then
            SlideWindow()
        End If

        While lookahead < DeflaterConstants.MIN_LOOKAHEAD AndAlso inputOff < inputEnd
            Dim more As Integer = 2 * WSIZE - lookahead - strstart

            If more > inputEnd - inputOff Then
                more = inputEnd - inputOff
            End If

            System.Array.Copy(inputBuf, inputOff, window, strstart + lookahead, more)
            m_adler.Update(inputBuf, inputOff, more)

            inputOff += more
            m_totalIn += more
            lookahead += more
        End While

        If lookahead >= MIN_MATCH Then
            UpdateHash()
        End If
    End Sub

    Private Sub UpdateHash()
        ins_h = (window(strstart) << HASH_SHIFT) Xor window(strstart + 1)
    End Sub


    Private Function InsertString() As Integer
        Dim match As Short
        Dim hash As Integer = ((ins_h << HASH_SHIFT) Xor window(strstart + (MIN_MATCH - 1))) And HASH_MASK
        prev(strstart And WMASK) = InlineAssignHelper(match, head(hash))
        head(hash) = CShort(strstart)
        ins_h = hash
        Return match And &HFFFF
    End Function

    Private Sub SlideWindow()
        Array.Copy(window, WSIZE, window, 0, WSIZE)
        matchStart -= WSIZE
        strstart -= WSIZE
        blockStart -= WSIZE

        For i As Integer = 0 To HASH_SIZE - 1
            Dim m As Integer = head(i) And &HFFFF
            head(i) = CShort(If(m >= WSIZE, (m - WSIZE), 0))
        Next

        For i As Integer = 0 To WSIZE - 1
            Dim m As Integer = prev(i) And &HFFFF
            prev(i) = CShort(If(m >= WSIZE, (m - WSIZE), 0))
        Next
    End Sub

    Private Function FindLongestMatch(curMatch As Integer) As Boolean
        Dim chainLength As Integer = Me.p_max_chain
        Dim niceLength As Integer = Me.niceLength
        Dim prev As Short() = Me.prev
        Dim scan As Integer = Me.strstart
        Dim match As Integer
        Dim best_end As Integer = Me.strstart + matchLen
        Dim best_len As Integer = Math.Max(matchLen, MIN_MATCH - 1)

        Dim limit As Integer = Math.Max(strstart - MAX_DIST, 0)

        Dim strend As Integer = strstart + MAX_MATCH - 1
        Dim scan_end1 As Byte = window(best_end - 1)
        Dim scan_end As Byte = window(best_end)

        If best_len >= Me.goodLength Then
            chainLength >>= 2
        End If


        If niceLength > lookahead Then
            niceLength = lookahead
        End If



        Do
            If window(curMatch + best_len) <> scan_end OrElse window(curMatch + best_len - 1) <> scan_end1 OrElse window(curMatch) <> window(scan) OrElse window(curMatch + 1) <> window(scan + 1) Then
                Continue Do
            End If

            match = curMatch + 2
            scan += 2

            '        				while (
            '	window[++scan] == window[++match] &&
            '	window[++scan] == window[++match] &&
            '	window[++scan] == window[++match] &&
            '	window[++scan] == window[++match] &&
            '	window[++scan] == window[++match] &&
            '	window[++scan] == window[++match] &&
            '	window[++scan] == window[++match] &&
            '	window[++scan] == window[++match] &&
            '	(scan < strend)) 
            '                {
            '	// Do nothing
            '}


            While window(scan) = window(match) AndAlso _
                window(PreI(scan)) = window(PreI(match)) AndAlso _
                window(PreI(scan)) = window(PreI(match)) AndAlso _
                window(PreI(scan)) = window(PreI(match)) AndAlso _
                window(PreI(scan)) = window(PreI(match)) AndAlso _
                window(PreI(scan)) = window(PreI(match)) AndAlso _
                window(PreI(scan)) = window(PreI(match)) AndAlso _
                window(PreI(scan)) = window(PreI(match)) AndAlso (scan < strend)
            End While

            If scan > best_end Then
                matchStart = curMatch
                best_end = scan
                best_len = scan - strstart

                If best_len >= niceLength Then
                    Exit Do
                End If

                scan_end1 = window(best_end - 1)
                scan_end = window(best_end)
            End If
            scan = strstart
        Loop While (InlineAssignHelper(curMatch, (prev(curMatch And WMASK) And &HFFFF))) > limit AndAlso PreD(chainLength) <> 0
        '} while ((curMatch = (prev[curMatch & WMASK] & 0xffff)) > limit && --chainLength != 0);

        matchLen = Math.Min(best_len, lookahead)
        Return matchLen >= MIN_MATCH
    End Function

    Private Function DeflateStored(flush As Boolean, finish As Boolean) As Boolean
        If Not flush AndAlso (lookahead = 0) Then
            Return False
        End If

        strstart += lookahead
        lookahead = 0

        Dim storedLength As Integer = strstart - blockStart

        If (storedLength >= DeflaterConstants.MAX_BLOCK_SIZE) OrElse (blockStart < WSIZE AndAlso storedLength >= MAX_DIST) OrElse flush Then
            Dim lastBlock As Boolean = finish
            If storedLength > DeflaterConstants.MAX_BLOCK_SIZE Then
                storedLength = DeflaterConstants.MAX_BLOCK_SIZE
                lastBlock = False
            End If
            huffman.FlushStoredBlock(window, blockStart, storedLength, lastBlock)
            blockStart += storedLength
            Return Not lastBlock
        End If
        Return True
    End Function

    Private Function DeflateFast(flush As Boolean, finish As Boolean) As Boolean
        If lookahead < MIN_LOOKAHEAD AndAlso Not flush Then
            Return False
        End If

        While lookahead >= MIN_LOOKAHEAD OrElse flush
            If lookahead = 0 Then
                huffman.FlushBlock(window, blockStart, strstart - blockStart, finish)
                blockStart = strstart
                Return False
            End If

            If strstart > 2 * WSIZE - MIN_LOOKAHEAD Then
                SlideWindow()
            End If

            Dim hashHead As Integer
            If lookahead >= MIN_MATCH AndAlso (InlineAssignHelper(hashHead, InsertString())) <> 0 AndAlso m_strategy <> DeflateStrategy.HuffmanOnly AndAlso strstart - hashHead <= MAX_DIST AndAlso FindLongestMatch(hashHead) Then

                Dim full As Boolean = huffman.TallyDist(strstart - matchStart, matchLen)

                lookahead -= matchLen
                If matchLen <= p_max_lazy AndAlso lookahead >= MIN_MATCH Then
                    'while (--matchLen > 0)
                    While PreD(matchLen) > 0
                        strstart += 1
                        InsertString()
                    End While
                    strstart += 1
                Else
                    strstart += matchLen
                    If lookahead >= MIN_MATCH - 1 Then
                        UpdateHash()
                    End If
                End If
                matchLen = MIN_MATCH - 1
                If Not full Then
                    Continue While
                End If
            Else
                huffman.TallyLit(window(strstart) And &HFF)
                strstart += 1
                lookahead -= 1
            End If

            If huffman.IsFull() Then
                Dim lastBlock As Boolean = finish AndAlso (lookahead = 0)
                huffman.FlushBlock(window, blockStart, strstart - blockStart, lastBlock)
                blockStart = strstart
                Return Not lastBlock
            End If
        End While
        Return True
    End Function

    Private Function DeflateSlow(flush As Boolean, finish As Boolean) As Boolean
        If lookahead < MIN_LOOKAHEAD AndAlso Not flush Then
            Return False
        End If

        While lookahead >= MIN_LOOKAHEAD OrElse flush
            If lookahead = 0 Then
                If prevAvailable Then
                    huffman.TallyLit(window(strstart - 1) And &HFF)
                End If
                prevAvailable = False
                huffman.FlushBlock(window, blockStart, strstart - blockStart, finish)
                blockStart = strstart
                Return False
            End If

            If strstart >= 2 * WSIZE - MIN_LOOKAHEAD Then
                SlideWindow()
            End If

            Dim prevMatch As Integer = matchStart
            Dim prevLen As Integer = matchLen
            If lookahead >= MIN_MATCH Then

                Dim hashHead As Integer = InsertString()

                If m_strategy <> DeflateStrategy.HuffmanOnly AndAlso hashHead <> 0 AndAlso strstart - hashHead <= MAX_DIST AndAlso FindLongestMatch(hashHead) Then
                    If matchLen <= 5 AndAlso (m_strategy = DeflateStrategy.Filtered OrElse (matchLen = MIN_MATCH AndAlso strstart - matchStart > TooFar)) Then
                        matchLen = MIN_MATCH - 1
                    End If
                End If
            End If

            If (prevLen >= MIN_MATCH) AndAlso (matchLen <= prevLen) Then
                huffman.TallyDist(strstart - 1 - prevMatch, prevLen)
                prevLen -= 2
                Do
                    strstart += 1
                    lookahead -= 1
                    If lookahead >= MIN_MATCH Then
                        InsertString()
                    End If
                Loop While PreD(prevLen) > 0
                '} while (--prevLen > 0);

                strstart += 1
                lookahead -= 1
                prevAvailable = False
                matchLen = MIN_MATCH - 1
            Else
                If prevAvailable Then
                    huffman.TallyLit(window(strstart - 1) And &HFF)
                End If
                prevAvailable = True
                strstart += 1
                lookahead -= 1
            End If

            If huffman.IsFull() Then
                Dim len As Integer = strstart - blockStart
                If prevAvailable Then
                    len -= 1
                End If
                Dim lastBlock As Boolean = (finish AndAlso (lookahead = 0) AndAlso Not prevAvailable)
                huffman.FlushBlock(window, blockStart, len, lastBlock)
                blockStart += len
                Return Not lastBlock
            End If
        End While
        Return True
    End Function

End Class
