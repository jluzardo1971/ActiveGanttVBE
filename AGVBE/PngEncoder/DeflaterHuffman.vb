Public Class DeflaterHuffman

    Const BUFSIZE As Integer = 1 << (DeflaterConstants.DEFAULT_MEM_LEVEL + 6)
    Const LITERAL_NUM As Integer = 286
    Const DIST_NUM As Integer = 30
    Const BITLEN_NUM As Integer = 19
    Const REP_3_6 As Integer = 16
    Const REP_3_10 As Integer = 17
    Const REP_11_138 As Integer = 18
    Const EOF_SYMBOL As Integer = 256
    Shared ReadOnly BL_ORDER As Integer() = {16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15}
    Shared ReadOnly bit4Reverse As Byte() = {0, 8, 4, 12, 2, 10, 6, 14, 1, 9, 5, 13, 3, 11, 7, 15}
    Shared staticLCodes As Short()
    Shared staticLLength As Byte()
    Shared staticDCodes As Short()
    Shared staticDLength As Byte()

    Public mp_oPending As DeflaterPending
    Private literalTree As Tree
    Private distTree As Tree
    Private blTree As Tree
    Private d_buf As Short()
    Private l_buf As Byte()
    Private last_lit As Integer
    Private extra_bits As Integer

    Private Class Tree
        Public freqs As Short()
        Public length As Byte()
        Public minNumCodes As Integer
        Public numCodes As Integer
        Private codes As Short()
        Private bl_counts As Integer()
        Private maxLength As Integer
        Private dh As DeflaterHuffman

        Public Sub New(dh As DeflaterHuffman, elems As Integer, minCodes As Integer, maxLength As Integer)
            Me.dh = dh
            Me.minNumCodes = minCodes
            Me.maxLength = maxLength
            freqs = New Short(elems - 1) {}
            bl_counts = New Integer(maxLength - 1) {}
        End Sub

        Public Sub Reset()
            For i As Integer = 0 To freqs.Length - 1
                freqs(i) = 0
            Next
            codes = Nothing
            length = Nothing
        End Sub

        Public Sub WriteSymbol(code As Integer)
            dh.mp_oPending.WriteBits(codes(code) And &HFFFF, length(code))
        End Sub

        Public Sub CheckEmpty()
            Dim empty As Boolean = True
            For i As Integer = 0 To freqs.Length - 1
                If freqs(i) <> 0 Then
                    empty = False
                End If
            Next

            If Not empty Then
                Throw New Exception("!Empty")
            End If
        End Sub

        Public Sub SetStaticCodes(staticCodes As Short(), staticLengths As Byte())
            codes = staticCodes
            length = staticLengths
        End Sub

        Public Sub BuildCodes()
            Dim numSymbols As Integer = freqs.Length
            Dim nextCode As Integer() = New Integer(maxLength - 1) {}
            Dim code As Integer = 0
            codes = New Short(freqs.Length - 1) {}
            For bits As Integer = 0 To maxLength - 1
                nextCode(bits) = code

                code += bl_counts(bits) << (15 - bits)
            Next
            For i As Integer = 0 To numCodes - 1
                Dim bits As Integer = length(i)
                If bits > 0 Then
                    codes(i) = BitReverse(nextCode(bits - 1))
                    nextCode(bits - 1) += 1 << (16 - bits)
                End If
            Next
        End Sub

        Public Sub BuildTree()
            Dim numSymbols As Integer = freqs.Length
            Dim heap As Integer() = New Integer(numSymbols - 1) {}
            Dim heapLen As Integer = 0
            Dim maxCode As Integer = 0
            For n As Integer = 0 To numSymbols - 1
                Dim freq As Integer = freqs(n)
                If freq <> 0 Then
                    'int pos = heapLen++;
                    Dim pos As Integer = PosI(heapLen)
                    Dim ppos As Integer
                    While pos > 0 AndAlso freqs(heap(InlineAssignHelper(ppos, (pos - 1) \ 2))) > freq
                        heap(pos) = heap(ppos)
                        pos = ppos
                    End While
                    heap(pos) = n

                    maxCode = n
                End If
            Next
            While heapLen < 2
                'int node = maxCode < 2 ? ++maxCode : 0;
                Dim node As Integer = If(maxCode < 2, PreI(maxCode), 0)
                'heap[heapLen++] = node;
                heap(PosI(heapLen)) = node
            End While
            numCodes = Math.Max(maxCode + 1, minNumCodes)
            Dim numLeafs As Integer = heapLen
            Dim childs As Integer() = New Integer(4 * heapLen - 3) {}
            Dim values As Integer() = New Integer(2 * heapLen - 2) {}
            Dim numNodes As Integer = numLeafs
            For i As Integer = 0 To heapLen - 1
                Dim node As Integer = heap(i)
                childs(2 * i) = node
                childs(2 * i + 1) = -1
                values(i) = freqs(node) << 8
                heap(i) = i
            Next
            Do
                Dim first As Integer = heap(0)
                'int last = heap[--heapLen];
                Dim last As Integer = heap(PreD(heapLen))
                Dim ppos As Integer = 0
                Dim path As Integer = 1

                While path < heapLen
                    If path + 1 < heapLen AndAlso values(heap(path)) > values(heap(path + 1)) Then
                        path += 1
                    End If

                    heap(ppos) = heap(path)
                    ppos = path
                    path = path * 2 + 1
                End While
                Dim lastVal As Integer = values(last)
                While (InlineAssignHelper(path, ppos)) > 0 AndAlso values(heap(InlineAssignHelper(ppos, (path - 1) \ 2))) > lastVal
                    heap(path) = heap(ppos)
                End While
                heap(path) = last


                Dim second As Integer = heap(0)
                'last = numNodes++;
                last = PosI(numNodes)
                childs(2 * last) = first
                childs(2 * last + 1) = second
                Dim mindepth As Integer = Math.Min(values(first) And &HFF, values(second) And &HFF)
                values(last) = InlineAssignHelper(lastVal, values(first) + values(second) - mindepth + 1)
                ppos = 0
                path = 1

                While path < heapLen
                    If path + 1 < heapLen AndAlso values(heap(path)) > values(heap(path + 1)) Then
                        path += 1
                    End If

                    heap(ppos) = heap(path)
                    ppos = path
                    path = ppos * 2 + 1
                End While
                While (InlineAssignHelper(path, ppos)) > 0 AndAlso values(heap(InlineAssignHelper(ppos, (path - 1) \ 2))) > lastVal
                    heap(path) = heap(ppos)
                End While
                heap(path) = last
            Loop While heapLen > 1

            If heap(0) <> childs.Length \ 2 - 1 Then
                Throw New Exception("Heap invariant violated")
            End If
            BuildLength(childs)
        End Sub

        Public Function GetEncodedLength() As Integer
            Dim len As Integer = 0
            For i As Integer = 0 To freqs.Length - 1
                len += freqs(i) * length(i)
            Next
            Return len
        End Function

        Public Sub CalcBLFreq(blTree As Tree)
            Dim max_count As Integer
            Dim min_count As Integer
            Dim count As Integer
            Dim curlen As Integer = -1

            Dim i As Integer = 0
            While i < numCodes
                count = 1
                Dim nextlen As Integer = length(i)
                If nextlen = 0 Then
                    max_count = 138
                    min_count = 3
                Else
                    max_count = 6
                    min_count = 3
                    If curlen <> nextlen Then
                        blTree.freqs(nextlen) += 1
                        count = 0
                    End If
                End If
                curlen = nextlen
                i += 1

                While i < numCodes AndAlso curlen = length(i)
                    i += 1
                    'if (++count >= max_count)
                    If PreI(count) >= max_count Then
                        Exit While
                    End If
                End While

                If count < min_count Then
                    blTree.freqs(curlen) += CShort(count)
                ElseIf curlen <> 0 Then
                    blTree.freqs(REP_3_6) += 1
                ElseIf count <= 10 Then
                    blTree.freqs(REP_3_10) += 1
                Else
                    blTree.freqs(REP_11_138) += 1
                End If
            End While
        End Sub

        Public Sub WriteTree(blTree As Tree)
            Dim max_count As Integer
            Dim min_count As Integer
            Dim count As Integer
            Dim curlen As Integer = -1

            Dim i As Integer = 0
            While i < numCodes
                count = 1
                Dim nextlen As Integer = length(i)
                If nextlen = 0 Then
                    max_count = 138
                    min_count = 3
                Else
                    max_count = 6
                    min_count = 3
                    If curlen <> nextlen Then
                        blTree.WriteSymbol(nextlen)
                        count = 0
                    End If
                End If
                curlen = nextlen
                i += 1

                While i < numCodes AndAlso curlen = length(i)
                    i += 1
                    'if (++count >= max_count)
                    If PreI(count) >= max_count Then
                        Exit While
                    End If
                End While

                If count < min_count Then
                    'while (count-- > 0)
                    While PosD(count) > 0
                        blTree.WriteSymbol(curlen)
                    End While
                ElseIf curlen <> 0 Then
                    blTree.WriteSymbol(REP_3_6)
                    dh.mp_oPending.WriteBits(count - 3, 2)
                ElseIf count <= 10 Then
                    blTree.WriteSymbol(REP_3_10)
                    dh.mp_oPending.WriteBits(count - 3, 3)
                Else
                    blTree.WriteSymbol(REP_11_138)
                    dh.mp_oPending.WriteBits(count - 11, 7)
                End If
            End While
        End Sub

        Private Sub BuildLength(childs As Integer())
            Me.length = New Byte(freqs.Length - 1) {}
            Dim numNodes As Integer = childs.Length \ 2
            Dim numLeafs As Integer = (numNodes + 1) \ 2
            Dim overflow As Integer = 0

            For i As Integer = 0 To maxLength - 1
                bl_counts(i) = 0
            Next

            Dim lengths As Integer() = New Integer(numNodes - 1) {}
            lengths(numNodes - 1) = 0

            For i As Integer = numNodes - 1 To 0 Step -1
                If childs(2 * i + 1) <> -1 Then
                    Dim bitLength As Integer = lengths(i) + 1
                    If bitLength > maxLength Then
                        bitLength = maxLength
                        overflow += 1
                    End If
                    lengths(childs(2 * i)) = InlineAssignHelper(lengths(childs(2 * i + 1)), bitLength)
                Else
                    Dim bitLength As Integer = lengths(i)
                    bl_counts(bitLength - 1) += 1
                    Me.length(childs(2 * i)) = CByte(lengths(i))
                End If
            Next


            If overflow = 0 Then
                Return
            End If

            Dim incrBitLen As Integer = maxLength - 1
            Do
                'while (bl_counts[--incrBitLen] == 0)
                While bl_counts(PreD(incrBitLen)) = 0


                End While

                Do
                    bl_counts(incrBitLen) -= 1
                    'bl_counts[++incrBitLen]++;
                    bl_counts(PreI(incrBitLen)) += 1
                    overflow -= 1 << (maxLength - 1 - incrBitLen)
                Loop While overflow > 0 AndAlso incrBitLen < maxLength - 1
            Loop While overflow > 0

            bl_counts(maxLength - 1) += overflow
            bl_counts(maxLength - 2) -= overflow


            Dim nodePtr As Integer = 2 * numLeafs
            Dim bits As Integer = maxLength
            While bits <> 0
                Dim n As Integer = bl_counts(bits - 1)
                While n > 0
                    'int childPtr = 2 * childs[nodePtr++];
                    Dim childPtr As Integer = 2 * childs(PosI(nodePtr))
                    If childs(childPtr + 1) = -1 Then
                        length(childs(childPtr)) = CByte(bits)
                        n -= 1
                    End If
                End While
                bits -= 1
            End While
        End Sub
        Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
            target = value
            Return value
        End Function

    End Class

    Shared Sub New()
        staticLCodes = New Short(LITERAL_NUM - 1) {}
        staticLLength = New Byte(LITERAL_NUM - 1) {}

        Dim i As Integer = 0
        While i < 144
            staticLCodes(i) = BitReverse((&H30 + i) << 8)
            staticLLength(PosI(i)) = 8
        End While

        While i < 256
            staticLCodes(i) = BitReverse((&H190 - 144 + i) << 7)
            staticLLength(PosI(i)) = 9
        End While

        While i < 280
            staticLCodes(i) = BitReverse((&H0 - 256 + i) << 9)
            staticLLength(PosI(i)) = 7
        End While

        While i < LITERAL_NUM
            staticLCodes(i) = BitReverse((&HC0 - 280 + i) << 8)
            staticLLength(PosI(i)) = 8
        End While
        staticDCodes = New Short(DIST_NUM - 1) {}
        staticDLength = New Byte(DIST_NUM - 1) {}
        For i = 0 To DIST_NUM - 1
            staticDCodes(i) = BitReverse(i << 11)
            staticDLength(i) = 5
        Next
    End Sub

    Public Sub New(pending As DeflaterPending)
        Me.mp_oPending = pending
        literalTree = New Tree(Me, LITERAL_NUM, 257, 15)
        distTree = New Tree(Me, DIST_NUM, 1, 15)
        blTree = New Tree(Me, BITLEN_NUM, 4, 7)
        d_buf = New Short(BUFSIZE - 1) {}
        l_buf = New Byte(BUFSIZE - 1) {}
    End Sub

    Public Sub Reset()
        last_lit = 0
        extra_bits = 0
        literalTree.Reset()
        distTree.Reset()
        blTree.Reset()
    End Sub

    Public Sub SendAllTrees(blTreeCodes As Integer)
        blTree.BuildCodes()
        literalTree.BuildCodes()
        distTree.BuildCodes()
        mp_oPending.WriteBits(literalTree.numCodes - 257, 5)
        mp_oPending.WriteBits(distTree.numCodes - 1, 5)
        mp_oPending.WriteBits(blTreeCodes - 4, 4)
        For rank As Integer = 0 To blTreeCodes - 1
            mp_oPending.WriteBits(blTree.length(BL_ORDER(rank)), 3)
        Next
        literalTree.WriteTree(blTree)
        distTree.WriteTree(blTree)
    End Sub

    Public Sub CompressBlock()
        For i As Integer = 0 To last_lit - 1
            Dim litlen As Integer = l_buf(i) And &HFF
            Dim dist As Integer = d_buf(i)
            'if (dist-- != 0)
            If PosD(dist) <> 0 Then

                Dim lc As Integer = Lcode(litlen)
                literalTree.WriteSymbol(lc)

                Dim bits As Integer = (lc - 261) \ 4
                If bits > 0 AndAlso bits <= 5 Then
                    mp_oPending.WriteBits(litlen And ((1 << bits) - 1), bits)
                End If

                Dim dc As Integer = Dcode(dist)
                distTree.WriteSymbol(dc)

                bits = dc \ 2 - 1
                If bits > 0 Then
                    mp_oPending.WriteBits(dist And ((1 << bits) - 1), bits)
                End If
            Else
                literalTree.WriteSymbol(litlen)
            End If
        Next
        literalTree.WriteSymbol(EOF_SYMBOL)
    End Sub


    Public Sub FlushStoredBlock(stored As Byte(), storedOffset As Integer, storedLength As Integer, lastBlock As Boolean)
        mp_oPending.WriteBits((DeflaterConstants.STORED_BLOCK << 1) + (If(lastBlock, 1, 0)), 3)
        mp_oPending.AlignToByte()
        mp_oPending.WriteShort(storedLength)
        mp_oPending.WriteShort(Not storedLength)
        mp_oPending.WriteBlock(stored, storedOffset, storedLength)
        Reset()
    End Sub

    Public Sub FlushBlock(stored As Byte(), storedOffset As Integer, storedLength As Integer, lastBlock As Boolean)
        literalTree.freqs(EOF_SYMBOL) += 1
        literalTree.BuildTree()
        distTree.BuildTree()
        literalTree.CalcBLFreq(blTree)
        distTree.CalcBLFreq(blTree)
        blTree.BuildTree()

        Dim blTreeCodes As Integer = 4
        For i As Integer = 18 To blTreeCodes + 1 Step -1
            If blTree.length(BL_ORDER(i)) > 0 Then
                blTreeCodes = i + 1
            End If
        Next
        Dim opt_len As Integer = 14 + blTreeCodes * 3 + blTree.GetEncodedLength() + literalTree.GetEncodedLength() + distTree.GetEncodedLength() + extra_bits

        Dim static_len As Integer = extra_bits
        For i As Integer = 0 To LITERAL_NUM - 1
            static_len += literalTree.freqs(i) * staticLLength(i)
        Next
        For i As Integer = 0 To DIST_NUM - 1
            static_len += distTree.freqs(i) * staticDLength(i)
        Next
        If opt_len >= static_len Then
            opt_len = static_len
        End If

        If storedOffset >= 0 AndAlso storedLength + 4 < opt_len >> 3 Then
            FlushStoredBlock(stored, storedOffset, storedLength, lastBlock)
        ElseIf opt_len = static_len Then
            mp_oPending.WriteBits((DeflaterConstants.STATIC_TREES << 1) + (If(lastBlock, 1, 0)), 3)
            literalTree.SetStaticCodes(staticLCodes, staticLLength)
            distTree.SetStaticCodes(staticDCodes, staticDLength)
            CompressBlock()
            Reset()
        Else
            mp_oPending.WriteBits((DeflaterConstants.DYN_TREES << 1) + (If(lastBlock, 1, 0)), 3)
            SendAllTrees(blTreeCodes)
            CompressBlock()
            Reset()
        End If
    End Sub

    Public Function IsFull() As Boolean
        Return last_lit >= BUFSIZE
    End Function

    Public Function TallyLit(literal As Integer) As Boolean
        d_buf(last_lit) = 0
        'l_buf[last_lit++] = (byte)literal;
        l_buf(PosI(last_lit)) = CByte(literal)
        literalTree.freqs(literal) += 1
        Return IsFull()
    End Function

    Public Function TallyDist(distance As Integer, length As Integer) As Boolean

        d_buf(last_lit) = CShort(distance)
        'l_buf[last_lit++] = (byte)(length - 3);
        l_buf(PosI(last_lit)) = CByte(length - 3)

        Dim lc As Integer = Lcode(length - 3)
        literalTree.freqs(lc) += 1
        If lc >= 265 AndAlso lc < 285 Then
            extra_bits += (lc - 261) \ 4
        End If

        Dim dc As Integer = Dcode(distance - 1)
        distTree.freqs(dc) += 1
        If dc >= 4 Then
            extra_bits += dc \ 2 - 1
        End If
        Return IsFull()
    End Function

    Public Shared Function BitReverse(toReverse As Integer) As Short
        Return CShort(bit4Reverse(toReverse And &HF) << 12 Or bit4Reverse((toReverse >> 4) And &HF) << 8 Or bit4Reverse((toReverse >> 8) And &HF) << 4 Or bit4Reverse(toReverse >> 12))
    End Function

    Private Shared Function Lcode(length As Integer) As Integer
        If length = 255 Then
            Return 285
        End If

        Dim code As Integer = 257
        While length >= 8
            code += 4
            length >>= 1
        End While
        Return code + length
    End Function

    Private Shared Function Dcode(distance As Integer) As Integer
        Dim code As Integer = 0
        While distance >= 4
            code += 2
            distance >>= 1
        End While
        Return code + distance
    End Function

End Class

