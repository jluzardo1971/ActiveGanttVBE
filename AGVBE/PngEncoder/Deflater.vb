Public Class Deflater

    Public Const BEST_COMPRESSION As Integer = 9
    Public Const BEST_SPEED As Integer = 1
    Public Const DEFAULT_COMPRESSION As Integer = -1
    Public Const NO_COMPRESSION As Integer = 0
    Public Const DEFLATED As Integer = 8

    Private Const IS_SETDICT As Integer = &H1
    Private Const IS_FLUSHING As Integer = &H4
    Private Const IS_FINISHING As Integer = &H8

    Private Const INIT_STATE As Integer = &H0
    Private Const SETDICT_STATE As Integer = &H1
    '		private static  int INIT_FINISHING_STATE    = 0x08;
    '		private static  int SETDICT_FINISHING_STATE = 0x09;
    Private Const BUSY_STATE As Integer = &H10
    Private Const FLUSHING_STATE As Integer = &H14
    Private Const FINISHING_STATE As Integer = &H1C
    Private Const FINISHED_STATE As Integer = &H1E
    Private Const CLOSED_STATE As Integer = &H7F

    Private level As Integer
    Private noZlibHeaderOrFooter As Boolean
    Private state As Integer
    Private m_totalOut As Long
    Private pending As DeflaterPending
    Private engine As DeflaterEngine

    Public Sub New()
        Me.New(DEFAULT_COMPRESSION, False)
    End Sub

    Public Sub New(level As Integer)
        Me.New(level, False)
    End Sub

    Public Sub New(level As Integer, noZlibHeaderOrFooter As Boolean)
        If level = DEFAULT_COMPRESSION Then
            level = 6
        ElseIf level < NO_COMPRESSION OrElse level > BEST_COMPRESSION Then
            Throw New ArgumentOutOfRangeException("level")
        End If

        pending = New DeflaterPending()
        engine = New DeflaterEngine(pending)
        Me.noZlibHeaderOrFooter = noZlibHeaderOrFooter
        SetStrategy(DeflateStrategy.[Default])
        SetLevel(level)
        Reset()
    End Sub

    Public Sub Reset()
        state = (If(noZlibHeaderOrFooter, BUSY_STATE, INIT_STATE))
        m_totalOut = 0
        pending.Reset()
        engine.Reset()
    End Sub

    Public ReadOnly Property Adler() As Integer
        Get
            Return engine.Adler
        End Get
    End Property

    Public ReadOnly Property TotalIn() As Long
        Get
            Return engine.TotalIn
        End Get
    End Property

    Public ReadOnly Property TotalOut() As Long
        Get
            Return m_totalOut
        End Get
    End Property

    Public Sub Flush()
        state = state Or IS_FLUSHING
    End Sub

    Public Sub Finish()
        state = state Or (IS_FLUSHING Or IS_FINISHING)
    End Sub

    Public ReadOnly Property IsFinished() As Boolean
        Get
            Return (state = FINISHED_STATE) AndAlso pending.IsFlushed
        End Get
    End Property

    Public ReadOnly Property IsNeedingInput() As Boolean
        Get
            Return engine.NeedsInput()
        End Get
    End Property

    Public Sub SetInput(input As Byte())
        SetInput(input, 0, input.Length)
    End Sub

    Public Sub SetInput(input As Byte(), offset As Integer, count As Integer)
        If (state And IS_FINISHING) <> 0 Then
            Throw New InvalidOperationException("Finish() already called")
        End If
        engine.SetInput(input, offset, count)
    End Sub

    Public Sub SetLevel(level As Integer)
        If level = DEFAULT_COMPRESSION Then
            level = 6
        ElseIf level < NO_COMPRESSION OrElse level > BEST_COMPRESSION Then
            Throw New ArgumentOutOfRangeException("level")
        End If
        If Me.level <> level Then
            Me.level = level
            engine.SetLevel(level)
        End If
    End Sub

    Public Function GetLevel() As Integer
        Return level
    End Function

    Public Sub SetStrategy(strategy As DeflateStrategy)
        engine.Strategy = strategy
    End Sub

    Public Function Deflate(output As Byte()) As Integer
        Return Deflate(output, 0, output.Length)
    End Function

    Public Function Deflate(output As Byte(), offset As Integer, length As Integer) As Integer
        Dim origLength As Integer = length

        If state = CLOSED_STATE Then
            Throw New InvalidOperationException("Deflater closed")
        End If

        If state < BUSY_STATE Then
            Dim header As Integer = (DEFLATED + ((DeflaterConstants.MAX_WBITS - 8) << 4)) << 8
            Dim level_flags As Integer = (level - 1) >> 1
            If level_flags < 0 OrElse level_flags > 3 Then
                level_flags = 3
            End If
            header = header Or level_flags << 6
            If (state And IS_SETDICT) <> 0 Then
                header = header Or DeflaterConstants.PRESET_DICT
            End If
            header += 31 - (header Mod 31)
            pending.WriteShortMSB(header)
            If (state And IS_SETDICT) <> 0 Then
                Dim chksum As Integer = engine.Adler
                engine.ResetAdler()
                pending.WriteShortMSB(chksum >> 16)
                pending.WriteShortMSB(chksum And &HFFFF)
            End If
            state = BUSY_STATE Or (state And (IS_FLUSHING Or IS_FINISHING))
        End If

        While True
            Dim count As Integer = pending.Flush(output, offset, length)
            offset += count
            m_totalOut += count
            length -= count
            If length = 0 OrElse state = FINISHED_STATE Then
                Exit While
            End If
            If Not engine.Deflate((state And IS_FLUSHING) <> 0, (state And IS_FINISHING) <> 0) Then
                If state = BUSY_STATE Then
                    Return origLength - length
                ElseIf state = FLUSHING_STATE Then
                    If level <> NO_COMPRESSION Then
                        Dim neededbits As Integer = 8 + ((-pending.BitCount) And 7)
                        While neededbits > 0
                            pending.WriteBits(2, 10)
                            neededbits -= 10
                        End While
                    End If
                    state = BUSY_STATE
                ElseIf state = FINISHING_STATE Then
                    pending.AlignToByte()
                    If Not noZlibHeaderOrFooter Then
                        Dim adler As Integer = engine.Adler
                        pending.WriteShortMSB(adler >> 16)
                        pending.WriteShortMSB(adler And &HFFFF)
                    End If
                    state = FINISHED_STATE
                End If
            End If
        End While
        Return origLength - length
    End Function

    Public Sub SetDictionary(dictionary As Byte())
        SetDictionary(dictionary, 0, dictionary.Length)
    End Sub

    Public Sub SetDictionary(dictionary As Byte(), index As Integer, count As Integer)
        If state <> INIT_STATE Then
            Throw New InvalidOperationException()
        End If

        state = SETDICT_STATE
        engine.SetDictionary(dictionary, index, count)
    End Sub

End Class
