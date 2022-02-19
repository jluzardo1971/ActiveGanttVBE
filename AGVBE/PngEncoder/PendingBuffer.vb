Public Class PendingBuffer

    Private buffer_ As Byte()
    Private start As Integer
    Private [end] As Integer
    Private bits As UInteger
    Private m_bitCount As Integer

    Public Sub New()
        Me.New(4096)
    End Sub

    Public Sub New(bufferSize As Integer)
        buffer_ = New Byte(bufferSize - 1) {}
    End Sub

    Public Sub Reset()
        start = InlineAssignHelper([end], InlineAssignHelper(m_bitCount, 0))
    End Sub

    Public Sub WriteByte(value As Integer)
        'buffer_[end++] = unchecked((byte)value);
        buffer_(PosI([end])) = GetMSB(value)
    End Sub

    Public Sub WriteShort(value As Integer)
        'buffer_[end++] = unchecked((byte)value);
        buffer_(PosI([end])) = GetMSB(value)
        'buffer_[end++] = unchecked((byte)(value >> 8));
        buffer_(PosI([end])) = GetMSB(value >> 8)
    End Sub

    Public Sub WriteInt(value As Integer)
        'buffer_[end++] = unchecked((byte)value);
        buffer_(PosI([end])) = GetMSB(value)
        'buffer_[end++] = unchecked((byte)(value >> 8));
        buffer_(PosI([end])) = GetMSB(value >> 8)
        'buffer_[end++] = unchecked((byte)(value >> 16));
        buffer_(PosI([end])) = GetMSB(value >> 16)
        'buffer_[end++] = unchecked((byte)(value >> 24));
        buffer_(PosI([end])) = GetMSB(value >> 24)
    End Sub

    Public Sub WriteBlock(block As Byte(), offset As Integer, length As Integer)
        System.Array.Copy(block, offset, buffer_, [end], length)
        [end] += length
    End Sub

    Public ReadOnly Property BitCount() As Integer
        Get
            Return m_bitCount
        End Get
    End Property

    Public Sub AlignToByte()
        If m_bitCount > 0 Then
            'buffer_[end++] = unchecked((byte)bits);
            buffer_(PosI([end])) = GetMSB(bits)
            If m_bitCount > 8 Then
                'buffer_[end++] = unchecked((byte)(bits >> 8));
                buffer_(PosI([end])) = GetMSB(bits >> 8)
            End If
        End If
        bits = 0
        m_bitCount = 0
    End Sub

    Public Sub WriteBits(b As Integer, count As Integer)
        bits = bits Or CUInt(b << m_bitCount)
        m_bitCount += count
        If m_bitCount >= 16 Then
            'buffer_[end++] = unchecked((byte)bits);
            buffer_(PosI([end])) = GetMSB(bits)
            'buffer_[end++] = unchecked((byte)(bits >> 8));
            buffer_(PosI([end])) = GetMSB(bits >> 8)
            bits >>= 16
            m_bitCount -= 16
        End If
    End Sub

    Public Sub WriteShortMSB(s As Integer)
        'buffer_[end++] = unchecked((byte)(s >> 8));
        buffer_(PosI([end])) = GetMSB(s >> 8)
        'buffer_[end++] = unchecked((byte)s);
        buffer_(PosI([end])) = GetMSB(s)
    End Sub

    Public ReadOnly Property IsFlushed() As Boolean
        Get
            Return [end] = 0
        End Get
    End Property

    Public Function Flush(output As Byte(), offset As Integer, length As Integer) As Integer
        If m_bitCount >= 8 Then
            'buffer_[end++] = unchecked((byte)bits);
            buffer_(PosI([end])) = GetMSB(bits)
            bits >>= 8
            m_bitCount -= 8
        End If

        If length > [end] - start Then
            length = [end] - start
            System.Array.Copy(buffer_, start, output, offset, length)
            start = 0
            [end] = 0
        Else
            System.Array.Copy(buffer_, start, output, offset, length)
            start += length
        End If
        Return length
    End Function

    Public Function ToByteArray() As Byte()
        Dim result As Byte() = New Byte([end] - start - 1) {}
        System.Array.Copy(buffer_, start, result, 0, result.Length)
        start = 0
        [end] = 0
        Return result
    End Function

    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function

End Class

