Public NotInheritable Class Adler32
    Implements IChecksum

    Const BASE As UInteger = 65521
    Private checksum As UInteger

    Public Property Value() As Long Implements IChecksum.Value
        Get
            Return checksum
        End Get
        Set(value As Long)
            checksum = CUInt(value)
        End Set
    End Property

    Public Sub New()
        Reset()
    End Sub

    Public Sub Reset() Implements IChecksum.Reset
        checksum = 1
    End Sub

    Public Sub Update(value As Integer) Implements IChecksum.Update
        Dim s1 As UInteger = checksum And &HFFFF
        Dim s2 As UInteger = checksum >> 16

        s1 = (s1 + (CUInt(value) And &HFF)) Mod BASE
        s2 = (s1 + s2) Mod BASE

        checksum = (s2 << 16) + s1
    End Sub

    Public Sub Update(buffer As Byte()) Implements IChecksum.Update
        If buffer Is Nothing Then
            Throw New ArgumentNullException("buffer")
        End If
        Update(buffer, 0, buffer.Length)
    End Sub

    Public Sub Update(buffer As Byte(), offset As Integer, count As Integer) Implements IChecksum.Update
        If buffer Is Nothing Then
            Throw New ArgumentNullException("buffer")
        End If

        If offset < 0 Then
            Throw New ArgumentOutOfRangeException("offset")
        End If

        If count < 0 Then
            Throw New ArgumentOutOfRangeException("count")
        End If

        If offset >= buffer.Length Then
            Throw New ArgumentOutOfRangeException("offset")
        End If

        If offset + count > buffer.Length Then
            Throw New ArgumentOutOfRangeException("count")
        End If

        Dim s1 As UInteger = checksum And &HFFFF
        Dim s2 As UInteger = checksum >> 16

        While count > 0
            Dim n As Integer = 3800
            If n > count Then
                n = count
            End If
            count -= n
            While PreD(n) >= 0
                s1 = s1 + CUInt(buffer(PosI(offset)) And &HFF)
                s2 = s2 + s1
            End While
            s1 = s1 Mod BASE
            s2 = s2 Mod BASE
        End While
        checksum = (s2 << 16) Or s1
    End Sub
End Class

