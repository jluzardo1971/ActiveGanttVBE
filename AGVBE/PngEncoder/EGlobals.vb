Module EGlobals

    Public Function PreI(ByRef i As Integer) As Integer
        i = i + 1
        Return i
    End Function

    Public Function PosI(ByRef i As Integer) As Integer
        i = i + 1
        Return i - 1
    End Function

    Public Function PreD(ByRef i As Integer) As Integer
        i = i - 1
        Return i
    End Function

    Public Function PosD(ByRef i As Integer) As Integer
        i = i - 1
        Return i + 1
    End Function

    Public Function ForceToInt32(ByVal lLongValue As Long) As Integer
        Dim aBytes64 = BitConverter.GetBytes(lLongValue)
        Dim lReturn As Int32
        lReturn = BitConverter.ToInt32(aBytes64, 0)
        Return lReturn
    End Function

    Public Function GetMSB(ByVal i As UInt32) As Byte
        Dim bytes As Byte() = BitConverter.GetBytes(i)
        If BitConverter.IsLittleEndian Then
            Return bytes(0)
        Else
            Return bytes(bytes.Length - 1)
        End If
    End Function

    Public Function GetMSB(ByVal i As Int32) As Byte
        Dim bytes As Byte() = BitConverter.GetBytes(i)
        If BitConverter.IsLittleEndian Then
            Return bytes(0)
        Else
            Return bytes(bytes.Length - 1)
        End If
    End Function

End Module
