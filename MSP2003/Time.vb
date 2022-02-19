Public Class Time

    Private mp_yHour As Byte
    Private mp_yMinute As Byte
    Private mp_ySecond As Byte


    Public Sub New()
        mp_yHour = 0
        mp_yMinute = 0
        mp_ySecond = 0
    End Sub

    Public Overrides Function ToString() As String
        Return mp_yHour.ToString("00") & ":" & mp_yMinute.ToString("00") & ":" & mp_ySecond.ToString("00")
    End Function

    Public Sub FromString(ByVal sString As String)
        mp_yHour = System.Convert.ToByte(sString.Substring(0, 2))
        mp_yMinute = System.Convert.ToByte(sString.Substring(3, 2))
        mp_ySecond = System.Convert.ToByte(sString.Substring(6, 2))
    End Sub

    Public Property Hour() As Byte
        Get
            Return mp_yHour
        End Get
        Set(ByVal value As Byte)
            mp_yHour = value
        End Set
    End Property

    Public Property Minute() As Byte
        Get
            Return mp_yMinute
        End Get
        Set(ByVal value As Byte)
            mp_yMinute = value
        End Set
    End Property

    Public Property Second() As Byte
        Get
            Return mp_ySecond
        End Get
        Set(ByVal value As Byte)
            mp_ySecond = value
        End Set
    End Property

    Public Function ToDateTime() As System.DateTime
        Dim dtReturn As System.DateTime = New System.DateTime(0, 0, 0, mp_yHour, mp_yMinute, mp_ySecond)
        Return dtReturn
    End Function

    Public Sub FromDateTime(ByVal dtDate As System.DateTime)
        mp_yHour = System.Convert.ToByte(dtDate.Hour)
        mp_yMinute = System.Convert.ToByte(dtDate.Minute)
        mp_ySecond = System.Convert.ToByte(dtDate.Second)
    End Sub

    Public Function IsNull() As Boolean
        Dim bReturn As Boolean = True
        If mp_yHour <> 0 Then
            bReturn = False
        End If
        If mp_yMinute <> 0 Then
            bReturn = False
        End If
        If mp_ySecond <> 0 Then
            bReturn = False
        End If
        Return bReturn
    End Function

End Class
