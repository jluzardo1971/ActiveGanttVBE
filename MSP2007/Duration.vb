Public Class Duration

    Private mp_lYear As Integer
    Private mp_lMonth As Integer
    Private mp_lDay As Integer
    Private mp_lHour As Integer
    Private mp_lMinute As Integer
    Private mp_lSecond As Integer

    Public Sub New()
        mp_lYear = 0
        mp_lMonth = 0
        mp_lDay = 0
        mp_lHour = 0
        mp_lMinute = 0
        mp_lSecond = 0
    End Sub

    Public Function IsNull() As Boolean
        Dim bReturn As Boolean = True
        If mp_lYear <> 0 Then
            bReturn = False
        End If
        If mp_lMonth <> 0 Then
            bReturn = False
        End If
        If mp_lDay <> 0 Then
            bReturn = False
        End If
        If mp_lHour <> 0 Then
            bReturn = False
        End If
        If mp_lMinute <> 0 Then
            bReturn = False
        End If
        If mp_lSecond <> 0 Then
            bReturn = False
        End If

        Return bReturn
    End Function

    Public Overrides Function ToString() As String
        Dim sReturn As String = "P"
        If mp_lYear <> 0 Then
            sReturn = sReturn & mp_lYear.ToString() & "Y"
        End If
        If mp_lMonth <> 0 Then
            sReturn = sReturn & mp_lMonth.ToString() & "M"
        End If
        If mp_lDay <> 0 Then
            sReturn = sReturn & mp_lDay.ToString() & "D"
        End If
        sReturn = sReturn & "T"
        sReturn = sReturn & mp_lHour.ToString() & "H"
        sReturn = sReturn & mp_lMinute.ToString() & "M"
        sReturn = sReturn & mp_lSecond.ToString() & "S"
        Return sReturn
    End Function

    Public Sub FromString(ByVal sString As String)
        Dim sTime As String
        Dim sDate As String
        Dim sBuff As String
        If sString.StartsWith("P") = False Or sString.Length = 0 Then
            Return
        End If
        If sString.IndexOf("T") > -1 Then
            sTime = sString.Substring(sString.IndexOf("T") + 1, sString.Length - sString.IndexOf("T") - 1)
            sDate = sString.Replace("T" & sTime, "")
        Else
            sTime = ""
            sDate = sString
        End If
        sDate = sDate.Substring(1, sDate.Length - 1)
        If sTime.Length > 0 Then
            If sTime.IndexOf("H") > -1 Then
                sBuff = sTime.Substring(0, sTime.IndexOf("H"))
                sTime = sTime.Replace(sBuff & "H", "")
                mp_lHour = System.Convert.ToInt16(sBuff)
            End If
            If sTime.IndexOf("M") > -1 Then
                sBuff = sTime.Substring(0, sTime.IndexOf("M"))
                sTime = sTime.Replace(sBuff & "M", "")
                mp_lMinute = System.Convert.ToInt16(sBuff)
            End If
            If sTime.IndexOf("S") > -1 Then
                sBuff = sTime.Substring(0, sTime.IndexOf("S"))
                sTime = sTime.Replace(sBuff & "S", "")
                mp_lSecond = System.Convert.ToInt16(sBuff)
            End If
        End If
        If sDate.Length > 0 Then
            If sDate.IndexOf("Y") > -1 Then
                sBuff = sDate.Substring(0, sDate.IndexOf("Y"))
                sDate = sDate.Replace(sBuff & "Y", "")
                mp_lYear = System.Convert.ToInt16(sBuff)
            End If
            If sDate.IndexOf("M") > -1 Then
                sBuff = sDate.Substring(0, sDate.IndexOf("M"))
                sDate = sDate.Replace(sBuff & "M", "")
                mp_lMonth = System.Convert.ToInt16(sBuff)
            End If
            If sDate.IndexOf("D") > -1 Then
                sBuff = sDate.Substring(0, sDate.IndexOf("D"))
                sDate = sDate.Replace(sBuff & "D", "")
                mp_lDay = System.Convert.ToInt16(sBuff)
            End If
        End If
    End Sub

    Public Property Year() As Integer
        Get
            Return mp_lYear
        End Get
        Set(ByVal value As Integer)
            mp_lYear = value
        End Set
    End Property

    Public Property Month() As Integer
        Get
            Return mp_lMonth
        End Get
        Set(ByVal value As Integer)
            mp_lMonth = value
        End Set
    End Property

    Public Property Day() As Integer
        Get
            Return mp_lDay
        End Get
        Set(ByVal value As Integer)
            mp_lDay = value
        End Set
    End Property

    Public Property Hour() As Integer
        Get
            Return mp_lHour
        End Get
        Set(ByVal value As Integer)
            mp_lHour = value
        End Set
    End Property

    Public Property Minute() As Integer
        Get
            Return mp_lMinute
        End Get
        Set(ByVal value As Integer)
            mp_lMinute = value
        End Set
    End Property

    Public Property Second() As Integer
        Get
            Return mp_lSecond
        End Get
        Set(ByVal value As Integer)
            mp_lSecond = value
        End Set
    End Property

End Class
