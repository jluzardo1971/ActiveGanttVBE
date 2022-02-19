Option Explicit On

Public Class DateTime

    Private mp_dtDateTime As System.DateTime
    Private mp_lSecondFraction As Integer

    Public Sub New()
        mp_dtDateTime = New System.DateTime(0)
        mp_lSecondFraction = 0
    End Sub

    Public Sub New(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer)
        Initialize(Year, Month, Day, 0, 0, 0, 0, 0, 0)
    End Sub

    Public Sub New(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer)
        Initialize(Year, Month, Day, Hour, Minute, Second, 0, 0, 0)
    End Sub

    Public Sub New(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer, ByVal Millisecond As Integer, ByVal Microsecond As Integer, ByVal Nanosecond As Integer)
        Initialize(Year, Month, Day, Hour, Minute, Second, Millisecond, Microsecond, Nanosecond)
    End Sub

    Public Shared ReadOnly Property Now() As AGVBE.DateTime
        Get
            Dim dtResult As AGVBE.DateTime = New AGVBE.DateTime()
            dtResult.SetToCurrentDateTime()
            Return dtResult
        End Get
    End Property

    Public Sub Initialize(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer, ByVal Millisecond As Integer, ByVal Microsecond As Integer, ByVal Nanosecond As Integer)
        If Year < 1 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Year > 9999 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Month < 1 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Month > 12 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Day < 1 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Day > 31 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Hour < 0 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Hour > 23 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Minute < 0 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Minute > 59 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Second < 0 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Second > 59 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Millisecond < 0 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Millisecond > 999 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Microsecond < 0 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Microsecond > 999 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Nanosecond < 0 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        ElseIf Nanosecond > 999 Then
            mp_dtDateTime = New System.DateTime(0)
            mp_lSecondFraction = 0
            Return
        End If
        mp_dtDateTime = New System.DateTime(Year, Month, Day, Hour, Minute, Second, Millisecond)
        mp_lSecondFraction = (Millisecond * 1000000) + (Microsecond * 1000) + Nanosecond
    End Sub

    Public Sub SetDateTime(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer)
        Initialize(Year, Month, Day, Hour, Minute, Second, 0, 0, 0)
    End Sub

    Public Sub SetDate(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer)
        Initialize(Year, Month, Day, 0, 0, 0, 0, 0, 0)
    End Sub

    Public Sub SetTime(ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer)
        Initialize(1, 1, 1, Hour, Minute, Second, 0, 0, 0)
    End Sub

    Public Sub SetSecondFraction(ByVal Millisecond As Integer, ByVal Microsecond As Integer, ByVal Nanosecond As Integer)
        Initialize(1, 1, 1, 0, 0, 0, Millisecond, Microsecond, Nanosecond)
    End Sub

    Public Function Nanosecond() As Integer
        Dim lDateTime As Integer = mp_lSecondFraction
        Dim lMillisecond As Integer = Millisecond()
        Dim lMicrosecond As Integer = Microsecond()
        lDateTime = lDateTime - (lMillisecond * 1000000)
        lDateTime = lDateTime - (lMicrosecond * 1000)
        Return lDateTime
    End Function

    Public Function Microsecond() As Integer
        Dim lDateTime As Integer = mp_lSecondFraction
        Dim lMillisecond As Integer = Millisecond()
        lDateTime = lDateTime - (lMillisecond * 1000000)
        Return System.Convert.ToInt32(System.Math.Floor(lDateTime / 1000))
    End Function

    Public Function Millisecond() As Integer
        Return System.Convert.ToInt32(System.Math.Floor(mp_lSecondFraction / 1000000))
    End Function

    Public Function Second() As Integer
        Return mp_dtDateTime.Second
    End Function

    Public Function Minute() As Integer
        Return mp_dtDateTime.Minute
    End Function

    Public Function Hour() As Integer
        Return mp_dtDateTime.Hour
    End Function

    Public Function Day() As Integer
        Return mp_dtDateTime.Day
    End Function

    Public Function DayOfWeek() As Integer
        Return mp_dtDateTime.DayOfWeek + 1
    End Function

    Public Function DayOfYear() As Integer
        Return mp_dtDateTime.DayOfYear
    End Function

    Public Function Week() As Integer
        Dim lWeekDay As Integer
        Dim lDayOfYear As Integer
        lWeekDay = mp_dtDateTime.DayOfWeek
        lDayOfYear = mp_dtDateTime.DayOfYear
        Return System.Convert.ToInt32((System.Math.Floor(lDayOfYear - lWeekDay) / 7) + 1)
    End Function

    Public Function Month() As Integer
        Return mp_dtDateTime.Month
    End Function

    Public Function Quarter() As Integer
        Select Case mp_dtDateTime.Month
            Case 1 To 3
                Return 1
            Case 4 To 6
                Return 2
            Case 7 To 9
                Return 3
            Case 10 To 12
                Return 4
        End Select
        Return -1
    End Function

    Public Function Year() As Integer
        Return mp_dtDateTime.Year
    End Function

    Public Sub SetToCurrentDateTime()
        mp_dtDateTime = System.DateTime.Now
        mp_lSecondFraction = (mp_dtDateTime.Millisecond * 1000000)
    End Sub

    Public Shared Operator =(ByVal DateTime1 As AGVBE.DateTime, ByVal DateTime2 As AGVBE.DateTime) As Boolean
        If ((DateTime1.mp_dtDateTime = DateTime2.mp_dtDateTime) And (DateTime1.mp_lSecondFraction = DateTime2.mp_lSecondFraction)) Then
            Return True
        Else
            Return False
        End If
    End Operator

    Public Overloads Overrides Function Equals(ByVal obj As Object) As Boolean
        If obj Is Nothing Or Not Me.GetType() Is obj.GetType() Then
            Return False
        End If
        Dim DateTime2 As AGVBE.DateTime = CType(obj, AGVBE.DateTime)
        If ((mp_dtDateTime = DateTime2.mp_dtDateTime) And (mp_lSecondFraction = DateTime2.mp_lSecondFraction)) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Overloads Function Equals(ByVal DateTime2 As AGVBE.DateTime) As Boolean
        If ((mp_dtDateTime = DateTime2.mp_dtDateTime) And (mp_lSecondFraction = DateTime2.mp_lSecondFraction)) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return System.Convert.ToInt32(mp_dtDateTime.GetHashCode() ^ mp_lSecondFraction)
    End Function

    Public Shared Operator <>(ByVal DateTime1 As AGVBE.DateTime, ByVal DateTime2 As AGVBE.DateTime) As Boolean
        If ((DateTime1 Is Nothing) Or (DateTime2 Is Nothing)) Then
            Return False
        End If
        If ((DateTime1.mp_dtDateTime = DateTime2.mp_dtDateTime) And (DateTime1.mp_lSecondFraction = DateTime2.mp_lSecondFraction)) Then
            Return False
        Else
            Return True
        End If
    End Operator

    Public Shared Operator <=(ByVal DateTime1 As AGVBE.DateTime, ByVal DateTime2 As AGVBE.DateTime) As Boolean
        If ((DateTime1.mp_dtDateTime = DateTime2.mp_dtDateTime) And (DateTime1.mp_lSecondFraction = DateTime2.mp_lSecondFraction)) Then
            Return True
        End If
        If (DateTime1.mp_dtDateTime < DateTime2.mp_dtDateTime) Then
            Return True
        End If
        If ((DateTime1.mp_dtDateTime = DateTime2.mp_dtDateTime) And (DateTime1.mp_lSecondFraction < DateTime2.mp_lSecondFraction)) Then
            Return True
        End If
        Return False
    End Operator

    Public Shared Operator >=(ByVal DateTime1 As AGVBE.DateTime, ByVal DateTime2 As AGVBE.DateTime) As Boolean
        If ((DateTime1.mp_dtDateTime = DateTime2.mp_dtDateTime) And (DateTime1.mp_lSecondFraction = DateTime2.mp_lSecondFraction)) Then
            Return True
        End If
        If (DateTime1.mp_dtDateTime > DateTime2.mp_dtDateTime) Then
            Return True
        End If
        If ((DateTime1.mp_dtDateTime = DateTime2.mp_dtDateTime) And (DateTime1.mp_lSecondFraction > DateTime2.mp_lSecondFraction)) Then
            Return True
        End If
        Return False
    End Operator

    Public Shared Operator <(ByVal DateTime1 As AGVBE.DateTime, ByVal DateTime2 As AGVBE.DateTime) As Boolean
        If (DateTime1.mp_dtDateTime < DateTime2.mp_dtDateTime) Then
            Return True
        End If
        If ((DateTime1.mp_dtDateTime = DateTime2.mp_dtDateTime) And (DateTime1.mp_lSecondFraction < DateTime2.mp_lSecondFraction)) Then
            Return True
        End If
        Return False
    End Operator

    Public Shared Operator >(ByVal DateTime1 As AGVBE.DateTime, ByVal DateTime2 As AGVBE.DateTime) As Boolean
        If (DateTime1.mp_dtDateTime > DateTime2.mp_dtDateTime) Then
            Return True
        End If
        If ((DateTime1.mp_dtDateTime = DateTime2.mp_dtDateTime) And (DateTime1.mp_lSecondFraction > DateTime2.mp_lSecondFraction)) Then
            Return True
        End If
        Return False
    End Operator

    Public Sub Clear()
        mp_dtDateTime = New System.DateTime(0)
        mp_lSecondFraction = 0
    End Sub

    Public Property DateTimePart() As System.DateTime
        Get
            Return mp_dtDateTime
        End Get
        Set(ByVal value As System.DateTime)
            mp_dtDateTime = value
            If (mp_dtDateTime.Millisecond <> 0) Then
                mp_lSecondFraction = (mp_dtDateTime.Millisecond * 1000000)
            End If
        End Set
    End Property

    Public Property SecondFractionPart() As Integer
        Get
            Return mp_lSecondFraction
        End Get
        Set(ByVal value As Integer)
            mp_lSecondFraction = value
            mp_dtDateTime = New System.DateTime(mp_dtDateTime.Year, mp_dtDateTime.Month, mp_dtDateTime.Day, mp_dtDateTime.Hour, mp_dtDateTime.Minute, mp_dtDateTime.Second, Millisecond)
        End Set
    End Property

    Public Overloads Function ToString(ByVal Format As String) As String
        Return ToString(Format, System.Globalization.CultureInfo.InvariantCulture)
    End Function

    Public Overloads Function ToString(ByVal Format As String, ByVal provider As System.IFormatProvider) As String
        Dim sReturn As String = ""
        Dim i As Integer
        If Format.Length = 0 Then
            Return mp_dtDateTime.ToString(Format)
        End If
        For i = 9 To 1 Step -1
            Format = GetSecondFractionFormat(i, False, Format)
        Next
        For i = 9 To 1 Step -1
            Format = GetSecondFractionFormat(i, True, Format)
        Next
        If IsNumeric(Format) Then
            sReturn = Format
        Else
            sReturn = mp_dtDateTime.ToString(Format, provider)
        End If
        Return sReturn
    End Function

    Private Function GetSecondFractionFormat(ByVal lDigits As Integer, ByVal bCapital As Boolean, ByVal Format As String) As String
        Dim sReturn As String = ""
        Dim sBuff As String = ""
        Dim sFormatSpec As String = ""
        If bCapital = False Then
            While sFormatSpec.Length < lDigits
                sFormatSpec = sFormatSpec & "f"
            End While
            If Format.Contains(sFormatSpec) = False Then
                Return Format
            End If
            sBuff = mp_lSecondFraction.ToString()
            While sBuff.Length < 9
                sBuff = "0" & sBuff
            End While
            sBuff = sBuff.Substring(0, lDigits)
            sReturn = sBuff
        Else
            While sFormatSpec.Length < lDigits
                sFormatSpec = sFormatSpec & "F"
            End While
            If Format.Contains(sFormatSpec) = False Then
                Return Format
            End If
            sBuff = mp_lSecondFraction.ToString()
            While sBuff.Length < 9
                sBuff = "0" & sBuff
            End While
            sBuff = sBuff.Substring(0, lDigits)
            Dim i As Integer = System.Convert.ToInt32(sBuff)
            If i > 0 Then
                sReturn = sBuff
            Else
                sReturn = ""
            End If
        End If
        sReturn = Format.Replace(sFormatSpec, sReturn)
        Return sReturn
    End Function

    Private Function IsNumeric(ByVal Expression As String) As Boolean
        Dim dDummy As Double
        Return Double.TryParse(Expression, dDummy)
    End Function

End Class
