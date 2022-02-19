Option Explicit On 

Public Class clsTierFormat

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_sMicrosecondIntervalFormat As String
    Private mp_sMillisecondIntervalFormat As String
    Private mp_sSecondIntervalFormat As String
    Private mp_sMinuteIntervalFormat As String
    Private mp_sHourIntervalFormat As String
    Private mp_sDayIntervalFormat As String
    Private mp_sDayOfWeekIntervalFormat As String
    Private mp_sDayOfYearIntervalFormat As String
    Private mp_sWeekIntervalFormat As String
    Private mp_sMonthIntervalFormat As String
    Private mp_sQuarterIntervalFormat As String
    Private mp_sYearIntervalFormat As String

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_sMicrosecondIntervalFormat = "ffffff"
        mp_sMillisecondIntervalFormat = "fff"
        mp_sSecondIntervalFormat = "ss"
        mp_sMinuteIntervalFormat = "mm"
        mp_sHourIntervalFormat = "HH:mm"
        mp_sDayIntervalFormat = "d"
        mp_sDayOfWeekIntervalFormat = "dddd d"
        mp_sDayOfYearIntervalFormat = "y"
        mp_sWeekIntervalFormat = "ww"
        mp_sMonthIntervalFormat = "MMMM yyyy"
        mp_sQuarterIntervalFormat = "q""Q"" yyyy"
        mp_sYearIntervalFormat = "yyyy"
    End Sub

    Public Property MicrosecondIntervalFormat() As String
        Get
            Return mp_sMicrosecondIntervalFormat
        End Get
        Set(ByVal Value As String)
            mp_sMicrosecondIntervalFormat = Value
        End Set
    End Property

    Public Property MillisecondIntervalFormat() As String
        Get
            Return mp_sMillisecondIntervalFormat
        End Get
        Set(ByVal Value As String)
            mp_sMillisecondIntervalFormat = Value
        End Set
    End Property

    Public Property SecondIntervalFormat() As String
        Get
            Return mp_sSecondIntervalFormat
        End Get
        Set(ByVal Value As String)
            mp_sSecondIntervalFormat = Value
        End Set
    End Property

    
    Public Property MinuteIntervalFormat() As String
        Get
            Return mp_sMinuteIntervalFormat
        End Get
        Set(ByVal Value As String)
            mp_sMinuteIntervalFormat = Value
        End Set
    End Property

    
    Public Property HourIntervalFormat() As String
        Get
            Return mp_sHourIntervalFormat
        End Get
        Set(ByVal Value As String)
            mp_sHourIntervalFormat = Value
        End Set
    End Property

    
    Public Property DayIntervalFormat() As String
        Get
            Return mp_sDayIntervalFormat
        End Get
        Set(ByVal Value As String)
            mp_sDayIntervalFormat = Value
        End Set
    End Property

    
    Public Property DayOfWeekIntervalFormat() As String
        Get
            Return mp_sDayOfWeekIntervalFormat
        End Get
        Set(ByVal Value As String)
            mp_sDayOfWeekIntervalFormat = Value
        End Set
    End Property

    
    Public Property DayOfYearIntervalFormat() As String
        Get
            Return mp_sDayOfYearIntervalFormat
        End Get
        Set(ByVal Value As String)
            mp_sDayOfYearIntervalFormat = Value
        End Set
    End Property

    
    Public Property WeekIntervalFormat() As String
        Get
            Return mp_sWeekIntervalFormat
        End Get
        Set(ByVal Value As String)
            mp_sWeekIntervalFormat = Value
        End Set
    End Property

    
    Public Property MonthIntervalFormat() As String
        Get
            Return mp_sMonthIntervalFormat
        End Get
        Set(ByVal Value As String)
            mp_sMonthIntervalFormat = Value
        End Set
    End Property

    
    Public Property QuarterIntervalFormat() As String
        Get
            Return mp_sQuarterIntervalFormat
        End Get
        Set(ByVal Value As String)
            mp_sQuarterIntervalFormat = Value
        End Set
    End Property

    
    Public Property YearIntervalFormat() As String
        Get
            Return mp_sYearIntervalFormat
        End Get
        Set(ByVal Value As String)
            mp_sYearIntervalFormat = Value
        End Set
    End Property

    
    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "TierFormat")
        oXML.InitializeWriter()
        oXML.WriteProperty("DayIntervalFormat", mp_sDayIntervalFormat)
        oXML.WriteProperty("DayOfWeekIntervalFormat", mp_sDayOfWeekIntervalFormat)
        oXML.WriteProperty("DayOfYearIntervalFormat", mp_sDayOfYearIntervalFormat)
        oXML.WriteProperty("HourIntervalFormat", mp_sHourIntervalFormat)
        oXML.WriteProperty("MinuteIntervalFormat", mp_sMinuteIntervalFormat)
        oXML.WriteProperty("SecondIntervalFormat", mp_sSecondIntervalFormat)
        oXML.WriteProperty("MillisecondIntervalFormat", mp_sMillisecondIntervalFormat)
        oXML.WriteProperty("MicrosecondIntervalFormat", mp_sMicrosecondIntervalFormat)
        oXML.WriteProperty("MonthIntervalFormat", mp_sMonthIntervalFormat)
        oXML.WriteProperty("QuarterIntervalFormat", mp_sQuarterIntervalFormat)
        oXML.WriteProperty("WeekIntervalFormat", mp_sWeekIntervalFormat)
        oXML.WriteProperty("YearIntervalFormat", mp_sYearIntervalFormat)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "TierFormat")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("DayIntervalFormat", mp_sDayIntervalFormat)
        oXML.ReadProperty("DayOfWeekIntervalFormat", mp_sDayOfWeekIntervalFormat)
        oXML.ReadProperty("DayOfYearIntervalFormat", mp_sDayOfYearIntervalFormat)
        oXML.ReadProperty("HourIntervalFormat", mp_sHourIntervalFormat)
        oXML.ReadProperty("MinuteIntervalFormat", mp_sMinuteIntervalFormat)
        oXML.ReadProperty("SecondIntervalFormat", mp_sSecondIntervalFormat)
        oXML.ReadProperty("MillisecondIntervalFormat", mp_sMillisecondIntervalFormat)
        oXML.ReadProperty("MicrosecondIntervalFormat", mp_sMicrosecondIntervalFormat)
        oXML.ReadProperty("MonthIntervalFormat", mp_sMonthIntervalFormat)
        oXML.ReadProperty("QuarterIntervalFormat", mp_sQuarterIntervalFormat)
        oXML.ReadProperty("WeekIntervalFormat", mp_sWeekIntervalFormat)
        oXML.ReadProperty("YearIntervalFormat", mp_sYearIntervalFormat)
    End Sub

End Class
