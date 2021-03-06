Option Explicit On

Public Class clsTierAppearance

    Private mp_oControl As ActiveGanttVBECtl
    Public MicrosecondColors As clsTierColors
    Public MillisecondColors As clsTierColors
    Public SecondColors As clsTierColors
    Public MinuteColors As clsTierColors
    Public HourColors As clsTierColors
    Public DayColors As clsTierColors
    Public DayOfWeekColors As clsTierColors
    Public DayOfYearColors As clsTierColors
    Public WeekColors As clsTierColors
    Public MonthColors As clsTierColors
    Public QuarterColors As clsTierColors
    Public YearColors As clsTierColors

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        Dim Colors_DimGray As Color = Color.FromArgb(255, 105, 105, 105)
        Dim Colors_Silver As Color = Color.FromArgb(255, 192, 192, 192)
        Dim Colors_CornflowerBlue As Color = Color.FromArgb(255, 100, 149, 237)
        Dim Colors_MediumSlateBlue As Color = Color.FromArgb(255, 123, 104, 238)
        Dim Colors_SlateBlue As Color = Color.FromArgb(255, 106, 90, 205)
        Dim Colors_RoyalBlue As Color = Color.FromArgb(255, 65, 105, 225)
        Dim Colors_SteelBlue As Color = Color.FromArgb(255, 70, 130, 180)
        Dim Colors_CadetBlue As Color = Color.FromArgb(255, 95, 158, 160)
        Dim Colors_DodgerBlue As Color = Color.FromArgb(255, 30, 144, 255)

        mp_oControl = Value
        MicrosecondColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_MICROSECOND)
        MicrosecondColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        MicrosecondColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        MicrosecondColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        MicrosecondColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        MicrosecondColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        MicrosecondColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        MicrosecondColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        MicrosecondColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        MicrosecondColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        MicrosecondColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        MillisecondColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_MILLISECOND)
        MillisecondColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        MillisecondColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        MillisecondColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        MillisecondColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        MillisecondColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        MillisecondColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        MillisecondColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        MillisecondColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        MillisecondColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        MillisecondColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        SecondColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_SECOND)
        SecondColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        SecondColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        SecondColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        SecondColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        SecondColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        SecondColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        SecondColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        SecondColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        SecondColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        SecondColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        MinuteColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_MINUTE)
        MinuteColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        MinuteColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        MinuteColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        MinuteColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        MinuteColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        MinuteColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        MinuteColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        MinuteColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        MinuteColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        MinuteColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        HourColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_HOUR)
        HourColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        HourColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        HourColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        HourColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        HourColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        HourColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        HourColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        HourColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        HourColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        HourColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        DayColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_DAY)
        DayColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        DayColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        DayColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        DayColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        DayColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        DayColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        DayColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        DayColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        DayColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        DayColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        DayOfWeekColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_DAYOFWEEK)
        DayOfWeekColors.Add(Colors_CornflowerBlue, Colors.Black, Colors_CornflowerBlue, Colors.Black, Colors_CornflowerBlue, Colors.Black, "Sunday")
        DayOfWeekColors.Add(Colors_MediumSlateBlue, Colors.Black, Colors_MediumSlateBlue, Colors.Black, Colors_MediumSlateBlue, Colors.Black, "Monday")
        DayOfWeekColors.Add(Colors_SlateBlue, Colors.Black, Colors_SlateBlue, Colors.Black, Colors_SlateBlue, Colors.Black, "Tuesday")
        DayOfWeekColors.Add(Colors_RoyalBlue, Colors.Black, Colors_RoyalBlue, Colors.Black, Colors_RoyalBlue, Colors.Black, "Wednesday")
        DayOfWeekColors.Add(Colors_SteelBlue, Colors.Black, Colors_SteelBlue, Colors.Black, Colors_SteelBlue, Colors.Black, "Thursday")
        DayOfWeekColors.Add(Colors_CadetBlue, Colors.Black, Colors_CadetBlue, Colors.Black, Colors_CadetBlue, Colors.Black, "Friday")
        DayOfWeekColors.Add(Colors_DodgerBlue, Colors.Black, Colors_DodgerBlue, Colors.Black, Colors_DodgerBlue, Colors.Black, "Saturday")
        DayOfYearColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_DAYOFYEAR)
        DayOfYearColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        DayOfYearColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        DayOfYearColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        DayOfYearColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        DayOfYearColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        DayOfYearColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        DayOfYearColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        DayOfYearColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        DayOfYearColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        DayOfYearColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        WeekColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_WEEK)
        WeekColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        WeekColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        WeekColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        WeekColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        WeekColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        WeekColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        WeekColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        WeekColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        WeekColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        WeekColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        MonthColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_MONTH)
        MonthColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, "January")
        MonthColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, "February")
        MonthColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, "March")
        MonthColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, "April")
        MonthColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, "May")
        MonthColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, "June")
        MonthColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, "July")
        MonthColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, "August")
        MonthColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, "September")
        MonthColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, "October")
        MonthColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, "November")
        MonthColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, "December")
        QuarterColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_QUARTER)
        QuarterColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        QuarterColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        QuarterColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        QuarterColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        YearColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_YEAR)
        YearColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        YearColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        YearColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        YearColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        YearColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        YearColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)
        YearColors.Add(Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black, Colors.DarkGray, Colors.Black)
        YearColors.Add(Colors_Silver, Colors.Black, Colors_Silver, Colors.Black, Colors_Silver, Colors.Black)
        YearColors.Add(Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black, Colors_DimGray, Colors.Black)
        YearColors.Add(Colors.Gray, Colors.Black, Colors.Gray, Colors.Black, Colors.Gray, Colors.Black)

    End Sub

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "TierAppearance")
        oXML.InitializeWriter()
        oXML.WriteObject(DayColors.GetXML())
        oXML.WriteObject(DayOfWeekColors.GetXML())
        oXML.WriteObject(DayOfYearColors.GetXML())
        oXML.WriteObject(HourColors.GetXML())
        oXML.WriteObject(MinuteColors.GetXML())
        oXML.WriteObject(SecondColors.GetXML())
        oXML.WriteObject(MillisecondColors.GetXML())
        oXML.WriteObject(MicrosecondColors.GetXML())
        oXML.WriteObject(MonthColors.GetXML())
        oXML.WriteObject(QuarterColors.GetXML())
        oXML.WriteObject(WeekColors.GetXML())
        oXML.WriteObject(YearColors.GetXML())
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "TierAppearance")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        DayColors.SetXML(oXML.ReadObject("DayColors"))
        DayOfWeekColors.SetXML(oXML.ReadObject("DayOfWeekColors"))
        DayOfYearColors.SetXML(oXML.ReadObject("DayOfYearColors"))
        HourColors.SetXML(oXML.ReadObject("HourColors"))
        MinuteColors.SetXML(oXML.ReadObject("MinuteColors"))
        SecondColors.SetXML(oXML.ReadObject("SecondColors"))
        MillisecondColors.SetXML(oXML.ReadObject("MillisecondColors"))
        MicrosecondColors.SetXML(oXML.ReadObject("MicrosecondColors"))
        MonthColors.SetXML(oXML.ReadObject("MonthColors"))
        QuarterColors.SetXML(oXML.ReadObject("QuarterColors"))
        WeekColors.SetXML(oXML.ReadObject("WeekColors"))
        YearColors.SetXML(oXML.ReadObject("YearColors"))
    End Sub

End Class
