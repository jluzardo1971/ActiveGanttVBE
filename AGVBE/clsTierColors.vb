Option Explicit On

Public Class clsTierColors

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_oCollection As clsCollectionBase
    Private mp_yTierType As E_TIERTYPE

    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal yTierType As E_TIERTYPE)
        mp_oControl = Value
        mp_oCollection = New clsCollectionBase(Value, "TierColor")
        mp_yTierType = yTierType
    End Sub

    
    Public ReadOnly Property Count() As Integer
        Get
            Return mp_oCollection.m_lCount
        End Get
    End Property

    
    Public Function Item(ByVal Index As String) As clsTierColor
        Return DirectCast(mp_oCollection.m_oItem(Index, SYS_ERRORS.TIERCOLORS_ITEM_1, SYS_ERRORS.TIERCOLORS_ITEM_2, SYS_ERRORS.TIERCOLORS_ITEM_3, SYS_ERRORS.TIERCOLORS_ITEM_4), clsTierColor)
    End Function

    Friend ReadOnly Property oCollection() As clsCollectionBase
        Get
            Return mp_oCollection
        End Get
    End Property

    Friend Sub Add(ByVal BackColor As Color, ByVal ForeColor As Color, ByVal StartGradientColor As Color, ByVal EndGradientColor As Color, ByVal HatchBackColor As Color, ByVal HatchForeColor As Color, Optional ByVal Key As String = "")
        mp_oCollection.AddMode = True
        Dim oTierColor As New clsTierColor(mp_oControl, Me)
        oTierColor.BackColor = BackColor
        oTierColor.ForeColor = ForeColor
        oTierColor.StartGradientColor = StartGradientColor
        oTierColor.EndGradientColor = EndGradientColor
        oTierColor.HatchBackColor = HatchBackColor
        oTierColor.HatchForeColor = HatchForeColor
        oTierColor.Key = Key
        mp_oCollection.m_Add(oTierColor, Key, SYS_ERRORS.TIERCOLORS_ADD_1, SYS_ERRORS.TIERCOLORS_ADD_2, False, SYS_ERRORS.TIERCOLORS_ADD_3)
    End Sub

    Private Function mp_CollectionName() As String
        Select Case mp_yTierType
            Case E_TIERTYPE.ST_MICROSECOND
                Return "MicrosecondColors"
            Case E_TIERTYPE.ST_MILLISECOND
                Return "MillisecondColors"
            Case E_TIERTYPE.ST_SECOND
                Return "SecondColors"
            Case E_TIERTYPE.ST_MINUTE
                Return "MinuteColors"
            Case E_TIERTYPE.ST_HOUR
                Return "HourColors"
            Case E_TIERTYPE.ST_DAY
                Return "DayColors"
            Case E_TIERTYPE.ST_DAYOFWEEK
                Return "DayOfWeekColors"
            Case E_TIERTYPE.ST_DAYOFYEAR
                Return "DayOfYearColors"
            Case E_TIERTYPE.ST_WEEK
                Return "WeekColors"
            Case E_TIERTYPE.ST_MONTH
                Return "MonthColors"
            Case E_TIERTYPE.ST_QUARTER
                Return "QuarterColors"
            Case E_TIERTYPE.ST_YEAR
                Return "YearColors"
        End Select
        Return ""
    End Function

    
    Public Function GetXML() As String
        Dim lIndex As Integer
        Dim oTierColor As clsTierColor
        Dim oXML As New clsXML(mp_oControl, mp_CollectionName)
        oXML.InitializeWriter()
        For lIndex = 1 To Count
            oTierColor = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsTierColor)
            oXML.WriteObject(oTierColor.GetXML)
        Next lIndex
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim lIndex As Integer
        Dim oXML As New clsXML(mp_oControl, mp_CollectionName)
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        mp_oCollection.m_Clear()
        For lIndex = 1 To oXML.ReadCollectionCount
            Dim oTierColor As New clsTierColor(mp_oControl, Me)
            oTierColor.SetXML(oXML.ReadCollectionObject(lIndex))
            mp_oCollection.AddMode = True
            mp_oCollection.m_Add(oTierColor, oTierColor.Key, SYS_ERRORS.TIERCOLORS_ADD_1, SYS_ERRORS.TIERCOLORS_ADD_2, False, SYS_ERRORS.TIERCOLORS_ADD_3)
        Next lIndex
    End Sub

End Class
