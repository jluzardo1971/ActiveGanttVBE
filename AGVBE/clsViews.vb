Option Explicit On 

Public Class clsViews

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_oCollection As clsCollectionBase
    Private mp_oDefaultView As clsView

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_oCollection = New clsCollectionBase(Value, "View")
        mp_oDefaultView = New clsView(mp_oControl)
        mp_oDefaultView.Interval = E_INTERVAL.IL_MINUTE
        mp_oDefaultView.Factor = 1
        mp_oDefaultView.TimeLine.TierArea.UpperTier.TierType = E_TIERTYPE.ST_MONTH
        mp_oDefaultView.TimeLine.TierArea.MiddleTier.Visible = False
        mp_oDefaultView.TimeLine.TierArea.LowerTier.TierType = E_TIERTYPE.ST_DAYOFWEEK
        mp_oDefaultView.ClientArea.ToolTipFormat = "hh:mmtt"
        mp_oDefaultView.TimeLine.TickMarkArea.TickMarks.Add(E_INTERVAL.IL_HOUR, 1, E_TICKMARKTYPES.TLT_BIG, True, "hh:mmtt")
        mp_oDefaultView.TimeLine.TickMarkArea.TickMarks.Add(E_INTERVAL.IL_MINUTE, 30, E_TICKMARKTYPES.TLT_BIG)
        mp_oDefaultView.TimeLine.TickMarkArea.TickMarks.Add(E_INTERVAL.IL_MINUTE, 15, E_TICKMARKTYPES.TLT_MEDIUM)
        mp_oDefaultView.TimeLine.TickMarkArea.TickMarks.Add(E_INTERVAL.IL_MINUTE, 5, E_TICKMARKTYPES.TLT_SMALL)
        mp_oDefaultView.TimeLine.TickMarkArea.Height = 30
    End Sub

    
    Public ReadOnly Property Count() As Integer
        Get
            Return mp_oCollection.m_lCount
        End Get
    End Property

    
    Public Function Item(ByVal Index As String) As clsView
        Return DirectCast(mp_oCollection.m_oItem(Index, SYS_ERRORS.VIEWS_ITEM_1, SYS_ERRORS.VIEWS_ITEM_2, SYS_ERRORS.VIEWS_ITEM_3, SYS_ERRORS.VIEWS_ITEM_4), clsView)
    End Function

    Friend Function FItem(ByVal Index As String) As clsView
        If Index = "0" Then
            FItem = mp_oDefaultView
        Else
            FItem = DirectCast(mp_oCollection.m_oItem(Index, SYS_ERRORS.VIEWS_ITEM_1, SYS_ERRORS.VIEWS_ITEM_2, SYS_ERRORS.VIEWS_ITEM_3, SYS_ERRORS.VIEWS_ITEM_4), clsView)
        End If
    End Function

    Friend ReadOnly Property oCollection() As clsCollectionBase
        Get
            Return mp_oCollection
        End Get
    End Property

    
    Public Function Add(ByVal Interval As E_INTERVAL, ByVal Factor As Integer, ByVal UpperTierType As E_TIERTYPE, ByVal MiddleTierType As E_TIERTYPE, ByVal LowerTierType As E_TIERTYPE, Optional ByVal Key As String = "") As clsView
        mp_oCollection.AddMode = True
        Dim oView As New clsView(mp_oControl)
        oView.Interval = Interval
        oView.Factor = Factor
        oView.TimeLine.TierArea.UpperTier.TierType = UpperTierType
        oView.TimeLine.TierArea.MiddleTier.TierType = MiddleTierType
        oView.TimeLine.TierArea.LowerTier.TierType = LowerTierType
        oView.Key = Key
        mp_oCollection.m_Add(oView, Key, SYS_ERRORS.VIEWS_ADD_1, SYS_ERRORS.VIEWS_ADD_2, False, SYS_ERRORS.VIEWS_ADD_3)
        Return oView
    End Function

    
    Public Sub Clear()
        mp_oControl.CurrentView = "0"
        mp_oCollection.m_Clear()
    End Sub

    
    Public Sub Remove(ByVal Index As String)
        mp_oControl.CurrentView = "0"
        mp_oCollection.m_Remove(Index, SYS_ERRORS.VIEWS_REMOVE_1, SYS_ERRORS.VIEWS_REMOVE_2, SYS_ERRORS.VIEWS_REMOVE_3, SYS_ERRORS.VIEWS_REMOVE_4)
    End Sub

    
    Public Function GetXML() As String
        Dim lIndex As Integer
        Dim oView As clsView
        Dim oXML As New clsXML(mp_oControl, "Views")
        oXML.InitializeWriter()
        For lIndex = 1 To Count
            oView = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsView)
            oXML.WriteObject(oView.GetXML)
        Next lIndex
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim lIndex As Integer
        Dim oXML As New clsXML(mp_oControl, "Views")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        mp_oCollection.m_Clear()
        For lIndex = 1 To oXML.ReadCollectionCount
            Dim oView As New clsView(mp_oControl)
            oView.SetXML(oXML.ReadCollectionObject(lIndex))
            mp_oCollection.AddMode = True
            mp_oCollection.m_Add(oView, oView.Key, SYS_ERRORS.VIEWS_ADD_1, SYS_ERRORS.VIEWS_ADD_2, False, SYS_ERRORS.VIEWS_ADD_3)
            oView = Nothing
        Next lIndex
    End Sub

End Class
