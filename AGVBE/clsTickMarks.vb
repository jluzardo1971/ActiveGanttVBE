Option Explicit On 

Public Class clsTickMarks

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_oCollection As clsCollectionBase

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_oCollection = New clsCollectionBase(Value, "TickMark")
    End Sub

    
    Public ReadOnly Property Count() As Integer
        Get
            Return mp_oCollection.m_lCount
        End Get
    End Property

    
    Public Function Item(ByVal Index As String) As clsTickMark
        Return DirectCast(mp_oCollection.m_oItem(Index, SYS_ERRORS.TICKMARKS_ITEM_1, SYS_ERRORS.TICKMARKS_ITEM_2, SYS_ERRORS.TICKMARKS_ITEM_3, SYS_ERRORS.TICKMARKS_ITEM_4), clsTickMark)
    End Function

    Friend ReadOnly Property oCollection() As clsCollectionBase
        Get
            Return mp_oCollection
        End Get
    End Property

    
    Public Function Add(ByVal Interval As E_INTERVAL, ByVal Factor As Integer, ByVal TickMarkType As E_TICKMARKTYPES) As clsTickMark
        Return Add(Interval, Factor, TickMarkType, False, "", "")
    End Function

    Public Function Add(ByVal Interval As E_INTERVAL, ByVal Factor As Integer, ByVal TickMarkType As E_TICKMARKTYPES, ByVal DisplayText As Boolean, ByVal TextFormat As String) As clsTickMark
        Return Add(Interval, Factor, TickMarkType, DisplayText, TextFormat, "")
    End Function

    Public Function Add(ByVal Interval As E_INTERVAL, ByVal Factor As Integer, ByVal TickMarkType As E_TICKMARKTYPES, ByVal DisplayText As Boolean, ByVal TextFormat As String, ByVal Key As String) As clsTickMark
        mp_oCollection.AddMode = True
        Dim oTickMark As New clsTickMark(mp_oControl, Me)
        oTickMark.Interval = Interval
        oTickMark.Factor = Factor
        oTickMark.TickMarkType = TickMarkType
        oTickMark.DisplayText = DisplayText
        oTickMark.TextFormat = TextFormat
        oTickMark.Key = Key
        mp_oCollection.m_Add(oTickMark, Key, SYS_ERRORS.TICKMARKS_ADD_1, SYS_ERRORS.TICKMARKS_ADD_2, False, SYS_ERRORS.TICKMARKS_ADD_3)
        Return oTickMark
    End Function

    
    Public Sub Clear()
        mp_oCollection.m_Clear()
    End Sub

    
    Public Sub Remove(ByVal Index As String)
        mp_oCollection.m_Remove(Index, SYS_ERRORS.TICKMARKS_REMOVE_1, SYS_ERRORS.TICKMARKS_REMOVE_2, SYS_ERRORS.TICKMARKS_REMOVE_3, SYS_ERRORS.TICKMARKS_REMOVE_4)
    End Sub

    
    Public Function GetXML() As String
        Dim lIndex As Integer
        Dim oTickMark As clsTickMark
        Dim oXML As New clsXML(mp_oControl, "TickMarks")
        oXML.InitializeWriter()
        For lIndex = 1 To Count
            oTickMark = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsTickMark)
            oXML.WriteObject(oTickMark.GetXML)
        Next lIndex
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim lIndex As Integer
        Dim oXML As New clsXML(mp_oControl, "TickMarks")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        mp_oCollection.m_Clear()
        For lIndex = 1 To oXML.ReadCollectionCount
            Dim oTickMark As New clsTickMark(mp_oControl, Me)
            oTickMark.SetXML(oXML.ReadCollectionObject(lIndex))
            mp_oCollection.AddMode = True
            mp_oCollection.m_Add(oTickMark, oTickMark.Key, SYS_ERRORS.TICKMARKS_ADD_1, SYS_ERRORS.TICKMARKS_ADD_2, False, SYS_ERRORS.TICKMARKS_ADD_3)
            oTickMark = Nothing
        Next lIndex
    End Sub

End Class
