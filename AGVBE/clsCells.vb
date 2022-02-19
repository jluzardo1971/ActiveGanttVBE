Option Explicit On 

Public Class clsCells

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_oCollection As clsCollectionBase
    Private mp_oRow As clsRow

    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal oRow As clsRow)
        mp_oControl = Value
        mp_oCollection = New clsCollectionBase(Value, "Cell")
        mp_oRow = oRow
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    
    Public ReadOnly Property Count() As Integer
        Get
            Return mp_oCollection.m_lCount
        End Get
    End Property

    
    Public Function Item(ByVal Index As String) As clsCell
        Return DirectCast(mp_oCollection.m_oItem(Index, SYS_ERRORS.CELLS_ITEM_1, SYS_ERRORS.CELLS_ITEM_2, SYS_ERRORS.CELLS_ITEM_3, SYS_ERRORS.CELLS_ITEM_4), clsCell)
    End Function

    Friend ReadOnly Property oCollection() As clsCollectionBase
        Get
            Return mp_oCollection
        End Get
    End Property

    Friend Sub Add()
        mp_oCollection.AddMode = True
        Dim oCell As New clsCell(mp_oControl, Me)
        mp_oCollection.m_Add(oCell, "", SYS_ERRORS.CELLS_ADD_1, SYS_ERRORS.CELLS_ADD_2, False, SYS_ERRORS.CELLS_ADD_3)
    End Sub

    Friend Sub Clear()
        mp_oCollection.m_Clear()
    End Sub

    Friend Sub Remove(ByVal Index As String)
        mp_oCollection.m_Remove(Index, SYS_ERRORS.CELLS_REMOVE_1, SYS_ERRORS.CELLS_REMOVE_2, SYS_ERRORS.CELLS_REMOVE_3, SYS_ERRORS.CELLS_REMOVE_4)
    End Sub

    Friend Function Row() As clsRow
        Return mp_oRow
    End Function

    
    Public Function GetXML() As String
        Dim lIndex As Integer
        Dim oCell As clsCell
        Dim oXML As New clsXML(mp_oControl, "Cells")
        oXML.InitializeWriter()
        For lIndex = 1 To Count
            oCell = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsCell)
            oXML.WriteObject(oCell.GetXML)
        Next lIndex
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim lIndex As Integer
        Dim oXML As New clsXML(mp_oControl, "Cells")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        mp_oCollection.m_Clear()
        For lIndex = 1 To oXML.ReadCollectionCount
            Dim oCell As New clsCell(mp_oControl, Me)
            oCell.SetXML(oXML.ReadCollectionObject(lIndex))
            mp_oCollection.AddMode = True
            mp_oCollection.m_Add(oCell, "", SYS_ERRORS.CELLS_ADD_1, SYS_ERRORS.CELLS_ADD_2, False, SYS_ERRORS.CELLS_ADD_3)
            oCell = Nothing
        Next lIndex
    End Sub

End Class
