Option Explicit On 

Public Class clsLayers

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_oCollection As clsCollectionBase
    Private mp_oDefaultLayer As clsLayer

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_oCollection = New clsCollectionBase(Value, "Layer")
        mp_oDefaultLayer = New clsLayer(Value)
    End Sub

    
    Public ReadOnly Property Count() As Integer
        Get
            Return mp_oCollection.m_lCount
        End Get
    End Property

    
    Public Function Item(ByVal Index As String) As clsLayer
        Return DirectCast(mp_oCollection.m_oItem(Index, SYS_ERRORS.LAYERS_ITEM_1, SYS_ERRORS.LAYERS_ITEM_2, SYS_ERRORS.LAYERS_ITEM_3, SYS_ERRORS.LAYERS_ITEM_4), clsLayer)
    End Function

    Friend Function FItem(ByVal Index As String) As clsLayer
        If Index = "0" Then
            Return mp_oDefaultLayer
        Else
            Return DirectCast(mp_oCollection.m_oItem(Index, SYS_ERRORS.LAYERS_ITEM_1, SYS_ERRORS.LAYERS_ITEM_2, SYS_ERRORS.LAYERS_ITEM_3, SYS_ERRORS.LAYERS_ITEM_4), clsLayer)
        End If
    End Function

    Friend ReadOnly Property oCollection() As clsCollectionBase
        Get
            Return mp_oCollection
        End Get
    End Property

    
    Public Function Add(ByVal Key As String, Optional ByVal Visible As Boolean = True) As clsLayer
        mp_oCollection.AddMode = True
        Dim oLayer As New clsLayer(mp_oControl)
        oLayer.Key = Key
        oLayer.Visible = Visible
        mp_oCollection.m_Add(oLayer, Key, SYS_ERRORS.LAYERS_ADD_1, SYS_ERRORS.LAYERS_ADD_2, False, SYS_ERRORS.LAYERS_ADD_3)
        Return oLayer
    End Function

    
    Public Sub Clear()
        mp_oControl.Tasks.oCollection.m_CollectionRemoveWhereNot("LayerIndex", "0")
        mp_oControl.CurrentLayer = "0"
        mp_oCollection.m_Clear()
    End Sub

    
    Public Sub Remove(ByVal Index As String)
        Dim sRIndex As String = ""
        Dim sRKey As String = ""
        mp_oCollection.m_GetKeyAndIndex(Index, sRKey, sRIndex)
        mp_oControl.Tasks.oCollection.m_CollectionRemoveWhere("LayerIndex", sRKey, sRIndex)
        mp_oControl.CurrentLayer = "0"
        mp_oCollection.m_Remove(Index, SYS_ERRORS.LAYERS_REMOVE_1, SYS_ERRORS.LAYERS_REMOVE_2, SYS_ERRORS.LAYERS_REMOVE_3, SYS_ERRORS.LAYERS_REMOVE_4)
    End Sub

    
    Public Function GetXML() As String
        Dim lIndex As Integer
        Dim oLayer As clsLayer
        Dim oXML As New clsXML(mp_oControl, "Layers")
        oXML.InitializeWriter()
        For lIndex = 1 To Count
            oLayer = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsLayer)
            oXML.WriteObject(oLayer.GetXML)
        Next lIndex
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim lIndex As Integer
        Dim oXML As New clsXML(mp_oControl, "Layers")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        mp_oCollection.m_Clear()
        For lIndex = 1 To oXML.ReadCollectionCount
            Dim oLayer As New clsLayer(mp_oControl)
            oLayer.SetXML(oXML.ReadCollectionObject(lIndex))
            mp_oCollection.AddMode = True
            mp_oCollection.m_Add(oLayer, oLayer.Key, SYS_ERRORS.LAYERS_ADD_1, SYS_ERRORS.LAYERS_ADD_2, False, SYS_ERRORS.LAYERS_ADD_3)
            oLayer = Nothing
        Next lIndex
    End Sub

End Class
