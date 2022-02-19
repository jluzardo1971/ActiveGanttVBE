Option Explicit On 

Public Class clsPercentages

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_oCollection As clsCollectionBase

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_oCollection = New clsCollectionBase(Value, "Percentage")
    End Sub

    
    Public ReadOnly Property Count() As Integer
        Get
            Return mp_oCollection.m_lCount
        End Get
    End Property

    
    Public Function Item(ByVal Index As String) As clsPercentage
        Return DirectCast(mp_oCollection.m_oItem(Index, SYS_ERRORS.PERCENTAGES_ITEM_1, SYS_ERRORS.PERCENTAGES_ITEM_2, SYS_ERRORS.PERCENTAGES_ITEM_3, SYS_ERRORS.PERCENTAGES_ITEM_4), clsPercentage)
    End Function

    Friend ReadOnly Property oCollection() As clsCollectionBase
        Get
            Return mp_oCollection
        End Get
    End Property

    
    Public Function Add(ByVal TaskKey As String, ByVal StyleIndex As String, ByVal Percent As Single, Optional ByVal Key As String = "") As clsPercentage
        mp_oCollection.AddMode = True
        Dim oPercentage As New clsPercentage(mp_oControl)
        Key = mp_oControl.StrLib.StrTrim(Key)
        TaskKey = mp_oControl.StrLib.StrTrim(TaskKey)
        oPercentage.Key = Key
        oPercentage.TaskKey = TaskKey
        oPercentage.Percent = Percent
        oPercentage.StyleIndex = StyleIndex
        mp_oCollection.m_Add(oPercentage, Key, SYS_ERRORS.PERCENTAGES_ADD_1, SYS_ERRORS.PERCENTAGES_ADD_2, False, SYS_ERRORS.PERCENTAGES_ADD_3)
        Return oPercentage
    End Function

    
    Public Sub Clear()
        mp_oCollection.m_Clear()
    End Sub

    
    Public Sub Remove(ByVal Index As String)
        mp_oCollection.m_Remove(Index, SYS_ERRORS.PERCENTAGES_REMOVE_1, SYS_ERRORS.PERCENTAGES_REMOVE_2, SYS_ERRORS.PERCENTAGES_REMOVE_3, SYS_ERRORS.PERCENTAGES_REMOVE_4)
    End Sub

    Friend Sub Draw()
        Dim lIndex As Integer
        Dim oPercentage As clsPercentage
        If Count = 0 Then
            Return
        End If
        If Count = 0 Then
            Return
        End If
        For lIndex = 1 To Count
            oPercentage = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsPercentage)
            If oPercentage.Visible = True Then
                mp_oControl.clsG.ClipRegion(oPercentage.LeftTrim, oPercentage.Top, oPercentage.RightTrim, oPercentage.Bottom, True)
                mp_oControl.DrawEventArgs.Clear()
                mp_oControl.DrawEventArgs.CustomDraw = False
                mp_oControl.DrawEventArgs.EventTarget = E_EVENTTARGET.EVT_PERCENTAGE
                mp_oControl.DrawEventArgs.ObjectIndex = lIndex
                mp_oControl.DrawEventArgs.ParentObjectIndex = 0
                'mp_oControl.DrawEventArgs.Graphics = mp_oControl.clsG.oGraphics
                mp_oControl.FireDraw()
                If mp_oControl.DrawEventArgs.CustomDraw = False Then
                    mp_oControl.clsG.mp_DrawItem(oPercentage.Left, oPercentage.Top, oPercentage.Right, oPercentage.Bottom, "", mp_oControl.StrLib.StrFormat(oPercentage.Percent, oPercentage.Format), oPercentage.Index = mp_oControl.SelectedPercentageIndex, Nothing, oPercentage.LeftTrim, oPercentage.RightTrim, oPercentage.Style)
                End If
            End If
        Next lIndex
    End Sub

    
    Public Function GetXML() As String
        Dim lIndex As Integer
        Dim oPercentage As clsPercentage
        Dim oXML As New clsXML(mp_oControl, "Percentages")
        oXML.InitializeWriter()
        For lIndex = 1 To Count
            oPercentage = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsPercentage)
            oXML.WriteObject(oPercentage.GetXML)
        Next lIndex
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim lIndex As Integer
        Dim oXML As New clsXML(mp_oControl, "Percentages")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        mp_oCollection.m_Clear()
        For lIndex = 1 To oXML.ReadCollectionCount
            Dim oPercentage As New clsPercentage(mp_oControl)
            oPercentage.SetXML(oXML.ReadCollectionObject(lIndex))
            mp_oCollection.AddMode = True
            mp_oCollection.m_Add(oPercentage, oPercentage.Key, SYS_ERRORS.PERCENTAGES_ADD_1, SYS_ERRORS.PERCENTAGES_ADD_2, False, SYS_ERRORS.PERCENTAGES_ADD_3)
            oPercentage = Nothing
        Next lIndex
    End Sub

End Class
