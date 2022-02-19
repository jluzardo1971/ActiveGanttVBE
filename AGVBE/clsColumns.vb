Option Explicit On 

Public Class clsColumns

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_oCollection As clsCollectionBase

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_oCollection = New clsCollectionBase(Value, "Column")
    End Sub

    
    Public ReadOnly Property Count() As Integer
        Get
            Return mp_oCollection.m_lCount
        End Get
    End Property

    Public Function Item(ByVal Index As String) As clsColumn
        Return DirectCast(mp_oCollection.m_oItem(Index, SYS_ERRORS.COLUMNS_ITEM_1, SYS_ERRORS.COLUMNS_ITEM_2, SYS_ERRORS.COLUMNS_ITEM_3, SYS_ERRORS.COLUMNS_ITEM_4), clsColumn)
    End Function

    Friend ReadOnly Property oCollection() As clsCollectionBase
        Get
            Return mp_oCollection
        End Get
    End Property

    Public Function Add(ByVal Text As String, Optional ByVal Key As String = "", Optional ByVal Width As Integer = 125, Optional ByVal StyleIndex As String = "") As clsColumn
        mp_oCollection.AddMode = True
        Dim oColumn As New clsColumn(mp_oControl)
        Text = mp_oControl.StrLib.StrTrim(Text)
        oColumn.Text = Text
        oColumn.Width = Width
        oColumn.StyleIndex = StyleIndex
        oColumn.Key = Key
        Dim lIndex As Integer
        Dim oRow As clsRow
        For lIndex = 1 To mp_oControl.Rows.Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oRow.Cells.Add()
        Next lIndex
        mp_oCollection.m_Add(oColumn, oColumn.Key, SYS_ERRORS.COLUMNS_ADD_1, SYS_ERRORS.COLUMNS_ADD_2, False, SYS_ERRORS.COLUMNS_ADD_3)
        mp_oControl.Splitter.f_AdjustPosition()
        Return oColumn
    End Function

    
    Public Sub Clear()
        mp_oControl.Rows.ClearCells()
        mp_oCollection.m_Clear()
        mp_oControl.SelectedColumnIndex = 0
        mp_oControl.SelectedCellIndex = 0
        mp_oControl.Splitter.f_AdjustPosition()
    End Sub

    
    Public Sub Remove(ByVal Index As String)
        Dim lIndex As Integer
        Dim oRow As clsRow
        For lIndex = 1 To mp_oControl.Rows.Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oRow.Cells.Remove(Index)
        Next lIndex
        mp_oCollection.m_Remove(Index, SYS_ERRORS.COLUMNS_REMOVE_1, SYS_ERRORS.COLUMNS_REMOVE_2, SYS_ERRORS.COLUMNS_REMOVE_3, SYS_ERRORS.COLUMNS_REMOVE_4)
        mp_oControl.SelectedColumnIndex = 0
        mp_oControl.SelectedCellIndex = 0
        mp_oControl.Splitter.f_AdjustPosition()
    End Sub

    
    Public Sub MoveColumn(ByVal OriginColumnIndex As Integer, ByVal DestColumnIndex As Integer)
        Dim oColumn As clsColumn
        Dim oRow As clsRow
        Dim lIndex As Integer
        If OriginColumnIndex < 1 Or OriginColumnIndex > Count Then
            Return
        End If
        If DestColumnIndex < 1 Or DestColumnIndex > Count Then
            Return
        End If
        If DestColumnIndex = OriginColumnIndex Then
            Return
        End If
        If mp_oControl.TreeviewColumnIndex > 0 Then
            For lIndex = 1 To mp_oControl.Columns.Count
                oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(lIndex), clsColumn)
                If lIndex = mp_oControl.TreeviewColumnIndex Then
                    oColumn.mp_bTreeViewColumnIndex = True
                Else
                    oColumn.mp_bTreeViewColumnIndex = False
                End If
            Next
        End If
        mp_oCollection.m_lCopyAndMoveItems(OriginColumnIndex, DestColumnIndex)
        For lIndex = 1 To mp_oControl.Rows.Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oRow.Cells.oCollection.m_lCopyAndMoveItems(OriginColumnIndex, DestColumnIndex)
        Next lIndex
        If mp_oControl.TreeviewColumnIndex > 0 Then
            For lIndex = 1 To mp_oControl.Columns.Count
                oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(lIndex), clsColumn)
                If oColumn.mp_bTreeViewColumnIndex = True Then
                    mp_oControl.TreeviewColumnIndex = lIndex
                End If
            Next
        End If
    End Sub

    Friend ReadOnly Property Width() As Integer
        Get
            Dim lIndex As Integer
            Dim lResult As Integer
            For lIndex = 1 To Count
                Dim oColumn As clsColumn
                oColumn = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsColumn)
                lResult = lResult + oColumn.Width
            Next lIndex
            Return lResult
        End Get
    End Property

    Friend Sub Position()
        Dim lIndex As Integer
        Dim oColumn As clsColumn
        Dim lLeft As Integer
        lLeft = -mp_oControl.HorizontalScrollBar.Value + mp_oControl.mt_LeftMargin
        For lIndex = 1 To Count
            oColumn = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsColumn)
            oColumn.f_lLeft = lLeft
            oColumn.f_lRight = lLeft + oColumn.Width
            If oColumn.Right < mp_oControl.mt_LeftMargin Then
                oColumn.f_bVisible = False
            ElseIf oColumn.Left > mp_oControl.Splitter.Left Then
                oColumn.f_bVisible = False
            Else
                oColumn.f_bVisible = True
            End If
            If oColumn.Right > oColumn.Left Then
                oColumn.f_bVisible = True
            Else
                oColumn.f_bVisible = False
            End If
            lLeft = lLeft + oColumn.Width
        Next lIndex
    End Sub

    Friend Sub Draw()
        Dim oColumn As clsColumn
        Dim lIndex As Integer
        If Count = 0 Then
            Return
        End If
        If (mp_oControl.CurrentViewObject.TimeLine.Height = 0) Then
            Return
        End If
        For lIndex = 1 To Count
            oColumn = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsColumn)
            If oColumn.Visible = True Then
                mp_oControl.clsG.ClipRegion(oColumn.LeftTrim, oColumn.Top, oColumn.RightTrim, oColumn.Bottom, True)
                mp_oControl.DrawEventArgs.Clear()
                mp_oControl.DrawEventArgs.CustomDraw = False
                mp_oControl.DrawEventArgs.EventTarget = E_EVENTTARGET.EVT_COLUMN
                mp_oControl.DrawEventArgs.ObjectIndex = lIndex
                mp_oControl.DrawEventArgs.ParentObjectIndex = 0
                'mp_oControl.DrawEventArgs.Graphics = mp_oControl.clsG.oGraphics
                mp_oControl.FireDraw()
                If mp_oControl.DrawEventArgs.CustomDraw = False Then
                    mp_oControl.clsG.mp_DrawItem(oColumn.Left, oColumn.Top, oColumn.Right - 1, oColumn.Bottom, "", oColumn.Text, (lIndex = mp_oControl.SelectedColumnIndex), oColumn.Image, oColumn.LeftTrim, oColumn.RightTrim, oColumn.Style)
                    If oColumn.Text.Length > 0 Then
                        oColumn.mp_lTextLeft = mp_oControl.clsG.mp_oTextFinalLayout.Left
                        oColumn.mp_lTextTop = mp_oControl.clsG.mp_oTextFinalLayout.Top
                        oColumn.mp_lTextRight = mp_oControl.clsG.mp_oTextFinalLayout.Left + mp_oControl.clsG.mp_oTextFinalLayout.Width - 1
                        oColumn.mp_lTextBottom = mp_oControl.clsG.mp_oTextFinalLayout.Top + mp_oControl.clsG.mp_oTextFinalLayout.Height - 1
                    Else
                        oColumn.mp_lTextLeft = oColumn.Left
                        oColumn.mp_lTextTop = oColumn.Top
                        oColumn.mp_lTextRight = oColumn.Right
                        oColumn.mp_lTextBottom = oColumn.Bottom
                    End If
                End If
            End If
        Next lIndex
    End Sub

    
    Public Function GetXML() As String
        Dim lIndex As Integer
        Dim oColumn As clsColumn
        Dim oXML As New clsXML(mp_oControl, "Columns")
        oXML.InitializeWriter()
        For lIndex = 1 To Count
            oColumn = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsColumn)
            oXML.WriteObject(oColumn.GetXML)
        Next lIndex
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim lIndex As Integer
        Dim oXML As New clsXML(mp_oControl, "Columns")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        mp_oCollection.m_Clear()
        For lIndex = 1 To oXML.ReadCollectionCount
            Dim oColumn As New clsColumn(mp_oControl)
            oColumn.SetXML(oXML.ReadCollectionObject(lIndex))
            mp_oCollection.AddMode = True
            mp_oCollection.m_Add(oColumn, oColumn.Key, SYS_ERRORS.COLUMNS_ADD_1, SYS_ERRORS.COLUMNS_ADD_2, False, SYS_ERRORS.COLUMNS_ADD_3)
            oColumn = Nothing
        Next lIndex
    End Sub

End Class
