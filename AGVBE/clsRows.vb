Option Explicit On

Public Class clsRows

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_oCollection As clsCollectionBase
    Private mp_lTopOffset As Integer
    Private mp_lRealFirstVisibleRow As Integer

    Private mp_lLoadIndex As Integer
    Friend mp_oTempCollection As ArrayList
    Friend mp_oTempDictionary As clsDictionary
    Private mp_oTempNodeList As ArrayList

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_oCollection = New clsCollectionBase(Value, "Row")
    End Sub

    
    Public ReadOnly Property Count() As Integer
        Get
            Return mp_oCollection.m_lCount()
        End Get
    End Property

    
    Public Function Item(ByVal Index As String) As clsRow
        Return DirectCast(mp_oCollection.m_oItem(Index, SYS_ERRORS.ROWS_ITEM_1, SYS_ERRORS.ROWS_ITEM_2, SYS_ERRORS.ROWS_ITEM_3, SYS_ERRORS.ROWS_ITEM_4), clsRow)
    End Function

    Friend ReadOnly Property oCollection() As clsCollectionBase
        Get
            Return mp_oCollection
        End Get
    End Property

    
    Public Function Add(ByVal Key As String) As clsRow
        Return Add(Key, "", False, True, "")
    End Function

    Public Function Add(ByVal Key As String, ByVal Text As String) As clsRow
        Return Add(Key, Text, False, True, "")
    End Function

    Public Function Add(ByVal Key As String, ByVal Text As String, ByVal MergeCells As Boolean) As clsRow
        Return Add(Key, Text, MergeCells, True, "")
    End Function

    Public Function Add(ByVal Key As String, ByVal Text As String, ByVal MergeCells As Boolean, ByVal Container As Boolean) As clsRow
        Return Add(Key, Text, MergeCells, Container, "")
    End Function

    Public Function Add(ByVal Key As String, ByVal Text As String, ByVal MergeCells As Boolean, ByVal Container As Boolean, ByVal StyleIndex As String) As clsRow
        mp_oCollection.AddMode = True
        Dim oRow As New clsRow(mp_oControl)
        Dim lIndex As Integer
        oRow.Key = Key
        oRow.Text = Text
        oRow.MergeCells = MergeCells
        oRow.Container = Container
        oRow.StyleIndex = StyleIndex
        mp_oCollection.m_Add(oRow, Key, SYS_ERRORS.ROWS_ADD_1, SYS_ERRORS.ROWS_ADD_2, True, SYS_ERRORS.ROWS_ADD_3)
        For lIndex = 1 To mp_oControl.Columns.Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_oControl.Rows.Count), clsRow)
            oRow.Cells.Add()
        Next lIndex
        mp_oControl.VerticalScrollBar.Update()
        Return oRow
    End Function

    
    Public Sub Clear()
        mp_oControl.Tasks.Clear()
        mp_oCollection.m_Clear()
        mp_oControl.SelectedRowIndex = 0
        mp_oControl.VerticalScrollBar.Reset()
    End Sub

    
    Public Sub Remove(ByVal Index As String)
        Dim sRIndex As String = ""
        Dim sRKey As String = ""
        mp_oCollection.m_GetKeyAndIndex(Index, sRKey, sRIndex)
        mp_oControl.Tasks.oCollection.m_CollectionRemoveWhere("RowKey", sRKey, sRIndex)
        mp_oCollection.m_Remove(Index, SYS_ERRORS.ROWS_REMOVE_1, SYS_ERRORS.ROWS_REMOVE_2, SYS_ERRORS.ROWS_REMOVE_3, SYS_ERRORS.ROWS_REMOVE_4)
        mp_oControl.SelectedRowIndex = 0
        mp_oControl.SelectedTaskIndex = 0
        mp_oControl.VerticalScrollBar.Update()
    End Sub

    
    Public Sub MoveRow(ByVal OriginRowIndex As Integer, ByVal DestRowIndex As Integer)
        If OriginRowIndex < 1 Or OriginRowIndex > Count Then
            Return
        End If
        If DestRowIndex < 1 Or DestRowIndex > Count Then
            Return
        End If
        If DestRowIndex = OriginRowIndex Then
            Return
        End If
        mp_oCollection.m_lCopyAndMoveItems(OriginRowIndex, DestRowIndex)
    End Sub

    
    Public Sub SortRows(ByVal PropertyName As String, ByVal Descending As Boolean, ByVal SortType As E_SORTTYPE)
        SortRows(PropertyName, Descending, SortType, -1, -1)
    End Sub

    Public Sub SortRows(ByVal PropertyName As String, ByVal Descending As Boolean, ByVal SortType As E_SORTTYPE, ByVal StartIndex As Integer, ByVal EndIndex As Integer)
        If StartIndex = -1 Then
            StartIndex = 1
        End If
        If EndIndex = -1 Then
            EndIndex = Count
        End If
        If Count = 0 Then
            Return
        End If
        If StartIndex < 1 Or StartIndex > Count Then
            Return
        End If
        If EndIndex < 1 Or EndIndex > Count Then
            Return
        End If
        If EndIndex = StartIndex Then
            Return
        End If
        mp_oCollection.m_Sort(PropertyName, Descending, SortType, StartIndex, EndIndex)
    End Sub

    
    Public Sub SortCells(ByVal CellIndex As Integer, ByVal PropertyName As String, ByVal Descending As Boolean, ByVal SortType As E_SORTTYPE)
        SortCells(CellIndex, PropertyName, Descending, SortType, -1, -1)
    End Sub

    Public Sub SortCells(ByVal CellIndex As Integer, ByVal PropertyName As String, ByVal Descending As Boolean, ByVal SortType As E_SORTTYPE, ByVal StartIndex As Integer, ByVal EndIndex As Integer)
        If StartIndex = -1 Then
            StartIndex = 1
        End If
        If EndIndex = -1 Then
            EndIndex = Count
        End If
        If Count = 0 Then
            Return
        End If
        If StartIndex < 1 Or StartIndex > Count Then
            Return
        End If
        If EndIndex < 1 Or EndIndex > Count Then
            Return
        End If
        If EndIndex = StartIndex Then
            Return
        End If
        If CellIndex < 1 Or CellIndex > mp_oControl.Columns.Count Then
            Return
        End If
        mp_oCollection.m_SortCells(CellIndex, PropertyName, Descending, SortType, StartIndex, EndIndex)
    End Sub

    Friend Sub ClearCells()
        Dim lIndex As Integer
        Dim oRow As clsRow = Nothing
        For lIndex = 1 To mp_oCollection.m_lCount()
            oRow = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oRow.Cells.Clear()
        Next lIndex
    End Sub

    Friend Function Height() As Integer
        Dim lBuffer As Integer = 0
        Dim lIndex As Integer
        Dim oRow As clsRow = Nothing
        If Count = 0 Then
            Return 0
        End If

        Dim bChildrenHidden As Boolean = False
        Dim lCurrentDepth As Integer = 0
        For lIndex = 1 To Count()
            Dim bHidden As Boolean = False
            oRow = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsRow)
            If oRow.Node.Depth = 0 Then
                bHidden = False
            End If
            If bChildrenHidden = True Then
                bHidden = True
            End If
            If oRow.Node.Depth < lCurrentDepth Then
                lCurrentDepth = oRow.Node.Depth
                bHidden = False
                bChildrenHidden = False
            End If
            If bHidden = False Then
                lBuffer = lBuffer + oRow.Height + 1
            End If
            If oRow.Node.Expanded = False And bChildrenHidden = False Then
                lCurrentDepth = oRow.Node.Depth + 1
                bChildrenHidden = True
            End If
        Next

        Return lBuffer
    End Function

    Friend Function CalculateHeight(ByVal StartIndex As Integer, ByVal EndIndex As Integer) As Integer
        Dim lBuffer As Integer
        Dim lIndex As Integer
        Dim oRow As clsRow = Nothing
        If StartIndex = 0 Then
            Return 0
        End If
        For lIndex = StartIndex To EndIndex
            oRow = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsRow)
            lBuffer = lBuffer + oRow.Height
        Next lIndex
        Return lBuffer
    End Function

    Friend Function CalculateRows(ByVal StartIndex As Integer, ByVal Height As Integer) As Integer
        Dim lBuffer As Integer
        Dim lIndex As Integer
        Dim oRow As clsRow = Nothing
        Dim lRows As Integer
        lRows = 1
        If StartIndex = 0 Then
            Return lRows
        End If
        For lIndex = StartIndex To Count
            oRow = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsRow)
            lBuffer = lBuffer + oRow.Height
            If lBuffer > Height Then
                Exit For
            End If
            lRows = lRows + 1
        Next lIndex
        Return lRows
    End Function

    Friend Sub PositionRows()
        Dim oRow As clsRow = Nothing
        Dim lRowIndex As Integer
        Dim lBottomBuff As Integer
        Dim oClientArea As clsClientArea = Nothing
        oClientArea = mp_oControl.CurrentViewObject.ClientArea
        If Count = 0 Then
            oClientArea.f_LastVisibleRow = 0
            mp_lTopOffset = mp_oControl.CurrentViewObject.ClientArea.Top
            Return
        Else
            mp_lTopOffset = 0
        End If
        For lRowIndex = (mp_lRealFirstVisibleRow - 1) To 1 Step -1
            oRow = DirectCast(mp_oCollection.m_oReturnArrayElement(lRowIndex), clsRow)
            If lRowIndex = (mp_lRealFirstVisibleRow - 1) Then
                oRow.f_lBottom = mp_oControl.CurrentViewObject.ClientArea.Top - 1
            Else
                oRow.f_lBottom = lBottomBuff - 1
            End If
            oRow.f_lTop = oRow.Bottom - oRow.Height
            lBottomBuff = oRow.Top
            oRow.ClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_ABOVEVISIBLEAREA
        Next lRowIndex
        For lRowIndex = mp_lRealFirstVisibleRow To Count
            oRow = DirectCast(mp_oCollection.m_oReturnArrayElement(lRowIndex), clsRow)
            If mp_lTopOffset < mp_oControl.mt_TableBottom Then
                If lRowIndex = mp_lRealFirstVisibleRow Then
                    oRow.f_lTop = mp_oControl.CurrentViewObject.ClientArea.Top
                Else
                    oRow.f_lTop = lBottomBuff + 1
                End If
                oRow.f_lBottom = oRow.Top + oRow.Height
                lBottomBuff = oRow.Bottom
                mp_lTopOffset = oRow.Bottom
                oClientArea.f_LastVisibleRow = lRowIndex
                oRow.ClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_INSIDEVISIBLEAREA
            Else
                Exit For
            End If
        Next lRowIndex
        For lRowIndex = (oClientArea.LastVisibleRow + 1) To Count
            oRow = DirectCast(mp_oCollection.m_oReturnArrayElement(lRowIndex), clsRow)
            oRow.f_lTop = lBottomBuff + 1
            oRow.f_lBottom = oRow.Top + oRow.Height
            lBottomBuff = oRow.Bottom
            oRow.ClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_BELOWVISIBLEAREA
        Next lRowIndex
    End Sub

    Friend Property TopOffset() As Integer
        Get
            Return mp_lTopOffset
        End Get
        Set(ByVal Value As Integer)
            mp_lTopOffset = Value
        End Set
    End Property

    Friend Sub InitializePosition()
        mp_lRealFirstVisibleRow = RealFirstVisibleRow
    End Sub

    Friend Sub NodesDrawBackground()
        Dim lIndex As Integer
        Dim oNode As clsNode = Nothing
        Dim oRow As clsRow = Nothing
        Dim oCell As clsCell = Nothing
        If Count = 0 Then
            Return
        End If
        For lIndex = mp_lRealFirstVisibleRow To Count
            oRow = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oCell = DirectCast(oRow.Cells.oCollection.m_oReturnArrayElement(mp_oControl.TreeviewColumnIndex), clsCell)
            oNode = oRow.Node
            If mp_oControl.Treeview.FullColumnSelect = True Then
                mp_oControl.clsG.mp_DrawItem(oCell.Left, oCell.Top, oCell.Right - 1, oCell.Bottom, "", "", (oRow.Index = mp_oControl.SelectedRowIndex And mp_oControl.TreeviewColumnIndex = mp_oControl.SelectedCellIndex), oCell.Image, oCell.LeftTrim, oCell.RightTrim, oNode.Style)
            Else
                mp_oControl.clsG.mp_DrawItem(oCell.Left, oCell.Top, oCell.Right - 1, oCell.Bottom, "", "", False, oCell.Image, oCell.LeftTrim, oCell.RightTrim, oNode.Style)
                If (oRow.Index = mp_oControl.SelectedRowIndex And mp_oControl.TreeviewColumnIndex = mp_oControl.SelectedCellIndex And oNode.Style.SelectionRectangleStyle.Visible = True) Then
                    mp_oControl.clsG.mp_DrawSelectionRectangle(oNode.mt_TextLeft, oNode.Top, mp_oControl.Splitter.Left, oNode.Bottom - 1, oNode.Style)
                End If
            End If
        Next
    End Sub

    Friend Sub NodesDrawTreeLines()
        Dim lIndex As Integer
        Dim oNode As clsNode = Nothing
        Dim oRow As clsRow = Nothing
        Dim oRelative As clsNode = Nothing
        If Count = 0 Or mp_oControl.Treeview.TreeLines = False Then
            Return
        End If
        For lIndex = mp_lRealFirstVisibleRow To Count
            oRow = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oNode = oRow.Node
            If (oRow.ClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_INSIDEVISIBLEAREA Or oRow.ClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_BELOWVISIBLEAREA) And oNode.Hidden = False Then
                If lIndex <= mp_oControl.CurrentViewObject.ClientArea.LastVisibleRow Then
                    If oNode.CheckBoxVisible = True Then
                        mp_oControl.clsG.DrawLine(oNode.Left, oNode.YCenter, oNode.Left + 15, oNode.YCenter, GRE_LINETYPE.LT_NORMAL, mp_oControl.Treeview.TreeLineColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    ElseIf oNode.ImageVisible = True Then
                        mp_oControl.clsG.DrawLine(oNode.Left, oNode.YCenter, oNode.Left + 15, oNode.YCenter, GRE_LINETYPE.LT_NORMAL, mp_oControl.Treeview.TreeLineColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    ElseIf oNode.ImageVisible = False And oNode.CheckBoxVisible = False Then
                        mp_oControl.clsG.DrawLine(oNode.Left, oNode.YCenter, oNode.mt_TextLeft, oNode.YCenter, GRE_LINETYPE.LT_NORMAL, mp_oControl.Treeview.TreeLineColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    End If
                End If
                If oNode.Index = 1 Then
                    'Dont Draw anything
                Else
                    oRelative = Nothing
                    If oNode.IsFirstSibling() Then
                        oRelative = oNode.Parent()
                        If oRelative.Row.Index > mp_oControl.CurrentViewObject.ClientArea.LastVisibleRow Then
                            oRelative = Nothing
                        Else
                            mp_oControl.clsG.DrawLine(oNode.Left, oNode.YCenter, oNode.Left, oRelative.mt_TextBottom, GRE_LINETYPE.LT_NORMAL, mp_oControl.Treeview.TreeLineColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                        End If
                    ElseIf oNode.IsLastSibling() Then
                        oRelative = oNode.FirstSibling
                        If oRelative.Row.Index > mp_oControl.CurrentViewObject.ClientArea.LastVisibleRow Then
                            oRelative = Nothing
                        Else
                            mp_oControl.clsG.DrawLine(oNode.Left, oNode.YCenter, oNode.Left, oRelative.YCenter, GRE_LINETYPE.LT_NORMAL, mp_oControl.Treeview.TreeLineColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                        End If
                    End If
                End If
            End If
        Next lIndex
    End Sub

    Friend Sub NodesDraw()
        Dim lIndex As Integer
        Dim oNode As clsNode = Nothing
        Dim oRow As clsRow = Nothing
        If Count = 0 Then
            Return
        End If
        For lIndex = mp_lRealFirstVisibleRow To Count
            oRow = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oNode = oRow.Node
            If oRow.ClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_INSIDEVISIBLEAREA And oNode.Hidden = False Then
                mp_oControl.clsG.DrawAlignedText(oNode.mt_TextLeft, oNode.mt_TextTop, oNode.mt_TextRight, oNode.mt_TextBottom, oNode.Text, GRE_HORIZONTALALIGNMENT.HAL_LEFT, GRE_VERTICALALIGNMENT.VAL_CENTER, oNode.Style.ForeColor, oNode.Style.Font)
                If oNode.Text.Length > 0 Then
                    oNode.mp_lTextLeft = mp_oControl.clsG.mp_oTextFinalLayout.Left
                    oNode.mp_lTextTop = mp_oControl.clsG.mp_oTextFinalLayout.Top
                    oNode.mp_lTextRight = mp_oControl.clsG.mp_oTextFinalLayout.Left + mp_oControl.clsG.mp_oTextFinalLayout.Width - 1
                    oNode.mp_lTextBottom = mp_oControl.clsG.mp_oTextFinalLayout.Top + mp_oControl.clsG.mp_oTextFinalLayout.Height - 1
                Else
                    oNode.mp_lTextLeft = oNode.mt_TextLeft
                    oNode.mp_lTextTop = oNode.mt_TextTop
                    oNode.mp_lTextRight = oNode.mt_TextRight
                    oNode.mp_lTextBottom = oNode.mt_TextBottom
                End If
            End If
        Next lIndex
    End Sub

    Private Sub mp_DrawSign(ByVal bExpanded As Boolean, ByVal X As Integer, ByVal Y As Integer, ByVal BorderColor As Color, ByVal SignColor As Color)
        If mp_oControl.Treeview.PlusMinusSigns = False Then
            Return
        End If
        Y = Y - 1
        mp_oControl.clsG.DrawLine(X - 4, Y - 3, X + 4, Y + 4, GRE_LINETYPE.LT_FILLED, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
        mp_oControl.clsG.DrawLine(X - 4, Y - 3, X + 4, Y + 5, GRE_LINETYPE.LT_BORDER, BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
        mp_oControl.clsG.DrawLine(X - 2, Y + 1, X + 2, Y + 1, GRE_LINETYPE.LT_NORMAL, SignColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
        If bExpanded = False Then
            mp_oControl.clsG.DrawLine(X, Y - 1, X, Y + 3, GRE_LINETYPE.LT_NORMAL, SignColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
        End If
    End Sub

    Friend Function HiddenRows() As Integer
        Dim lIndex As Integer
        Dim oRow As clsRow = Nothing
        Dim lReturn As Integer = 0

        Dim bChildrenHidden As Boolean = False
        Dim lCurrentDepth As Integer = 0
        For lIndex = 1 To Count()
            Dim bHidden As Boolean = False
            oRow = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsRow)
            If oRow.Node.Depth = 0 Then
                bHidden = False
            End If
            If bChildrenHidden = True Then
                bHidden = True
            End If
            If oRow.Node.Depth < lCurrentDepth Then
                lCurrentDepth = oRow.Node.Depth
                bHidden = False
                bChildrenHidden = False
            End If
            If bHidden = True Then
                lReturn = lReturn + 1
            End If
            If oRow.Node.Expanded = False And bChildrenHidden = False Then
                lCurrentDepth = oRow.Node.Depth + 1
                bChildrenHidden = True
            End If
        Next

        Return lReturn
    End Function

    Friend ReadOnly Property RealFirstVisibleRow() As Integer
        Get
            Return RealIndex(mp_oControl.VerticalScrollBar.Value)
        End Get
    End Property

    Friend Function RealIndex(ByVal Index As Integer) As Integer
        Dim lIndex As Integer
        Dim oRow As clsRow = Nothing
        Dim oNode As clsNode = Nothing
        Dim lCount As Integer = 0
        For lIndex = 1 To Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oNode = oRow.Node
            If oNode.Hidden = False Then
                lCount = lCount + 1
                If lCount = Index Then
                    Return lIndex
                End If
            End If
        Next
        Return -1
    End Function

    Private Sub mp_DrawCheckBox(ByRef oNode As clsNode)
        If mp_oControl.Treeview.CheckBoxes = False Then
            Return
        End If
        If oNode.CheckBoxVisible = True Then
            mp_oControl.clsG.DrawLine(oNode.CheckBoxLeft + 1, oNode.YCenter - 5, oNode.CheckBoxLeft + 11, oNode.YCenter + 5, GRE_LINETYPE.LT_FILLED, mp_oControl.Treeview.CheckBoxColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            mp_oControl.clsG.DrawLine(oNode.CheckBoxLeft + 1, oNode.YCenter - 5, oNode.CheckBoxLeft + 11, oNode.YCenter + 5, GRE_LINETYPE.LT_BORDER, mp_oControl.Treeview.CheckBoxBorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            If oNode.Checked = True Then
                mp_oControl.clsG.DrawLine(oNode.CheckBoxLeft + 3, oNode.YCenter, oNode.CheckBoxLeft + 3, oNode.YCenter + 2, GRE_LINETYPE.LT_NORMAL, mp_oControl.Treeview.CheckBoxMarkColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                mp_oControl.clsG.DrawLine(oNode.CheckBoxLeft + 4, oNode.YCenter + 1, oNode.CheckBoxLeft + 4, oNode.YCenter + 3, GRE_LINETYPE.LT_NORMAL, mp_oControl.Treeview.CheckBoxMarkColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                mp_oControl.clsG.DrawLine(oNode.CheckBoxLeft + 5, oNode.YCenter + 2, oNode.CheckBoxLeft + 5, oNode.YCenter + 4, GRE_LINETYPE.LT_NORMAL, mp_oControl.Treeview.CheckBoxMarkColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                mp_oControl.clsG.DrawLine(oNode.CheckBoxLeft + 6, oNode.YCenter + 1, oNode.CheckBoxLeft + 6, oNode.YCenter + 3, GRE_LINETYPE.LT_NORMAL, mp_oControl.Treeview.CheckBoxMarkColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                mp_oControl.clsG.DrawLine(oNode.CheckBoxLeft + 7, oNode.YCenter, oNode.CheckBoxLeft + 7, oNode.YCenter + 2, GRE_LINETYPE.LT_NORMAL, mp_oControl.Treeview.CheckBoxMarkColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                mp_oControl.clsG.DrawLine(oNode.CheckBoxLeft + 8, oNode.YCenter - 1, oNode.CheckBoxLeft + 8, oNode.YCenter + 1, GRE_LINETYPE.LT_NORMAL, mp_oControl.Treeview.CheckBoxMarkColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                mp_oControl.clsG.DrawLine(oNode.CheckBoxLeft + 9, oNode.YCenter - 2, oNode.CheckBoxLeft + 9, oNode.YCenter, GRE_LINETYPE.LT_NORMAL, mp_oControl.Treeview.CheckBoxMarkColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            End If
        End If
    End Sub

    Private Sub mp_DrawImage(ByRef oNode As clsNode)
        Dim oImage As Image = Nothing
        Dim lImageWidth As Integer
        Dim lImageHeight As Integer
        If mp_oControl.Treeview.Images = False Then
            Return
        End If
        If oNode.ImageVisible = True Then
            If oNode.Expanded = True And oNode.Children() > 0 And Not (oNode.ExpandedImage Is Nothing) Then
                oImage = oNode.ExpandedImage
            ElseIf oNode.Selected = True And Not (oNode.SelectedImage Is Nothing) Then
                oImage = oNode.SelectedImage
            ElseIf Not (oNode.Image Is Nothing) Then
                oImage = oNode.Image
            End If
            If Not oImage Is Nothing Then
                lImageWidth = System.Convert.ToInt32(oImage.Width)
                lImageHeight = System.Convert.ToInt32(oImage.Height)
                mp_oControl.clsG.PaintImage(oImage, oNode.ImageLeft, oNode.ImageTop, oNode.ImageRight, oNode.ImageBottom, 0, 0, True)
            End If
        End If
    End Sub

    Friend Sub Draw()
        Dim lCellIndex As Integer
        Dim lRowIndex As Integer
        Dim oRow As clsRow = Nothing
        Dim oColumn As clsColumn = Nothing
        Dim oCell As clsCell = Nothing
        Dim lTableBottom As Integer = mp_oControl.mt_TableBottom
        Dim lBottom As Integer
        mp_oControl.clsG.ClipRegion(mp_oControl.mt_LeftMargin, mp_oControl.CurrentViewObject.ClientArea.Top, mp_oControl.Splitter.Left, mp_oControl.mt_TableBottom, False)
        mp_oControl.DrawEventArgs.Clear()
        mp_oControl.DrawEventArgs.CustomDraw = False
        mp_oControl.DrawEventArgs.EventTarget = E_EVENTTARGET.EVT_GRID
        mp_oControl.DrawEventArgs.ObjectIndex = 0
        mp_oControl.DrawEventArgs.ParentObjectIndex = 0
        'mp_oControl.DrawEventArgs.Graphics = mp_oControl.clsG.oGraphics
        mp_oControl.FireDraw()
        If mp_oControl.DrawEventArgs.CustomDraw = True Then
            Return
        End If
        If Count = 0 Then
            Return
        End If
        For lRowIndex = mp_lRealFirstVisibleRow To mp_oControl.CurrentViewObject.ClientArea.LastVisibleRow
            oRow = DirectCast(mp_oCollection.m_oReturnArrayElement(lRowIndex), clsRow)
            If oRow.Visible = True And oRow.Height > -1 Then
                If oRow.MergeCells = True Then
                    If oRow.Bottom > lTableBottom Then
                        lBottom = lTableBottom
                    Else
                        lBottom = oRow.Bottom
                    End If
                    mp_oControl.clsG.ClipRegion(oRow.Left, oRow.Top, oRow.Right, lBottom, True)
                    mp_oControl.DrawEventArgs.Clear()
                    mp_oControl.DrawEventArgs.CustomDraw = False
                    mp_oControl.DrawEventArgs.EventTarget = E_EVENTTARGET.EVT_ROW
                    mp_oControl.DrawEventArgs.ObjectIndex = lRowIndex
                    mp_oControl.DrawEventArgs.ParentObjectIndex = 0
                    'mp_oControl.DrawEventArgs.Graphics = mp_oControl.clsG.oGraphics
                    mp_oControl.FireDraw()
                    If mp_oControl.DrawEventArgs.CustomDraw = False Then
                        mp_oControl.clsG.mp_DrawItem(oRow.Left, oRow.Top, oRow.Right, oRow.Bottom, "", oRow.Text, (lRowIndex = mp_oControl.SelectedRowIndex), oRow.Image, 0, 0, oRow.Style)
                        If oRow.Text.Length > 0 Then
                            oRow.mp_lTextLeft = mp_oControl.clsG.mp_oTextFinalLayout.Left
                            oRow.mp_lTextTop = mp_oControl.clsG.mp_oTextFinalLayout.Top
                            oRow.mp_lTextRight = mp_oControl.clsG.mp_oTextFinalLayout.Left + mp_oControl.clsG.mp_oTextFinalLayout.Width - 1
                            oRow.mp_lTextBottom = mp_oControl.clsG.mp_oTextFinalLayout.Top + mp_oControl.clsG.mp_oTextFinalLayout.Height - 1
                        Else
                            oRow.mp_lTextLeft = oRow.Left
                            oRow.mp_lTextTop = oRow.Top
                            oRow.mp_lTextRight = oRow.Right
                            oRow.mp_lTextBottom = oRow.Bottom
                        End If
                    End If
                Else
                    For lCellIndex = 1 To mp_oControl.Columns.Count
                        If lCellIndex <> mp_oControl.TreeviewColumnIndex Then
                            oCell = DirectCast(oRow.Cells.oCollection.m_oReturnArrayElement(lCellIndex), clsCell)
                            oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(lCellIndex), clsColumn)
                            If oColumn.Visible = True Then
                                If oCell.Bottom > lTableBottom Then
                                    lBottom = lTableBottom
                                Else
                                    lBottom = oCell.Bottom
                                End If
                                mp_oControl.clsG.ClipRegion(oCell.LeftTrim, oCell.Top, oCell.RightTrim, lBottom, True)
                                mp_oControl.DrawEventArgs.Clear()
                                mp_oControl.DrawEventArgs.CustomDraw = False
                                mp_oControl.DrawEventArgs.EventTarget = E_EVENTTARGET.EVT_CELL
                                mp_oControl.DrawEventArgs.ObjectIndex = lCellIndex
                                mp_oControl.DrawEventArgs.ParentObjectIndex = lRowIndex
                                'mp_oControl.DrawEventArgs.Graphics = mp_oControl.clsG.oGraphics
                                mp_oControl.FireDraw()
                                If mp_oControl.DrawEventArgs.CustomDraw = False Then
                                    mp_oControl.clsG.mp_DrawItem(oCell.Left, oCell.Top, oCell.Right - 1, oCell.Bottom, "", oCell.Text, (lRowIndex = mp_oControl.SelectedRowIndex And lCellIndex = mp_oControl.SelectedCellIndex), oCell.Image, oCell.LeftTrim, oCell.RightTrim, oCell.Style)
                                    If oCell.Text.Length > 0 Then
                                        oCell.mp_lTextLeft = mp_oControl.clsG.mp_oTextFinalLayout.Left
                                        oCell.mp_lTextTop = mp_oControl.clsG.mp_oTextFinalLayout.Top
                                        oCell.mp_lTextRight = mp_oControl.clsG.mp_oTextFinalLayout.Left + mp_oControl.clsG.mp_oTextFinalLayout.Width - 1
                                        oCell.mp_lTextBottom = mp_oControl.clsG.mp_oTextFinalLayout.Top + mp_oControl.clsG.mp_oTextFinalLayout.Height - 1
                                    Else
                                        oCell.mp_lTextLeft = oCell.Left
                                        oCell.mp_lTextTop = oCell.Top
                                        oCell.mp_lTextRight = oCell.Right
                                        oCell.mp_lTextBottom = oCell.Bottom
                                    End If
                                End If
                            End If
                        End If
                    Next lCellIndex
                End If
            End If
        Next lRowIndex
    End Sub

    Public Function GetXML() As String
        Dim lIndex As Integer
        Dim oRow As clsRow = Nothing
        Dim oXML As New clsXML(mp_oControl, "Rows")
        oXML.InitializeWriter()
        For lIndex = 1 To Count
            oRow = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oXML.WriteObject(oRow.GetXML())
        Next lIndex
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim lIndex As Integer
        Dim oXML As New clsXML(mp_oControl, "Rows")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        mp_oCollection.m_Clear()
        For lIndex = 1 To oXML.ReadCollectionCount()
            Dim oRow As New clsRow(mp_oControl)
            oRow.SetXML(oXML.ReadCollectionObject(lIndex))
            mp_oCollection.AddMode = True
            mp_oCollection.m_Add(oRow, oRow.Key, SYS_ERRORS.ROWS_ADD_1, SYS_ERRORS.ROWS_ADD_2, True, SYS_ERRORS.ROWS_ADD_3)
            oRow = Nothing
        Next lIndex
        mp_oControl.VerticalScrollBar.Update()
        mp_oControl.VerticalScrollBar.Value = 1
    End Sub

    Public Sub BeginLoad(ByVal Preserve As Boolean)
        mp_oTempNodeList = New ArrayList
        If Preserve = False Then
            mp_lLoadIndex = 1
            mp_oTempCollection = New ArrayList
            mp_oTempDictionary = New clsDictionary
            mp_oCollection.mp_aoCollection.Clear()
            mp_oCollection.mp_oKeys.Clear()
        Else
            mp_oTempCollection = mp_oCollection.mp_aoCollection
            mp_oCollection.mp_aoCollection = Nothing
            mp_oCollection.mp_aoCollection = New ArrayList
            mp_oTempDictionary = mp_oCollection.mp_oKeys
            mp_oCollection.mp_oKeys = Nothing
            mp_oCollection.mp_oKeys = New clsDictionary
            mp_lLoadIndex = mp_oTempCollection.Count + 1
        End If
    End Sub

    Public Function Load(ByVal sKey As String) As clsRow
        Dim lIndex As Integer
        Dim oRow As New clsRow(mp_oControl)
        oRow.Key = sKey
        For lIndex = 1 To mp_oControl.Columns.Count
            oRow.Cells.Add()
        Next
        oRow.Index = mp_lLoadIndex
        If oRow.Node.Depth = 0 Then
            oRow.Node.mp_oParent = Nothing
        Else
            oRow.Node.mp_oParent = DirectCast(mp_oTempNodeList(oRow.Node.Depth - 1), clsNode)
        End If
        If oRow.Node.Depth > (mp_oTempNodeList.Count - 1) Then
            mp_oTempNodeList.Add(oRow.Node)
        Else
            mp_oTempNodeList.Item(oRow.Node.Depth) = oRow.Node
        End If
        mp_oTempCollection.Add(oRow)
        mp_oTempDictionary.Add(mp_lLoadIndex, sKey)
        mp_lLoadIndex = mp_lLoadIndex + 1
        Return oRow
    End Function

    Public Sub EndLoad()
        mp_oCollection.mp_aoCollection = mp_oTempCollection
        mp_oCollection.mp_oKeys = mp_oTempDictionary
        mp_oControl.VerticalScrollBar.Update()
        mp_oTempCollection = Nothing
        mp_oTempDictionary = Nothing
        mp_oTempNodeList = Nothing
        UpdateTree()
    End Sub

    Public Sub UpdateTree()
        Dim lIndex As Integer
        Dim oRow As clsRow
        mp_oTempNodeList = New ArrayList
        For lIndex = 1 To Count
            oRow = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsRow)
            If oRow.Node.Depth = 0 Then
                oRow.Node.mp_oParent = Nothing
            Else
                oRow.Node.mp_oParent = DirectCast(mp_oTempNodeList(oRow.Node.Depth - 1), clsNode)
            End If
            If oRow.Node.Depth > (mp_oTempNodeList.Count - 1) Then
                mp_oTempNodeList.Add(oRow.Node)
            Else
                mp_oTempNodeList.Item(oRow.Node.Depth) = oRow.Node
            End If
        Next
    End Sub

    Friend Sub NodesDrawElements()
        Dim lIndex As Integer
        Dim oNode As clsNode = Nothing
        Dim oRow As clsRow = Nothing
        If Count = 0 Then
            Return
        End If
        For lIndex = mp_lRealFirstVisibleRow To Count
            oRow = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oNode = oRow.Node
            If oRow.ClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_INSIDEVISIBLEAREA And oNode.Hidden = False Then
                If oNode.Children() > 0 Then
                    mp_DrawSign(oNode.Expanded, oNode.Left, oNode.YCenter, mp_oControl.Treeview.PlusMinusBorderColor, mp_oControl.Treeview.PlusMinusSignColor)
                End If
                mp_DrawCheckBox(oNode)
                mp_DrawImage(oNode)
            End If
        Next lIndex
    End Sub

End Class
