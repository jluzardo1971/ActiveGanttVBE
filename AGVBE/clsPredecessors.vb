Option Explicit On 

Public Class clsPredecessors

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_oCollection As clsCollectionBase

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_oCollection = New clsCollectionBase(Value, "Predecessor")
    End Sub

    
    Public ReadOnly Property Count() As Integer
        Get
            Return mp_oCollection.m_lCount
        End Get
    End Property

    
    Public Function Item(ByVal Index As String) As clsPredecessor
        Return DirectCast(mp_oCollection.m_oItem(Index, SYS_ERRORS.PREDECESSORS_ITEM_1, SYS_ERRORS.PREDECESSORS_ITEM_2, SYS_ERRORS.PREDECESSORS_ITEM_3, SYS_ERRORS.PREDECESSORS_ITEM_4), clsPredecessor)
    End Function

    Friend ReadOnly Property oCollection() As clsCollectionBase
        Get
            Return mp_oCollection
        End Get
    End Property

    Public Function Add(ByVal SuccessorKey As String, ByVal PredecessorKey As String, Optional ByVal PredecessorType As E_CONSTRAINTTYPE = E_CONSTRAINTTYPE.PCT_START_TO_END, Optional ByVal Key As String = "", Optional ByVal StyleIndex As String = "") As clsPredecessor
        mp_oCollection.AddMode = True
        Dim oPredecessor As New clsPredecessor(mp_oControl, Me)
        oPredecessor.PredecessorType = PredecessorType
        oPredecessor.PredecessorKey = PredecessorKey
        oPredecessor.StyleIndex = StyleIndex
        oPredecessor.Key = Key
        oPredecessor.SuccessorKey = SuccessorKey
        mp_oCollection.m_Add(oPredecessor, Key, SYS_ERRORS.PREDECESSORS_ADD_1, SYS_ERRORS.PREDECESSORS_ADD_2, False, SYS_ERRORS.PREDECESSORS_ADD_3)
        Return oPredecessor
    End Function

    
    Public Sub Clear()
        mp_oCollection.m_Clear()
    End Sub

    
    Public Sub Remove(ByVal Index As String)
        mp_oCollection.m_Remove(Index, SYS_ERRORS.PREDECESSORS_REMOVE_1, SYS_ERRORS.PREDECESSORS_REMOVE_2, SYS_ERRORS.PREDECESSORS_REMOVE_3, SYS_ERRORS.PREDECESSORS_REMOVE_4)
    End Sub

    Friend Sub Draw()
        Dim lIndex As Integer
        Dim oPredecessor As clsPredecessor
        For lIndex = 1 To Count
            oPredecessor = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsPredecessor)
            oPredecessor.ClearRectangles()
            If oPredecessor.SuccessorTask.Row.Height > -1 And oPredecessor.PredecessorTask.Row.Height > -1 Then
                Select Case oPredecessor.PredecessorType
                    Case E_CONSTRAINTTYPE.PCT_END_TO_END
                        mp_DrawEndToEnd(oPredecessor)
                    Case E_CONSTRAINTTYPE.PCT_START_TO_START
                        mp_DrawStartToStart(oPredecessor)
                    Case E_CONSTRAINTTYPE.PCT_START_TO_END
                        mp_DrawStartToEnd(oPredecessor)
                    Case E_CONSTRAINTTYPE.PCT_END_TO_START
                        mp_DrawEndToStart(oPredecessor)
                End Select
            End If
        Next
    End Sub

    Private Sub mp_DrawEndToEnd(ByVal oPredecessor As clsPredecessor)
        Dim oPredecessorTask As clsTask
        Dim oSuccessorTask As clsTask
        Dim lPredecessorCtr As Integer
        Dim lSuccessorCtr As Integer
        Dim lXOffset As Integer
        Dim lYOffset As Integer
        Dim oStyle As clsPredecessorStyle
        oStyle = mp_GetPredecessorStyle(oPredecessor)
        lXOffset = oStyle.XOffset
        lYOffset = oStyle.YOffset
        oPredecessorTask = oPredecessor.PredecessorTask
        oSuccessorTask = oPredecessor.SuccessorTask
        If mp_DrawPredecessor(oPredecessorTask, oSuccessorTask) = False Then
            Return
        End If
        lPredecessorCtr = System.Convert.ToInt32(oPredecessorTask.Top + ((oPredecessorTask.Bottom - oPredecessorTask.Top) / 2))
        lSuccessorCtr = System.Convert.ToInt32(oSuccessorTask.Top + ((oSuccessorTask.Bottom - oSuccessorTask.Top) / 2))
        If oPredecessorTask.Right >= oSuccessorTask.Right Then
            Dim oPoints(3) As S_Point
            oPoints(0).X = oPredecessorTask.Right
            oPoints(0).Y = lPredecessorCtr
            oPoints(1).X = oPredecessorTask.Right + lXOffset
            oPoints(1).Y = lPredecessorCtr
            oPoints(2).X = oPredecessorTask.Right + lXOffset
            oPoints(2).Y = lSuccessorCtr
            oPoints(3).X = oSuccessorTask.Right
            oPoints(3).Y = lSuccessorCtr
            mp_DrawPredecessorLines(oPoints, oPredecessor)
            mp_oControl.clsG.mp_DrawArrow(oPoints(3).X + 1, oPoints(3).Y, GRE_ARROWDIRECTION.AWD_LEFT, oStyle.ArrowSize, oStyle.LineColor)
        Else
            Dim oPoints(3) As S_Point
            oPoints(0).X = oSuccessorTask.Right
            oPoints(0).Y = lSuccessorCtr
            oPoints(1).X = oSuccessorTask.Right + lXOffset
            oPoints(1).Y = lSuccessorCtr
            oPoints(2).X = oSuccessorTask.Right + lXOffset
            oPoints(2).Y = lPredecessorCtr
            oPoints(3).X = oPredecessorTask.Right
            oPoints(3).Y = lPredecessorCtr
            mp_DrawPredecessorLines(oPoints, oPredecessor)
            mp_oControl.clsG.mp_DrawArrow(oPoints(0).X + 1, oPoints(0).Y, GRE_ARROWDIRECTION.AWD_LEFT, oStyle.ArrowSize, oStyle.LineColor)
        End If
    End Sub

    Private Sub mp_DrawStartToStart(ByVal oPredecessor As clsPredecessor)
        Dim oPredecessorTask As clsTask
        Dim oSuccessorTask As clsTask
        Dim lPredecessorCtr As Integer
        Dim lSuccessorCtr As Integer
        Dim lXOffset As Integer
        Dim lYOffset As Integer
        Dim oStyle As clsPredecessorStyle
        oStyle = mp_GetPredecessorStyle(oPredecessor)
        lXOffset = oStyle.XOffset
        lYOffset = oStyle.YOffset
        oPredecessorTask = oPredecessor.PredecessorTask
        oSuccessorTask = oPredecessor.SuccessorTask
        If mp_DrawPredecessor(oPredecessorTask, oSuccessorTask) = False Then
            Return
        End If
        lPredecessorCtr = System.Convert.ToInt32(oPredecessorTask.Top + ((oPredecessorTask.Bottom - oPredecessorTask.Top) / 2))
        lSuccessorCtr = System.Convert.ToInt32(oSuccessorTask.Top + ((oSuccessorTask.Bottom - oSuccessorTask.Top) / 2))
        If oPredecessorTask.Left <= oSuccessorTask.Left Then
            Dim oPoints(3) As S_Point
            oPoints(0).X = oPredecessorTask.Left
            oPoints(0).Y = lPredecessorCtr
            oPoints(1).X = oPredecessorTask.Left - lXOffset
            oPoints(1).Y = lPredecessorCtr
            oPoints(2).X = oPredecessorTask.Left - lXOffset
            oPoints(2).Y = lSuccessorCtr
            oPoints(3).X = oSuccessorTask.Left
            oPoints(3).Y = lSuccessorCtr
            mp_DrawPredecessorLines(oPoints, oPredecessor)
            mp_oControl.clsG.mp_DrawArrow(oPoints(3).X - 1, oPoints(3).Y, GRE_ARROWDIRECTION.AWD_RIGHT, oStyle.ArrowSize, oStyle.LineColor)
        Else
            Dim oPoints(3) As S_Point
            oPoints(0).X = oSuccessorTask.Left
            oPoints(0).Y = lSuccessorCtr
            oPoints(1).X = oSuccessorTask.Left - lXOffset
            oPoints(1).Y = lSuccessorCtr
            oPoints(2).X = oSuccessorTask.Left - lXOffset
            oPoints(2).Y = lPredecessorCtr
            oPoints(3).X = oPredecessorTask.Left
            oPoints(3).Y = lPredecessorCtr
            mp_DrawPredecessorLines(oPoints, oPredecessor)
            mp_oControl.clsG.mp_DrawArrow(oPoints(0).X - 1, oPoints(0).Y, GRE_ARROWDIRECTION.AWD_RIGHT, oStyle.ArrowSize, oStyle.LineColor)
        End If
    End Sub

    Private Sub mp_DrawStartToEnd(ByVal oPredecessor As clsPredecessor)
        Dim oPredecessorTask As clsTask
        Dim oSuccessorTask As clsTask
        Dim lPredecessorCtr As Integer
        Dim lSuccessorCtr As Integer
        Dim lXOffset As Integer
        Dim lYOffset As Integer
        Dim oStyle As clsPredecessorStyle
        oStyle = mp_GetPredecessorStyle(oPredecessor)
        lXOffset = oStyle.XOffset
        lYOffset = oStyle.YOffset
        oPredecessorTask = oPredecessor.PredecessorTask
        oSuccessorTask = oPredecessor.SuccessorTask
        If mp_DrawPredecessor(oPredecessorTask, oSuccessorTask) = False Then
            Return
        End If
        lPredecessorCtr = System.Convert.ToInt32(oPredecessorTask.Top + ((oPredecessorTask.Bottom - oPredecessorTask.Top) / 2))
        lSuccessorCtr = System.Convert.ToInt32(oSuccessorTask.Top + ((oSuccessorTask.Bottom - oSuccessorTask.Top) / 2))

        If lPredecessorCtr < lSuccessorCtr Then 'Down
            Dim oPoints(5) As S_Point
            oPoints(0).X = oPredecessorTask.Left
            oPoints(0).Y = lPredecessorCtr
            oPoints(1).X = oPredecessorTask.Left - lXOffset
            oPoints(1).Y = lPredecessorCtr
            oPoints(2).X = oPredecessorTask.Left - lXOffset
            oPoints(2).Y = oPredecessorTask.Bottom + lYOffset
            oPoints(3).X = oSuccessorTask.Right + lXOffset
            oPoints(3).Y = oPredecessorTask.Bottom + lYOffset
            oPoints(4).X = oSuccessorTask.Right + lXOffset
            oPoints(4).Y = lSuccessorCtr
            oPoints(5).X = oSuccessorTask.Right
            oPoints(5).Y = lSuccessorCtr
            mp_DrawPredecessorLines(oPoints, oPredecessor)
            mp_oControl.clsG.mp_DrawArrow(oPoints(5).X + 1, oPoints(5).Y, GRE_ARROWDIRECTION.AWD_LEFT, oStyle.ArrowSize, oStyle.LineColor)
        Else
            Dim oPoints(5) As S_Point
            oPoints(0).X = oPredecessorTask.Left
            oPoints(0).Y = lPredecessorCtr
            oPoints(1).X = oPredecessorTask.Left - lXOffset
            oPoints(1).Y = lPredecessorCtr
            oPoints(2).X = oPredecessorTask.Left - lXOffset
            oPoints(2).Y = oPredecessorTask.Top - lYOffset
            oPoints(3).X = oSuccessorTask.Right + lXOffset
            oPoints(3).Y = oPredecessorTask.Top - lYOffset
            oPoints(4).X = oSuccessorTask.Right + lXOffset
            oPoints(4).Y = lSuccessorCtr
            oPoints(5).X = oSuccessorTask.Right
            oPoints(5).Y = lSuccessorCtr
            mp_DrawPredecessorLines(oPoints, oPredecessor)
            mp_oControl.clsG.mp_DrawArrow(oPoints(5).X + 1, oPoints(5).Y, GRE_ARROWDIRECTION.AWD_LEFT, oStyle.ArrowSize, oStyle.LineColor)
        End If
    End Sub

    Private Sub mp_DrawEndToStart(ByVal oPredecessor As clsPredecessor)
        Dim oPredecessorTask As clsTask
        Dim oSuccessorTask As clsTask
        Dim lPredecessorCtr As Integer
        Dim lSuccessorCtr As Integer
        Dim lXOffset As Integer
        Dim lYOffset As Integer
        Dim oStyle As clsPredecessorStyle
        oStyle = mp_GetPredecessorStyle(oPredecessor)
        lXOffset = oStyle.XOffset
        lYOffset = oStyle.YOffset
        oPredecessorTask = oPredecessor.PredecessorTask
        oSuccessorTask = oPredecessor.SuccessorTask
        If mp_DrawPredecessor(oPredecessorTask, oSuccessorTask) = False Then
            Return
        End If
        lPredecessorCtr = System.Convert.ToInt32(oPredecessorTask.Top + ((oPredecessorTask.Bottom - oPredecessorTask.Top) / 2))
        lSuccessorCtr = System.Convert.ToInt32(oSuccessorTask.Top + ((oSuccessorTask.Bottom - oSuccessorTask.Top) / 2))
        If oPredecessor.PredecessorTask.Right <= oPredecessor.SuccessorTask.Left Then
            '//With Lag
            If lPredecessorCtr < lSuccessorCtr Then '//Down
                Dim oPoints(2) As S_Point
                oPoints(0).X = oPredecessorTask.Right
                oPoints(0).Y = lPredecessorCtr
                oPoints(1).X = oSuccessorTask.Left + lXOffset
                oPoints(1).Y = lPredecessorCtr
                oPoints(2).X = oSuccessorTask.Left + lXOffset
                oPoints(2).Y = oSuccessorTask.Top
                mp_DrawPredecessorLines(oPoints, oPredecessor)
                mp_oControl.clsG.mp_DrawArrow(oPoints(2).X, oPoints(2).Y - 1, GRE_ARROWDIRECTION.AWD_DOWN, oStyle.ArrowSize, oStyle.LineColor)
            Else
                Dim oPoints(2) As S_Point
                oPoints(0).X = oPredecessorTask.Right
                oPoints(0).Y = lPredecessorCtr
                oPoints(1).X = oSuccessorTask.Left + lXOffset
                oPoints(1).Y = lPredecessorCtr
                oPoints(2).X = oSuccessorTask.Left + lXOffset
                oPoints(2).Y = oSuccessorTask.Bottom
                mp_DrawPredecessorLines(oPoints, oPredecessor)
                mp_oControl.clsG.mp_DrawArrow(oPoints(2).X, oPoints(2).Y + 1, GRE_ARROWDIRECTION.AWD_UP, oStyle.ArrowSize, oStyle.LineColor)
            End If
        Else
            '//With Lead
            If lPredecessorCtr < lSuccessorCtr Then '//Down
                Dim oPoints(5) As S_Point
                oPoints(0).X = oPredecessorTask.Right
                oPoints(0).Y = lPredecessorCtr
                oPoints(1).X = oPredecessorTask.Right + lXOffset
                oPoints(1).Y = lPredecessorCtr
                oPoints(2).X = oPredecessorTask.Right + lXOffset
                oPoints(2).Y = lSuccessorCtr - lYOffset
                oPoints(3).X = oSuccessorTask.Left - lXOffset
                oPoints(3).Y = lSuccessorCtr - lYOffset
                oPoints(4).X = oSuccessorTask.Left - lXOffset
                oPoints(4).Y = lSuccessorCtr
                oPoints(5).X = oSuccessorTask.Left
                oPoints(5).Y = lSuccessorCtr
                mp_DrawPredecessorLines(oPoints, oPredecessor)
                mp_oControl.clsG.mp_DrawArrow(oPoints(5).X - 1, oPoints(5).Y, GRE_ARROWDIRECTION.AWD_RIGHT, oStyle.ArrowSize, oStyle.LineColor)
            Else '// Up
                Dim oPoints(5) As S_Point
                oPoints(0).X = oPredecessorTask.Right
                oPoints(0).Y = lPredecessorCtr
                oPoints(1).X = oPredecessorTask.Right + lXOffset
                oPoints(1).Y = lPredecessorCtr
                oPoints(2).X = oPredecessorTask.Right + lXOffset
                oPoints(2).Y = lSuccessorCtr + lYOffset
                oPoints(3).X = oSuccessorTask.Left - lXOffset
                oPoints(3).Y = lSuccessorCtr + lYOffset
                oPoints(4).X = oSuccessorTask.Left - lXOffset
                oPoints(4).Y = lSuccessorCtr
                oPoints(5).X = oSuccessorTask.Left
                oPoints(5).Y = lSuccessorCtr
                mp_DrawPredecessorLines(oPoints, oPredecessor)
                mp_oControl.clsG.mp_DrawArrow(oPoints(5).X - 1, oPoints(5).Y, GRE_ARROWDIRECTION.AWD_RIGHT, oStyle.ArrowSize, oStyle.LineColor)
            End If
        End If
    End Sub

    Private Sub mp_DrawPredecessorLines(ByVal oPoints As S_Point(), ByVal oPredecessor As clsPredecessor)
        Dim i As Integer
        For i = 0 To oPoints.GetUpperBound(0) - 1
            Dim oStyle As clsPredecessorStyle
            oStyle = mp_GetPredecessorStyle(oPredecessor)
            mp_oControl.clsG.DrawLine(oPoints(i).X, oPoints(i).Y, oPoints(i + 1).X, oPoints(i + 1).Y, GRE_LINETYPE.LT_NORMAL, oStyle.LineColor, oStyle.LineStyle, oStyle.LineWidth)
            Dim oRectangle As S_Rectangle
            If oPoints(i).X = oPoints(i + 1).X Then '//Vertical Line
                oRectangle.X1 = oPoints(i).X - mp_oControl.CurrentViewObject.ClientArea.PredecessorSelectionOffset
                oRectangle.X2 = oPoints(i + 1).X + mp_oControl.CurrentViewObject.ClientArea.PredecessorSelectionOffset
                If oPoints(i).Y < oPoints(i + 1).Y Then
                    oRectangle.Y1 = oPoints(i).Y
                    oRectangle.Y2 = oPoints(i + 1).Y
                Else
                    oRectangle.Y1 = oPoints(i + 1).Y
                    oRectangle.Y2 = oPoints(i).Y
                End If
            ElseIf oPoints(i).Y = oPoints(i + 1).Y Then  '//Horizontal Line
                oRectangle.Y1 = oPoints(i).Y - mp_oControl.CurrentViewObject.ClientArea.PredecessorSelectionOffset
                oRectangle.Y2 = oPoints(i + 1).Y + mp_oControl.CurrentViewObject.ClientArea.PredecessorSelectionOffset
                If oPoints(i).X < oPoints(i + 1).X Then
                    oRectangle.X1 = oPoints(i).X
                    oRectangle.X2 = oPoints(i + 1).X
                Else
                    oRectangle.X1 = oPoints(i + 1).X
                    oRectangle.X2 = oPoints(i).X
                End If
            End If
            oPredecessor.AddRectangle(oRectangle)
        Next
    End Sub

    Private Function mp_GetPredecessorStyle(ByVal oPredecessor As clsPredecessor) As clsPredecessorStyle
        Dim oStyle As clsPredecessorStyle
        If oPredecessor.mp_bWarning = False Then
            If oPredecessor.Index = mp_oControl.SelectedPredecessorIndex Then
                oStyle = oPredecessor.SelectedStyle.PredecessorStyle
            Else
                oStyle = oPredecessor.Style.PredecessorStyle
            End If
        Else
            oStyle = oPredecessor.WarningStyle.PredecessorStyle
        End If
        Return oStyle
    End Function

    Private Function mp_DrawPredecessor(ByVal oTask1 As clsTask, ByVal oTask2 As clsTask) As Boolean
        If oTask1.Row.ClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_ABOVEVISIBLEAREA And oTask2.Row.ClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_ABOVEVISIBLEAREA Then
            Return False
        End If
        If oTask1.Row.ClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_BELOWVISIBLEAREA And oTask2.Row.ClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_BELOWVISIBLEAREA Then
            Return False
        End If
        If oTask1.ClientAreaVisiblity = E_CLIENTAREAVISIBILITY.VS_RIGHTOFVISIBLEAREA And oTask2.ClientAreaVisiblity = E_CLIENTAREAVISIBILITY.VS_RIGHTOFVISIBLEAREA Then
            Return False
        End If
        If oTask1.ClientAreaVisiblity = E_CLIENTAREAVISIBILITY.VS_LEFTOFVISIBLEAREA And oTask2.ClientAreaVisiblity = E_CLIENTAREAVISIBILITY.VS_LEFTOFVISIBLEAREA Then
            Return False
        End If
        Return True
    End Function



    
    Public Function GetXML() As String
        Dim lIndex As Integer
        Dim oPredecessor As clsPredecessor
        Dim oXML As New clsXML(mp_oControl, "Predecessors")
        oXML.InitializeWriter()
        For lIndex = 1 To Count
            oPredecessor = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsPredecessor)
            oXML.WriteObject(oPredecessor.GetXML)
        Next lIndex
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim lIndex As Integer
        Dim oXML As New clsXML(mp_oControl, "Predecessors")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        mp_oCollection.m_Clear()
        For lIndex = 1 To oXML.ReadCollectionCount
            Dim oPredecessor As New clsPredecessor(mp_oControl, Me)
            oPredecessor.SetXML(oXML.ReadCollectionObject(lIndex))
            mp_oCollection.AddMode = True
            mp_oCollection.m_Add(oPredecessor, oPredecessor.Key, SYS_ERRORS.PREDECESSORS_ADD_1, SYS_ERRORS.PREDECESSORS_ADD_2, False, SYS_ERRORS.PREDECESSORS_ADD_3)
            oPredecessor = Nothing
        Next lIndex
    End Sub

End Class
