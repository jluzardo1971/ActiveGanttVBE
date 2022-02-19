Option Explicit On 

Imports System.Math
Imports System.Convert
Imports System.Diagnostics

Public Class clsMath

    Private mp_oControl As ActiveGanttVBECtl

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
    End Sub

    Public Function DateTimeAdd(ByVal Interval As E_INTERVAL, ByVal Number As Integer, ByVal dtDate As AGVBE.DateTime) As AGVBE.DateTime
        Dim dtReturn As AGVBE.DateTime = New AGVBE.DateTime()
        Dim dtBuff As System.DateTime = New System.DateTime(0)
        Dim lBuff As Integer = 0
        Dim lDay As Long = 0
        Dim lHour As Long = 0
        Dim lMinute As Long = 0
        Dim lSecond As Long = 0
        Dim lMillisecond As Long = 0
        Dim lMicrosecond As Long = 0
        Dim lNanosecond As Long = 0
        Dim bNegative As Boolean = False

        If (Number < 0) Then
            bNegative = True
            Number = System.Math.Abs(Number)
        End If

        dtBuff = New System.DateTime(dtDate.Year(), dtDate.Month(), dtDate.Day(), dtDate.Hour(), dtDate.Minute(), dtDate.Second())
        Select Case Interval
            Case E_INTERVAL.IL_NANOSECOND
                lDay = Convert.ToInt64(System.Math.Floor(Convert.ToDouble(Number) / Convert.ToDouble(86400000000000)))
                Number = Convert.ToInt32(Convert.ToInt64(Number) - (lDay * Convert.ToInt64(86400000000000)))
                lHour = Convert.ToInt64(System.Math.Floor(Convert.ToDouble(Number) / Convert.ToDouble(3600000000000)))
                Number = Convert.ToInt32(Convert.ToInt64(Number) - (lHour * Convert.ToInt64(3600000000000)))
                lMinute = Convert.ToInt64(System.Math.Floor(Convert.ToDouble(Number) / Convert.ToDouble(60000000000)))
                Number = Convert.ToInt32(Convert.ToInt64(Number) - (lMinute * Convert.ToInt64(60000000000)))
                lSecond = Convert.ToInt64(System.Math.Floor(Convert.ToDouble(Number) / Convert.ToDouble(1000000000)))
                Number = Convert.ToInt32(Convert.ToInt64(Number) - (lSecond * Convert.ToInt64(1000000000)))
                lMillisecond = Convert.ToInt64(System.Math.Floor(Convert.ToDouble(Number) / Convert.ToDouble(1000000)))
                Number = Convert.ToInt32(Convert.ToInt64(Number) - (lMillisecond * Convert.ToInt64(1000000)))
                lMicrosecond = Convert.ToInt64(System.Math.Floor(Convert.ToDouble(Number) / Convert.ToDouble(1000)))
                Number = Convert.ToInt32(Convert.ToInt64(Number) - (lMicrosecond * Convert.ToInt64(1000)))
                lNanosecond = Convert.ToInt64(Number)
            Case E_INTERVAL.IL_MICROSECOND
                lDay = Convert.ToInt64(System.Math.Floor(Convert.ToDouble(Number) / Convert.ToDouble(86400000000)))
                Number = Convert.ToInt32(Convert.ToInt64(Number) - (lDay * Convert.ToInt64(86400000000)))
                lHour = Convert.ToInt64(System.Math.Floor(Convert.ToDouble(Number) / Convert.ToDouble(3600000000)))
                Number = Convert.ToInt32(Convert.ToInt64(Number) - (lHour * Convert.ToInt64(3600000000)))
                lMinute = Convert.ToInt64(System.Math.Floor(Convert.ToDouble(Number) / Convert.ToDouble(60000000)))
                Number = Convert.ToInt32(Convert.ToInt64(Number) - (lMinute * Convert.ToInt64(60000000)))
                lSecond = Convert.ToInt64(System.Math.Floor(Convert.ToDouble(Number) / Convert.ToDouble(1000000)))
                Number = Convert.ToInt32(Convert.ToInt64(Number) - (lSecond * Convert.ToInt64(1000000)))
                lMillisecond = Convert.ToInt64(System.Math.Floor(Convert.ToDouble(Number) / Convert.ToDouble(1000)))
                Number = Convert.ToInt32(Convert.ToInt64(Number) - (lMillisecond * Convert.ToInt64(1000)))
                lMicrosecond = Convert.ToInt64(Number)
            Case E_INTERVAL.IL_MILLISECOND
                lDay = Convert.ToInt64(System.Math.Floor(Convert.ToDouble(Number) / Convert.ToDouble(86400000)))
                Number = Convert.ToInt32(Convert.ToInt64(Number) - (lDay * Convert.ToInt64(86400000)))
                lHour = Convert.ToInt64(System.Math.Floor(Convert.ToDouble(Number) / Convert.ToDouble(3600000)))
                Number = Convert.ToInt32(Convert.ToInt64(Number) - (lHour * Convert.ToInt64(3600000)))
                lMinute = Convert.ToInt64(System.Math.Floor(Convert.ToDouble(Number) / Convert.ToDouble(60000)))
                Number = Convert.ToInt32(Convert.ToInt64(Number) - (lMinute * Convert.ToInt64(60000)))
                lSecond = Convert.ToInt64(System.Math.Floor(Convert.ToDouble(Number) / Convert.ToDouble(1000)))
                Number = Convert.ToInt32(Convert.ToInt64(Number) - (lSecond * Convert.ToInt64(1000)))
                lMillisecond = Convert.ToInt64(Number)
            Case E_INTERVAL.IL_SECOND
                If bNegative = True Then
                    Number = Number * (-1)
                End If
                dtBuff = dtBuff.AddSeconds(Number)
                lBuff = dtDate.SecondFractionPart
            Case E_INTERVAL.IL_MINUTE
                If bNegative = True Then
                    Number = Number * (-1)
                End If
                dtBuff = dtBuff.AddMinutes(Number)
                lBuff = dtDate.SecondFractionPart
            Case E_INTERVAL.IL_HOUR
                If bNegative = True Then
                    Number = Number * (-1)
                End If
                dtBuff = dtBuff.AddHours(Number)
                lBuff = dtDate.SecondFractionPart
            Case E_INTERVAL.IL_DAY
                If bNegative = True Then
                    Number = Number * (-1)
                End If
                dtBuff = dtBuff.AddDays(Number)
                lBuff = dtDate.SecondFractionPart
            Case E_INTERVAL.IL_WEEK
                If bNegative = True Then
                    Number = Number * (-1)
                End If
                dtBuff = dtBuff.AddDays(Number * 7)
                lBuff = dtDate.SecondFractionPart
            Case E_INTERVAL.IL_MONTH
                If bNegative = True Then
                    Number = Number * (-1)
                End If
                dtBuff = dtBuff.AddMonths(Number)
                lBuff = dtDate.SecondFractionPart
            Case E_INTERVAL.IL_QUARTER
                If bNegative = True Then
                    Number = Number * (-1)
                End If
                dtBuff = dtBuff.AddMonths(Number * 3)
                lBuff = dtDate.SecondFractionPart
            Case E_INTERVAL.IL_YEAR
                If bNegative = True Then
                    Number = Number * (-1)
                End If
                dtBuff = dtBuff.AddYears(Number)
                lBuff = dtDate.SecondFractionPart
        End Select

        If Interval = E_INTERVAL.IL_MILLISECOND Or Interval = E_INTERVAL.IL_MICROSECOND Or Interval = E_INTERVAL.IL_NANOSECOND Then
            If bNegative = False Then
                dtBuff = dtBuff.AddDays(lDay)
                dtBuff = dtBuff.AddHours(lHour)
                dtBuff = dtBuff.AddMinutes(lMinute)
                dtBuff = dtBuff.AddSeconds(lSecond)
            Else
                dtBuff = dtBuff.AddDays(-lDay)
                dtBuff = dtBuff.AddHours(-lHour)
                dtBuff = dtBuff.AddMinutes(-lMinute)
                dtBuff = dtBuff.AddSeconds(-lSecond)
            End If
            If (bNegative = False) Then
                lBuff = dtDate.SecondFractionPart + (Convert.ToInt32(lMillisecond) * 1000000) + (Convert.ToInt32(lMicrosecond) * 1000) + Convert.ToInt32(lNanosecond)
                If (lBuff > 99999999) Then
                    Dim lAdditionalSeconds As Integer = Convert.ToInt32(System.Math.Floor(Convert.ToDouble(lBuff) / Convert.ToDouble(1000000000)))
                    dtBuff = dtBuff.AddSeconds(Convert.ToDouble(lAdditionalSeconds))
                    lBuff = lBuff - (lAdditionalSeconds * 1000000000)
                End If
            Else
                lBuff = dtDate.SecondFractionPart - (Convert.ToInt32(lMillisecond) * 1000000) - (Convert.ToInt32(lMicrosecond) * 1000) - Convert.ToInt32(lNanosecond)
                If (lBuff < 0) Then
                    Dim lAdditionalSeconds As Integer = Convert.ToInt32(System.Math.Floor(Convert.ToDouble(-lBuff) / Convert.ToDouble(1000000000)) - 1)
                    dtBuff = dtBuff.AddSeconds(Convert.ToDouble(lAdditionalSeconds))
                    lBuff = lBuff + (lAdditionalSeconds * 1000000000 * (-1))
                End If
            End If
        End If

        dtReturn = New AGVBE.DateTime(dtBuff.Year, dtBuff.Month, dtBuff.Day, dtBuff.Hour, dtBuff.Minute, dtBuff.Second)
        dtReturn.SecondFractionPart = lBuff

        Return dtReturn

    End Function

    Public Function DateTimeDiff(ByVal Interval As E_INTERVAL, ByVal dtDate1 As AGVBE.DateTime, ByVal dtDate2 As AGVBE.DateTime) As Integer
        Dim lSecondFractionSpan As Integer = 0
        Dim lReturn As Integer = 0
        Dim lReturn64 As Long = 0
        Dim tsResult As System.TimeSpan = New System.TimeSpan(0)
        Dim lYearDiff As Integer = 0
        If dtDate1 = dtDate2 Then
            Return 0
        End If
        If Interval = E_INTERVAL.IL_MILLISECOND Or Interval = E_INTERVAL.IL_MICROSECOND Or Interval = E_INTERVAL.IL_NANOSECOND Then
            If dtDate1 < dtDate2 Then
                If dtDate1.SecondFractionPart > 0 Then
                    lSecondFractionSpan = 1000000000 - dtDate1.SecondFractionPart + dtDate2.SecondFractionPart
                    If lSecondFractionSpan >= 1000000000 Then
                        lSecondFractionSpan = lSecondFractionSpan - 1000000000
                    End If
                Else
                    lSecondFractionSpan = dtDate2.SecondFractionPart
                End If
            Else
                If dtDate2.SecondFractionPart > 0 Then
                    lSecondFractionSpan = 1000000000 - dtDate2.SecondFractionPart + dtDate1.SecondFractionPart
                    If lSecondFractionSpan >= 1000000000 Then
                        lSecondFractionSpan = lSecondFractionSpan - 1000000000
                    End If
                Else
                    lSecondFractionSpan = dtDate1.SecondFractionPart
                End If
                lSecondFractionSpan = -lSecondFractionSpan
            End If
        End If
        Select Case Interval
            Case E_INTERVAL.IL_NANOSECOND
                tsResult = dtDate2.DateTimePart.Subtract(dtDate1.DateTimePart)
                lReturn = Convert.ToInt32(System.Math.Floor(tsResult.TotalSeconds))
                lReturn64 = Convert.ToInt64(lReturn) * Convert.ToInt64(1000000000)
                lReturn64 = Convert.ToInt64(lReturn64) + Convert.ToInt64(lSecondFractionSpan)
                lReturn = Convert.ToInt32(lReturn64)
                Return lReturn
            Case E_INTERVAL.IL_MICROSECOND
                tsResult = dtDate2.DateTimePart.Subtract(dtDate1.DateTimePart)
                lReturn = Convert.ToInt32(System.Math.Floor(tsResult.TotalSeconds))
                lReturn64 = Convert.ToInt64(lReturn) * Convert.ToInt64(1000000000)
                lReturn64 = Convert.ToInt64(lReturn64) + Convert.ToInt64(lSecondFractionSpan)
                lReturn = Convert.ToInt32(System.Math.Floor(ToDouble(lReturn64) / ToDouble(1000)))
                Return lReturn
            Case E_INTERVAL.IL_MILLISECOND
                tsResult = dtDate2.DateTimePart.Subtract(dtDate1.DateTimePart)
                lReturn = Convert.ToInt32(System.Math.Floor(tsResult.TotalSeconds))
                lReturn64 = Convert.ToInt64(lReturn) * Convert.ToInt64(1000000000)
                lReturn64 = Convert.ToInt64(lReturn64) + Convert.ToInt64(lSecondFractionSpan)
                lReturn = Convert.ToInt32(System.Math.Floor(ToDouble(lReturn64) / ToDouble(1000000)))
                Return lReturn
            Case E_INTERVAL.IL_SECOND
                tsResult = dtDate2.DateTimePart.Subtract(dtDate1.DateTimePart)
                lReturn = Convert.ToInt32(System.Math.Floor(tsResult.TotalSeconds))
                Return lReturn
            Case E_INTERVAL.IL_MINUTE
                tsResult = dtDate2.DateTimePart.Subtract(dtDate1.DateTimePart)
                lReturn = Convert.ToInt32(System.Math.Floor(tsResult.TotalMinutes))
                Return lReturn
            Case E_INTERVAL.IL_HOUR
                tsResult = dtDate2.DateTimePart.Subtract(dtDate1.DateTimePart)
                lReturn = Convert.ToInt32(System.Math.Floor(tsResult.TotalHours))
                Return lReturn
            Case E_INTERVAL.IL_DAY
                tsResult = dtDate2.DateTimePart.Subtract(dtDate1.DateTimePart)
                lReturn = Convert.ToInt32(System.Math.Floor(tsResult.TotalDays))
                Return lReturn
            Case E_INTERVAL.IL_WEEK
                tsResult = dtDate2.DateTimePart.Subtract(dtDate1.DateTimePart)
                lReturn = Convert.ToInt32(System.Math.Floor(tsResult.TotalDays / 7))
                Return lReturn
            Case E_INTERVAL.IL_MONTH
                lYearDiff = dtDate2.Year() - dtDate1.Year()
                Dim lMonthDiff As Integer = dtDate2.Month() - dtDate1.Month()
                lYearDiff = (lYearDiff * 12)
                Return lYearDiff + lMonthDiff
            Case E_INTERVAL.IL_YEAR
                lYearDiff = dtDate2.Year() - dtDate1.Year()
                Return lYearDiff
        End Select
        Return 0
    End Function

    Public Function GetXCoordinateFromDate(ByVal dtCoordinate As AGVBE.DateTime) As Integer
        Return System.Convert.ToInt32((DateTimeDiff(mp_oControl.CurrentViewObject.Interval, mp_oControl.CurrentViewObject.TimeLine.StartDate, dtCoordinate) / mp_oControl.CurrentViewObject.Factor) + mp_oControl.CurrentViewObject.TimeLine.f_lStart)
    End Function

    Public Function GetDateFromXCoordinate(ByVal v_lXCoordinate As Integer) As AGVBE.DateTime
        Return DateTimeAdd(mp_oControl.CurrentViewObject.Interval, (v_lXCoordinate - mp_oControl.CurrentViewObject.TimeLine.f_lStart) * mp_oControl.CurrentViewObject.Factor, mp_oControl.CurrentViewObject.TimeLine.StartDate)
    End Function

    Public Function GetRowIndexByPosition(ByVal Y As Integer) As Integer
        Dim oRow As clsRow
        Dim lVisRowIndex As Integer
        If (mp_oControl.Rows.Count = 0) Then
            Return -1
        End If
        For lVisRowIndex = mp_oControl.VerticalScrollBar.Value To mp_oControl.CurrentViewObject.ClientArea.LastVisibleRow
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lVisRowIndex), clsRow)
            If Y >= oRow.Top And Y <= oRow.Bottom And oRow.Visible = True Then
                Return lVisRowIndex
            End If
        Next lVisRowIndex
        Return -1
    End Function

    Public Function GetCellIndexByPosition(ByVal X As Integer) As Integer
        Dim oColumn As clsColumn
        Dim lIndex As Integer
        For lIndex = 1 To mp_oControl.Columns.Count
            oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(lIndex), clsColumn)
            If X > oColumn.Left And X < oColumn.Right Then
                Return lIndex
            End If
        Next lIndex
        Return -1
    End Function

    Public Function GetColumnIndexByPosition(ByVal X As Integer, ByVal Y As Integer) As Integer
        Dim oColumn As clsColumn
        Dim lIndex As Integer
        For lIndex = 1 To mp_oControl.Columns.Count
            oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(lIndex), clsColumn)
            If X >= oColumn.Left And X <= oColumn.Right And Y >= oColumn.Top And Y <= oColumn.Bottom Then
                Return lIndex
            End If
        Next lIndex
        Return -1
    End Function

    Public Function GetTaskIndexByPosition(ByVal X As Integer, ByVal Y As Integer) As Integer
        Dim oTask As clsTask
        Dim lIndex As Integer
        For lIndex = mp_oControl.Tasks.Count To 1 Step -1
            oTask = DirectCast(mp_oControl.Tasks.oCollection.m_oReturnArrayElement(lIndex), clsTask)
            If oTask.Visible = True And InCurrentLayer(oTask.LayerIndex) Then
                If X >= oTask.Left And X <= oTask.Right And Y >= oTask.Top And Y <= oTask.Bottom Then
                    Return lIndex
                End If
            End If
        Next lIndex
        Return -1
    End Function

    Public Function GetPredecessorIndexByPosition(ByVal X As Integer, ByVal Y As Integer) As Integer
        Dim oPredecessor As clsPredecessor
        Dim lIndex As Integer
        For lIndex = mp_oControl.Predecessors.Count To 1 Step -1
            oPredecessor = DirectCast(mp_oControl.Predecessors.oCollection.m_oReturnArrayElement(lIndex), clsPredecessor)
            If oPredecessor.Visible = True And oPredecessor.HitTest(X, Y) = True Then
                Return lIndex
            End If
        Next lIndex
        Return -1
    End Function

    Public Function GetPercentageIndexByPosition(ByVal X As Integer, ByVal Y As Integer) As Integer
        Dim oPercentage As clsPercentage
        Dim lIndex As Integer
        For lIndex = mp_oControl.Percentages.Count To 1 Step -1
            oPercentage = DirectCast(mp_oControl.Percentages.oCollection.m_oReturnArrayElement(lIndex), clsPercentage)
            If oPercentage.Visible = True Then
                If X >= oPercentage.Left And X <= oPercentage.RightSel And Y >= oPercentage.Top And Y <= oPercentage.Bottom Then
                    Return lIndex
                End If
            End If
        Next lIndex
        Return -1
    End Function

    Public Function GetNodeIndexByCheckBoxPosition(ByVal X As Integer, ByVal Y As Integer) As Integer
        Dim lIndex As Integer
        Dim oNode As clsNode
        Dim oRow As clsRow
        Dim lReturn As Integer
        If mp_oControl.Treeview.CheckBoxes = False Then
            Return -1
        End If
        lReturn = -1
        For lIndex = 1 To mp_oControl.Rows.Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oNode = oRow.Node
            If oRow.ClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_INSIDEVISIBLEAREA And X >= (oNode.CheckBoxLeft) And X <= (oNode.CheckBoxLeft + 13) And Y <= (oNode.YCenter + 6) And Y >= (oNode.YCenter - 7) Then
                lReturn = oNode.Index
            End If
        Next lIndex
        Return lReturn
    End Function

    Public Function GetNodeIndexBySignPosition(ByVal X As Integer, ByVal Y As Integer) As Integer
        Dim lIndex As Integer
        Dim oNode As clsNode
        Dim oRow As clsRow
        Dim lReturn As Integer
        If mp_oControl.Treeview.PlusMinusSigns = False Then
            Return -1
        End If
        lReturn = -1
        For lIndex = 1 To mp_oControl.Rows.Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oNode = oRow.Node
            If oRow.ClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_INSIDEVISIBLEAREA And X >= (oNode.Left - 5) And X <= (oNode.Left + 5) And Y <= (oNode.YCenter + 5) And Y >= (oNode.YCenter - 5) Then
                lReturn = oNode.Index
            End If
        Next lIndex
        Return lReturn
    End Function

    Friend Function InCurrentLayer(ByVal sLayer As String) As Boolean
        If (mp_oControl.LayerEnableObjects = E_LAYEROBJECTENABLE.EC_INALLLAYERS) Then
            Return True
        Else
            Dim lLayerIndex As Integer
            Dim lCurrentLayerIndex As Integer
            lLayerIndex = mp_oControl.Layers.oCollection.m_lReturnIndex(sLayer, True)
            lCurrentLayerIndex = mp_oControl.Layers.oCollection.m_lReturnIndex(mp_oControl.CurrentLayer, True)
            If (lLayerIndex <> lCurrentLayerIndex) Then
                Return False
            Else
                Return True
            End If
        End If
    End Function

    Public Function DetectConflict(ByVal StartDate As AGVBE.DateTime, ByVal EndDate As AGVBE.DateTime, ByVal RowKey As String, ByVal ExcludeIndex As Integer, ByVal LayerIndex As String) As Boolean
        Dim oTask As clsTask
        Dim oTimeBlock As clsTimeBlock
        Dim lIndex As Integer
        Dim lLayerIndex As Integer
        Dim lRowIndex As Integer
        If EndDate < StartDate Then
            Dim dtEndDate As AGVBE.DateTime = StartDate
            Dim dtStartDate As AGVBE.DateTime = EndDate
            StartDate = dtStartDate
            EndDate = dtEndDate
        End If
        lLayerIndex = mp_oControl.Layers.oCollection.m_lReturnIndex(LayerIndex, True)
        If (lLayerIndex = -1) Then
            mp_oControl.mp_ErrorReport(SYS_ERRORS.INVALID_LAYER_INDEX, "Invalid Layer Index", "ActiveGanttVBECtl.mp_bDetectConflict")
            Return False
        End If
        '// Compare to other Task Objects
        For lIndex = 1 To mp_oControl.Tasks.Count
            oTask = DirectCast(mp_oControl.Tasks.oCollection.m_oReturnArrayElement(lIndex), clsTask)
            If RowKey = oTask.RowKey And ExcludeIndex <> lIndex Then
                If (mp_oControl.Layers.oCollection.m_lReturnIndex(oTask.LayerIndex, True) = lLayerIndex) Then
                    '// mp_aoTasks              S------------------E
                    '// interval                S------------------E
                    If (StartDate = oTask.StartDate Or EndDate = oTask.EndDate) And (StartDate <> EndDate) Then
                        Return True
                    End If
                    '// mp_aoTasks              S------------------E
                    '// interval                             S------------------E
                    If StartDate > oTask.StartDate And StartDate < oTask.EndDate Then
                        Return True
                    End If
                    '// mp_aoTasks              S------------------E
                    '// interval            S------------------E
                    If EndDate > oTask.StartDate And EndDate < oTask.EndDate Then
                        Return True
                    End If
                    '// mp_aoTasks              S------------------E
                    '// interval                 S-------------------------E
                    If StartDate < oTask.StartDate And EndDate > oTask.EndDate Then
                        Return True
                    End If
                    '// mp_aoTasks         S--------------------------E
                    '// interval                     S------------------E
                    If StartDate > oTask.StartDate And EndDate < oTask.EndDate Then
                        Return True
                    End If
                End If
            End If
        Next lIndex
        '// Compare to TimeBlock Objects 
        For lIndex = 1 To mp_oControl.TimeBlocks.Count
            oTimeBlock = DirectCast(mp_oControl.TimeBlocks.oCollection.m_oReturnArrayElement(lIndex), clsTimeBlock)
            lRowIndex = mp_oControl.Rows.oCollection.m_lFindIndexByKey(RowKey)

            If oTimeBlock.GenerateConflict = True Then
                '// mp_aoTimeBlocks              S------------------E
                '// interval                     S------------------E
                If (StartDate = oTimeBlock.StartDate Or EndDate = oTimeBlock.EndDate) And (StartDate <> EndDate) Then
                    Return True
                End If
                '// mp_aoTimeBlocks              S------------------E
                '// interval                             S------------------E
                If StartDate > oTimeBlock.StartDate And StartDate < oTimeBlock.EndDate Then
                    Return True
                End If
                '// mp_aoTimeBlocks              S------------------E
                '// interval            S------------------E
                If EndDate > oTimeBlock.StartDate And EndDate < oTimeBlock.EndDate Then
                    Return True
                End If
                '// mp_aoTimeBlocks             S------------------E
                '// interval                 S-------------------------E
                If StartDate < oTimeBlock.StartDate And EndDate > oTimeBlock.EndDate Then
                    Return True
                End If
                '// mp_aoTimeBlocks        S--------------------------E
                '// interval                     S------------------E
                If StartDate > oTimeBlock.StartDate And EndDate < oTimeBlock.EndDate Then
                    Return True
                End If
            End If
        Next lIndex
        '// Compare to Temporary TimeBlock Objects 
        For lIndex = 1 To mp_oControl.TempTimeBlocks.Count
            oTimeBlock = DirectCast(mp_oControl.TempTimeBlocks.oCollection.m_oReturnArrayElement(lIndex), clsTimeBlock)
            lRowIndex = mp_oControl.Rows.oCollection.m_lFindIndexByKey(RowKey)
            If oTimeBlock.GenerateConflict = True Then
                '// mp_aoTimeBlocks              S------------------E
                '// interval                     S------------------E
                If (StartDate = oTimeBlock.StartDate Or EndDate = oTimeBlock.EndDate) And (StartDate <> EndDate) Then
                    Return True
                End If
                '// mp_aoTimeBlocks              S------------------E
                '// interval                             S------------------E
                If StartDate > oTimeBlock.StartDate And StartDate < oTimeBlock.EndDate Then
                    Return True
                End If
                '// mp_aoTimeBlocks              S------------------E
                '// interval            S------------------E
                If EndDate > oTimeBlock.StartDate And EndDate < oTimeBlock.EndDate Then
                    Return True
                End If
                '// mp_aoTimeBlocks             S------------------E
                '// interval                 S-------------------------E
                If StartDate < oTimeBlock.StartDate And EndDate > oTimeBlock.EndDate Then
                    Return True
                End If
                '// mp_aoTimeBlocks        S--------------------------E
                '// interval                     S------------------E
                If StartDate > oTimeBlock.StartDate And EndDate < oTimeBlock.EndDate Then
                    Return True
                End If
            End If
        Next lIndex
        Return False
    End Function

    Friend Function PercentageComplete(ByVal X1 As Integer, ByVal X2 As Integer, ByVal X As Integer) As Single
        X2 = X2 - X1
        X = X - X1
        If X = 0 Then
            Return 0
        ElseIf X = X2 Then
            Return 1
        Else
            Return System.Convert.ToSingle(X / X2)
        End If
    End Function

    Private Function NewDateTime(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer) As AGVBE.DateTime
        Dim dtReturn As AGVBE.DateTime
        dtReturn = New AGVBE.DateTime(Year, Month, Day, Hour, Minute, Second)
        Return dtReturn
    End Function

    Public Function RoundDate(ByVal Interval As E_INTERVAL, ByVal Number As Integer, ByVal dtDate As AGVBE.DateTime) As AGVBE.DateTime
        Dim lBuffer As Integer
        Dim lBuffer2 As Integer

        If (Interval = E_INTERVAL.IL_NANOSECOND) Then
            lBuffer = dtDate.Nanosecond
            lBuffer2 = Round(lBuffer, Number)
            Return DateTimeAdd(E_INTERVAL.IL_NANOSECOND, lBuffer2 - lBuffer, dtDate)
        ElseIf (Interval = E_INTERVAL.IL_MICROSECOND) Then
            lBuffer = dtDate.Microsecond
            lBuffer2 = Round(lBuffer, Number)
            Return DateTimeAdd(E_INTERVAL.IL_MICROSECOND, lBuffer2 - lBuffer, dtDate)
        ElseIf (Interval = E_INTERVAL.IL_MILLISECOND) Then
            lBuffer = dtDate.Millisecond
            lBuffer2 = Round(lBuffer, Number)
            Return DateTimeAdd(E_INTERVAL.IL_MILLISECOND, lBuffer2 - lBuffer, dtDate)
        ElseIf (Interval = E_INTERVAL.IL_SECOND) Then
            lBuffer = dtDate.Second
            lBuffer2 = Round(lBuffer, Number)
            dtDate.SecondFractionPart = 0
            Return DateTimeAdd(E_INTERVAL.IL_SECOND, lBuffer2 - lBuffer, dtDate)
        ElseIf (Interval = E_INTERVAL.IL_MINUTE) Then
            Select Case Number
                Case 1
                    dtDate = RoundDate(E_INTERVAL.IL_SECOND, 60, dtDate)
                    lBuffer = dtDate.Second
                    lBuffer2 = Round(lBuffer, 60)
                    dtDate.SecondFractionPart = 0
                    Return DateTimeAdd(E_INTERVAL.IL_SECOND, lBuffer2 - lBuffer, dtDate)
                Case Else
                    dtDate = RoundDate(E_INTERVAL.IL_MINUTE, 1, dtDate)
                    lBuffer = dtDate.Minute
                    lBuffer2 = Round(lBuffer, Number)
                    dtDate.SecondFractionPart = 0
                    Return DateTimeAdd(E_INTERVAL.IL_MINUTE, lBuffer2 - lBuffer, dtDate)
            End Select
        ElseIf (Interval = E_INTERVAL.IL_HOUR) Then
            Select Case Number
                Case 1
                    dtDate = RoundDate(E_INTERVAL.IL_MINUTE, 1, dtDate)
                    lBuffer = dtDate.Minute
                    lBuffer2 = Round(lBuffer, 60)
                    dtDate.SecondFractionPart = 0
                    Return DateTimeAdd(E_INTERVAL.IL_MINUTE, lBuffer2 - lBuffer, dtDate)
                Case Else
                    dtDate = RoundDate(E_INTERVAL.IL_HOUR, 1, dtDate)
                    lBuffer = dtDate.Hour
                    lBuffer2 = Round(lBuffer, Number)
                    dtDate.SecondFractionPart = 0
                    Return DateTimeAdd(E_INTERVAL.IL_HOUR, lBuffer2 - lBuffer, dtDate)
            End Select
        ElseIf (Interval = E_INTERVAL.IL_DAY) Then
            Select Case Number
                Case 1
                    dtDate = RoundDate(E_INTERVAL.IL_HOUR, 1, dtDate)
                    lBuffer = dtDate.Hour
                    lBuffer2 = Round(lBuffer, 24)
                    dtDate.SecondFractionPart = 0
                    Return DateTimeAdd(E_INTERVAL.IL_HOUR, lBuffer2 - lBuffer, dtDate)
                Case Else
                    dtDate = RoundDate(E_INTERVAL.IL_DAY, 1, dtDate)
                    lBuffer = dtDate.Day
                    lBuffer2 = Round(lBuffer, Number)
                    dtDate.SecondFractionPart = 0
                    Return DateTimeAdd(E_INTERVAL.IL_DAY, lBuffer2 - lBuffer, dtDate)
            End Select
        ElseIf (Interval = E_INTERVAL.IL_WEEK) Then
            Select Case Number
                Case 1
                    dtDate = RoundDate(E_INTERVAL.IL_DAY, 1, dtDate)
                    lBuffer = dtDate.DayOfWeek
                    If lBuffer <= 3 Then
                        dtDate = DateTimeAdd(E_INTERVAL.IL_DAY, -(lBuffer - 1), dtDate)
                    ElseIf lBuffer >= 4 Then
                        dtDate = DateTimeAdd(E_INTERVAL.IL_DAY, 8 - lBuffer, dtDate)
                    End If
                    dtDate.SecondFractionPart = 0
                    Return dtDate
                Case Else
                    dtDate = RoundDate(E_INTERVAL.IL_WEEK, 1, dtDate)
                    lBuffer = dtDate.Day
                    lBuffer2 = Round(lBuffer, Number)
                    dtDate.SecondFractionPart = 0
                    Return DateTimeAdd(E_INTERVAL.IL_WEEK, lBuffer2 - lBuffer, dtDate)
            End Select
        ElseIf (Interval = E_INTERVAL.IL_MONTH) Then
            Select Case Number
                Case 1
                    Dim dtNextMonth As AGVBE.DateTime
                    dtDate = RoundDate(E_INTERVAL.IL_DAY, 1, dtDate)
                    lBuffer = dtDate.Day
                    dtDate.SecondFractionPart = 0
                    If lBuffer = 1 Then
                        Return dtDate
                    ElseIf lBuffer >= 15 Then
                        dtNextMonth = DateTimeAdd(E_INTERVAL.IL_MONTH, 1, dtDate)
                        Return NewDateTime(dtNextMonth.Year, dtNextMonth.Month, 1, 0, 0, 0)
                    Else
                        Return NewDateTime(dtDate.Year, dtDate.Month, 1, 0, 0, 0)
                    End If
                Case Else
                    dtDate = RoundDate(E_INTERVAL.IL_MONTH, 1, dtDate)
                    lBuffer = dtDate.Month
                    Dim i As Integer = Round(1, 3)
                    lBuffer2 = Round(lBuffer - 1, Number) + 1
                    dtDate.SecondFractionPart = 0
                    Return DateTimeAdd(E_INTERVAL.IL_MONTH, lBuffer2 - lBuffer, dtDate)
            End Select
        ElseIf (Interval = E_INTERVAL.IL_QUARTER) Then
            dtDate = RoundDate(E_INTERVAL.IL_DAY, 1, dtDate)
            dtDate.SecondFractionPart = 0
            Return RoundDate(E_INTERVAL.IL_MONTH, 3, dtDate)
        ElseIf (Interval = E_INTERVAL.IL_YEAR) Then
            Select Case Number
                Case 1
                    dtDate = RoundDate(E_INTERVAL.IL_MONTH, 1, dtDate)
                    lBuffer = dtDate.Month
                    lBuffer2 = Round(lBuffer, 11) + 1
                    If lBuffer2 = 1 Then
                        Return NewDateTime(dtDate.Year, 1, 1, 0, 0, 0)
                    ElseIf lBuffer2 = 12 Then
                        Return NewDateTime(dtDate.Year + 1, 1, 1, 0, 0, 0)
                    End If
                Case Else
                    dtDate = RoundDate(E_INTERVAL.IL_YEAR, 1, dtDate)
                    lBuffer = dtDate.Year
                    lBuffer2 = Round(lBuffer, Number)
                    dtDate.SecondFractionPart = 0
                    Return DateTimeAdd(E_INTERVAL.IL_YEAR, lBuffer2 - lBuffer, dtDate)
            End Select
        End If
        Return Nothing
    End Function

    Public Function RoundDouble(ByVal dParam As Double) As Integer
        Dim dInt As Double
        Dim dFract As Double
        If (dParam > 0) Then
            dInt = System.Math.Floor(dParam)
            dFract = dParam - dInt
            If (dFract >= 0.5) Then
                Return System.Convert.ToInt32(dInt) + 1
            Else
                Return System.Convert.ToInt32(dInt)
            End If
        ElseIf (dParam < 0) Then
            dInt = System.Math.Ceiling(dParam)
            dFract = dParam - dInt
            If (dFract >= -0.5) Then
                Return System.Convert.ToInt32(dInt)
            Else
                Return System.Convert.ToInt32(dInt) - 1
            End If
        End If
        Return 0
    End Function

    Friend Function Round(ByVal v_lNumberToRound As Integer, ByVal v_lRoundTo As Integer) As Integer
        Dim lRoundToHalf As Integer
        Dim lMultiplier As Integer
        Do While v_lNumberToRound > v_lRoundTo
            v_lNumberToRound = v_lNumberToRound - v_lRoundTo
            lMultiplier = lMultiplier + 1
        Loop
        lRoundToHalf = System.Convert.ToInt32(System.Math.Abs(System.Math.Floor(-(v_lRoundTo / 2))))
        If v_lNumberToRound >= lRoundToHalf Then
            v_lNumberToRound = v_lRoundTo
        Else
            v_lNumberToRound = 0
        End If
        Return (v_lRoundTo * lMultiplier) + v_lNumberToRound
    End Function

    Friend Function RoundUpper(ByVal dParam As Double) As Integer
        Dim iParam As Integer
        Dim iAdd As Integer = 0
        iParam = RoundDouble(dParam)
        If (dParam - iParam) <> 0 Then
            iAdd = 1
        End If
        Return iParam + iAdd
    End Function

    Friend Function mp_DateBlockVisible(ByVal dtIntStart As AGVBE.DateTime, ByVal dtIntEnd As AGVBE.DateTime, ByVal dtBaseDate As AGVBE.DateTime, ByVal yInterval As E_INTERVAL, ByVal lFactor As Integer) As Boolean
        Dim dtStart As AGVBE.DateTime
        Dim dtEnd As AGVBE.DateTime
        If lFactor > 0 Then
            dtStart = dtBaseDate
            dtEnd = DateTimeAdd(yInterval, lFactor, dtBaseDate)
        Else
            dtStart = DateTimeAdd(yInterval, lFactor, dtBaseDate)
            dtEnd = dtBaseDate
        End If
        If dtStart > dtIntStart And dtStart < dtIntEnd Then
            Return True
        End If
        If dtEnd > dtIntStart And dtEnd < dtIntEnd Then
            Return True
        End If
        If dtStart <= dtIntStart And dtEnd >= dtIntEnd Then
            Return True
        End If
        Return False
    End Function

    Friend Sub mp_GenerateTimeBlocks(ByRef aTimeBlocks As ArrayList, ByVal dtIntStart As AGVBE.DateTime, ByVal dtIntEnd As AGVBE.DateTime)
        Dim i As Integer
        Dim oTimeBlock As clsTimeBlock
        For i = 1 To mp_oControl.TimeBlocks.Count
            oTimeBlock = DirectCast(mp_oControl.TimeBlocks.oCollection.m_oReturnArrayElement(i), clsTimeBlock)
            If mp_bIsValidTimeBlock(oTimeBlock) = True Then
                If oTimeBlock.TimeBlockType = E_TIMEBLOCKTYPE.TBT_SINGLE_OCURRENCE Then
                    'If mp_DateBlockVisible(dtIntStart, dtIntEnd, oTimeBlock.StartDate, oTimeBlock.DurationInterval, oTimeBlock.DurationFactor) = True Then
                    mp_AddTimeBlock(aTimeBlocks, oTimeBlock.StartDate, oTimeBlock)
                    'End If
                ElseIf oTimeBlock.TimeBlockType = E_TIMEBLOCKTYPE.TBT_RECURRING Then
                    Dim dtCurrent As AGVBE.DateTime
                    Dim dtBase As AGVBE.DateTime
                    Dim dtTimeLineStart As AGVBE.DateTime
                    Dim dtTimeLineEnd As AGVBE.DateTime
                    Select Case oTimeBlock.RecurringType
                        Case E_RECURRINGTYPE.RCT_DAY
                            dtTimeLineStart = dtIntStart
                            dtTimeLineEnd = dtIntEnd
                            dtTimeLineStart = DateTimeAdd(E_INTERVAL.IL_DAY, -7, dtTimeLineStart)
                            dtTimeLineEnd = DateTimeAdd(E_INTERVAL.IL_DAY, 7, dtTimeLineEnd)
                            dtCurrent = New AGVBE.DateTime(dtTimeLineStart.Year, dtTimeLineStart.Month, dtTimeLineStart.Day, 0, 0, 0)
                            Do While dtCurrent < dtTimeLineEnd
                                dtBase = New AGVBE.DateTime(dtCurrent.Year, dtCurrent.Month, dtCurrent.Day, oTimeBlock.BaseDate.Hour, oTimeBlock.BaseDate.Minute, oTimeBlock.BaseDate.Second)
                                dtCurrent = DateTimeAdd(E_INTERVAL.IL_DAY, 1, dtCurrent)
                                If mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock.DurationInterval, oTimeBlock.DurationFactor) Then
                                    mp_AddTimeBlock(aTimeBlocks, dtBase, oTimeBlock)
                                End If
                            Loop
                        Case E_RECURRINGTYPE.RCT_WEEK
                            dtTimeLineStart = dtIntStart
                            dtTimeLineEnd = dtIntEnd
                            dtTimeLineStart = DateTimeAdd(E_INTERVAL.IL_DAY, -7, dtTimeLineStart)
                            dtTimeLineEnd = DateTimeAdd(E_INTERVAL.IL_DAY, 7, dtTimeLineEnd)
                            dtCurrent = New AGVBE.DateTime(dtTimeLineStart.Year, dtTimeLineStart.Month, dtTimeLineStart.Day, 0, 0, 0)
                            Do While dtCurrent < dtTimeLineEnd
                                dtBase = New AGVBE.DateTime(dtCurrent.Year, dtCurrent.Month, dtCurrent.Day, oTimeBlock.BaseDate.Hour, oTimeBlock.BaseDate.Minute, oTimeBlock.BaseDate.Second)
                                dtCurrent = DateTimeAdd(E_INTERVAL.IL_DAY, 1, dtCurrent)
                                If System.Convert.ToInt32(oTimeBlock.BaseWeekDay) = System.Convert.ToInt32(dtBase.DayOfWeek) Then
                                    If mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock.DurationInterval, oTimeBlock.DurationFactor) Then
                                        mp_AddTimeBlock(aTimeBlocks, dtBase, oTimeBlock)
                                    End If
                                End If
                            Loop
                        Case E_RECURRINGTYPE.RCT_MONTH
                            dtTimeLineStart = dtIntStart
                            dtTimeLineEnd = dtIntEnd
                            dtTimeLineStart = DateTimeAdd(E_INTERVAL.IL_MONTH, -1, dtTimeLineStart)
                            dtTimeLineEnd = DateTimeAdd(E_INTERVAL.IL_MONTH, 1, dtTimeLineEnd)
                            dtCurrent = New AGVBE.DateTime(dtTimeLineStart.Year, dtTimeLineStart.Month, dtTimeLineStart.Day, 0, 0, 0)
                            Do While dtCurrent < dtTimeLineEnd
                                If oTimeBlock.BaseDate.Day = dtCurrent.Day Then
                                    dtBase = New AGVBE.DateTime(dtCurrent.Year, dtCurrent.Month, dtCurrent.Day, oTimeBlock.BaseDate.Hour, oTimeBlock.BaseDate.Minute, oTimeBlock.BaseDate.Second)
                                    dtCurrent = DateTimeAdd(E_INTERVAL.IL_DAY, 1, dtCurrent)
                                    If mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock.DurationInterval, oTimeBlock.DurationFactor) Then
                                        mp_AddTimeBlock(aTimeBlocks, dtBase, oTimeBlock)
                                    End If
                                Else
                                    dtCurrent = DateTimeAdd(E_INTERVAL.IL_DAY, 1, dtCurrent)
                                End If
                            Loop
                        Case E_RECURRINGTYPE.RCT_YEAR
                            dtTimeLineStart = dtIntStart
                            dtTimeLineEnd = dtIntEnd
                            dtTimeLineStart = DateTimeAdd(E_INTERVAL.IL_YEAR, -1, dtTimeLineStart)
                            dtTimeLineEnd = DateTimeAdd(E_INTERVAL.IL_YEAR, 1, dtTimeLineEnd)
                            dtCurrent = New AGVBE.DateTime(dtTimeLineStart.Year, dtTimeLineStart.Month, dtTimeLineStart.Day, 0, 0, 0)
                            Do While dtCurrent < dtTimeLineEnd
                                If oTimeBlock.BaseDate.Month = dtCurrent.Month Then
                                    If oTimeBlock.BaseDate.Day = dtCurrent.Day Then
                                        dtBase = New AGVBE.DateTime(dtCurrent.Year, dtCurrent.Month, dtCurrent.Day, oTimeBlock.BaseDate.Hour, oTimeBlock.BaseDate.Minute, oTimeBlock.BaseDate.Second)
                                        dtCurrent = DateTimeAdd(E_INTERVAL.IL_DAY, 1, dtCurrent)
                                        If mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock.DurationInterval, oTimeBlock.DurationFactor) Then
                                            mp_AddTimeBlock(aTimeBlocks, dtBase, oTimeBlock)
                                        End If
                                    Else
                                        dtCurrent = DateTimeAdd(E_INTERVAL.IL_DAY, 1, dtCurrent)
                                    End If
                                Else
                                    dtCurrent = DateTimeAdd(E_INTERVAL.IL_MONTH, 1, dtCurrent)
                                End If
                            Loop
                    End Select
                End If
            End If
        Next i
        If aTimeBlocks.Count > 0 Then
            mp_QuickSortTB(aTimeBlocks, 0, aTimeBlocks.Count - 1)
            mp_MergeTB(aTimeBlocks)
        End If
    End Sub

    Private Sub mp_AddTimeBlock(ByRef aTimeBlocks As ArrayList, ByVal dtBase As AGVBE.DateTime, ByVal oTimeBlock As clsTimeBlock)
        Dim oTB As S_TIMEBLOCK
        If oTimeBlock.DurationFactor > 0 Then
            oTB.dtStart = dtBase
            oTB.dtEnd = DateTimeAdd(oTimeBlock.DurationInterval, oTimeBlock.DurationFactor, dtBase)
        Else
            oTB.dtEnd = dtBase
            oTB.dtStart = DateTimeAdd(oTimeBlock.DurationInterval, oTimeBlock.DurationFactor, dtBase)
        End If
        aTimeBlocks.Add(oTB)
    End Sub

    Friend Function mp_bIsValidTimeBlock(ByVal oTimeBlock As clsTimeBlock) As Boolean
        If oTimeBlock.DurationFactor = 0 Then
            Return False
        End If
        If oTimeBlock.NonWorking = True Then
            Return True
        End If
        Return False
    End Function

    Private Sub mp_MergeTB(ByRef aTimeBlocks As ArrayList)
        Dim i As Integer
        Dim lStart As Integer = 0
        Dim bFinished As Boolean = False
        While bFinished = False
            For i = lStart To aTimeBlocks.Count - 2
                Dim oTB1 As S_TIMEBLOCK
                Dim oTB2 As S_TIMEBLOCK
                Dim lTB1 As Integer
                Dim lTB2 As Integer
                lTB1 = i
                lTB2 = i + 1
                oTB1 = DirectCast(aTimeBlocks(lTB1), AGVBE.S_TIMEBLOCK)
                oTB2 = DirectCast(aTimeBlocks(lTB2), AGVBE.S_TIMEBLOCK)
                If (oTB2.dtStart > oTB1.dtStart) And (oTB2.dtEnd < oTB1.dtEnd) Then
                    'Case 1
                    'xxxxxxxxxxxxxxxxx           TB1
                    '   xxxxxxxxxxxx             TB2
                    aTimeBlocks.RemoveAt(lTB2)
                    bFinished = False
                    If i >= 1 Then lStart = i - 1
                    Exit For
                ElseIf (oTB2.dtStart = oTB1.dtStart) And (oTB2.dtEnd = oTB1.dtEnd) Then
                    'Case 2
                    'xxxxxxxxxxxxxxxxx           TB1
                    'xxxxxxxxxxxxxxxxx           TB2
                    aTimeBlocks.RemoveAt(lTB2)
                    bFinished = False
                    If i >= 1 Then lStart = i - 1
                    Exit For
                ElseIf (oTB2.dtStart = oTB1.dtStart) And (oTB2.dtEnd > oTB1.dtEnd) Then
                    'Case 3
                    'xxxxxxxxxxxx                TB1
                    'xxxxxxxxxxxxxxxxx           TB2
                    oTB1.dtEnd = oTB2.dtEnd
                    aTimeBlocks(lTB1) = oTB1
                    aTimeBlocks.RemoveAt(lTB2)
                    bFinished = False
                    If i >= 1 Then lStart = i - 1
                    Exit For
                ElseIf (oTB2.dtStart = oTB1.dtStart) And (oTB2.dtEnd < oTB1.dtEnd) Then
                    'Case 4
                    'xxxxxxxxxxxxxxxxx           TB1
                    'xxxxxxxxxxx                 TB2
                    aTimeBlocks.RemoveAt(lTB2)
                    bFinished = False
                    If i >= 1 Then lStart = i - 1
                    Exit For
                ElseIf (oTB2.dtStart > oTB1.dtStart) And (oTB2.dtStart < oTB1.dtEnd) Then
                    'Case 5
                    'xxxxxxxxxxxxxxxxx            TB1
                    '              xxxxxxxxxxxxxx TB2
                    oTB1.dtEnd = oTB2.dtEnd
                    aTimeBlocks(lTB1) = oTB1
                    aTimeBlocks.RemoveAt(lTB2)
                    bFinished = False
                    If i >= 1 Then lStart = i - 1
                    Exit For
                ElseIf (oTB1.dtEnd = oTB2.dtStart) Then
                    'Case 6
                    'xxxxxxxxxxxxxx               TB1
                    '              xxxxxxxxxxxxxx TB2
                    oTB1.dtEnd = oTB2.dtEnd
                    aTimeBlocks(lTB1) = oTB1
                    aTimeBlocks.RemoveAt(lTB2)
                    bFinished = False
                    If i >= 1 Then lStart = i - 1
                    Exit For
                End If
                bFinished = True
            Next
        End While
    End Sub

    Friend Function mp_CheckDuration(ByRef aTimeBlocks As ArrayList, ByVal dtStartDate As AGVBE.DateTime) As Integer
        Dim i As Integer
        Dim lSeconds As Integer = 0
        Dim bInside As Boolean = False
        For i = 0 To aTimeBlocks.Count - 2
            Dim oTB1 As S_TIMEBLOCK
            Dim oTB2 As S_TIMEBLOCK
            oTB1 = DirectCast(aTimeBlocks(i), AGVBE.S_TIMEBLOCK)
            oTB2 = DirectCast(aTimeBlocks(i + 1), AGVBE.S_TIMEBLOCK)
            Dim lSecondDiff As Integer = 0
            If dtStartDate >= oTB1.dtEnd And dtStartDate < oTB2.dtStart Then
                lSecondDiff = DateTimeDiff(E_INTERVAL.IL_SECOND, dtStartDate, oTB2.dtStart)
                bInside = True
            ElseIf bInside = True Then
                lSecondDiff = DateTimeDiff(E_INTERVAL.IL_SECOND, oTB1.dtEnd, oTB2.dtStart)
            End If
            If bInside = True And lSecondDiff <= 0 Then
                mp_oControl.mp_ErrorReport(SYS_ERRORS.CHECK_DURATION_ERROR, "Inconsistent State in Check Duration", "clsMath.mp_CheckDuration")
                Return -1
            End If
            lSeconds = lSeconds + lSecondDiff
        Next
        Return lSeconds
    End Function

    Friend Function mp_GetStartDate(ByRef aTimeBlocks As ArrayList, ByRef bStartDateVerified As Boolean, ByRef dtStartDate As AGVBE.DateTime) As Boolean
        If bStartDateVerified = True Then
            Return True
        End If
        Dim i As Integer
        For i = 0 To aTimeBlocks.Count - 1
            Dim oTB1 As S_TIMEBLOCK
            oTB1 = DirectCast(aTimeBlocks(i), AGVBE.S_TIMEBLOCK)
            If dtStartDate >= oTB1.dtStart And dtStartDate < oTB1.dtEnd Then
                dtStartDate = oTB1.dtEnd
                bStartDateVerified = True
                Return False
            End If
        Next
        bStartDateVerified = True
        Return True
    End Function

    Friend Sub mp_GetEndDate(ByRef aTimeBlocks As ArrayList, ByVal lDurationInSeconds As Integer, ByRef dtStartDate As AGVBE.DateTime, ByRef dtEndDate As AGVBE.DateTime)
        Dim i As Integer
        Dim bInside As Boolean = False
        For i = 0 To aTimeBlocks.Count - 2
            Dim oTB1 As S_TIMEBLOCK
            Dim oTB2 As S_TIMEBLOCK
            oTB1 = DirectCast(aTimeBlocks(i), AGVBE.S_TIMEBLOCK)
            oTB2 = DirectCast(aTimeBlocks(i + 1), AGVBE.S_TIMEBLOCK)
            Dim lSecondDiff As Integer
            If dtStartDate >= oTB1.dtEnd And dtStartDate < oTB2.dtStart And bInside = False Then
                lSecondDiff = DateTimeDiff(E_INTERVAL.IL_SECOND, dtStartDate, oTB2.dtStart)
                bInside = True
                If lDurationInSeconds <= lSecondDiff Then
                    dtEndDate = DateTimeAdd(E_INTERVAL.IL_SECOND, lDurationInSeconds, dtStartDate)
                    Return
                End If
            ElseIf bInside = True Then
                lSecondDiff = DateTimeDiff(E_INTERVAL.IL_SECOND, oTB1.dtEnd, oTB2.dtStart)
            End If
            If bInside = True Then
                lDurationInSeconds = lDurationInSeconds - lSecondDiff
                If lDurationInSeconds <= 0 Then
                    lDurationInSeconds = lDurationInSeconds + lSecondDiff
                    dtEndDate = DateTimeAdd(E_INTERVAL.IL_SECOND, lDurationInSeconds, oTB1.dtEnd)
                    Return
                End If
            End If
        Next
    End Sub

    Friend Sub mp_ValidateStartDate(ByRef aTimeBlocks As ArrayList, ByRef dtStartDate As AGVBE.DateTime)
        Dim i As Integer
        For i = 0 To aTimeBlocks.Count - 1
            Dim oTB1 As S_TIMEBLOCK
            oTB1 = DirectCast(aTimeBlocks(i), AGVBE.S_TIMEBLOCK)
            If dtStartDate >= oTB1.dtStart And dtStartDate < oTB1.dtEnd Then
                dtStartDate = oTB1.dtEnd
                Return
            End If
        Next
    End Sub

    Friend Sub mp_ValidateEndDate(ByRef aTimeBlocks As ArrayList, ByRef dtEndDate As AGVBE.DateTime)
        Dim i As Integer
        For i = 0 To aTimeBlocks.Count - 1
            Dim oTB1 As S_TIMEBLOCK
            oTB1 = DirectCast(aTimeBlocks(i), AGVBE.S_TIMEBLOCK)
            If dtEndDate >= oTB1.dtStart And dtEndDate < oTB1.dtEnd Then
                dtEndDate = oTB1.dtStart
                Return
            End If
        Next
    End Sub

    Friend Sub mp_StandarizeInterval(ByRef dtIntStart As AGVBE.DateTime, ByRef dtIntEnd As AGVBE.DateTime)
        If dtIntStart < dtIntEnd Then
            Return
        End If
        Dim dtIntBuff As AGVBE.DateTime
        dtIntBuff = dtIntStart
        dtIntStart = dtIntEnd
        dtIntEnd = dtIntBuff
    End Sub

    Public Function GetEndDate(ByRef StartDate As AGVBE.DateTime, ByVal DurationInterval As E_INTERVAL, ByVal DurationFactor As Integer) As AGVBE.DateTime
        Dim EndDate As AGVBE.DateTime = New AGVBE.DateTime()
        Dim dtIntStart As AGVBE.DateTime
        Dim dtIntEnd As AGVBE.DateTime
        Dim bFinished As Boolean = False
        Dim bStartDateVerified As Boolean = False
        Dim lIntFactor As Integer = 2
        Dim aTimeBlocks As ArrayList
        Dim lDurationInSeconds As Integer
        Dim lCheckDuration As Integer
        Dim lPass As Integer = 0
        Dim i As Integer
        Dim oTimeBlock As clsTimeBlock
        Dim lSecondsInDay As Integer = 86400
        Dim lSecondsInWeek As Integer = 604800
        Dim lSecondsInMonth As Integer = 2419200
        Dim lSecondsInYear As Integer = 31449600
        Dim lEstimatedDuration As Integer
        If Not (DurationInterval = E_INTERVAL.IL_SECOND Or DurationInterval = E_INTERVAL.IL_MINUTE Or DurationInterval = E_INTERVAL.IL_HOUR Or DurationInterval = E_INTERVAL.IL_DAY) Then
            mp_oControl.mp_ErrorReport(SYS_ERRORS.INVALID_DURATION_INTERVAL, "Interval is invalid for a duration", "clsMath.GetEndDate")
            Return EndDate
        End If
        lDurationInSeconds = mp_oControl.MathLib.mp_GetSeconds(DurationInterval, DurationFactor)
        lEstimatedDuration = lDurationInSeconds
        For i = 1 To mp_oControl.TimeBlocks.Count
            oTimeBlock = DirectCast(mp_oControl.TimeBlocks.oCollection.m_oReturnArrayElement(i), clsTimeBlock)
            If mp_oControl.MathLib.mp_bIsValidTimeBlock(oTimeBlock) Then
                If oTimeBlock.TimeBlockType = E_TIMEBLOCKTYPE.TBT_RECURRING Then
                    Select Case oTimeBlock.RecurringType
                        Case E_RECURRINGTYPE.RCT_DAY
                            lSecondsInDay = lSecondsInDay - mp_oControl.MathLib.mp_GetSeconds(oTimeBlock.DurationInterval, oTimeBlock.DurationFactor)
                        Case E_RECURRINGTYPE.RCT_WEEK
                            lSecondsInWeek = lSecondsInWeek - mp_oControl.MathLib.mp_GetSeconds(oTimeBlock.DurationInterval, oTimeBlock.DurationFactor)
                        Case E_RECURRINGTYPE.RCT_MONTH
                            lSecondsInMonth = lSecondsInMonth - mp_oControl.MathLib.mp_GetSeconds(oTimeBlock.DurationInterval, oTimeBlock.DurationFactor)
                        Case E_RECURRINGTYPE.RCT_YEAR
                            lSecondsInYear = lSecondsInYear - mp_oControl.MathLib.mp_GetSeconds(oTimeBlock.DurationInterval, oTimeBlock.DurationFactor)
                    End Select
                End If
            End If
        Next
        If lDurationInSeconds > 31449600 Then
            lEstimatedDuration = System.Convert.ToInt32(lEstimatedDuration + (System.Math.Ceiling(lDurationInSeconds / lSecondsInYear) * 31449600))
        End If
        If lDurationInSeconds > 2419200 Then
            lEstimatedDuration = System.Convert.ToInt32(lEstimatedDuration + (System.Math.Ceiling(lDurationInSeconds / lSecondsInMonth) * 2419200))
        End If
        If lDurationInSeconds > 604800 Then
            lEstimatedDuration = System.Convert.ToInt32(lEstimatedDuration + (System.Math.Ceiling(lDurationInSeconds / lSecondsInWeek) * 604800))
        End If
        If lDurationInSeconds > 86400 Then
            lEstimatedDuration = System.Convert.ToInt32(lEstimatedDuration + (System.Math.Ceiling(lDurationInSeconds / lSecondsInDay) * 86400))
        End If
        Do While bFinished = False
            lPass = lPass + 1
            dtIntStart = StartDate
            dtIntEnd = mp_oControl.MathLib.DateTimeAdd(E_INTERVAL.IL_SECOND, lEstimatedDuration * lIntFactor, StartDate)
            mp_oControl.MathLib.mp_StandarizeInterval(dtIntStart, dtIntEnd)
            If mp_oControl.TimeBlocks.IntervalType = E_TBINTERVALTYPE.TBIT_AUTOMATIC Then
                aTimeBlocks = New ArrayList()
                mp_oControl.MathLib.mp_GenerateTimeBlocks(aTimeBlocks, dtIntStart, dtIntEnd)
            Else
                aTimeBlocks = mp_oControl.TimeBlocks.mp_aTimeBlocks
            End If
            If aTimeBlocks.Count = 0 Then
                EndDate = mp_oControl.MathLib.DateTimeAdd(DurationInterval, DurationFactor, StartDate)
                Return EndDate
            Else
                If mp_oControl.MathLib.mp_GetStartDate(aTimeBlocks, bStartDateVerified, StartDate) = True Then
                    lCheckDuration = mp_oControl.MathLib.mp_CheckDuration(aTimeBlocks, StartDate)
                    If lCheckDuration < lDurationInSeconds Then
                        If lCheckDuration = 0 Then
                            lCheckDuration = lDurationInSeconds
                        End If
                        lIntFactor = System.Convert.ToInt32((System.Math.Ceiling(lDurationInSeconds / lCheckDuration) * 2) + lIntFactor)
                    Else
                        mp_oControl.MathLib.mp_GetEndDate(aTimeBlocks, lDurationInSeconds, StartDate, EndDate)
                        bFinished = True
                    End If
                End If
            End If
        Loop
        Return EndDate
    End Function

    Friend Sub mp_GetTimeBlocks(ByRef aTimeBlocks As ArrayList, ByRef dtStartDate As AGVBE.DateTime, ByRef dtEndDate As AGVBE.DateTime)
        If mp_oControl.TimeBlocks.IntervalType = E_TBINTERVALTYPE.TBIT_AUTOMATIC Then
            mp_GenerateTimeBlocks(aTimeBlocks, dtStartDate, dtEndDate)
        Else
            aTimeBlocks = mp_oControl.TimeBlocks.mp_aTimeBlocks
        End If
    End Sub

    Public Function CalculateDuration(ByRef dtStartDate As AGVBE.DateTime, ByRef dtEndDate As AGVBE.DateTime, ByVal DurationInterval As E_INTERVAL) As Integer
        Dim aTimeBlocks As ArrayList = New ArrayList()
        Dim lDurationInSeconds As Integer
        Dim lReturn As Integer = 0
        mp_StandarizeInterval(dtStartDate, dtEndDate)
        If Not (DurationInterval = E_INTERVAL.IL_SECOND Or DurationInterval = E_INTERVAL.IL_MINUTE Or DurationInterval = E_INTERVAL.IL_HOUR Or DurationInterval = E_INTERVAL.IL_DAY) Then
            mp_oControl.mp_ErrorReport(SYS_ERRORS.INVALID_DURATION_INTERVAL, "Interval is invalid for a duration", "clsMath.CalculateDuration")
            Return -1
        End If
        mp_GetTimeBlocks(aTimeBlocks, dtStartDate, dtEndDate)
        If aTimeBlocks.Count = 0 Then
            Return DateTimeDiff(DurationInterval, dtStartDate, dtEndDate)
        Else
            mp_ValidateStartDate(aTimeBlocks, dtStartDate)
            mp_ValidateEndDate(aTimeBlocks, dtEndDate)
            lDurationInSeconds = mp_GetDuration(aTimeBlocks, dtStartDate, dtEndDate)
            Select Case DurationInterval
                Case E_INTERVAL.IL_SECOND
                    lReturn = lDurationInSeconds
                Case E_INTERVAL.IL_MINUTE
                    lReturn = System.Convert.ToInt32(System.Math.Floor(lDurationInSeconds / 60))
                Case E_INTERVAL.IL_HOUR
                    lReturn = System.Convert.ToInt32(System.Math.Floor(lDurationInSeconds / 3600))
                Case E_INTERVAL.IL_DAY
                    lReturn = System.Convert.ToInt32(System.Math.Floor(lDurationInSeconds / 86400))
            End Select
            Return lReturn
        End If
    End Function

    Private Function mp_GetDuration(ByRef aTimeBlocks As ArrayList, ByVal dtStartDate As AGVBE.DateTime, ByVal dtEndDate As AGVBE.DateTime) As Integer
        Dim i As Integer
        Dim bInside As Boolean = False
        Dim lReturn As Integer = 0
        For i = 0 To aTimeBlocks.Count - 2
            Dim oTB1 As S_TIMEBLOCK
            Dim oTB2 As S_TIMEBLOCK
            oTB1 = DirectCast(aTimeBlocks(i), AGVBE.S_TIMEBLOCK)
            oTB2 = DirectCast(aTimeBlocks(i + 1), AGVBE.S_TIMEBLOCK)
            If dtStartDate >= oTB1.dtEnd And dtStartDate <= oTB2.dtStart And dtEndDate >= oTB1.dtEnd And dtEndDate <= oTB2.dtStart Then
                lReturn = DateTimeDiff(E_INTERVAL.IL_SECOND, dtStartDate, dtEndDate)
            ElseIf dtStartDate >= oTB1.dtEnd And dtStartDate <= oTB2.dtStart Then
                lReturn = lReturn + DateTimeDiff(E_INTERVAL.IL_SECOND, dtStartDate, oTB2.dtStart)
                bInside = True
            ElseIf dtEndDate >= oTB1.dtEnd And dtEndDate <= oTB2.dtStart And bInside = True Then
                lReturn = lReturn + DateTimeDiff(E_INTERVAL.IL_SECOND, oTB1.dtEnd, dtEndDate)
                Exit For
            ElseIf bInside = True Then
                lReturn = lReturn + DateTimeDiff(E_INTERVAL.IL_SECOND, oTB1.dtEnd, oTB2.dtStart)
            End If
        Next
        Return lReturn
    End Function

    Friend Sub mp_DumpTB(ByVal oTB As S_TIMEBLOCK)
        Debug.WriteLine("StartDate: " & oTB.dtStart.ToString("yyyy/MM/dd HH:mm:ss"))
        Debug.WriteLine("EndDate: " & oTB.dtEnd.ToString("yyyy/MM/dd HH:mm:ss"))
    End Sub

    Friend Sub mp_DumpTimeBlocks(ByRef aTimeBlocks As ArrayList, ByVal sCaption As String)
        Debug.WriteLine(sCaption & " *************************** Dumping TimeBlocks:")
        Dim i As Integer
        For i = 0 To aTimeBlocks.Count - 1
            Dim oTB As S_TIMEBLOCK
            oTB = DirectCast(aTimeBlocks(i), AGVBE.S_TIMEBLOCK)
            Debug.WriteLine(i.ToString() & ":")
            Debug.WriteLine("StartDate: " & oTB.dtStart.ToString("yyyy/MM/dd HH:mm:ss"))
            Debug.WriteLine("EndDate: " & oTB.dtEnd.ToString("yyyy/MM/dd HH:mm:ss"))
        Next
    End Sub

    Friend Function mp_GetSeconds(ByVal yInterval As E_INTERVAL, ByVal lFactor As Integer) As Integer
        If lFactor < 0 Then
            lFactor = lFactor * -1
        End If
        Select Case yInterval
            Case E_INTERVAL.IL_SECOND
                Return lFactor
            Case E_INTERVAL.IL_MINUTE
                Return lFactor * 60
            Case E_INTERVAL.IL_HOUR
                Return lFactor * 3600
            Case E_INTERVAL.IL_DAY
                Return lFactor * 86400
        End Select
        Return -1
    End Function

    Friend Sub mp_QuickSortTB(ByRef aTimeBlocks As ArrayList, ByVal StartIndex As Integer, ByVal EndIndex As Integer)
        ' StartIndex = subscript of beginning of array
        ' EndIndex = subscript of end of array

        Dim MiddleIndex As Integer
        If StartIndex < EndIndex Then
            MiddleIndex = mp_QSPartitionTB(aTimeBlocks, StartIndex, EndIndex)
            mp_QuickSortTB(aTimeBlocks, StartIndex, MiddleIndex) ' sort first section
            mp_QuickSortTB(aTimeBlocks, MiddleIndex + 1, EndIndex) ' sort second section
        End If
        Return
    End Sub

    Friend Function mp_QSPartitionTB(ByRef aTimeBlocks As ArrayList, ByVal StartIndex As Integer, ByVal EndIndex As Integer) As Integer
        Dim oX As S_TIMEBLOCK = DirectCast(aTimeBlocks(StartIndex), S_TIMEBLOCK)
        Dim x As AGVBE.DateTime = oX.dtStart
        Dim i As Integer = StartIndex - 1
        Dim j As Integer = EndIndex + 1
        Dim temp As S_TIMEBLOCK
        Do
            Do
                j -= 1
                oX = DirectCast(aTimeBlocks(j), S_TIMEBLOCK)
            Loop While x < oX.dtStart 'Change to > for descending
            Do
                i += 1
                oX = DirectCast(aTimeBlocks(i), S_TIMEBLOCK)
            Loop While x > oX.dtStart 'Change to < for descending
            If i < j Then
                temp = DirectCast(aTimeBlocks(i), S_TIMEBLOCK)
                aTimeBlocks(i) = aTimeBlocks(j)
                aTimeBlocks(j) = temp

            End If
        Loop While i < j
        Return j ' returns middle subscript  
    End Function

    Friend Function GetTierIndex(ByVal Number As Integer) As String
        Dim sNumber As String
        Dim lLastDigit As Integer
        sNumber = System.Convert.ToString(Number)
        lLastDigit = System.Convert.ToInt32(sNumber.Substring(sNumber.Length() - 1, 1))
        If lLastDigit = 0 Then
            Return "10"
        Else
            Return lLastDigit.ToString()
        End If
    End Function

End Class
