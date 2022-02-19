Option Explicit On 

Public Class clsTimeBlocks

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_oCollection As clsCollectionBase
    Friend mp_aTimeBlocks As ArrayList
    Private mp_dtIntervalStart As AGVBE.DateTime
    Private mp_dtIntervalEnd As AGVBE.DateTime
    Private mp_yIntervalType As E_TBINTERVALTYPE

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_oCollection = New clsCollectionBase(Value, "TimeBlock")
        mp_dtIntervalStart = New AGVBE.DateTime()
        mp_dtIntervalEnd = New AGVBE.DateTime()
        mp_yIntervalType = E_TBINTERVALTYPE.TBIT_AUTOMATIC
    End Sub

    
    Public ReadOnly Property Count() As Integer
        Get
            Return mp_oCollection.m_lCount()
        End Get
    End Property

    
    Public Function Item(ByVal Index As String) As clsTimeBlock
        Return DirectCast(mp_oCollection.m_oItem(Index, SYS_ERRORS.TIMEBLOCKS_ITEM_1, SYS_ERRORS.TIMEBLOCKS_ITEM_2, SYS_ERRORS.TIMEBLOCKS_ITEM_3, SYS_ERRORS.TIMEBLOCKS_ITEM_4), clsTimeBlock)
    End Function

    Friend ReadOnly Property oCollection() As clsCollectionBase
        Get
            Return mp_oCollection
        End Get
    End Property

    Public Function Add(ByVal Key As String) As clsTimeBlock
        mp_oCollection.AddMode = True
        Dim oTimeBlock As New clsTimeBlock(mp_oControl)
        Key = mp_oControl.StrLib.StrTrim(Key)
        oTimeBlock.Key = Key
        mp_oCollection.m_Add(oTimeBlock, Key, SYS_ERRORS.TIMEBLOCKS_ADD_1, SYS_ERRORS.TIMEBLOCKS_ADD_2, False, SYS_ERRORS.TIMEBLOCKS_ADD_3)
        Return oTimeBlock
    End Function

    
    Public Sub Clear()
        mp_oCollection.m_Clear()
    End Sub

    
    Public Sub Remove(ByVal Index As String)
        mp_oCollection.m_Remove(Index, SYS_ERRORS.TIMEBLOCKS_REMOVE_1, SYS_ERRORS.TIMEBLOCKS_REMOVE_2, SYS_ERRORS.TIMEBLOCKS_REMOVE_3, SYS_ERRORS.TIMEBLOCKS_REMOVE_4)
    End Sub

    Friend Sub CreateTemporaryTimeBlocks()

        Dim lIndex As Integer = 0
        mp_oControl.TempTimeBlocks().Clear()

        For lIndex = 1 To Count
            Dim oTimeBlock As clsTimeBlock = Nothing
            Dim oTempTimeBlock As clsTimeBlock = Nothing
            Dim dtTimeLineStart As AGVBE.DateTime = New AGVBE.DateTime()
            Dim dtTimeLineEnd As AGVBE.DateTime = New AGVBE.DateTime()
            Dim dtCurrent As AGVBE.DateTime = New AGVBE.DateTime()
            Dim dtStartBuff As AGVBE.DateTime = New AGVBE.DateTime()
            Dim dtEndBuff As AGVBE.DateTime = New AGVBE.DateTime()
            Dim dtBase As AGVBE.DateTime
            Dim dtStartDate As AGVBE.DateTime = New AGVBE.DateTime()
            Dim dtEndDate As AGVBE.DateTime = New AGVBE.DateTime()
            oTimeBlock = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsTimeBlock)
            If oTimeBlock.TimeBlockType = E_TIMEBLOCKTYPE.TBT_RECURRING Then
                dtTimeLineStart = mp_oControl.CurrentViewObject.TimeLine.StartDate
                dtTimeLineEnd = mp_oControl.CurrentViewObject.TimeLine.EndDate
                Select Case oTimeBlock.RecurringType
                    Case E_RECURRINGTYPE.RCT_DAY
                        dtTimeLineStart = mp_oControl.MathLib.DateTimeAdd(E_INTERVAL.IL_DAY, -1, dtTimeLineStart)
                        dtCurrent = New AGVBE.DateTime(dtTimeLineStart.Year, dtTimeLineStart.Month, dtTimeLineStart.Day, 0, 0, 0)
                        Do While dtCurrent < dtTimeLineEnd
                            dtBase = New AGVBE.DateTime(dtCurrent.Year, dtCurrent.Month, dtCurrent.Day, oTimeBlock.BaseDate.Hour, oTimeBlock.BaseDate.Minute, oTimeBlock.BaseDate.Second)
                            dtCurrent = mp_oControl.MathLib.DateTimeAdd(E_INTERVAL.IL_DAY, 1, dtCurrent)
                            If mp_oControl.MathLib.mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock.DurationInterval, oTimeBlock.DurationFactor) Then
                                oTempTimeBlock = mp_oControl.TempTimeBlocks().Add("")
                                oTempTimeBlock.BaseDate = dtBase
                                oTempTimeBlock.DurationInterval = oTimeBlock.DurationInterval
                                oTempTimeBlock.DurationFactor = oTimeBlock.DurationFactor
                                CopyTimeBlock(oTempTimeBlock, oTimeBlock)
                            End If
                        Loop
                    Case E_RECURRINGTYPE.RCT_WEEK
                        dtTimeLineStart = mp_oControl.MathLib.DateTimeAdd(E_INTERVAL.IL_DAY, -7, dtTimeLineStart)
                        dtCurrent = New AGVBE.DateTime(dtTimeLineStart.Year, dtTimeLineStart.Month, dtTimeLineStart.Day, 0, 0, 0)
                        Do While dtCurrent < dtTimeLineEnd
                            dtBase = New AGVBE.DateTime(dtCurrent.Year, dtCurrent.Month, dtCurrent.Day, oTimeBlock.BaseDate.Hour, oTimeBlock.BaseDate.Minute, oTimeBlock.BaseDate.Second)
                            dtCurrent = mp_oControl.MathLib.DateTimeAdd(E_INTERVAL.IL_DAY, 1, dtCurrent)
                            If System.Convert.ToInt32(oTimeBlock.BaseWeekDay) = System.Convert.ToInt32(dtBase.DayOfWeek) Then
                                If mp_oControl.MathLib.mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock.DurationInterval, oTimeBlock.DurationFactor) Then
                                    oTempTimeBlock = mp_oControl.TempTimeBlocks().Add("")
                                    oTempTimeBlock.BaseDate = dtBase
                                    oTempTimeBlock.DurationInterval = oTimeBlock.DurationInterval
                                    oTempTimeBlock.DurationFactor = oTimeBlock.DurationFactor
                                    CopyTimeBlock(oTempTimeBlock, oTimeBlock)
                                End If
                            End If
                        Loop
                    Case E_RECURRINGTYPE.RCT_MONTH
                        dtTimeLineStart = mp_oControl.MathLib.DateTimeAdd(E_INTERVAL.IL_MONTH, -1, dtTimeLineStart)
                        dtCurrent = New AGVBE.DateTime(dtTimeLineStart.Year, dtTimeLineStart.Month, dtTimeLineStart.Day, 0, 0, 0)
                        Do While dtCurrent < dtTimeLineEnd
                            If oTimeBlock.BaseDate.Day = dtCurrent.Day Then
                                dtBase = New AGVBE.DateTime(dtCurrent.Year, dtCurrent.Month, dtCurrent.Day, oTimeBlock.BaseDate.Hour, oTimeBlock.BaseDate.Minute, oTimeBlock.BaseDate.Second)
                                dtCurrent = mp_oControl.MathLib.DateTimeAdd(E_INTERVAL.IL_DAY, 1, dtCurrent)
                                If mp_oControl.MathLib.mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock.DurationInterval, oTimeBlock.DurationFactor) Then
                                    oTempTimeBlock = mp_oControl.TempTimeBlocks().Add("")
                                    oTempTimeBlock.BaseDate = dtBase
                                    oTempTimeBlock.DurationInterval = oTimeBlock.DurationInterval
                                    oTempTimeBlock.DurationFactor = oTimeBlock.DurationFactor
                                    CopyTimeBlock(oTempTimeBlock, oTimeBlock)
                                End If
                            Else
                                dtCurrent = mp_oControl.MathLib.DateTimeAdd(E_INTERVAL.IL_DAY, 1, dtCurrent)
                            End If
                        Loop
                    Case E_RECURRINGTYPE.RCT_YEAR
                        dtTimeLineStart = mp_oControl.MathLib.DateTimeAdd(E_INTERVAL.IL_YEAR, -1, dtTimeLineStart)
                        dtCurrent = New AGVBE.DateTime(dtTimeLineStart.Year, dtTimeLineStart.Month, dtTimeLineStart.Day, 0, 0, 0)
                        Do While dtCurrent < dtTimeLineEnd
                            If oTimeBlock.BaseDate.Month = dtCurrent.Month Then
                                If oTimeBlock.BaseDate.Day = dtCurrent.Day Then
                                    dtBase = New AGVBE.DateTime(dtCurrent.Year, dtCurrent.Month, dtCurrent.Day, oTimeBlock.BaseDate.Hour, oTimeBlock.BaseDate.Minute, oTimeBlock.BaseDate.Second)
                                    dtCurrent = mp_oControl.MathLib.DateTimeAdd(E_INTERVAL.IL_DAY, 1, dtCurrent)
                                    If mp_oControl.MathLib.mp_DateBlockVisible(dtTimeLineStart, dtTimeLineEnd, dtBase, oTimeBlock.DurationInterval, oTimeBlock.DurationFactor) Then
                                        oTempTimeBlock = mp_oControl.TempTimeBlocks().Add("")
                                        oTempTimeBlock.BaseDate = dtBase
                                        oTempTimeBlock.DurationInterval = oTimeBlock.DurationInterval
                                        oTempTimeBlock.DurationFactor = oTimeBlock.DurationFactor
                                        CopyTimeBlock(oTempTimeBlock, oTimeBlock)
                                    End If
                                Else
                                    dtCurrent = mp_oControl.MathLib.DateTimeAdd(E_INTERVAL.IL_DAY, 1, dtCurrent)
                                End If
                            Else
                                dtCurrent = mp_oControl.MathLib.DateTimeAdd(E_INTERVAL.IL_MONTH, 1, dtCurrent)
                            End If
                        Loop
                End Select
            End If
        Next
    End Sub

    Private Sub CopyTimeBlock(ByVal oDestination As clsTimeBlock, ByVal oOriginal As clsTimeBlock)
        oDestination.TimeBlockType = E_TIMEBLOCKTYPE.TBT_SINGLE_OCURRENCE
        oDestination.StyleIndex = oOriginal.StyleIndex
        oDestination.GenerateConflict = oOriginal.GenerateConflict
        oDestination.Tag = oOriginal.Tag
        oDestination.NonWorking = oOriginal.NonWorking
        oDestination.f_Visible = oOriginal.f_Visible
    End Sub

    Friend Sub Draw()
        DrawClass(Me)
        DrawClass(mp_oControl.TempTimeBlocks)
    End Sub

    Friend Sub DrawClass(ByVal oTimeBlocks As clsTimeBlocks)
        Dim lIndex As Integer
        Dim oTimeBlock As clsTimeBlock = Nothing
        If oTimeBlocks.Count = 0 Then
            Return
        End If
        mp_oControl.clsG.ClipRegion(mp_oControl.Splitter.Right, mp_oControl.CurrentViewObject.ClientArea.Top, mp_oControl.mt_RightMargin, mp_oControl.CurrentViewObject.ClientArea.Bottom, True)
        For lIndex = 1 To oTimeBlocks.Count
            oTimeBlock = DirectCast(oTimeBlocks.mp_oCollection.m_oReturnArrayElement(lIndex), clsTimeBlock)
            If oTimeBlock.Visible = True Then
                mp_oControl.DrawEventArgs.Clear()
                mp_oControl.DrawEventArgs.CustomDraw = False
                mp_oControl.DrawEventArgs.EventTarget = E_EVENTTARGET.EVT_TIMEBLOCK
                mp_oControl.DrawEventArgs.ObjectIndex = lIndex
                mp_oControl.DrawEventArgs.ParentObjectIndex = 0
                'mp_oControl.DrawEventArgs.Graphics = mp_oControl.clsG.oGraphics
                mp_oControl.FireDraw()
                If mp_oControl.DrawEventArgs.CustomDraw = False Then
                    If (oTimeBlock.Right - oTimeBlock.Left) >= 1 Then
                        mp_oControl.clsG.mp_DrawItem(oTimeBlock.Left, oTimeBlock.Top, oTimeBlock.Right, oTimeBlock.Bottom, "", "", False, Nothing, oTimeBlock.LeftTrim, oTimeBlock.RightTrim, oTimeBlock.Style)
                    End If
                End If
            End If
        Next lIndex
    End Sub

    Public Property IntervalStart() As AGVBE.DateTime
        Get
            Return mp_dtIntervalStart
        End Get
        Set(ByVal value As AGVBE.DateTime)
            mp_dtIntervalStart = value
        End Set
    End Property

    Public Property IntervalEnd() As AGVBE.DateTime
        Get
            Return mp_dtIntervalEnd
        End Get
        Set(ByVal value As AGVBE.DateTime)
            mp_dtIntervalEnd = value
        End Set
    End Property

    Public Property IntervalType() As E_TBINTERVALTYPE
        Get
            Return mp_yIntervalType
        End Get
        Set(ByVal value As E_TBINTERVALTYPE)
            mp_yIntervalType = value
        End Set
    End Property

    Public Sub CalculateInterval()
        If mp_yIntervalType = E_TBINTERVALTYPE.TBIT_AUTOMATIC Then
            Return
        End If
        mp_aTimeBlocks = New ArrayList()
        mp_oControl.MathLib.mp_GenerateTimeBlocks(mp_aTimeBlocks, mp_dtIntervalStart, mp_dtIntervalEnd)
    End Sub

    Public Function CP_GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "CP_TimeBlocks")
        oXML.InitializeWriter()
        oXML.WriteProperty("IntervalStart", mp_dtIntervalStart)
        oXML.WriteProperty("IntervalEnd", mp_dtIntervalEnd)
        oXML.WriteProperty("IntervalType", mp_yIntervalType)
        Return oXML.GetXML()
    End Function

    Public Sub CP_SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "CP_TimeBlocks")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("IntervalStart", mp_dtIntervalStart)
        oXML.ReadProperty("IntervalEnd", mp_dtIntervalEnd)
        oXML.ReadProperty("IntervalType", mp_yIntervalType)
    End Sub

    Public Function GetXML() As String
        Dim lIndex As Integer
        Dim oTimeBlock As clsTimeBlock = Nothing
        Dim oXML As New clsXML(mp_oControl, "TimeBlocks")
        oXML.InitializeWriter()
        For lIndex = 1 To Count
            oTimeBlock = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsTimeBlock)
            oXML.WriteObject(oTimeBlock.GetXML)
        Next lIndex
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim lIndex As Integer
        Dim oXML As New clsXML(mp_oControl, "TimeBlocks")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        mp_oCollection.m_Clear()
        For lIndex = 1 To oXML.ReadCollectionCount
            Dim oTimeBlock As New clsTimeBlock(mp_oControl)
            oTimeBlock.SetXML(oXML.ReadCollectionObject(lIndex))
            mp_oCollection.AddMode = True
            mp_oCollection.m_Add(oTimeBlock, oTimeBlock.Key, SYS_ERRORS.TIMEBLOCKS_ADD_1, SYS_ERRORS.TIMEBLOCKS_ADD_2, False, SYS_ERRORS.TIMEBLOCKS_ADD_3)
        Next lIndex
    End Sub

End Class
