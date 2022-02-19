Option Explicit On 

Public Class clsTasks

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_oCollection As clsCollectionBase

    Dim mp_lLoadIndex As Integer
    Private mp_oTempCollection As ArrayList
    Private mp_oTempDictionary As clsDictionary

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_oCollection = New clsCollectionBase(Value, "Task")
    End Sub

    
    Public ReadOnly Property Count() As Integer
        Get
            Return mp_oCollection.m_lCount
        End Get
    End Property

    
    Public Function Item(ByVal Index As String) As clsTask
        Return DirectCast(mp_oCollection.m_oItem(Index, SYS_ERRORS.TASKS_ITEM_1, SYS_ERRORS.TASKS_ITEM_2, SYS_ERRORS.TASKS_ITEM_3, SYS_ERRORS.TASKS_ITEM_4), clsTask)
    End Function

    Friend ReadOnly Property oCollection() As clsCollectionBase
        Get
            Return mp_oCollection
        End Get
    End Property

    
    Public Function Add(ByVal Text As String, ByVal RowKey As String, ByVal StartDate As AGVBE.DateTime, ByVal EndDate As AGVBE.DateTime, Optional ByVal Key As String = "", Optional ByVal StyleIndex As String = "", Optional ByVal LayerIndex As String = "0") As clsTask
        mp_oCollection.AddMode = True
        Dim oTask As New clsTask(mp_oControl)
        Key = mp_oControl.StrLib.StrTrim(Key)
        Text = mp_oControl.StrLib.StrTrim(Text)
        RowKey = mp_oControl.StrLib.StrTrim(RowKey)
        oTask.Text = Text
        oTask.RowKey = RowKey
        oTask.StartDate = StartDate
        oTask.EndDate = EndDate
        oTask.Key = Key
        oTask.StyleIndex = StyleIndex
        oTask.LayerIndex = LayerIndex
        mp_oCollection.m_Add(oTask, Key, SYS_ERRORS.TASKS_ADD_1, SYS_ERRORS.TASKS_ADD_2, False, SYS_ERRORS.TASKS_ADD_3)
        Return oTask
    End Function

    Public Function DAdd(ByVal RowKey As String, ByVal StartDate As AGVBE.DateTime, ByVal DurationInterval As E_INTERVAL, ByVal DurationFactor As Integer, Optional ByVal Text As String = "", Optional ByVal Key As String = "", Optional ByVal StyleIndex As String = "", Optional ByVal LayerIndex As String = "0") As clsTask
        mp_oCollection.AddMode = True
        Dim oTask As New clsTask(mp_oControl)
        oTask.TaskType = E_TASKTYPE.TT_DURATION
        Key = mp_oControl.StrLib.StrTrim(Key)
        Text = mp_oControl.StrLib.StrTrim(Text)
        RowKey = mp_oControl.StrLib.StrTrim(RowKey)
        oTask.Text = Text
        oTask.RowKey = RowKey
        oTask.StartDate = StartDate
        oTask.DurationInterval = DurationInterval
        oTask.DurationFactor = DurationFactor
        oTask.Key = Key
        oTask.StyleIndex = StyleIndex
        oTask.LayerIndex = LayerIndex
        mp_oCollection.m_Add(oTask, Key, SYS_ERRORS.TASKS_ADD_1, SYS_ERRORS.TASKS_ADD_2, False, SYS_ERRORS.TASKS_ADD_3)
        Return oTask
    End Function

    Public Sub Clear()
        mp_oControl.Predecessors.Clear()
        mp_oControl.Percentages.Clear()
        mp_oCollection.m_Clear()
        mp_oControl.SelectedTaskIndex = 0
    End Sub

    Public Sub Remove(ByVal Index As String)
        Dim lIndex As Integer
        Dim sRIndex As String = ""
        Dim sRKey As String = ""
        Dim oPredecessor As clsPredecessor
        mp_oCollection.m_GetKeyAndIndex(Index, sRKey, sRIndex)
        mp_oControl.Percentages.oCollection.m_CollectionRemoveWhere("TaskKey", sRKey, sRIndex)
        If sRKey.Length > 0 Then
            For lIndex = mp_oControl.Predecessors.Count To 1 Step -1
                oPredecessor = DirectCast(mp_oControl.Predecessors.oCollection.m_oReturnArrayElement(lIndex), clsPredecessor)
                If oPredecessor.SuccessorKey = sRKey Or oPredecessor.PredecessorKey = sRKey Then
                    mp_oControl.Predecessors.Remove(lIndex.ToString())
                End If
            Next lIndex
        End If
        mp_oCollection.m_Remove(Index, SYS_ERRORS.TASKS_REMOVE_1, SYS_ERRORS.TASKS_REMOVE_2, SYS_ERRORS.TASKS_REMOVE_3, SYS_ERRORS.TASKS_REMOVE_4)
        mp_oControl.SelectedTaskIndex = 0
    End Sub

    
    Public Sub Sort(ByVal PropertyName As String, ByVal Descending As Boolean, ByVal SortType As E_SORTTYPE, Optional ByVal StartIndex As Integer = -1, Optional ByVal EndIndex As Integer = -1)
        If StartIndex = -1 Then
            StartIndex = 1
        End If
        If EndIndex = -1 Then
            EndIndex = Count
        End If
        If Count = 0 Then Return
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

    Friend Sub Draw()
        Dim lIndex As Integer
        Dim oTask As clsTask
        If Count = 0 Then
            Return
        End If
        For lIndex = 1 To Count
            oTask = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsTask)
            If oTask.Visible = True And oTask.InsideVisibleTimeLineArea = True Then
                mp_oControl.DrawEventArgs.Clear()
                mp_oControl.DrawEventArgs.CustomDraw = False
                mp_oControl.DrawEventArgs.EventTarget = E_EVENTTARGET.EVT_TASK
                mp_oControl.DrawEventArgs.ObjectIndex = lIndex
                mp_oControl.DrawEventArgs.ParentObjectIndex = 0
                'mp_oControl.DrawEventArgs.Graphics = mp_oControl.clsG.oGraphics
                mp_oControl.FireDraw()
                If mp_oControl.DrawEventArgs.CustomDraw = False Then
                    If oTask.Type = E_OBJECTTYPE.OT_MILESTONE Then
                        mp_oControl.clsG.mp_DrawItemI(oTask, "", lIndex = mp_oControl.SelectedTaskIndex, mp_GetTaskStyle(oTask))
                    Else
                        mp_oControl.clsG.mp_DrawItem(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom, "", oTask.Text, (lIndex = mp_oControl.SelectedTaskIndex), oTask.Image, oTask.LeftTrim, oTask.RightTrim, mp_GetTaskStyle(oTask))
                    End If
                    If oTask.Text.Length > 0 Then
                        oTask.mp_lTextLeft = mp_oControl.clsG.mp_oTextFinalLayout.Left
                        oTask.mp_lTextTop = mp_oControl.clsG.mp_oTextFinalLayout.Top
                        oTask.mp_lTextRight = mp_oControl.clsG.mp_oTextFinalLayout.Left + mp_oControl.clsG.mp_oTextFinalLayout.Width - 1
                        oTask.mp_lTextBottom = mp_oControl.clsG.mp_oTextFinalLayout.Top + mp_oControl.clsG.mp_oTextFinalLayout.Height - 1
                    Else
                        If oTask.Style.TextPlacement = E_TEXTPLACEMENT.SCP_EXTERIORPLACEMENT Then
                            If oTask.Style.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT Then
                                oTask.mp_lTextLeft = oTask.Left - mp_oControl.mp_lStrWidth("WWWWW", oTask.Style.Font) - oTask.Style.TextXMargin
                                oTask.mp_lTextRight = oTask.Left - oTask.Style.TextXMargin + 1
                            End If
                            If oTask.Style.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT Then
                                oTask.mp_lTextLeft = oTask.Right + oTask.Style.TextXMargin
                                oTask.mp_lTextRight = oTask.Right + mp_oControl.mp_lStrWidth("WWWWW", oTask.Style.Font) + oTask.Style.TextXMargin + 1
                            End If
                            If oTask.Style.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_CENTER Then
                                oTask.mp_lTextLeft = oTask.Left
                                oTask.mp_lTextRight = oTask.Right + 1
                            End If
                            If oTask.Style.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP Then
                                oTask.mp_lTextTop = oTask.Top - mp_oControl.mp_lStrHeight("WWWWW", oTask.Style.Font) - oTask.Style.TextYMargin
                                oTask.mp_lTextBottom = oTask.Top - oTask.Style.TextYMargin + 3
                            End If
                            If oTask.Style.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM Then
                                oTask.mp_lTextTop = oTask.Bottom + oTask.Style.TextYMargin
                                oTask.mp_lTextBottom = oTask.Bottom + mp_oControl.mp_lStrHeight("WWWWW", oTask.Style.Font) + oTask.Style.TextYMargin + 3
                            End If
                            If oTask.Style.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_CENTER Then
                                oTask.mp_lTextTop = oTask.Top
                                oTask.mp_lTextBottom = oTask.Bottom + 3
                            End If
                        Else
                            oTask.mp_lTextLeft = oTask.Left
                            oTask.mp_lTextTop = oTask.Top
                            oTask.mp_lTextRight = oTask.Right
                            oTask.mp_lTextBottom = oTask.Bottom
                        End If
                    End If
                End If
            End If
        Next lIndex
    End Sub

    Private Function mp_GetTaskStyle(ByVal oTask As clsTask) As clsStyle
        Dim oStyle As clsStyle
        If oTask.mp_bWarning = True Then
            oStyle = oTask.WarningStyle
        Else
            oStyle = oTask.Style
        End If
        Return oStyle
    End Function

    
    Public Function GetXML() As String
        Dim lIndex As Integer
        Dim oTask As clsTask
        Dim oXML As New clsXML(mp_oControl, "Tasks")
        oXML.InitializeWriter()
        For lIndex = 1 To Count
            oTask = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsTask)
            oXML.WriteObject(oTask.GetXML)
        Next lIndex
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim lIndex As Integer
        Dim oXML As New clsXML(mp_oControl, "Tasks")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        mp_oCollection.m_Clear()
        For lIndex = 1 To oXML.ReadCollectionCount
            Dim oTask As New clsTask(mp_oControl)
            oTask.SetXML(oXML.ReadCollectionObject(lIndex))
            mp_oCollection.AddMode = True
            mp_oCollection.m_Add(oTask, oTask.Key, SYS_ERRORS.TASKS_ADD_1, SYS_ERRORS.TASKS_ADD_2, False, SYS_ERRORS.TASKS_ADD_3)
            oTask = Nothing
        Next lIndex
    End Sub

    Public Sub BeginLoad(ByVal Preserve As Boolean)
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

    Public Function Load(ByVal sRowKey As String, ByVal sKey As String) As clsTask
        Dim oTask As New clsTask(mp_oControl)
        If sKey.Length > 0 Then
            oTask.Key = sKey
        End If
        oTask.Index = mp_lLoadIndex
        If mp_oControl.Rows.mp_oTempCollection.Count > 0 Then
            Dim lRowIndex As Object
            lRowIndex = mp_oControl.Rows.mp_oTempDictionary.Item(sRowKey)
            If Not lRowIndex Is Nothing Then
                oTask.mp_oRow = DirectCast(mp_oControl.Rows.mp_oTempCollection(System.Convert.ToInt32(lRowIndex) - 1), clsRow)
            Else
                lRowIndex = mp_oControl.Rows.oCollection.mp_oKeys.Item(sRowKey)
                oTask.mp_oRow = DirectCast(mp_oControl.Rows.oCollection.mp_aoCollection(System.Convert.ToInt32(lRowIndex) - 1), clsRow)
            End If
        Else
            Dim lRowIndex As Integer
            lRowIndex = mp_oControl.Rows.oCollection.mp_oKeys.Item(sRowKey)
            oTask.mp_oRow = DirectCast(mp_oControl.Rows.oCollection.mp_aoCollection(lRowIndex - 1), clsRow)
        End If
        mp_oTempCollection.Add(oTask)
        mp_oTempDictionary.Add(mp_lLoadIndex, sKey)
        mp_lLoadIndex = mp_lLoadIndex + 1
        Return oTask
    End Function

    Public Sub EndLoad()
        mp_oCollection.mp_aoCollection = mp_oTempCollection
        mp_oCollection.mp_oKeys = mp_oTempDictionary
        mp_oTempCollection = Nothing
        mp_oTempDictionary = Nothing
    End Sub

End Class
