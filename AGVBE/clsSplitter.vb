Option Explicit On

Public Class clsSplitter

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_lPosition As Integer
    Private mp_yAppearance As E_BORDERSTYLE
    Private mp_lWidth As Integer
    Private mp_yType As E_SPLITTERTYPE
    Private mp_aColors As ArrayList
    Private mp_sStyleIndex As String
    Private mp_oStyle As clsStyle
    Private mp_lOffset As Integer

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_lPosition = 125
        mp_aColors = New ArrayList
        mp_yAppearance = E_BORDERSTYLE.TLB_3D
        mp_yType = E_SPLITTERTYPE.SA_APPEARANCE
        Me.Width = 6
        mp_sStyleIndex = "DS_SPLITTER"
        mp_oStyle = mp_oControl.Styles.FItem("DS_SPLITTER")
        mp_lOffset = -1
    End Sub

    Public Sub SetColor(ByVal Index As Integer, ByVal oColor As Color)
        Index = Index - 1
        If mp_yType <> E_SPLITTERTYPE.SA_USERDEFINED Then
            mp_oControl.mp_ErrorReport(SYS_ERRORS.SPLITTER_INVALIDOP, "Invalid Operation", "ActiveGanttVBECtl.clsSplitter.SetColor")
            Return
        End If
        If Index < 0 Then
            mp_oControl.mp_ErrorReport(SYS_ERRORS.SPLITTER_INVALID_INDEX, "Invalid Index", "ActiveGanttVBECtl.clsSplitter.SetColor")
            Return
        End If
        If Index > mp_aColors.Count - 1 Then
            mp_oControl.mp_ErrorReport(SYS_ERRORS.SPLITTER_INVALID_INDEX, "Invalid Index", "ActiveGanttVBECtl.clsSplitter.SetColor")
            Return
        End If
        mp_aColors.Item(Index) = oColor
    End Sub

    Public Function GetColor(ByVal Index As Integer) As Color
        Index = Index - 1
        If mp_yType <> E_SPLITTERTYPE.SA_USERDEFINED Then
            mp_oControl.mp_ErrorReport(SYS_ERRORS.SPLITTER_INVALIDOP, "Invalid Operation", "ActiveGanttVBECtl.clsSplitter.GetColor")
            Return Nothing
        End If
        If Index < 0 Then
            mp_oControl.mp_ErrorReport(SYS_ERRORS.SPLITTER_INVALID_INDEX, "Invalid Index", "ActiveGanttVBECtl.clsSplitter.GetColor")
            Return Nothing
        End If
        If Index > mp_aColors.Count - 1 Then
            mp_oControl.mp_ErrorReport(SYS_ERRORS.SPLITTER_INVALID_INDEX, "Invalid Index", "ActiveGanttVBECtl.clsSplitter.GetColor")
            Return Nothing
        End If
        Return DirectCast(mp_aColors.Item(Index), Color)
    End Function

    Public Property Type() As E_SPLITTERTYPE
        Get
            Return mp_yType
        End Get
        Set(ByVal Value As E_SPLITTERTYPE)
            mp_yType = Value
        End Set
    End Property

    
    Public Property Appearance() As E_BORDERSTYLE
        Get
            Return mp_yAppearance
        End Get
        Set(ByVal Value As E_BORDERSTYLE)
            mp_yAppearance = Value
        End Set
    End Property

    Public Property Width() As Integer
        Get
            If mp_yType = E_SPLITTERTYPE.SA_APPEARANCE Then
                Return 6
            Else
                Return mp_lWidth
            End If
        End Get
        Set(ByVal Value As Integer)
            If mp_yType = E_SPLITTERTYPE.SA_APPEARANCE Then
                mp_lWidth = 6
            Else
                If Value < 0 Then
                    mp_oControl.mp_ErrorReport(SYS_ERRORS.SPLITTER_INVALID_WIDTH, "Invalid Width Value", "ActiveGanttVBECtl.clsSplitter.Width")
                    Return
                End If
                mp_lWidth = Value
            End If
            Dim i As Integer
            mp_aColors.Clear()
            For i = 0 To mp_lWidth - 1
                mp_aColors.Add(Colors.White)
            Next
        End Set
    End Property

    
    Public Property Position() As Integer
        Get
            Return mp_lPosition
        End Get
        Set(ByVal Value As Integer)
            If (Value <= 0) Then
                Exit Property
            End If
            mp_lPosition = Value
            If (mp_lPosition > (mp_oControl.Columns.Width + mp_lOffset)) Then
                mp_lPosition = mp_oControl.Columns.Width + mp_lOffset
                mp_oControl.HorizontalScrollBar.Value = 0
            End If
        End Set
    End Property

    Friend ReadOnly Property Left() As Integer
        Get
            If (mp_oControl.Columns.Count <> 0) Then
                Return Position + mp_oControl.mt_BorderThickness - 1
            Else
                If (mp_oControl.f_UserMode = True) Then
                    Return 0
                Else
                    Return 125 + mp_oControl.mt_BorderThickness - 1
                End If
            End If
        End Get
    End Property

    Friend ReadOnly Property Right() As Integer
        Get
            If (mp_oControl.Columns.Count <> 0) Then
                Return Position + mp_oControl.mt_BorderThickness + Me.Width()
            Else
                If (mp_oControl.f_UserMode = True) Then
                    Return mp_oControl.mt_BorderThickness
                Else
                    Return 125 + mp_oControl.mt_BorderThickness + Me.Width()
                End If
            End If
        End Get
    End Property

    Public Property StyleIndex() As String
        Get
            If mp_sStyleIndex = "DS_SPLITTER" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_SPLITTER"
            mp_sStyleIndex = Value
            mp_oStyle = mp_oControl.Styles.FItem(Value)
        End Set
    End Property

    Public ReadOnly Property Style() As clsStyle
        Get
            Return mp_oStyle
        End Get
    End Property

    Friend Sub Draw()
        If mp_oControl.Columns.Count = 0 And mp_oControl.f_UserMode = True Then
            Return
        End If
        mp_oControl.clsG.ClipRegion(Left() + 1, 0, Left() + Me.Width + 1, mp_oControl.mt_BottomMargin, True)
        If mp_yType = E_SPLITTERTYPE.SA_STYLE Then
            mp_oControl.clsG.mp_DrawItem(Left() + 1, 0, Left() + Me.Width() + 1, mp_oControl.clsG.Height, "", "", False, Nothing, 0, 0, Me.Style)
        Else
            Dim i As Integer
            If mp_yType = E_SPLITTERTYPE.SA_APPEARANCE Then
                mp_aColors.Clear()
                For i = 0 To 5
                    mp_aColors.Add(Colors.White)
                Next
                Select Case mp_yAppearance
                    Case E_BORDERSTYLE.TLB_3D
                        mp_aColors.Item(0) = Color.FromArgb(255, 192, 192, 192)
                        mp_aColors.Item(1) = Colors.White
                        mp_aColors.Item(2) = Color.FromArgb(255, 192, 192, 192)
                        mp_aColors.Item(3) = Color.FromArgb(255, 192, 192, 192)
                        mp_aColors.Item(4) = Color.FromArgb(255, 64, 64, 64)
                        mp_aColors.Item(5) = Color.FromArgb(255, 66, 66, 66)
                    Case E_BORDERSTYLE.TLB_NONE
                        mp_aColors.Item(0) = Color.FromArgb(255, 192, 192, 192)
                        mp_aColors.Item(1) = Color.FromArgb(255, 192, 192, 192)
                        mp_aColors.Item(2) = Color.FromArgb(255, 192, 192, 192)
                        mp_aColors.Item(3) = Color.FromArgb(255, 192, 192, 192)
                        mp_aColors.Item(4) = Color.FromArgb(255, 192, 192, 192)
                        mp_aColors.Item(5) = Color.FromArgb(255, 192, 192, 192)
                    Case E_BORDERSTYLE.TLB_SINGLE
                        mp_aColors.Item(0) = Colors.Black
                        mp_aColors.Item(1) = Color.FromArgb(255, 192, 192, 192)
                        mp_aColors.Item(2) = Color.FromArgb(255, 192, 192, 192)
                        mp_aColors.Item(3) = Color.FromArgb(255, 192, 192, 192)
                        mp_aColors.Item(4) = Color.FromArgb(255, 192, 192, 192)
                        mp_aColors.Item(5) = Colors.Black
                End Select
            End If
            For i = 0 To mp_aColors.Count - 1
                mp_oControl.clsG.DrawLine(Left() + i + 1, 0, Left() + i + 1, mp_oControl.clsG.Height, GRE_LINETYPE.LT_NORMAL, DirectCast(mp_aColors.Item(i), Color), GRE_LINEDRAWSTYLE.LDS_SOLID)
            Next
        End If
    End Sub

    Friend Sub f_AdjustPosition()
        Dim lWidth As Integer
        lWidth = mp_oControl.Columns.Width + mp_lOffset
        If lWidth < (mp_oControl.clsG.Width - 25) Then
            If mp_lPosition < lWidth Then
                mp_lPosition = lWidth
            End If
        End If
        If mp_lPosition > lWidth Then
            mp_lPosition = lWidth
            mp_oControl.HorizontalScrollBar.Value = 0
        End If
    End Sub

    Public Property Offset() As Integer
        Get
            Return mp_lOffset
        End Get
        Set(ByVal value As Integer)
            mp_lOffset = value
        End Set
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "Splitter")
        oXML.InitializeWriter()
        oXML.WriteProperty("Appearance", mp_yAppearance)
        oXML.WriteProperty("Offset", mp_lOffset)
        oXML.WriteProperty("Position", mp_lPosition)
        oXML.WriteProperty("Type", mp_yType)
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        oXML.WriteProperty("Width", mp_lWidth)
        If mp_yType = E_SPLITTERTYPE.SA_USERDEFINED Then
            oXML.WriteProperty("ColorCount", mp_aColors.Count)
            Dim i As Integer
            For i = 0 To mp_aColors.Count - 1
                oXML.WriteProperty("Color" & i.ToString(), DirectCast(mp_aColors.Item(i), System.Windows.Media.Color))
            Next
        Else
            oXML.WriteProperty("ColorCount", 0)
        End If
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "Splitter")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("Appearance", mp_yAppearance)
        oXML.ReadProperty("Offset", mp_lOffset)
        oXML.ReadProperty("Position", mp_lPosition)
        Position = mp_lPosition
        oXML.ReadProperty("Type", mp_yType)
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        StyleIndex = mp_sStyleIndex
        oXML.ReadProperty("Width", mp_lWidth)
        mp_aColors.Clear()
        Dim lColorCount As Integer
        oXML.ReadProperty("ColorCount", lColorCount)
        If lColorCount > 0 Then
            Dim i As Integer
            For i = 0 To lColorCount - 1
                Dim oColor As Color
                oXML.ReadProperty("Color" & i.ToString(), oColor)
                mp_aColors.Add(oColor)
            Next
        End If
        StyleIndex = mp_sStyleIndex
    End Sub

End Class
