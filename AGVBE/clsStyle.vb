Option Explicit On

Public Class clsStyle
    Inherits clsItemBase

    Private mp_oControl As ActiveGanttVBECtl
    Public TaskStyle As clsTaskStyle
    Public MilestoneStyle As clsMilestoneStyle
    Public PredecessorStyle As clsPredecessorStyle
    Public TextFlags As clsTextFlags
    Public CustomBorderStyle As clsCustomBorderStyle
    Public ScrollBarStyle As clsScrollBarStyle
    Public SelectionRectangleStyle As clsSelectionRectangleStyle
    Public ButtonBorderStyle As clsButtonBorderStyle
    Private mp_bTextVisible As Boolean
    Private mp_yBackgroundMode As GRE_BACKGROUNDMODE
    Private mp_bClipText As Boolean
    Private mp_bDrawTextInVisibleArea As Boolean
    Private mp_bUseMask As Boolean
    Private mp_clrBackColor As Color
    Private mp_clrForeColor As Color
    Private mp_clrBorderColor As Color
    Private mp_clrEndGradientColor As Color
    Private mp_clrStartGradientColor As Color
    Private mp_oFont As Font
    Private mp_lPatternFactor As Integer
    Private mp_lTextXMargin As Integer
    Private mp_lTextYMargin As Integer
    Private mp_lOffsetBottom As Integer
    Private mp_lOffsetTop As Integer
    Private mp_lImageXMargin As Integer
    Private mp_lImageYMargin As Integer
    Private mp_sTag As String
    Private mp_lBorderWidth As Integer
    Private mp_yAppearance As E_STYLEAPPEARANCE
    Private mp_yPattern As GRE_PATTERN
    Private mp_yBorderStyle As GRE_BORDERSTYLE
    Private mp_yButtonStyle As GRE_BUTTONSTYLE
    Private mp_yTextAlignmentHorizontal As GRE_HORIZONTALALIGNMENT
    Private mp_yTextAlignmentVertical As GRE_VERTICALALIGNMENT
    Private mp_yTextPlacement As E_TEXTPLACEMENT
    Private mp_yGradientFillMode As GRE_GRADIENTFILLMODE
    Private mp_yImageAlignmentHorizontal As GRE_HORIZONTALALIGNMENT
    Private mp_yImageAlignmentVertical As GRE_VERTICALALIGNMENT
    Private mp_yPlacement As E_PLACEMENT
    Private mp_yFillMode As GRE_FILLMODE
    Private mp_yHatchStyle As GRE_HATCHSTYLE
    Private mp_clrHatchBackColor As Color
    Private mp_clrHatchForeColor As Color
    Private mp_clrTextEditBackColor As Color
    Private mp_clrTextEditForeColor As Color

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        TaskStyle = New clsTaskStyle(mp_oControl)
        MilestoneStyle = New clsMilestoneStyle(mp_oControl)
        PredecessorStyle = New clsPredecessorStyle(mp_oControl)
        TextFlags = New clsTextFlags(mp_oControl)
        CustomBorderStyle = New clsCustomBorderStyle(mp_oControl)
        ScrollBarStyle = New clsScrollBarStyle(mp_oControl)
        SelectionRectangleStyle = New clsSelectionRectangleStyle(mp_oControl)
        ButtonBorderStyle = New clsButtonBorderStyle(mp_oControl)
        mp_bTextVisible = True
        mp_bClipText = True
        mp_bDrawTextInVisibleArea = False
        mp_bUseMask = True
        mp_clrBackColor = Color.FromArgb(255, 192, 192, 192)
        mp_clrBorderColor = Colors.Black
        mp_clrEndGradientColor = Colors.Black
        mp_clrForeColor = Colors.Black
        mp_clrStartGradientColor = Colors.Black
        mp_oFont = New Font("Tahoma", 8, E_FONTSIZEUNITS.FSU_POINTS)
        mp_lPatternFactor = 5
        mp_lTextXMargin = 0
        mp_lTextYMargin = 0
        mp_lOffsetBottom = 10
        mp_lOffsetTop = 10
        mp_lImageXMargin = 3
        mp_lImageYMargin = 3
        mp_sTag = ""
        mp_lBorderWidth = 1
        mp_yAppearance = E_STYLEAPPEARANCE.SA_RAISED
        mp_yPattern = GRE_PATTERN.FP_DOWNWARDDIAGONAL
        mp_yBorderStyle = GRE_BORDERSTYLE.SBR_NONE
        mp_yButtonStyle = GRE_BUTTONSTYLE.BT_NORMALWINDOWS
        mp_yTextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_CENTER
        mp_yTextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_CENTER
        mp_yTextPlacement = E_TEXTPLACEMENT.SCP_OBJECTEXTENTSPLACEMENT
        mp_yGradientFillMode = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL
        mp_yImageAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_CENTER
        mp_yImageAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_CENTER
        mp_yPlacement = E_PLACEMENT.PLC_ROWEXTENTSPLACEMENT
        mp_yFillMode = GRE_FILLMODE.FM_COMPLETELYFILLED
        mp_yBackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        mp_yHatchStyle = GRE_HATCHSTYLE.HS_HORIZONTAL
        mp_clrHatchBackColor = Colors.White
        mp_clrHatchForeColor = Colors.Black
        mp_clrTextEditBackColor = Colors.White
        mp_clrTextEditForeColor = Colors.Black
    End Sub

    Public Property TextEditBackColor() As Color
        Get
            Return mp_clrTextEditBackColor
        End Get
        Set(ByVal value As Color)
            mp_clrTextEditBackColor = value
        End Set
    End Property

    Public Property TextEditForeColor() As Color
        Get
            Return mp_clrTextEditForeColor
        End Get
        Set(ByVal value As Color)
            mp_clrTextEditForeColor = value
        End Set
    End Property

    Public Property HatchBackColor() As Color
        Get
            Return mp_clrHatchBackColor
        End Get
        Set(ByVal Value As Color)
            mp_clrHatchBackColor = Value
        End Set
    End Property

    
    Public Property HatchForeColor() As Color
        Get
            Return mp_clrHatchForeColor
        End Get
        Set(ByVal Value As Color)
            mp_clrHatchForeColor = Value
        End Set
    End Property

    
    Public Property Key() As String
        Get
            Return mp_sKey
        End Get
        Set(ByVal Value As String)
            mp_oControl.Styles.oCollection.mp_SetKey(mp_sKey, Value, SYS_ERRORS.STYLES_SET_KEY)
        End Set
    End Property

    
    Public Property Appearance() As E_STYLEAPPEARANCE
        Get
            Return mp_yAppearance
        End Get
        Set(ByVal Value As E_STYLEAPPEARANCE)
            mp_yAppearance = Value
        End Set
    End Property

    
    Public Property BackColor() As Color
        Get
            Return mp_clrBackColor
        End Get
        Set(ByVal Value As Color)
            mp_clrBackColor = Value
        End Set
    End Property

    
    Public Property Pattern() As GRE_PATTERN
        Get
            Return mp_yPattern
        End Get
        Set(ByVal Value As GRE_PATTERN)
            mp_yPattern = Value
        End Set
    End Property

    
    Public Property BorderColor() As Color
        Get
            Return mp_clrBorderColor
        End Get
        Set(ByVal Value As Color)
            mp_clrBorderColor = Value
        End Set
    End Property

    
    Public Property BorderStyle() As GRE_BORDERSTYLE
        Get
            Return mp_yBorderStyle
        End Get
        Set(ByVal Value As GRE_BORDERSTYLE)
            mp_yBorderStyle = Value
        End Set
    End Property

    
    Public Property ButtonStyle() As GRE_BUTTONSTYLE
        Get
            Return mp_yButtonStyle
        End Get
        Set(ByVal Value As GRE_BUTTONSTYLE)
            mp_yButtonStyle = Value
        End Set
    End Property

    
    Public Property TextAlignmentHorizontal() As GRE_HORIZONTALALIGNMENT
        Get
            Return mp_yTextAlignmentHorizontal
        End Get
        Set(ByVal Value As GRE_HORIZONTALALIGNMENT)
            mp_yTextAlignmentHorizontal = Value
        End Set
    End Property

    
    Public Property TextAlignmentVertical() As GRE_VERTICALALIGNMENT
        Get
            Return mp_yTextAlignmentVertical
        End Get
        Set(ByVal Value As GRE_VERTICALALIGNMENT)
            mp_yTextAlignmentVertical = Value
        End Set
    End Property

    
    Public Property TextVisible() As Boolean
        Get
            Return mp_bTextVisible
        End Get
        Set(ByVal Value As Boolean)
            mp_bTextVisible = Value
        End Set
    End Property

    
    Public Property TextXMargin() As Integer
        Get
            Return mp_lTextXMargin
        End Get
        Set(ByVal Value As Integer)
            mp_lTextXMargin = Value
        End Set
    End Property

    
    Public Property TextYMargin() As Integer
        Get
            Return mp_lTextYMargin
        End Get
        Set(ByVal Value As Integer)
            mp_lTextYMargin = Value
        End Set
    End Property

    
    Public Property ClipText() As Boolean
        Get
            Return mp_bClipText
        End Get
        Set(ByVal Value As Boolean)
            mp_bClipText = Value
        End Set
    End Property

    
    Public Property Font() As Font
        Get
            Return mp_oFont
        End Get
        Set(ByVal Value As Font)
            mp_oFont = Value
        End Set
    End Property

    
    Public Property ForeColor() As Color
        Get
            Return mp_clrForeColor
        End Get
        Set(ByVal Value As Color)
            mp_clrForeColor = Value
        End Set
    End Property

    
    Public Property OffsetBottom() As Integer
        Get
            Return mp_lOffsetBottom
        End Get
        Set(ByVal Value As Integer)
            mp_lOffsetBottom = Value
        End Set
    End Property

    
    Public Property OffsetTop() As Integer
        Get
            Return mp_lOffsetTop
        End Get
        Set(ByVal Value As Integer)
            mp_lOffsetTop = Value
        End Set
    End Property

    
    Public Property ImageAlignmentHorizontal() As GRE_HORIZONTALALIGNMENT
        Get
            Return mp_yImageAlignmentHorizontal
        End Get
        Set(ByVal Value As GRE_HORIZONTALALIGNMENT)
            mp_yImageAlignmentHorizontal = Value
        End Set
    End Property

    
    Public Property ImageAlignmentVertical() As GRE_VERTICALALIGNMENT
        Get
            Return mp_yImageAlignmentVertical
        End Get
        Set(ByVal Value As GRE_VERTICALALIGNMENT)
            mp_yImageAlignmentVertical = Value
        End Set
    End Property

    
    Public Property ImageXMargin() As Integer
        Get
            Return mp_lImageXMargin
        End Get
        Set(ByVal Value As Integer)
            mp_lImageXMargin = Value
        End Set
    End Property

    
    Public Property ImageYMargin() As Integer
        Get
            Return mp_lImageYMargin
        End Get
        Set(ByVal Value As Integer)
            mp_lImageYMargin = Value
        End Set
    End Property

    
    Public Property Placement() As E_PLACEMENT
        Get
            Return mp_yPlacement
        End Get
        Set(ByVal Value As E_PLACEMENT)
            mp_yPlacement = Value
        End Set
    End Property
    
    Public Property UseMask() As Boolean
        Get
            Return mp_bUseMask
        End Get
        Set(ByVal Value As Boolean)
            mp_bUseMask = Value
        End Set
    End Property

    
    Public Property TextPlacement() As E_TEXTPLACEMENT
        Get
            Return mp_yTextPlacement
        End Get
        Set(ByVal Value As E_TEXTPLACEMENT)
            mp_yTextPlacement = Value
        End Set
    End Property

    
    Public Property PatternFactor() As Integer
        Get
            Return mp_lPatternFactor
        End Get
        Set(ByVal Value As Integer)
            mp_lPatternFactor = Value
        End Set
    End Property

    
    Public Property HatchStyle() As GRE_HATCHSTYLE
        Get
            Return mp_yHatchStyle
        End Get
        Set(ByVal Value As GRE_HATCHSTYLE)
            mp_yHatchStyle = Value
        End Set
    End Property

    
    Public Property StartGradientColor() As Color
        Get
            Return mp_clrStartGradientColor
        End Get
        Set(ByVal Value As Color)
            mp_clrStartGradientColor = Value
        End Set
    End Property

    
    Public Property EndGradientColor() As Color
        Get
            Return mp_clrEndGradientColor
        End Get
        Set(ByVal Value As Color)
            mp_clrEndGradientColor = Value
        End Set
    End Property

    
    Public Property GradientFillMode() As GRE_GRADIENTFILLMODE
        Get
            Return mp_yGradientFillMode
        End Get
        Set(ByVal Value As GRE_GRADIENTFILLMODE)
            mp_yGradientFillMode = Value
        End Set
    End Property

    
    Public Property FillMode() As GRE_FILLMODE
        Get
            Return mp_yFillMode
        End Get
        Set(ByVal Value As GRE_FILLMODE)
            mp_yFillMode = Value
        End Set
    End Property

    
    Public Property BackgroundMode() As GRE_BACKGROUNDMODE
        Get
            Return mp_yBackgroundMode
        End Get
        Set(ByVal Value As GRE_BACKGROUNDMODE)
            mp_yBackgroundMode = Value
        End Set
    End Property

    
    Public Property DrawTextInVisibleArea() As Boolean
        Get
            Return mp_bDrawTextInVisibleArea
        End Get
        Set(ByVal Value As Boolean)
            mp_bDrawTextInVisibleArea = Value
        End Set
    End Property

    
    Public Property Tag() As String
        Get
            Return mp_sTag
        End Get
        Set(ByVal Value As String)
            mp_sTag = Value
        End Set
    End Property

    Public Property BorderWidth() As Integer
        Get
            Return mp_lBorderWidth
        End Get
        Set(ByVal Value As Integer)
            mp_lBorderWidth = Value
        End Set
    End Property

    
    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "Style")
        oXML.InitializeWriter()
        oXML.WriteProperty("Appearance", mp_yAppearance)
        oXML.WriteProperty("BackColor", mp_clrBackColor)
        oXML.WriteProperty("BackgroundMode", mp_yBackgroundMode)
        oXML.WriteProperty("BorderColor", mp_clrBorderColor)
        oXML.WriteProperty("BorderStyle", mp_yBorderStyle)
        oXML.WriteProperty("ButtonStyle", mp_yButtonStyle)
        oXML.WriteProperty("ClipText", mp_bClipText)
        oXML.WriteProperty("DrawTextInVisibleArea", mp_bDrawTextInVisibleArea)
        oXML.WriteProperty("EndGradientColor", mp_clrEndGradientColor)
        oXML.WriteProperty("FillMode", mp_yFillMode)
        oXML.WriteProperty("Font", mp_oFont)
        oXML.WriteProperty("ForeColor", mp_clrForeColor)
        oXML.WriteProperty("GradientFillMode", mp_yGradientFillMode)
        oXML.WriteProperty("HatchBackColor", mp_clrHatchBackColor)
        oXML.WriteProperty("HatchForeColor", mp_clrHatchForeColor)
        oXML.WriteProperty("HatchStyle", mp_yHatchStyle)
        oXML.WriteProperty("ImageAlignmentHorizontal", mp_yImageAlignmentHorizontal)
        oXML.WriteProperty("ImageAlignmentVertical", mp_yImageAlignmentVertical)
        oXML.WriteProperty("ImageXMargin", mp_lImageXMargin)
        oXML.WriteProperty("ImageYMargin", mp_lImageYMargin)
        oXML.WriteProperty("Key", mp_sKey)
        oXML.WriteProperty("OffsetBottom", mp_lOffsetBottom)
        oXML.WriteProperty("OffsetTop", mp_lOffsetTop)
        oXML.WriteProperty("Pattern", mp_yPattern)
        oXML.WriteProperty("PatternFactor", mp_lPatternFactor)
        oXML.WriteProperty("Placement", mp_yPlacement)
        oXML.WriteProperty("StartGradientColor", mp_clrStartGradientColor)
        oXML.WriteProperty("Tag", mp_sTag)
        oXML.WriteProperty("BorderWidth", mp_lBorderWidth)
        oXML.WriteProperty("TextAlignmentHorizontal", mp_yTextAlignmentHorizontal)
        oXML.WriteProperty("TextAlignmentVertical", mp_yTextAlignmentVertical)
        oXML.WriteProperty("TextPlacement", mp_yTextPlacement)
        oXML.WriteProperty("TextVisible", mp_bTextVisible)
        oXML.WriteProperty("TextXMargin", mp_lTextXMargin)
        oXML.WriteProperty("TextYMargin", mp_lTextYMargin)
        oXML.WriteProperty("UseMask", mp_bUseMask)
        oXML.WriteProperty("TextEditBackColor", mp_clrTextEditBackColor)
        oXML.WriteProperty("TextEditForeColor", mp_clrTextEditForeColor)
        oXML.WriteObject(CustomBorderStyle.GetXML())
        oXML.WriteObject(MilestoneStyle.GetXML())
        oXML.WriteObject(PredecessorStyle.GetXML())
        oXML.WriteObject(TaskStyle.GetXML())
        oXML.WriteObject(TextFlags.GetXML())
        oXML.WriteObject(ScrollBarStyle.GetXML())
        oXML.WriteObject(SelectionRectangleStyle.GetXML())
        oXML.WriteObject(ButtonBorderStyle.GetXML())
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "Style")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("Appearance", mp_yAppearance)
        oXML.ReadProperty("BackColor", mp_clrBackColor)
        oXML.ReadProperty("BackgroundMode", mp_yBackgroundMode)
        oXML.ReadProperty("BorderColor", mp_clrBorderColor)
        oXML.ReadProperty("BorderStyle", mp_yBorderStyle)
        oXML.ReadProperty("ButtonStyle", mp_yButtonStyle)
        oXML.ReadProperty("ClipText", mp_bClipText)
        oXML.ReadProperty("DrawTextInVisibleArea", mp_bDrawTextInVisibleArea)
        oXML.ReadProperty("EndGradientColor", mp_clrEndGradientColor)
        oXML.ReadProperty("FillMode", mp_yFillMode)
        oXML.ReadProperty("Font", mp_oFont)
        oXML.ReadProperty("ForeColor", mp_clrForeColor)
        oXML.ReadProperty("GradientFillMode", mp_yGradientFillMode)
        oXML.ReadProperty("HatchBackColor", mp_clrHatchBackColor)
        oXML.ReadProperty("HatchForeColor", mp_clrHatchForeColor)
        oXML.ReadProperty("HatchStyle", mp_yHatchStyle)
        oXML.ReadProperty("ImageAlignmentHorizontal", mp_yImageAlignmentHorizontal)
        oXML.ReadProperty("ImageAlignmentVertical", mp_yImageAlignmentVertical)
        oXML.ReadProperty("ImageXMargin", mp_lImageXMargin)
        oXML.ReadProperty("ImageYMargin", mp_lImageYMargin)
        oXML.ReadProperty("Key", mp_sKey)
        oXML.ReadProperty("OffsetBottom", mp_lOffsetBottom)
        oXML.ReadProperty("OffsetTop", mp_lOffsetTop)
        oXML.ReadProperty("Pattern", mp_yPattern)
        oXML.ReadProperty("PatternFactor", mp_lPatternFactor)
        oXML.ReadProperty("Placement", mp_yPlacement)
        oXML.ReadProperty("StartGradientColor", mp_clrStartGradientColor)
        oXML.ReadProperty("Tag", mp_sTag)
        oXML.ReadProperty("BorderWidth", mp_lBorderWidth)
        oXML.ReadProperty("TextAlignmentHorizontal", mp_yTextAlignmentHorizontal)
        oXML.ReadProperty("TextAlignmentVertical", mp_yTextAlignmentVertical)
        oXML.ReadProperty("TextPlacement", mp_yTextPlacement)
        oXML.ReadProperty("TextVisible", mp_bTextVisible)
        oXML.ReadProperty("TextXMargin", mp_lTextXMargin)
        oXML.ReadProperty("TextYMargin", mp_lTextYMargin)
        oXML.ReadProperty("UseMask", mp_bUseMask)
        oXML.ReadProperty("TextEditBackColor", mp_clrTextEditBackColor)
        oXML.ReadProperty("TextEditForeColor", mp_clrTextEditForeColor)
        CustomBorderStyle.SetXML(oXML.ReadObject("CustomBorderStyle"))
        MilestoneStyle.SetXML(oXML.ReadObject("MilestoneStyle"))
        PredecessorStyle.SetXML(oXML.ReadObject("PredecessorStyle"))
        TaskStyle.SetXML(oXML.ReadObject("TaskStyle"))
        TextFlags.SetXML(oXML.ReadObject("TextFlags"))
        ScrollBarStyle.SetXML(oXML.ReadObject("ScrollBarStyle"))
        SelectionRectangleStyle.SetXML(oXML.ReadObject("SelectionRectangleStyle"))
        ButtonBorderStyle.SetXML(oXML.ReadObject("ButtonBorderStyle"))
    End Sub

End Class
