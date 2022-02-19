Option Explicit On

Public Class clsColumn
    Inherits clsItemBase

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_bAllowMove As Boolean
    Private mp_bAllowSize As Boolean
    Private mp_lWidth As Integer
    Private mp_sText As String
    Private mp_oImage As Image
    Private mp_sStyleIndex As String
    Private mp_sTag As String
    Private mp_lLeft As Integer
    Private mp_lRight As Integer
    Private mp_bVisible As Boolean
    Private mp_oStyle As clsStyle
    Private mp_sImageTag As String
    Friend mp_lTextLeft As Double
    Friend mp_lTextTop As Double
    Friend mp_lTextRight As Double
    Friend mp_lTextBottom As Double
    Private mp_bAllowTextEdit As Boolean
    Friend mp_bTreeViewColumnIndex As Boolean

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_bAllowMove = True
        mp_bAllowSize = True
        mp_lWidth = 125
        mp_sText = ""
        mp_oImage = Nothing
        mp_sStyleIndex = "DS_COLUMN"
        mp_oStyle = mp_oControl.Styles.FItem("DS_COLUMN")
        mp_sTag = ""
        mp_lLeft = 0
        mp_lRight = 0
        mp_bVisible = False
        mp_sImageTag = ""
        mp_bAllowTextEdit = False
    End Sub

    Public Property AllowTextEdit() As Boolean
        Get
            Return mp_bAllowTextEdit
        End Get
        Set(ByVal value As Boolean)
            mp_bAllowTextEdit = value
        End Set
    End Property

    Public Property AllowMove() As Boolean
        Get
            Return mp_bAllowMove
        End Get
        Set(ByVal Value As Boolean)
            mp_bAllowMove = Value
        End Set
    End Property

    
    Public Property AllowSize() As Boolean
        Get
            Return mp_bAllowSize
        End Get
        Set(ByVal Value As Boolean)
            mp_bAllowSize = Value
        End Set
    End Property

    
    Public Property Key() As String
        Get
            Return mp_sKey
        End Get
        Set(ByVal Value As String)
            mp_oControl.Columns.oCollection.mp_SetKey(mp_sKey, Value, SYS_ERRORS.COLUMNS_SET_KEY)
        End Set
    End Property

    
    Public Property Width() As Integer
        Get
            Return mp_lWidth
        End Get
        Set(ByVal Value As Integer)
            mp_lWidth = Value
        End Set
    End Property

    
    Public Property Text() As String
        Get
            Return mp_sText
        End Get
        Set(ByVal Value As String)
            mp_sText = Value
        End Set
    End Property

    
    Public Property Image() As Image
        Get
            Return mp_oImage
        End Get
        Set(ByVal Value As Image)
            mp_oImage = Value
        End Set
    End Property

    
    Public Property StyleIndex() As String
        Get
            If mp_sStyleIndex = "DS_COLUMN" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_COLUMN"
            mp_sStyleIndex = Value
            mp_oStyle = mp_oControl.Styles.FItem(Value)
        End Set
    End Property

    
    Public ReadOnly Property Style() As clsStyle
        Get
            Return mp_oStyle
        End Get
    End Property

    
    Public Property Tag() As String
        Get
            Return mp_sTag
        End Get
        Set(ByVal Value As String)
            mp_sTag = Value
        End Set
    End Property

    
    Public ReadOnly Property LeftTrim() As Integer
        Get
            If mp_lLeft < mp_oControl.mt_LeftMargin Then
                Return mp_oControl.mt_LeftMargin
            Else
                Return mp_lLeft
            End If
        End Get
    End Property

    
    Public ReadOnly Property RightTrim() As Integer
        Get
            If mp_lRight > mp_oControl.Splitter.Left Then
                Return mp_oControl.Splitter.Left
            Else
                Return mp_lRight
            End If
        End Get
    End Property

    
    Public ReadOnly Property Left() As Integer
        Get
            Return mp_lLeft
        End Get
    End Property

    Friend WriteOnly Property f_lLeft() As Integer
        Set(ByVal Value As Integer)
            mp_lLeft = Value
        End Set
    End Property

    
    Public ReadOnly Property Top() As Integer
        Get
            Return mp_oControl.mt_TopMargin
        End Get
    End Property

    
    Public ReadOnly Property Right() As Integer
        Get
            Return mp_lRight
        End Get
    End Property

    Friend WriteOnly Property f_lRight() As Integer
        Set(ByVal Value As Integer)
            mp_lRight = Value
        End Set
    End Property

    
    Public ReadOnly Property Bottom() As Integer
        Get
            Return mp_oControl.CurrentViewObject.TimeLine.Bottom
        End Get
    End Property

    
    Public ReadOnly Property Visible() As Boolean
        Get
            Return mp_bVisible
        End Get
    End Property

    Friend WriteOnly Property f_bVisible() As Boolean
        Set(ByVal Value As Boolean)
            If Value = False Then
                mp_lLeft = 0
                mp_lRight = 0
            End If
            mp_bVisible = Value
        End Set
    End Property

    Public Property ImageTag() As String
        Get
            Return mp_sImageTag
        End Get
        Set(ByVal Value As String)
            mp_sImageTag = Value
        End Set
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "Column")
        oXML.InitializeWriter()
        oXML.WriteProperty("AllowMove", mp_bAllowMove)
        oXML.WriteProperty("AllowSize", mp_bAllowSize)
        oXML.WriteProperty("Image", mp_oImage)
        oXML.WriteProperty("Key", mp_sKey)
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        oXML.WriteProperty("Tag", mp_sTag)
        oXML.WriteProperty("Text", mp_sText)
        oXML.WriteProperty("Width", mp_lWidth)
        oXML.WriteProperty("ImageTag", mp_sImageTag)
        oXML.WriteProperty("AllowTextEdit", mp_bAllowTextEdit)
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "Column")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("AllowMove", mp_bAllowMove)
        oXML.ReadProperty("AllowSize", mp_bAllowSize)
        oXML.ReadProperty("Image", mp_oImage)
        oXML.ReadProperty("Key", mp_sKey)
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        StyleIndex = mp_sStyleIndex
        oXML.ReadProperty("Tag", mp_sTag)
        oXML.ReadProperty("Text", mp_sText)
        oXML.ReadProperty("Width", mp_lWidth)
        oXML.ReadProperty("ImageTag", mp_sImageTag)
        oXML.ReadProperty("AllowTextEdit", mp_bAllowTextEdit)
    End Sub

End Class
