Option Explicit On

Public Class clsRow
    Inherits clsItemBase

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_bAllowSize As Boolean
    Private mp_bAllowMove As Boolean
    Private mp_bContainer As Boolean
    Private mp_bMergeCells As Boolean
    Private mp_lHeight As Integer
    
    Public Node As clsNode
    
    Public Cells As clsCells
    Private mp_sText As String
    Private mp_oImage As Image
    Private mp_sStyleIndex As String
    Private mp_sTag As String
    Private mp_sClientAreaStyleIndex As String
    Private mp_lTop As Integer
    Private mp_lBottom As Integer
    'Private mp_bVisible As Boolean
    Private mp_yClientAreaVisibility As E_CLIENTAREAVISIBILITY
    Private mp_oStyle As clsStyle
    Private mp_oClientAreaStyle As clsStyle
    Private mp_bUseNodeImages As Boolean
    Private mp_sImageTag As String
    Friend mp_lTextLeft As Double
    Friend mp_lTextTop As Double
    Friend mp_lTextRight As Double
    Friend mp_lTextBottom As Double
    Private mp_bAllowTextEdit As Boolean

    Public Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_bContainer = True
        mp_bMergeCells = False
        mp_lHeight = 40
        mp_sText = ""
        mp_oImage = Nothing
        mp_sStyleIndex = "DS_ROW"
        mp_oStyle = mp_oControl.Styles.FItem("DS_ROW")
        mp_sTag = ""
        Cells = New clsCells(mp_oControl, Me)
        mp_sClientAreaStyleIndex = "DS_CLIENTAREA"
        mp_oClientAreaStyle = mp_oControl.Styles.FItem("DS_CLIENTAREA")
        Node = New clsNode(mp_oControl, Me)
        '// Metrics
        mp_lTop = 0
        mp_lBottom = 0
        'mp_bVisible = False
        mp_bUseNodeImages = False
        mp_bAllowSize = True
        mp_bAllowMove = True
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
            mp_oControl.Rows.oCollection.mp_SetKey(mp_sKey, Value, SYS_ERRORS.ROWS_SET_KEY)
        End Set
    End Property

    
    Public Property Container() As Boolean
        Get
            Return mp_bContainer
        End Get
        Set(ByVal Value As Boolean)
            mp_bContainer = Value
        End Set
    End Property

    
    Public Property UseNodeImages() As Boolean
        Get
            Return mp_bUseNodeImages
        End Get
        Set(ByVal Value As Boolean)
            mp_bUseNodeImages = Value
        End Set
    End Property

    
    Public Property MergeCells() As Boolean
        Get
            If mp_oControl.TreeviewColumnIndex <> 0 Then
                Return False
            Else
                Return mp_bMergeCells
            End If
        End Get
        Set(ByVal Value As Boolean)
            mp_bMergeCells = Value
        End Set
    End Property

    
    Public Property Height() As Integer
        Get
            If Node.Hidden = False Then
                Return mp_lHeight
            Else
                Return -1
            End If
        End Get
        Set(ByVal Value As Integer)
            mp_lHeight = Value
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
            If mp_bUseNodeImages = False Then
                Return mp_oImage
            Else
                Dim oImage As Image = Nothing
                If Node.Expanded = True And Node.Children() > 0 And Not (Node.ExpandedImage Is Nothing) Then
                    oImage = Node.ExpandedImage
                ElseIf Node.Selected = True And Not (Node.SelectedImage Is Nothing) Then
                    oImage = Node.SelectedImage
                ElseIf Not (Node.Image Is Nothing) Then
                    oImage = Node.Image
                End If
                Return oImage
            End If
        End Get
        Set(ByVal Value As Image)
            mp_oImage = Value
        End Set
    End Property

    
    Public Property StyleIndex() As String
        Get
            If mp_sStyleIndex = "DS_ROW" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_ROW"
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

    
    Public Property ClientAreaStyleIndex() As String
        Get
            If mp_sClientAreaStyleIndex = "DS_CLIENTAREA" Then
                Return ""
            Else
                Return mp_sClientAreaStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_CLIENTAREA"
            mp_sClientAreaStyleIndex = Value
            mp_oClientAreaStyle = mp_oControl.Styles.FItem(Value)
        End Set
    End Property

    
    Public ReadOnly Property ClientAreaStyle() As clsStyle
        Get
            Return mp_oClientAreaStyle
        End Get
    End Property

    
    Public ReadOnly Property Left() As Integer
        Get
            Return mp_oControl.mt_LeftMargin
        End Get
    End Property

    
    Public ReadOnly Property Top() As Integer
        Get
            Return mp_lTop
        End Get
    End Property

    Friend WriteOnly Property f_lTop() As Integer
        Set(ByVal Value As Integer)
            mp_lTop = Value
        End Set
    End Property

    
    Public ReadOnly Property Right() As Integer
        Get
            Return mp_oControl.Splitter.Left
        End Get
    End Property

    
    Public ReadOnly Property Bottom() As Integer
        Get
            Return mp_lBottom
        End Get
    End Property

    Friend WriteOnly Property f_lBottom() As Integer
        Set(ByVal Value As Integer)
            mp_lBottom = Value
        End Set
    End Property

    
    Public ReadOnly Property Visible() As Boolean
        Get
            If mp_yClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_INSIDEVISIBLEAREA Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Friend Property ClientAreaVisibility() As E_CLIENTAREAVISIBILITY
        Get
            Return mp_yClientAreaVisibility
        End Get
        Set(ByVal Value As E_CLIENTAREAVISIBILITY)
            mp_yClientAreaVisibility = Value
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
        Dim oXML As New clsXML(mp_oControl, "Row")
        oXML.InitializeWriter()
        oXML.WriteProperty("AllowMove", mp_bAllowMove)
        oXML.WriteProperty("AllowSize", mp_bAllowSize)
        oXML.WriteProperty("ClientAreaStyleIndex", mp_sClientAreaStyleIndex)
        oXML.WriteProperty("Container", mp_bContainer)
        oXML.WriteProperty("Height", mp_lHeight)
        oXML.WriteProperty("Image", mp_oImage)
        oXML.WriteProperty("Key", mp_sKey)
        oXML.WriteProperty("MergeCells", mp_bMergeCells)
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        oXML.WriteProperty("Tag", mp_sTag)
        oXML.WriteProperty("Text", mp_sText)
        oXML.WriteProperty("UseNodeImages", mp_bUseNodeImages)
        oXML.WriteProperty("ImageTag", mp_sImageTag)
        oXML.WriteProperty("AllowTextEdit", mp_bAllowTextEdit)
        oXML.WriteObject(Cells.GetXML())
        oXML.WriteObject(Node.GetXML())
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "Row")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("AllowMove", mp_bAllowMove)
        oXML.ReadProperty("AllowSize", mp_bAllowSize)
        oXML.ReadProperty("ClientAreaStyleIndex", mp_sClientAreaStyleIndex)
        ClientAreaStyleIndex = mp_sClientAreaStyleIndex
        oXML.ReadProperty("Container", mp_bContainer)
        oXML.ReadProperty("Height", mp_lHeight)
        oXML.ReadProperty("Image", mp_oImage)
        oXML.ReadProperty("Key", mp_sKey)
        oXML.ReadProperty("MergeCells", mp_bMergeCells)
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        StyleIndex = mp_sStyleIndex
        oXML.ReadProperty("Tag", mp_sTag)
        oXML.ReadProperty("Text", mp_sText)
        oXML.ReadProperty("UseNodeImages", mp_bUseNodeImages)
        oXML.ReadProperty("ImageTag", mp_sImageTag)
        oXML.ReadProperty("AllowTextEdit", mp_bAllowTextEdit)
        Cells.SetXML(oXML.ReadObject("Cells"))
        Node.SetXML(oXML.ReadObject("Node"))
    End Sub

End Class
