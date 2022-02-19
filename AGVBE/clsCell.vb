Option Explicit On

Public Class clsCell
    Inherits clsItemBase

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_sText As String
    Private mp_oImage As Image
    Private mp_sStyleIndex As String
    Private mp_sTag As String
    Private mp_oCells As clsCells
    Private mp_oStyle As clsStyle
    Private mp_sImageTag As String
    Friend mp_lTextLeft As Double
    Friend mp_lTextTop As Double
    Friend mp_lTextRight As Double
    Friend mp_lTextBottom As Double
    Private mp_bAllowTextEdit As Boolean

    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal oCells As clsCells)
        mp_oControl = Value
        mp_sText = ""
        mp_oImage = Nothing
        mp_sStyleIndex = "DS_CELL"
        mp_oStyle = mp_oControl.Styles.FItem("DS_CELL")
        mp_sTag = ""
        mp_oCells = oCells
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

    
    Public ReadOnly Property Row() As clsRow
        Get
            Return mp_oCells.Row()
        End Get
    End Property

    
    Public Property Key() As String
        Get
            Return mp_sKey
        End Get
        Set(ByVal Value As String)
            mp_oCells.oCollection.mp_SetKey(mp_sKey, Value, SYS_ERRORS.CELLS_SET_KEY)
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
            If mp_sStyleIndex = "DS_CELL" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_CELL"
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

    
    Public ReadOnly Property RowKey() As String
        Get
            Return mp_oCells.Row().Key
        End Get
    End Property

    
    Public ReadOnly Property Left() As Integer
        Get
            Return mp_oControl.Columns.Item(mp_lIndex.ToString()).Left
        End Get
    End Property

    
    Public ReadOnly Property Top() As Integer
        Get
            Return mp_oCells.Row().Top
        End Get
    End Property

    
    Public ReadOnly Property Right() As Integer
        Get
            Return mp_oControl.Columns.Item(mp_lIndex.ToString()).Right
        End Get
    End Property

    
    Public ReadOnly Property Bottom() As Integer
        Get
            Return mp_oCells.Row().Bottom
        End Get
    End Property

    
    Public ReadOnly Property LeftTrim() As Integer
        Get
            Return mp_oControl.Columns.Item(mp_lIndex.ToString()).LeftTrim
        End Get
    End Property

    
    Public ReadOnly Property RightTrim() As Integer
        Get
            Return mp_oControl.Columns.Item(mp_lIndex.ToString()).RightTrim
        End Get
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
        Dim oXML As New clsXML(mp_oControl, "Cell")
        oXML.InitializeWriter()
        oXML.WriteProperty("Image", mp_oImage)
        oXML.WriteProperty("Key", mp_sKey)
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        oXML.WriteProperty("Tag", mp_sTag)
        oXML.WriteProperty("Text", mp_sText)
        oXML.WriteProperty("ImageTag", mp_sTag)
        oXML.WriteProperty("AllowTextEdit", mp_bAllowTextEdit)
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "Cell")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("Image", mp_oImage)
        oXML.ReadProperty("Key", mp_sKey)
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        StyleIndex = mp_sStyleIndex
        oXML.ReadProperty("Tag", mp_sTag)
        oXML.ReadProperty("Text", mp_sText)
        oXML.ReadProperty("ImageTag", mp_sImageTag)
        oXML.ReadProperty("AllowTextEdit", mp_bAllowTextEdit)
    End Sub

End Class
