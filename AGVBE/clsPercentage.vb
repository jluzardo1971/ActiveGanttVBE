Option Explicit On 

Public Class clsPercentage
    Inherits clsItemBase

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_fPercent As Single
    Private mp_sTaskKey As String
    Private mp_sTag As String
    Private mp_bVisible As Boolean
    Private mp_bAllowSize As Boolean
    Private mp_oTask As clsTask
    Private mp_sStyleIndex As String
    Private mp_oStyle As clsStyle
    Private mp_sFormat As String

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_fPercent = 0
        mp_sTaskKey = ""
        mp_sTag = ""
        mp_bVisible = True
        mp_bAllowSize = True
        mp_sStyleIndex = "DS_PERCENTAGE"
        mp_oStyle = mp_oControl.Styles.FItem("DS_PERCENTAGE")
        mp_sFormat = ""
    End Sub

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
            mp_oControl.Percentages.oCollection.mp_SetKey(mp_sKey, Value, SYS_ERRORS.PERCENTAGES_SET_KEY)
        End Set
    End Property

    Public Property Percent() As Single
        Get
            Return mp_fPercent
        End Get
        Set(ByVal Value As Single)
            mp_fPercent = Value
        End Set
    End Property

    Public Property TaskKey() As String
        Get
            Return mp_sTaskKey
        End Get
        Set(ByVal Value As String)
            mp_sTaskKey = Value
            mp_oTask = mp_oControl.Tasks.Item(Value)
        End Set
    End Property

    Public ReadOnly Property Task() As clsTask
        Get
            Return mp_oTask
        End Get
    End Property

    Public ReadOnly Property Layer() As clsLayer
        Get
            Return mp_oTask.Layer
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
            If Left < mp_oControl.Splitter.Right Then
                Return mp_oControl.Splitter.Right
            Else
                Return Left
            End If
        End Get
    End Property


    Public ReadOnly Property RightTrim() As Integer
        Get
            If Right > mp_oControl.mt_RightMargin Then
                Return mp_oControl.mt_RightMargin
            Else
                Return Right
            End If
        End Get
    End Property

    Friend ReadOnly Property f_bLeftVisible() As Boolean
        Get
            If LeftTrim = Left Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Friend ReadOnly Property f_bRightVisible() As Boolean
        Get
            If RightTrim = Right Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property


    Public ReadOnly Property Left() As Integer
        Get
            Return mp_oTask.Left
        End Get
    End Property


    Public ReadOnly Property Top() As Integer
        Get
            If (mp_oTask.Row.Height <= -1) Then
                Return mp_oTask.Row.Top
            End If
            If mp_oStyle.Placement = E_PLACEMENT.PLC_ROWEXTENTSPLACEMENT Or mp_oStyle.Appearance = E_STYLEAPPEARANCE.SA_CELL Then
                Return mp_oTask.Row.Top
            End If
            If mp_oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT Then
                Return mp_oTask.Row.Top + mp_oStyle.OffsetTop
            End If
            Return 0
        End Get
    End Property


    Public ReadOnly Property Bottom() As Integer
        Get
            If (mp_oTask.Row.Height <= -1) Then
                Return mp_oTask.Row.Top
            End If
            If mp_oStyle.Placement = E_PLACEMENT.PLC_ROWEXTENTSPLACEMENT Or mp_oStyle.Appearance = E_STYLEAPPEARANCE.SA_CELL Then
                Return mp_oTask.Row.Bottom - 1
            End If
            If mp_oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT Then
                Return mp_oTask.Row.Top + mp_oStyle.OffsetTop + mp_oStyle.OffsetBottom
            End If
            Return 0
        End Get
    End Property


    Public ReadOnly Property Right() As Integer
        Get
            Return Left + System.Convert.ToInt32((mp_oTask.Right - mp_oTask.Left) * mp_fPercent)
        End Get
    End Property

    Friend ReadOnly Property RightSel() As Integer
        Get
            If Right = Left Then
                Return Left + 15
            Else
                Return Right
            End If
        End Get
    End Property


    Public Property Visible() As Boolean
        Get
            If mp_oTask.Row.Visible = False Then
                Return False
            End If
            If mp_oTask.Visible = False Or mp_oTask.Type = E_OBJECTTYPE.OT_MILESTONE Then
                Return False
            End If
            Return mp_bVisible
        End Get
        Set(ByVal Value As Boolean)
            mp_bVisible = Value
        End Set
    End Property

    Public Property StyleIndex() As String
        Get
            If mp_sStyleIndex = "DS_PERCENTAGE" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_PERCENTAGE"
            mp_sStyleIndex = Value
            mp_oStyle = mp_oControl.Styles.FItem(Value)
        End Set
    End Property

    Public ReadOnly Property Style() As clsStyle
        Get
            Return mp_oStyle
        End Get
    End Property

    Public Property Format() As String
        Get
            Return mp_sFormat
        End Get
        Set(ByVal Value As String)
            mp_sFormat = Value
        End Set
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "Percentage")
        oXML.InitializeWriter()
        oXML.WriteProperty("AllowSize", mp_bAllowSize)
        oXML.WriteProperty("Key", mp_sKey)
        oXML.WriteProperty("Percent", mp_fPercent)
        oXML.WriteProperty("Tag", mp_sTag)
        oXML.WriteProperty("TaskKey", mp_sTaskKey)
        oXML.WriteProperty("Visible", mp_bVisible)
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        oXML.WriteProperty("Format", mp_sFormat)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "Percentage")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("AllowSize", mp_bAllowSize)
        oXML.ReadProperty("Key", mp_sKey)
        oXML.ReadProperty("Percent", mp_fPercent)
        oXML.ReadProperty("Tag", mp_sTag)
        oXML.ReadProperty("TaskKey", mp_sTaskKey)
        mp_oTask = mp_oControl.Tasks.Item(mp_sTaskKey)
        oXML.ReadProperty("Visible", mp_bVisible)
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        StyleIndex = mp_sStyleIndex
        oXML.ReadProperty("Format", mp_sFormat)
    End Sub

End Class
