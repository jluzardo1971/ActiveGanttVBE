Option Explicit On

Public Class clsButtonState

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_sType As String
    Private mp_sNormalStyleIndex As String
    Private mp_sPressedStyleIndex As String
    Private mp_sDisabledStyleIndex As String
    Private mp_oNormalStyle As clsStyle
    Private mp_oPressedStyle As clsStyle
    Private mp_oDisabledStyle As clsStyle

    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal sType As String)
        mp_oControl = Value
        mp_sType = sType
        mp_sNormalStyleIndex = "DS_SB_NORMAL"
        mp_oNormalStyle = mp_oControl.Styles.FItem("DS_SB_NORMAL")
        mp_sPressedStyleIndex = "DS_SB_PRESSED"
        mp_oPressedStyle = mp_oControl.Styles.FItem("DS_SB_PRESSED")
        mp_sDisabledStyleIndex = "DS_SB_DISABLED"
        mp_oDisabledStyle = mp_oControl.Styles.FItem("DS_SB_DISABLED")
    End Sub

    Public Property NormalStyleIndex() As String
        Get
            If mp_sNormalStyleIndex = "DS_SB_NORMAL" Then
                Return ""
            Else
                Return mp_sNormalStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_SB_NORMAL"
            mp_sNormalStyleIndex = Value
            mp_oNormalStyle = mp_oControl.Styles.FItem(Value)
        End Set
    End Property

    Public ReadOnly Property NormalStyle() As clsStyle
        Get
            Return mp_oNormalStyle
        End Get
    End Property

    Public Property PressedStyleIndex() As String
        Get
            If mp_sPressedStyleIndex = "DS_SB_PRESSED" Then
                Return ""
            Else
                Return mp_sPressedStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_SB_PRESSED"
            mp_sPressedStyleIndex = Value
            mp_oPressedStyle = mp_oControl.Styles.FItem(Value)
        End Set
    End Property

    Public ReadOnly Property PressedStyle() As clsStyle
        Get
            Return mp_oPressedStyle
        End Get
    End Property

    Public Property DisabledStyleIndex() As String
        Get
            If mp_sDisabledStyleIndex = "DS_SB_DISABLED" Then
                Return ""
            Else
                Return mp_sDisabledStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_SB_DISABLED"
            mp_sDisabledStyleIndex = Value
            mp_oDisabledStyle = mp_oControl.Styles.FItem(Value)
        End Set
    End Property

    Public ReadOnly Property DisabledStyle() As clsStyle
        Get
            Return mp_oDisabledStyle
        End Get
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, mp_sType & "ButtonState")
        oXML.InitializeWriter()
        oXML.WriteProperty("NormalStyleIndex", mp_sNormalStyleIndex)
        oXML.WriteProperty("PressedStyleIndex", mp_sPressedStyleIndex)
        oXML.WriteProperty("DisabledStyleIndex", mp_sDisabledStyleIndex)

        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, mp_sType & "ButtonState")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("NormalStyleIndex", mp_sNormalStyleIndex)
        NormalStyleIndex = mp_sNormalStyleIndex
        oXML.ReadProperty("PressedStyleIndex", mp_sPressedStyleIndex)
        PressedStyleIndex = mp_sPressedStyleIndex
        oXML.ReadProperty("DisabledStyleIndex", mp_sDisabledStyleIndex)
        DisabledStyleIndex = mp_sDisabledStyleIndex
    End Sub



End Class
