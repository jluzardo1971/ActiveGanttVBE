Option Explicit On

Public Class clsScrollBarSeparator

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_sStyleIndex As String
    Private mp_oStyle As clsStyle

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_sStyleIndex = "DS_SB_SEPARATOR"
        mp_oStyle = mp_oControl.Styles.FItem("DS_SB_SEPARATOR")
    End Sub

    Public Property StyleIndex() As String
        Get
            If mp_sStyleIndex = "DS_SB_SEPARATOR" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_SB_SEPARATOR"
            mp_sStyleIndex = Value
            mp_oStyle = mp_oControl.Styles.FItem(Value)
        End Set
    End Property

    Public ReadOnly Property Style() As clsStyle
        Get
            Return mp_oStyle
        End Get
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "ScrollBarSeparator")
        oXML.InitializeWriter()
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "ScrollBarSeparator")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        StyleIndex = mp_sStyleIndex
    End Sub

End Class

