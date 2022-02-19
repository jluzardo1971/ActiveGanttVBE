Public Class Font

    Private mp_sFamily As String
    Public Size As Single = 10
    Friend Italic As Boolean = False
    Friend Underline As Boolean = False
    Public FontStyle As FontStyle
    Public FontWeight As FontWeight
    Public VerticalAlignment As GRE_VERTICALALIGNMENT = GRE_VERTICALALIGNMENT.VAL_TOP
    Public HorizontalAlignment As GRE_HORIZONTALALIGNMENT = GRE_HORIZONTALALIGNMENT.HAL_LEFT

    Public Sub New(ByVal FamilyName As String, ByVal lSize As Single, ByVal lSizeUnits As E_FONTSIZEUNITS)
        mp_sFamily = FamilyName
        If lSizeUnits = E_FONTSIZEUNITS.FSU_PIXELS Then
            Size = lSize
        ElseIf lSizeUnits = E_FONTSIZEUNITS.FSU_POINTS Then
            Size = (96 * lSize / 72)
        End If
    End Sub

    Public Sub New(ByVal FamilyName As String, ByVal lSize As Single, ByVal lSizeUnits As E_FONTSIZEUNITS, ByVal newStyle As FontWeight)
        mp_sFamily = FamilyName
        If lSizeUnits = E_FONTSIZEUNITS.FSU_PIXELS Then
            Size = lSize
        ElseIf lSizeUnits = E_FONTSIZEUNITS.FSU_POINTS Then
            Size = (96 * lSize / 72)
        End If
        FontWeight = newStyle
    End Sub

    Public Function GetSize(ByVal lSizeUnits As E_FONTSIZEUNITS) As Single
        If lSizeUnits = E_FONTSIZEUNITS.FSU_PIXELS Then
            Return Size
        ElseIf lSizeUnits = E_FONTSIZEUNITS.FSU_POINTS Then
            Return Size / 96 * 72
        End If
        Return 0
    End Function

    Public Function GetFontFamily() As FontFamily
        Dim oFontFamily As New FontFamily(mp_sFamily)
        Return oFontFamily
    End Function

    Public Function Clone() As Font
        Return CType(Me.MemberwiseClone(), Font)
    End Function

    Public ReadOnly Property Name() As String
        Get
            Return mp_sFamily
        End Get
    End Property

    Public Property FamilyName() As String
        Get
            Return mp_sFamily
        End Get
        Set(ByVal value As String)
            mp_sFamily = value
        End Set
    End Property

    Friend ReadOnly Property Bold() As Boolean
        Get
            If FontWeight = FontWeights.Bold Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

End Class
