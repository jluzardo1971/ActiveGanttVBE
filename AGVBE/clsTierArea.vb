Option Explicit On 

Public Class clsTierArea

    Private mp_oControl As ActiveGanttVBECtl
    Public UpperTier As clsTier
    Public MiddleTier As clsTier
    Public LowerTier As clsTier
    Public TierFormat As clsTierFormat
    Public TierAppearance As clsTierAppearance
    Private mp_oTimeLine As clsTimeLine

    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal oTimeLine As clsTimeLine)
        mp_oControl = Value
        mp_oTimeLine = oTimeLine
        UpperTier = New clsTier(mp_oControl, Me, E_TIERPOSITION.SP_UPPER)
        MiddleTier = New clsTier(mp_oControl, Me, E_TIERPOSITION.SP_MIDDLE)
        LowerTier = New clsTier(mp_oControl, Me, E_TIERPOSITION.SP_LOWER)
        TierFormat = New clsTierFormat(mp_oControl)
        TierAppearance = New clsTierAppearance(mp_oControl)
    End Sub

    Friend ReadOnly Property TimeLine() As clsTimeLine
        Get
            Return mp_oTimeLine
        End Get
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "TierArea")
        oXML.InitializeWriter()
        oXML.WriteObject(LowerTier.GetXML())
        oXML.WriteObject(MiddleTier.GetXML())
        oXML.WriteObject(TierAppearance.GetXML())
        oXML.WriteObject(TierFormat.GetXML())
        oXML.WriteObject(UpperTier.GetXML())
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "TierArea")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        LowerTier.SetXML(oXML.ReadObject("LowerTier"))
        MiddleTier.SetXML(oXML.ReadObject("MiddleTier"))
        TierAppearance.SetXML(oXML.ReadObject("TierAppearance"))
        TierFormat.SetXML(oXML.ReadObject("TierFormat"))
        UpperTier.SetXML(oXML.ReadObject("UpperTier"))
    End Sub

End Class
