Option Explicit On

Public Class ResourceRate
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_dtRatesFrom As System.DateTime
	Private mp_dtRatesTo As System.DateTime
	Private mp_yRateTable As E_RATETABLE
	Private mp_cStandardRate As Decimal
	Private mp_yStandardRateFormat As E_STANDARDRATEFORMAT_1
	Private mp_cOvertimeRate As Decimal
	Private mp_yOvertimeRateFormat As E_OVERTIMERATEFORMAT
	Private mp_cCostPerUse As Decimal

	Public Sub New()
		mp_dtRatesFrom = New System.DateTime(0)
		mp_dtRatesTo = New System.DateTime(0)
		mp_yRateTable = E_RATETABLE.RT_A
		mp_cStandardRate = 0
		mp_yStandardRateFormat = E_STANDARDRATEFORMAT_1.SRF_1_M
		mp_cOvertimeRate = 0
		mp_yOvertimeRateFormat = E_OVERTIMERATEFORMAT.ORF_M
		mp_cCostPerUse = 0
	End Sub

	Public Property dtRatesFrom() As System.DateTime
		Get
			Return mp_dtRatesFrom
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtRatesFrom = Value
		End Set
	End Property

	Public Property dtRatesTo() As System.DateTime
		Get
			Return mp_dtRatesTo
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtRatesTo = Value
		End Set
	End Property

	Public Property yRateTable() As E_RATETABLE
		Get
			Return mp_yRateTable
		End Get
		Set(ByVal Value As E_RATETABLE)
			mp_yRateTable = Value
		End Set
	End Property

	Public Property cStandardRate() As Decimal
		Get
			Return mp_cStandardRate
		End Get
		Set(ByVal Value As Decimal)
			mp_cStandardRate = Value
		End Set
	End Property

	Public Property yStandardRateFormat() As E_STANDARDRATEFORMAT_1
		Get
			Return mp_yStandardRateFormat
		End Get
		Set(ByVal Value As E_STANDARDRATEFORMAT_1)
			mp_yStandardRateFormat = Value
		End Set
	End Property

	Public Property cOvertimeRate() As Decimal
		Get
			Return mp_cOvertimeRate
		End Get
		Set(ByVal Value As Decimal)
			mp_cOvertimeRate = Value
		End Set
	End Property

	Public Property yOvertimeRateFormat() As E_OVERTIMERATEFORMAT
		Get
			Return mp_yOvertimeRateFormat
		End Get
		Set(ByVal Value As E_OVERTIMERATEFORMAT)
			mp_yOvertimeRateFormat = Value
		End Set
	End Property

	Public Property cCostPerUse() As Decimal
		Get
			Return mp_cCostPerUse
		End Get
		Set(ByVal Value As Decimal)
			mp_cCostPerUse = Value
		End Set
	End Property

	Public Property Key() As String
		Get
			Return mp_sKey
		End Get
		Set(ByVal Value As String)
			mp_oCollection.mp_SetKey(mp_sKey, Value, SYS_ERRORS.MP_SET_KEY)
		End Set
	End Property

	Public Function IsNull() As Boolean
		Dim bReturn As Boolean = True
		If mp_dtRatesFrom.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtRatesTo.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_yRateTable <> E_RATETABLE.RT_A Then
			bReturn = False
		End If
		If mp_cStandardRate <> 0 Then
			bReturn = False
		End If
		If mp_yStandardRateFormat <> E_STANDARDRATEFORMAT_1.SRF_1_M Then
			bReturn = False
		End If
		If mp_cOvertimeRate <> 0 Then
			bReturn = False
		End If
		If mp_yOvertimeRateFormat <> E_OVERTIMERATEFORMAT.ORF_M Then
			bReturn = False
		End If
		If mp_cCostPerUse <> 0 Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<Rate/>"
		End if
		Dim oXML As New clsXML("Rate")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		If mp_dtRatesFrom.Ticks <> 0 Then
			oXML.WriteProperty("RatesFrom", mp_dtRatesFrom)
		End If
		If mp_dtRatesTo.Ticks <> 0 Then
			oXML.WriteProperty("RatesTo", mp_dtRatesTo)
		End If
		oXML.WriteProperty("RateTable", mp_yRateTable)
		oXML.WriteProperty("StandardRate", mp_cStandardRate)
		oXML.WriteProperty("StandardRateFormat", mp_yStandardRateFormat)
		oXML.WriteProperty("OvertimeRate", mp_cOvertimeRate)
		oXML.WriteProperty("OvertimeRateFormat", mp_yOvertimeRateFormat)
		oXML.WriteProperty("CostPerUse", mp_cCostPerUse)
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("Rate")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("RatesFrom", mp_dtRatesFrom)
		oXML.ReadProperty("RatesTo", mp_dtRatesTo)
		oXML.ReadProperty("RateTable", mp_yRateTable)
		oXML.ReadProperty("StandardRate", mp_cStandardRate)
		oXML.ReadProperty("StandardRateFormat", mp_yStandardRateFormat)
		oXML.ReadProperty("OvertimeRate", mp_cOvertimeRate)
		oXML.ReadProperty("OvertimeRateFormat", mp_yOvertimeRateFormat)
		oXML.ReadProperty("CostPerUse", mp_cCostPerUse)
	End Sub

End Class
