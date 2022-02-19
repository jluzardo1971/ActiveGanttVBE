Option Explicit On

Public Class TimephasedData
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_yType As E_TYPE_5
	Private mp_lUID As Integer
	Private mp_dtStart As System.DateTime
	Private mp_dtFinish As System.DateTime
	Private mp_yUnit As E_UNIT
	Private mp_sValue As String

	Public Sub New()
		mp_yType = E_TYPE_5.T_5_ASSIGNMENT_REMAINING_WORK
		mp_lUID = 0
		mp_dtStart = New System.DateTime(0)
		mp_dtFinish = New System.DateTime(0)
		mp_yUnit = E_UNIT.U_M
		mp_sValue = ""
	End Sub

	Public Property yType() As E_TYPE_5
		Get
			Return mp_yType
		End Get
		Set(ByVal Value As E_TYPE_5)
			mp_yType = Value
		End Set
	End Property

	Public Property lUID() As Integer
		Get
			Return mp_lUID
		End Get
		Set(ByVal Value As Integer)
			mp_lUID = Value
		End Set
	End Property

	Public Property dtStart() As System.DateTime
		Get
			Return mp_dtStart
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtStart = Value
		End Set
	End Property

	Public Property dtFinish() As System.DateTime
		Get
			Return mp_dtFinish
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtFinish = Value
		End Set
	End Property

	Public Property yUnit() As E_UNIT
		Get
			Return mp_yUnit
		End Get
		Set(ByVal Value As E_UNIT)
			mp_yUnit = Value
		End Set
	End Property

	Public Property sValue() As String
		Get
			Return mp_sValue
		End Get
		Set(ByVal Value As String)
			mp_sValue = Value
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
		If mp_yType <> E_TYPE_5.T_5_ASSIGNMENT_REMAINING_WORK Then
			bReturn = False
		End If
		If mp_lUID <> 0 Then
			bReturn = False
		End If
		If mp_dtStart.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtFinish.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_yUnit <> E_UNIT.U_M Then
			bReturn = False
		End If
		If mp_sValue <> "" Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<TimephasedData/>"
		End if
		Dim oXML As New clsXML("TimephasedData")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		oXML.WriteProperty("Type", mp_yType)
		oXML.WriteProperty("UID", mp_lUID)
		If mp_dtStart.Ticks <> 0 Then
			oXML.WriteProperty("Start", mp_dtStart)
		End If
		If mp_dtFinish.Ticks <> 0 Then
			oXML.WriteProperty("Finish", mp_dtFinish)
		End If
		oXML.WriteProperty("Unit", mp_yUnit)
		If mp_sValue <> "" Then
			oXML.WriteProperty("Value", mp_sValue)
		End If
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("TimephasedData")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("Type", mp_yType)
		oXML.ReadProperty("UID", mp_lUID)
		oXML.ReadProperty("Start", mp_dtStart)
		oXML.ReadProperty("Finish", mp_dtFinish)
		oXML.ReadProperty("Unit", mp_yUnit)
		oXML.ReadProperty("Value", mp_sValue)
	End Sub

End Class
