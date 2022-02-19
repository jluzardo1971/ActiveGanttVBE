Option Explicit On

Public Class TaskBaseline
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_oTimephasedData_C As TimephasedData_C
	Private mp_lNumber As Integer
	Private mp_bInterim As Boolean
	Private mp_dtStart As System.DateTime
	Private mp_dtFinish As System.DateTime
	Private mp_oDuration As Duration
	Private mp_yDurationFormat As E_DURATIONFORMAT
	Private mp_bEstimatedDuration As Boolean
	Private mp_oWork As Duration
	Private mp_cCost As Decimal
	Private mp_fBCWS As Single
	Private mp_fBCWP As Single

	Public Sub New()
		mp_oTimephasedData_C = New TimephasedData_C()
		mp_lNumber = 0
		mp_bInterim = False
		mp_dtStart = New System.DateTime(0)
		mp_dtFinish = New System.DateTime(0)
		mp_oDuration = New Duration()
		mp_yDurationFormat = E_DURATIONFORMAT.DF_M
		mp_bEstimatedDuration = False
		mp_oWork = New Duration()
		mp_cCost = 0
		mp_fBCWS = 0
		mp_fBCWP = 0
	End Sub

	Public ReadOnly Property oTimephasedData_C() As TimephasedData_C
		Get
			Return mp_oTimephasedData_C
		End Get
	End Property

	Public Property lNumber() As Integer
		Get
			Return mp_lNumber
		End Get
		Set(ByVal Value As Integer)
			mp_lNumber = Value
		End Set
	End Property

	Public Property bInterim() As Boolean
		Get
			Return mp_bInterim
		End Get
		Set(ByVal Value As Boolean)
			mp_bInterim = Value
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

	Public ReadOnly Property oDuration() As Duration
		Get
			Return mp_oDuration
		End Get
	End Property

	Public Property yDurationFormat() As E_DURATIONFORMAT
		Get
			Return mp_yDurationFormat
		End Get
		Set(ByVal Value As E_DURATIONFORMAT)
			mp_yDurationFormat = Value
		End Set
	End Property

	Public Property bEstimatedDuration() As Boolean
		Get
			Return mp_bEstimatedDuration
		End Get
		Set(ByVal Value As Boolean)
			mp_bEstimatedDuration = Value
		End Set
	End Property

	Public ReadOnly Property oWork() As Duration
		Get
			Return mp_oWork
		End Get
	End Property

	Public Property cCost() As Decimal
		Get
			Return mp_cCost
		End Get
		Set(ByVal Value As Decimal)
			mp_cCost = Value
		End Set
	End Property

	Public Property fBCWS() As Single
		Get
			Return mp_fBCWS
		End Get
		Set(ByVal Value As Single)
			mp_fBCWS = Value
		End Set
	End Property

	Public Property fBCWP() As Single
		Get
			Return mp_fBCWP
		End Get
		Set(ByVal Value As Single)
			mp_fBCWP = Value
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
		If mp_oTimephasedData_C.IsNull() = False Then
			bReturn = False
		End If
		If mp_lNumber <> 0 Then
			bReturn = False
		End If
		If mp_bInterim <> false Then
			bReturn = False
		End If
		If mp_dtStart.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtFinish.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_oDuration.IsNull() = False Then
			bReturn = False
		End If
		If mp_yDurationFormat <> E_DURATIONFORMAT.DF_M Then
			bReturn = False
		End If
		If mp_bEstimatedDuration <> False Then
			bReturn = False
		End If
		If mp_oWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_cCost <> 0 Then
			bReturn = False
		End If
		If mp_fBCWS <> 0 Then
			bReturn = False
		End If
		If mp_fBCWP <> 0 Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<Baseline/>"
		End if
		Dim oXML As New clsXML("Baseline")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		If mp_oTimephasedData_C.IsNull() = False Then
			mp_oTimephasedData_C.WriteObjectProtected(oXML)
		End If
		oXML.WriteProperty("Number", mp_lNumber)
		oXML.WriteProperty("Interim", mp_bInterim)
		If mp_dtStart.Ticks <> 0 Then
			oXML.WriteProperty("Start", mp_dtStart)
		End If
		If mp_dtFinish.Ticks <> 0 Then
			oXML.WriteProperty("Finish", mp_dtFinish)
		End If
		oXML.WriteProperty("Duration", mp_oDuration)
		oXML.WriteProperty("DurationFormat", mp_yDurationFormat)
		oXML.WriteProperty("EstimatedDuration", mp_bEstimatedDuration)
		oXML.WriteProperty("Work", mp_oWork)
		oXML.WriteProperty("Cost", mp_cCost)
		oXML.WriteProperty("BCWS", mp_fBCWS)
		oXML.WriteProperty("BCWP", mp_fBCWP)
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("Baseline")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		mp_oTimephasedData_C.ReadObjectProtected(oXML)
		oXML.ReadProperty("Number", mp_lNumber)
		oXML.ReadProperty("Interim", mp_bInterim)
		oXML.ReadProperty("Start", mp_dtStart)
		oXML.ReadProperty("Finish", mp_dtFinish)
		oXML.ReadProperty("Duration", mp_oDuration)
		oXML.ReadProperty("DurationFormat", mp_yDurationFormat)
		oXML.ReadProperty("EstimatedDuration", mp_bEstimatedDuration)
		oXML.ReadProperty("Work", mp_oWork)
		oXML.ReadProperty("Cost", mp_cCost)
		oXML.ReadProperty("BCWS", mp_fBCWS)
		oXML.ReadProperty("BCWP", mp_fBCWP)
	End Sub

End Class
