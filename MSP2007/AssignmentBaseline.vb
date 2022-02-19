Option Explicit On

Public Class AssignmentBaseline
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_oTimephasedData_C As TimephasedData_C
	Private mp_sNumber As String
	Private mp_sStart As String
	Private mp_sFinish As String
	Private mp_oWork As Duration
	Private mp_sCost As String
	Private mp_fBCWS As Single
	Private mp_fBCWP As Single

	Public Sub New()
		mp_oTimephasedData_C = New TimephasedData_C()
		mp_sNumber = ""
		mp_sStart = ""
		mp_sFinish = ""
		mp_oWork = New Duration()
		mp_sCost = ""
		mp_fBCWS = 0
		mp_fBCWP = 0
	End Sub

	Public ReadOnly Property oTimephasedData_C() As TimephasedData_C
		Get
			Return mp_oTimephasedData_C
		End Get
	End Property

	Public Property sNumber() As String
		Get
			Return mp_sNumber
		End Get
		Set(ByVal Value As String)
			mp_sNumber = Value
		End Set
	End Property

	Public Property sStart() As String
		Get
			Return mp_sStart
		End Get
		Set(ByVal Value As String)
			mp_sStart = Value
		End Set
	End Property

	Public Property sFinish() As String
		Get
			Return mp_sFinish
		End Get
		Set(ByVal Value As String)
			mp_sFinish = Value
		End Set
	End Property

	Public ReadOnly Property oWork() As Duration
		Get
			Return mp_oWork
		End Get
	End Property

	Public Property sCost() As String
		Get
			Return mp_sCost
		End Get
		Set(ByVal Value As String)
			mp_sCost = Value
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
		If mp_sNumber <> "" Then
			bReturn = False
		End If
		If mp_sStart <> "" Then
			bReturn = False
		End If
		If mp_sFinish <> "" Then
			bReturn = False
		End If
		If mp_oWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_sCost <> "" Then
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
		oXML.WriteProperty("Number", mp_sNumber)
		If mp_sStart <> "" Then
			oXML.WriteProperty("Start", mp_sStart)
		End If
		If mp_sFinish <> "" Then
			oXML.WriteProperty("Finish", mp_sFinish)
		End If
		oXML.WriteProperty("Work", mp_oWork)
		If mp_sCost <> "" Then
			oXML.WriteProperty("Cost", mp_sCost)
		End If
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
		oXML.ReadProperty("Number", mp_sNumber)
		oXML.ReadProperty("Start", mp_sStart)
		oXML.ReadProperty("Finish", mp_sFinish)
		oXML.ReadProperty("Work", mp_oWork)
		oXML.ReadProperty("Cost", mp_sCost)
		oXML.ReadProperty("BCWS", mp_fBCWS)
		oXML.ReadProperty("BCWP", mp_fBCWP)
	End Sub

End Class
