Option Explicit On

Public Class ResourceBaseline
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_lNumber As Integer
	Private mp_oWork As Duration
	Private mp_fCost As Single
	Private mp_fBCWS As Single
	Private mp_fBCWP As Single

	Public Sub New()
		mp_lNumber = 0
		mp_oWork = New Duration()
		mp_fCost = 0
		mp_fBCWS = 0
		mp_fBCWP = 0
	End Sub

	Public Property lNumber() As Integer
		Get
			Return mp_lNumber
		End Get
		Set(ByVal Value As Integer)
			mp_lNumber = Value
		End Set
	End Property

	Public ReadOnly Property oWork() As Duration
		Get
			Return mp_oWork
		End Get
	End Property

	Public Property fCost() As Single
		Get
			Return mp_fCost
		End Get
		Set(ByVal Value As Single)
			mp_fCost = Value
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
		If mp_lNumber <> 0 Then
			bReturn = False
		End If
		If mp_oWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_fCost <> 0 Then
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
		oXML.WriteProperty("Number", mp_lNumber)
		oXML.WriteProperty("Work", mp_oWork)
		oXML.WriteProperty("Cost", mp_fCost)
		oXML.WriteProperty("BCWS", mp_fBCWS)
		oXML.WriteProperty("BCWP", mp_fBCWP)
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("Baseline")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("Number", mp_lNumber)
		oXML.ReadProperty("Work", mp_oWork)
		oXML.ReadProperty("Cost", mp_fCost)
		oXML.ReadProperty("BCWS", mp_fBCWS)
		oXML.ReadProperty("BCWP", mp_fBCWP)
	End Sub

End Class
