Option Explicit On

Public Class ResourceAvailabilityPeriod
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_dtAvailableFrom As System.DateTime
	Private mp_dtAvailableTo As System.DateTime
	Private mp_fAvailableUnits As Single

	Public Sub New()
		mp_dtAvailableFrom = New System.DateTime(0)
		mp_dtAvailableTo = New System.DateTime(0)
		mp_fAvailableUnits = 0
	End Sub

	Public Property dtAvailableFrom() As System.DateTime
		Get
			Return mp_dtAvailableFrom
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtAvailableFrom = Value
		End Set
	End Property

	Public Property dtAvailableTo() As System.DateTime
		Get
			Return mp_dtAvailableTo
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtAvailableTo = Value
		End Set
	End Property

	Public Property fAvailableUnits() As Single
		Get
			Return mp_fAvailableUnits
		End Get
		Set(ByVal Value As Single)
			mp_fAvailableUnits = Value
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
		If mp_dtAvailableFrom.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtAvailableTo.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_fAvailableUnits <> 0 Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<AvailabilityPeriod/>"
		End if
		Dim oXML As New clsXML("AvailabilityPeriod")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		If mp_dtAvailableFrom.Ticks <> 0 Then
			oXML.WriteProperty("AvailableFrom", mp_dtAvailableFrom)
		End If
		If mp_dtAvailableTo.Ticks <> 0 Then
			oXML.WriteProperty("AvailableTo", mp_dtAvailableTo)
		End If
		oXML.WriteProperty("AvailableUnits", mp_fAvailableUnits)
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("AvailabilityPeriod")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("AvailableFrom", mp_dtAvailableFrom)
		oXML.ReadProperty("AvailableTo", mp_dtAvailableTo)
		oXML.ReadProperty("AvailableUnits", mp_fAvailableUnits)
	End Sub

End Class
