Option Explicit On

Public Class WorkingTime
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_oFromTime As Time
	Private mp_oToTime As Time

	Public Sub New()
		mp_oFromTime = New Time()
		mp_oToTime = New Time()
	End Sub

	Public ReadOnly Property oFromTime() As Time
		Get
			Return mp_oFromTime
		End Get
	End Property

	Public ReadOnly Property oToTime() As Time
		Get
			Return mp_oToTime
		End Get
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
		If mp_oFromTime.IsNull() = False Then
			bReturn = False
		End If
		If mp_oToTime.IsNull() = False Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<WorkingTime/>"
		End if
		Dim oXML As New clsXML("WorkingTime")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		If mp_oFromTime.IsNull() = False Then
			oXML.WriteProperty("FromTime", mp_oFromTime)
		End If
		If mp_oToTime.IsNull() = False Then
			oXML.WriteProperty("ToTime", mp_oToTime)
		End If
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("WorkingTime")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("FromTime", mp_oFromTime)
		oXML.ReadProperty("ToTime", mp_oToTime)
	End Sub

End Class
