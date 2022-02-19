Option Explicit On

Public Class TimePeriod
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_dtFromDate As System.DateTime
	Private mp_dtToDate As System.DateTime

	Public Sub New()
		mp_dtFromDate = New System.DateTime(0)
		mp_dtToDate = New System.DateTime(0)
	End Sub

	Public Property dtFromDate() As System.DateTime
		Get
			Return mp_dtFromDate
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtFromDate = Value
		End Set
	End Property

	Public Property dtToDate() As System.DateTime
		Get
			Return mp_dtToDate
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtToDate = Value
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
		If mp_dtFromDate.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtToDate.Ticks <> 0 Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<TimePeriod/>"
		End if
		Dim oXML As New clsXML("TimePeriod")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		If mp_dtFromDate.Ticks <> 0 Then
			oXML.WriteProperty("FromDate", mp_dtFromDate)
		End If
		If mp_dtToDate.Ticks <> 0 Then
			oXML.WriteProperty("ToDate", mp_dtToDate)
		End If
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("TimePeriod")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("FromDate", mp_dtFromDate)
		oXML.ReadProperty("ToDate", mp_dtToDate)
	End Sub

End Class
