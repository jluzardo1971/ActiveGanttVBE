Option Explicit On

Public Class ResourceOutlineCode
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_sFieldID As String
	Private mp_lValueID As Integer
	Private mp_lValueGUID As Integer

	Public Sub New()
		mp_sFieldID = ""
		mp_lValueID = 0
		mp_lValueGUID = 0
	End Sub

	Public Property sFieldID() As String
		Get
			Return mp_sFieldID
		End Get
		Set(ByVal Value As String)
			mp_sFieldID = Value
		End Set
	End Property

	Public Property lValueID() As Integer
		Get
			Return mp_lValueID
		End Get
		Set(ByVal Value As Integer)
			mp_lValueID = Value
		End Set
	End Property

	Public Property lValueGUID() As Integer
		Get
			Return mp_lValueGUID
		End Get
		Set(ByVal Value As Integer)
			mp_lValueGUID = Value
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
		If mp_sFieldID <> "" Then
			bReturn = False
		End If
		If mp_lValueID <> 0 Then
			bReturn = False
		End If
		If mp_lValueGUID <> 0 Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<OutlineCode/>"
		End if
		Dim oXML As New clsXML("OutlineCode")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		If mp_sFieldID <> "" Then
			oXML.WriteProperty("FieldID", mp_sFieldID)
		End If
		oXML.WriteProperty("ValueID", mp_lValueID)
		oXML.WriteProperty("ValueGUID", mp_lValueGUID)
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("OutlineCode")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("FieldID", mp_sFieldID)
		oXML.ReadProperty("ValueID", mp_lValueID)
		oXML.ReadProperty("ValueGUID", mp_lValueGUID)
	End Sub

End Class
