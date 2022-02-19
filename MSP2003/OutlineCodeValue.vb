Option Explicit On

Public Class OutlineCodeValue
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_lValueID As Integer
	Private mp_lParentValueID As Integer
	Private mp_sValue As String
	Private mp_sDescription As String

	Public Sub New()
		mp_lValueID = 0
		mp_lParentValueID = 0
		mp_sValue = ""
		mp_sDescription = ""
	End Sub

	Public Property lValueID() As Integer
		Get
			Return mp_lValueID
		End Get
		Set(ByVal Value As Integer)
			mp_lValueID = Value
		End Set
	End Property

	Public Property lParentValueID() As Integer
		Get
			Return mp_lParentValueID
		End Get
		Set(ByVal Value As Integer)
			mp_lParentValueID = Value
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

	Public Property sDescription() As String
		Get
			Return mp_sDescription
		End Get
		Set(ByVal Value As String)
			mp_sDescription = Value
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
		If mp_lValueID <> 0 Then
			bReturn = False
		End If
		If mp_lParentValueID <> 0 Then
			bReturn = False
		End If
		If mp_sValue <> "" Then
			bReturn = False
		End If
		If mp_sDescription <> "" Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<Value/>"
		End if
		Dim oXML As New clsXML("Value")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		oXML.WriteProperty("ValueID", mp_lValueID)
		oXML.WriteProperty("ParentValueID", mp_lParentValueID)
		If mp_sValue <> "" Then
			oXML.WriteProperty("Value", mp_sValue)
		End If
		If mp_sDescription <> "" Then
			oXML.WriteProperty("Description", mp_sDescription)
		End If
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("Value")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("ValueID", mp_lValueID)
		oXML.ReadProperty("ParentValueID", mp_lParentValueID)
		oXML.ReadProperty("Value", mp_sValue)
		oXML.ReadProperty("Description", mp_sDescription)
	End Sub

End Class
