Option Explicit On

Public Class TaskPredecessorLink
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_lPredecessorUID As Integer
	Private mp_yType As E_TYPE_5
	Private mp_bCrossProject As Boolean
	Private mp_sCrossProjectName As String
	Private mp_lLinkLag As Integer
	Private mp_yLagFormat As E_LAGFORMAT

	Public Sub New()
		mp_lPredecessorUID = 0
		mp_yType = E_TYPE_5.T_5_FF
		mp_bCrossProject = False
		mp_sCrossProjectName = ""
		mp_lLinkLag = 0
		mp_yLagFormat = E_LAGFORMAT.LF_M
	End Sub

	Public Property lPredecessorUID() As Integer
		Get
			Return mp_lPredecessorUID
		End Get
		Set(ByVal Value As Integer)
			mp_lPredecessorUID = Value
		End Set
	End Property

	Public Property yType() As E_TYPE_5
		Get
			Return mp_yType
		End Get
		Set(ByVal Value As E_TYPE_5)
			mp_yType = Value
		End Set
	End Property

	Public Property bCrossProject() As Boolean
		Get
			Return mp_bCrossProject
		End Get
		Set(ByVal Value As Boolean)
			mp_bCrossProject = Value
		End Set
	End Property

	Public Property sCrossProjectName() As String
		Get
			Return mp_sCrossProjectName
		End Get
		Set(ByVal Value As String)
			mp_sCrossProjectName = Value
		End Set
	End Property

	Public Property lLinkLag() As Integer
		Get
			Return mp_lLinkLag
		End Get
		Set(ByVal Value As Integer)
			mp_lLinkLag = Value
		End Set
	End Property

	Public Property yLagFormat() As E_LAGFORMAT
		Get
			Return mp_yLagFormat
		End Get
		Set(ByVal Value As E_LAGFORMAT)
			mp_yLagFormat = Value
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
		If mp_lPredecessorUID <> 0 Then
			bReturn = False
		End If
		If mp_yType <> E_TYPE_5.T_5_FF Then
			bReturn = False
		End If
		If mp_bCrossProject <> False Then
			bReturn = False
		End If
		If mp_sCrossProjectName <> "" Then
			bReturn = False
		End If
		If mp_lLinkLag <> 0 Then
			bReturn = False
		End If
		If mp_yLagFormat <> E_LAGFORMAT.LF_M Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<PredecessorLink/>"
		End if
		Dim oXML As New clsXML("PredecessorLink")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		oXML.WriteProperty("PredecessorUID", mp_lPredecessorUID)
		oXML.WriteProperty("Type", mp_yType)
		oXML.WriteProperty("CrossProject", mp_bCrossProject)
		If mp_sCrossProjectName <> "" Then
			oXML.WriteProperty("CrossProjectName", mp_sCrossProjectName)
		End If
		oXML.WriteProperty("LinkLag", mp_lLinkLag)
		oXML.WriteProperty("LagFormat", mp_yLagFormat)
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("PredecessorLink")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("PredecessorUID", mp_lPredecessorUID)
		oXML.ReadProperty("Type", mp_yType)
		oXML.ReadProperty("CrossProject", mp_bCrossProject)
		oXML.ReadProperty("CrossProjectName", mp_sCrossProjectName)
		oXML.ReadProperty("LinkLag", mp_lLinkLag)
		oXML.ReadProperty("LagFormat", mp_yLagFormat)
	End Sub

End Class
