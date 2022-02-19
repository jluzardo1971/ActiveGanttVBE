Option Explicit On

Public Class CalendarWorkWeek
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_oTimePeriod As TimePeriod
	Private mp_sName As String
	Private mp_oWeekDay_C As WeekDay_C

	Public Sub New()
		mp_oTimePeriod = New TimePeriod()
		mp_sName = ""
		mp_oWeekDay_C = New WeekDay_C()
	End Sub

	Public ReadOnly Property oTimePeriod() As TimePeriod
		Get
			Return mp_oTimePeriod
		End Get
	End Property

	Public Property sName() As String
		Get
			Return mp_sName
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sName = Value
		End Set
	End Property

	Public ReadOnly Property oWeekDay_C() As WeekDay_C
		Get
			Return mp_oWeekDay_C
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
		If mp_oTimePeriod.IsNull() = False Then
			bReturn = False
		End If
		If mp_sName <> "" Then
			bReturn = False
		End If
		If mp_oWeekDay_C.IsNull() = False Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<WorkWeek/>"
		End if
		Dim oXML As New clsXML("WorkWeek")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		If mp_oTimePeriod.IsNull() = False Then
			oXML.WriteObject(mp_oTimePeriod.GetXML())
		End If
		If mp_sName <> "" Then
			oXML.WriteProperty("Name", mp_sName)
		End If
		If mp_oWeekDay_C.IsNull() = False Then
			mp_oWeekDay_C.WriteObjectProtected(oXML)
		End If
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("WorkWeek")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		mp_oTimePeriod.SetXML(oXML.ReadObject("TimePeriod"))
		oXML.ReadProperty("Name", mp_sName)
		If mp_sName.Length > 512 Then
			mp_sName = mp_sName.Substring(0, 512)
		End If
		mp_oWeekDay_C.ReadObjectProtected(oXML)
	End Sub

End Class
