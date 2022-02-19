Option Explicit On

Public Class Calendar
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_lUID As Integer
	Private mp_sName As String
	Private mp_bIsBaseCalendar As Boolean
	Private mp_lBaseCalendarUID As Integer
	Private mp_oWeekDays As WeekDays

	Public Sub New()
		mp_lUID = 0
		mp_sName = ""
		mp_bIsBaseCalendar = False
		mp_lBaseCalendarUID = 0
		mp_oWeekDays = New WeekDays()
	End Sub

	Public Property lUID() As Integer
		Get
			Return mp_lUID
		End Get
		Set(ByVal Value As Integer)
			mp_lUID = Value
		End Set
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

	Public Property bIsBaseCalendar() As Boolean
		Get
			Return mp_bIsBaseCalendar
		End Get
		Set(ByVal Value As Boolean)
			mp_bIsBaseCalendar = Value
		End Set
	End Property

	Public Property lBaseCalendarUID() As Integer
		Get
			Return mp_lBaseCalendarUID
		End Get
		Set(ByVal Value As Integer)
			mp_lBaseCalendarUID = Value
		End Set
	End Property

	Public ReadOnly Property oWeekDays() As WeekDays
		Get
			Return mp_oWeekDays
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
		If mp_lUID <> 0 Then
			bReturn = False
		End If
		If mp_sName <> "" Then
			bReturn = False
		End If
		If mp_bIsBaseCalendar <> False Then
			bReturn = False
		End If
		If mp_lBaseCalendarUID <> 0 Then
			bReturn = False
		End If
		If mp_oWeekDays.IsNull() = False Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<Calendar/>"
		End if
		Dim oXML As New clsXML("Calendar")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		oXML.WriteProperty("UID", mp_lUID)
		If mp_sName <> "" Then
			oXML.WriteProperty("Name", mp_sName)
		End If
		oXML.WriteProperty("IsBaseCalendar", mp_bIsBaseCalendar)
		oXML.WriteProperty("BaseCalendarUID", mp_lBaseCalendarUID)
		If mp_oWeekDays.IsNull() = False Then
			oXML.WriteObject(mp_oWeekDays.GetXML())
		End If
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("Calendar")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("UID", mp_lUID)
		oXML.ReadProperty("Name", mp_sName)
		If mp_sName.Length > 512 Then
			mp_sName = mp_sName.Substring(0, 512)
		End If
		oXML.ReadProperty("IsBaseCalendar", mp_bIsBaseCalendar)
		oXML.ReadProperty("BaseCalendarUID", mp_lBaseCalendarUID)
		mp_oWeekDays.SetXML(oXML.ReadObject("WeekDays"))
	End Sub

End Class
