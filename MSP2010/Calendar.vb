Option Explicit On

Public Class Calendar
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_lUID As Integer
	Private mp_sName As String
	Private mp_bIsBaseCalendar As Boolean
	Private mp_bIsBaselineCalendar As Boolean
	Private mp_lBaseCalendarUID As Integer
	Private mp_oWeekDays As CalendarWeekDays
	Private mp_oExceptions As CalendarExceptions
	Private mp_oWorkWeeks As CalendarWorkWeeks

	Public Sub New()
		mp_lUID = 0
		mp_sName = ""
		mp_bIsBaseCalendar = False
		mp_bIsBaselineCalendar = False
		mp_lBaseCalendarUID = 0
		mp_oWeekDays = New CalendarWeekDays()
		mp_oExceptions = New CalendarExceptions()
		mp_oWorkWeeks = New CalendarWorkWeeks()
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

	Public Property bIsBaselineCalendar() As Boolean
		Get
			Return mp_bIsBaselineCalendar
		End Get
		Set(ByVal Value As Boolean)
			mp_bIsBaselineCalendar = Value
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

	Public ReadOnly Property oWeekDays() As CalendarWeekDays
		Get
			Return mp_oWeekDays
		End Get
	End Property

	Public ReadOnly Property oExceptions() As CalendarExceptions
		Get
			Return mp_oExceptions
		End Get
	End Property

	Public ReadOnly Property oWorkWeeks() As CalendarWorkWeeks
		Get
			Return mp_oWorkWeeks
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
		If mp_bIsBaselineCalendar <> False Then
			bReturn = False
		End If
		If mp_lBaseCalendarUID <> 0 Then
			bReturn = False
		End If
		If mp_oWeekDays.IsNull() = False Then
			bReturn = False
		End If
		If mp_oExceptions.IsNull() = False Then
			bReturn = False
		End If
		If mp_oWorkWeeks.IsNull() = False Then
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
		oXML.WriteProperty("IsBaselineCalendar", mp_bIsBaselineCalendar)
		oXML.WriteProperty("BaseCalendarUID", mp_lBaseCalendarUID)
		If mp_oWeekDays.IsNull() = False Then
			oXML.WriteObject(mp_oWeekDays.GetXML())
		End If
		If mp_oExceptions.IsNull() = False Then
			oXML.WriteObject(mp_oExceptions.GetXML())
		End If
		If mp_oWorkWeeks.IsNull() = False Then
			oXML.WriteObject(mp_oWorkWeeks.GetXML())
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
		oXML.ReadProperty("IsBaselineCalendar", mp_bIsBaselineCalendar)
		oXML.ReadProperty("BaseCalendarUID", mp_lBaseCalendarUID)
		mp_oWeekDays.SetXML(oXML.ReadObject("WeekDays"))
		mp_oExceptions.SetXML(oXML.ReadObject("Exceptions"))
		mp_oWorkWeeks.SetXML(oXML.ReadObject("WorkWeeks"))
	End Sub

End Class
