Option Explicit On

Public Class CalendarException
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_bEnteredByOccurrences As Boolean
	Private mp_oTimePeriod As TimePeriod
	Private mp_lOccurrences As Integer
	Private mp_sName As String
	Private mp_yType As E_TYPE_3
	Private mp_lPeriod As Integer
	Private mp_lDaysOfWeek As Integer
	Private mp_yMonthItem As E_MONTHITEM
	Private mp_yMonthPosition As E_MONTHPOSITION
	Private mp_yMonth As E_MONTH
	Private mp_lMonthDay As Integer
	Private mp_bDayWorking As Boolean
	Private mp_oWorkingTimes As WorkingTimes

	Public Sub New()
		mp_bEnteredByOccurrences = False
		mp_oTimePeriod = New TimePeriod()
		mp_lOccurrences = 0
		mp_sName = ""
		mp_yType = E_TYPE_3.T_3_DAILY
		mp_lPeriod = 0
		mp_lDaysOfWeek = 0
		mp_yMonthItem = E_MONTHITEM.MI_DAY
		mp_yMonthPosition = E_MONTHPOSITION.MP_FIRST_POSITION
		mp_yMonth = E_MONTH.M_JANUARY
		mp_lMonthDay = 0
		mp_bDayWorking = False
		mp_oWorkingTimes = New WorkingTimes()
	End Sub

	Public Property bEnteredByOccurrences() As Boolean
		Get
			Return mp_bEnteredByOccurrences
		End Get
		Set(ByVal Value As Boolean)
			mp_bEnteredByOccurrences = Value
		End Set
	End Property

	Public ReadOnly Property oTimePeriod() As TimePeriod
		Get
			Return mp_oTimePeriod
		End Get
	End Property

	Public Property lOccurrences() As Integer
		Get
			Return mp_lOccurrences
		End Get
		Set(ByVal Value As Integer)
			mp_lOccurrences = Value
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

	Public Property yType() As E_TYPE_3
		Get
			Return mp_yType
		End Get
		Set(ByVal Value As E_TYPE_3)
			mp_yType = Value
		End Set
	End Property

	Public Property lPeriod() As Integer
		Get
			Return mp_lPeriod
		End Get
		Set(ByVal Value As Integer)
			mp_lPeriod = Value
		End Set
	End Property

	Public Property lDaysOfWeek() As Integer
		Get
			Return mp_lDaysOfWeek
		End Get
		Set(ByVal Value As Integer)
			mp_lDaysOfWeek = Value
		End Set
	End Property

	Public Property yMonthItem() As E_MONTHITEM
		Get
			Return mp_yMonthItem
		End Get
		Set(ByVal Value As E_MONTHITEM)
			mp_yMonthItem = Value
		End Set
	End Property

	Public Property yMonthPosition() As E_MONTHPOSITION
		Get
			Return mp_yMonthPosition
		End Get
		Set(ByVal Value As E_MONTHPOSITION)
			mp_yMonthPosition = Value
		End Set
	End Property

	Public Property yMonth() As E_MONTH
		Get
			Return mp_yMonth
		End Get
		Set(ByVal Value As E_MONTH)
			mp_yMonth = Value
		End Set
	End Property

	Public Property lMonthDay() As Integer
		Get
			Return mp_lMonthDay
		End Get
		Set(ByVal Value As Integer)
			mp_lMonthDay = Value
		End Set
	End Property

	Public Property bDayWorking() As Boolean
		Get
			Return mp_bDayWorking
		End Get
		Set(ByVal Value As Boolean)
			mp_bDayWorking = Value
		End Set
	End Property

	Public ReadOnly Property oWorkingTimes() As WorkingTimes
		Get
			Return mp_oWorkingTimes
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
		If mp_bEnteredByOccurrences <> False Then
			bReturn = False
		End If
		If mp_oTimePeriod.IsNull() = False Then
			bReturn = False
		End If
		If mp_lOccurrences <> 0 Then
			bReturn = False
		End If
		If mp_sName <> "" Then
			bReturn = False
		End If
		If mp_yType <> E_TYPE_3.T_3_DAILY Then
			bReturn = False
		End If
		If mp_lPeriod <> 0 Then
			bReturn = False
		End If
		If mp_lDaysOfWeek <> 0 Then
			bReturn = False
		End If
		If mp_yMonthItem <> E_MONTHITEM.MI_DAY Then
			bReturn = False
		End If
		If mp_yMonthPosition <> E_MONTHPOSITION.MP_FIRST_POSITION Then
			bReturn = False
		End If
		If mp_yMonth <> E_MONTH.M_JANUARY Then
			bReturn = False
		End If
		If mp_lMonthDay <> 0 Then
			bReturn = False
		End If
		If mp_bDayWorking <> False Then
			bReturn = False
		End If
		If mp_oWorkingTimes.IsNull() = False Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<Exception/>"
		End if
		Dim oXML As New clsXML("Exception")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		oXML.WriteProperty("EnteredByOccurrences", mp_bEnteredByOccurrences)
		If mp_oTimePeriod.IsNull() = False Then
			oXML.WriteObject(mp_oTimePeriod.GetXML())
		End If
		oXML.WriteProperty("Occurrences", mp_lOccurrences)
		If mp_sName <> "" Then
			oXML.WriteProperty("Name", mp_sName)
		End If
		oXML.WriteProperty("Type", mp_yType)
		oXML.WriteProperty("Period", mp_lPeriod)
		oXML.WriteProperty("DaysOfWeek", mp_lDaysOfWeek)
		oXML.WriteProperty("MonthItem", mp_yMonthItem)
		oXML.WriteProperty("MonthPosition", mp_yMonthPosition)
		oXML.WriteProperty("Month", mp_yMonth)
		oXML.WriteProperty("MonthDay", mp_lMonthDay)
		oXML.WriteProperty("DayWorking", mp_bDayWorking)
		If mp_oWorkingTimes.IsNull() = False Then
			oXML.WriteObject(mp_oWorkingTimes.GetXML())
		End If
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("Exception")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("EnteredByOccurrences", mp_bEnteredByOccurrences)
		mp_oTimePeriod.SetXML(oXML.ReadObject("TimePeriod"))
		oXML.ReadProperty("Occurrences", mp_lOccurrences)
		oXML.ReadProperty("Name", mp_sName)
		If mp_sName.Length > 512 Then
			mp_sName = mp_sName.Substring(0, 512)
		End If
		oXML.ReadProperty("Type", mp_yType)
		oXML.ReadProperty("Period", mp_lPeriod)
		oXML.ReadProperty("DaysOfWeek", mp_lDaysOfWeek)
		oXML.ReadProperty("MonthItem", mp_yMonthItem)
		oXML.ReadProperty("MonthPosition", mp_yMonthPosition)
		oXML.ReadProperty("Month", mp_yMonth)
		oXML.ReadProperty("MonthDay", mp_lMonthDay)
		oXML.ReadProperty("DayWorking", mp_bDayWorking)
		mp_oWorkingTimes.SetXML(oXML.ReadObject("WorkingTimes"))
	End Sub

End Class
