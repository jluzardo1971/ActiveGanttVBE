Option Explicit On

Public Class WeekDay
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_yDayType As E_DAYTYPE
	Private mp_bDayWorking As Boolean

	Public Sub New()
		mp_yDayType = E_DAYTYPE.DT_EXCEPTION
		mp_bDayWorking = False
	End Sub

	Public Property yDayType() As E_DAYTYPE
		Get
			Return mp_yDayType
		End Get
		Set(ByVal Value As E_DAYTYPE)
			mp_yDayType = Value
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
		If mp_yDayType <> E_DAYTYPE.DT_EXCEPTION Then
			bReturn = False
		End If
		If mp_bDayWorking <> False Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<WeekDay/>"
		End if
		Dim oXML As New clsXML("WeekDay")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		oXML.WriteProperty("DayType", mp_yDayType)
		oXML.WriteProperty("DayWorking", mp_bDayWorking)
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("WeekDay")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("DayType", mp_yDayType)
		oXML.ReadProperty("DayWorking", mp_bDayWorking)
	End Sub

End Class
