Option Explicit On

Public Class Calendars

	Private mp_oCollection As clsCollectionBase

	Public Sub New()
		mp_oCollection = New clsCollectionBase("Calendar")
	End Sub

	Public ReadOnly Property Count() As Integer
		Get
			Return mp_oCollection.m_lCount
		End Get
	End Property

	Public Function Item(ByVal Index As String) As Calendar
		Return mp_oCollection.m_oItem(Index, SYS_ERRORS.MP_ITEM_1, SYS_ERRORS.MP_ITEM_2, SYS_ERRORS.MP_ITEM_3, SYS_ERRORS.MP_ITEM_4)
	End Function

	Public Function Add() As Calendar
		mp_oCollection.AddMode = True
		Dim oCalendar As New Calendar()
		oCalendar.mp_oCollection = mp_oCollection
		mp_oCollection.m_Add(oCalendar, "", SYS_ERRORS.MP_ADD_1, SYS_ERRORS.MP_ADD_2, False, SYS_ERRORS.MP_ADD_3)
		Return oCalendar
	End Function

	Public Sub Clear()
		mp_oCollection.m_Clear()
	End Sub

	Public Sub Remove(ByVal Index As String)
		mp_oCollection.m_Remove(Index, SYS_ERRORS.MP_REMOVE_1, SYS_ERRORS.MP_REMOVE_2, SYS_ERRORS.MP_REMOVE_3, SYS_ERRORS.MP_REMOVE_4)
	End Sub

	Public Function IsNull() As Boolean
		Dim bReturn As Boolean = True
		If Count > 0 Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<Calendars/>"
		End if
		Dim lIndex As Integer
		Dim oCalendar As Calendar
		Dim oXML As New clsXML("Calendars")
		oXML.BoolsAreNumeric = True
		oXML.InitializeWriter()
		For lIndex = 1 To Count
			oCalendar = mp_oCollection.m_oReturnArrayElement(lIndex)
			oXML.WriteObject(oCalendar.GetXML)
		Next lIndex
		Return oXML.GetXML
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim lIndex As Integer
		Dim oXML As New clsXML("Calendars")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		mp_oCollection.m_Clear()
		If oXML.ReadCollectionCount = 0 Then
			Return
		End If
		For lIndex = 1 To oXML.ReadCollectionCount
			Dim oCalendar As New Calendar()
			oCalendar.SetXML(oXML.ReadCollectionObject(lIndex))
			mp_oCollection.AddMode = True
			Dim sKey As String = ""
			sKey = "K" & oCalendar.lUID.ToString()
				oCalendar.mp_oCollection = mp_oCollection
				oCalendar.Key = sKey
			mp_oCollection.m_Add(oCalendar, sKey, SYS_ERRORS.MP_ADD_1, SYS_ERRORS.MP_ADD_2, False, SYS_ERRORS.MP_ADD_3)
			oCalendar = Nothing
		Next lIndex
	End Sub

End Class
