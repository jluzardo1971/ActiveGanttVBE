Option Explicit On

Public Class TaskExtendedAttribute_C

	Private mp_oCollection As clsCollectionBase

	Public Sub New()
		mp_oCollection = New clsCollectionBase("TaskExtendedAttribute")
	End Sub

	Public ReadOnly Property Count() As Integer
		Get
			Return mp_oCollection.m_lCount
		End Get
	End Property

	Public Function Item(ByVal Index As String) As TaskExtendedAttribute
		Return mp_oCollection.m_oItem(Index, SYS_ERRORS.MP_ITEM_1, SYS_ERRORS.MP_ITEM_2, SYS_ERRORS.MP_ITEM_3, SYS_ERRORS.MP_ITEM_4)
	End Function

	Public Function Add() As TaskExtendedAttribute
		mp_oCollection.AddMode = True
		Dim oTaskExtendedAttribute As New TaskExtendedAttribute()
		oTaskExtendedAttribute.mp_oCollection = mp_oCollection
		mp_oCollection.m_Add(oTaskExtendedAttribute, "", SYS_ERRORS.MP_ADD_1, SYS_ERRORS.MP_ADD_2, False, SYS_ERRORS.MP_ADD_3)
		Return oTaskExtendedAttribute
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

	Friend Sub ReadObjectProtected(ByRef oXML As clsXML)
		Dim lIndex As Integer
		For lIndex = 1 To oXML.ReadCollectionCount
			If oXML.GetCollectionObjectName(lIndex) = "ExtendedAttribute" Then
				Dim oTaskExtendedAttribute As New TaskExtendedAttribute()
				oTaskExtendedAttribute.SetXML(oXML.ReadCollectionObject(lIndex))
				mp_oCollection.AddMode = True
				Dim sKey As String = ""
				oTaskExtendedAttribute.mp_oCollection = mp_oCollection
				mp_oCollection.m_Add(oTaskExtendedAttribute, sKey, SYS_ERRORS.MP_ADD_1, SYS_ERRORS.MP_ADD_2, False, SYS_ERRORS.MP_ADD_3)
				oTaskExtendedAttribute = Nothing
			End If
		Next
	End Sub

	Friend Sub WriteObjectProtected(ByRef oXML As clsXML)
		Dim lIndex As Integer
		Dim oTaskExtendedAttribute As TaskExtendedAttribute
		For lIndex = 1 To Count
		oTaskExtendedAttribute = mp_oCollection.m_oReturnArrayElement(lIndex)
		oXML.WriteObject(oTaskExtendedAttribute.GetXML)
		Next lIndex
	End Sub

End Class
