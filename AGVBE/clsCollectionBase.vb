Option Explicit On 

Friend Class clsCollectionBase

    Private mp_oControl As ActiveGanttVBECtl
    Private m_sObjectName As String
    Friend mp_aoCollection As ArrayList
    Private mp_bAddMode As Boolean
    Private mp_bDescending As Boolean
    Private mp_bIgnoreKeyChecks As Boolean
    Private mp_bSortCells As Boolean
    Private mp_lCellIndex As Integer
    Friend mp_oKeys As clsDictionary
    Private mp_sPropertyName As String
    Private mp_ySortType As E_SORTTYPE

    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal sObjectName As String)
        mp_oControl = Value
        mp_aoCollection = New ArrayList()
        mp_oKeys = New clsDictionary()
        m_sObjectName = sObjectName
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public ReadOnly Property m_oItem(ByVal v_lIndex As String, ByVal v_lErr1 As Integer, ByVal v_lErr2 As Integer, ByVal v_lErr3 As Integer, ByVal v_lErr4 As Integer) As Object
        Get
            Dim lIndex As Integer
            If Not g_IsNumeric(v_lIndex) Then
                If v_lIndex = "" Then
                    mp_oControl.mp_ErrorReport(v_lErr1, "Invalid " & m_sObjectName & " object key, key cannot be an empty string", "ActiveGanttVBECtl.cls" & m_sObjectName & "s.Get m_oItem")
                    Return Nothing
                End If
                lIndex = m_lFindIndexByKey(v_lIndex)
                If lIndex = -1 Then
                    mp_oControl.mp_ErrorReport(v_lErr2, m_sObjectName & " object not found, invalid key (""" & v_lIndex & """)", "ActiveGanttVBECtl.cls" & m_sObjectName & "s.Get m_oItem")
                    Return Nothing
                End If
            Else
                lIndex = System.Convert.ToInt32(v_lIndex)
                If System.Convert.ToString(lIndex) <> v_lIndex Then
                    mp_oControl.mp_ErrorReport(v_lErr3, m_sObjectName & " object not found, invalid Index", "ActiveGanttVBECtl.cls" & m_sObjectName & "s.Get m_oItem")
                    Return Nothing
                End If
                If lIndex < 1 Or lIndex > mp_aoCollection.Count() Then
                    mp_oControl.mp_ErrorReport(v_lErr4, m_sObjectName & " object not found, Index out of bounds", "ActiveGanttVBECtl.cls" & m_sObjectName & "s.Get m_oItem")
                    Return Nothing
                End If
            End If
            Return mp_aoCollection(lIndex - 1)
        End Get
    End Property

    Public Property m_bIgnoreKeyChecks() As Boolean
        Get
            Return mp_bIgnoreKeyChecks
        End Get
        Set(ByVal Value As Boolean)
            mp_bIgnoreKeyChecks = Value
        End Set
    End Property

    Public Sub m_Add(ByVal r_oObject As Object, ByVal v_sKey As String, ByVal v_lErr1 As Integer, ByVal v_lErr2 As Integer, Optional ByVal v_bKeyRequired As Boolean = False, Optional ByVal v_lKeyError As Integer = 0)
        Dim oItemBase As clsItemBase
        oItemBase = DirectCast(r_oObject, clsItemBase)
        Dim lUpperBounds As Integer
        If mp_bAddMode = False Then
            mp_oControl.mp_ErrorReport(SYS_ERRORS.ERR_ADDMODE_G, "AddMode must be set to true before executing oCollection.m_Add", "cls" & m_sObjectName & "s")
        End If
        If g_IsNumeric(v_sKey) Then
            mp_oControl.mp_ErrorReport(v_lErr1, "Invalid " & m_sObjectName & " object key, key cannot be numeric", "ActiveGanttVBECtl.cls" & m_sObjectName & "s.Add")
            Return
        End If
        If v_sKey <> "" Then
            If m_bIsKeyUnique(v_sKey) = False Then
                mp_oControl.mp_ErrorReport(v_lErr2, "Key is not unique in " & m_sObjectName & "s collection", "ActiveGanttVBECtl.cls" & m_sObjectName & "s.Add")
                Return
            End If
        End If
        If v_sKey = "" And v_bKeyRequired = True Then
            mp_oControl.mp_ErrorReport(v_lKeyError, "A Key is required for all objects in the " & m_sObjectName & "s collection", "ActiveGanttVBECtl.cls" & m_sObjectName & "s.Add")
            Return
        End If
        lUpperBounds = mp_aoCollection.Count() + 1
        oItemBase.Index = lUpperBounds
        mp_aoCollection.Add(r_oObject)
        If v_sKey <> "" Then
            mp_oKeys.Add(lUpperBounds, v_sKey)
        End If
        mp_bAddMode = False
    End Sub

    Public Function m_lCount() As Integer
        Return mp_aoCollection.Count()
    End Function

    Public Sub m_Clear()
        mp_aoCollection.Clear()
        m_ReconstKeys()
    End Sub

    Public Sub m_Remove(ByVal v_sIndex As String, ByVal v_lErr1 As Integer, ByVal v_lErr2 As Integer, ByVal v_lErr3 As Integer, ByVal v_lErr4 As Integer)
        Dim lIndex As Integer
        Dim lRemovedIndex As Integer
        If Not g_IsNumeric(v_sIndex) Then
            If v_sIndex = "" Then
                mp_oControl.mp_ErrorReport(v_lErr1, "Invalid " & m_sObjectName & " object key, key cannot be an empty string", "ActiveGanttVBECtl.cls" & m_sObjectName & "s.Remove")
                Return
            End If
            lIndex = m_lFindIndexByKey(v_sIndex)
            If lIndex = -1 Then
                mp_oControl.mp_ErrorReport(v_lErr2, m_sObjectName & " object not found, invalid key (""" & v_sIndex & """)", "ActiveGanttVBECtl.cls" & m_sObjectName & "s.Remove")
                Return
            End If
            lRemovedIndex = lIndex
        Else
            lIndex = System.Convert.ToInt32(v_sIndex)
            If System.Convert.ToString(lIndex) <> v_sIndex Then
                mp_oControl.mp_ErrorReport(v_lErr3, m_sObjectName & " object not found, invalid Index", "ActiveGanttVBECtl.cls" & m_sObjectName & "s.Remove")
                Return
            End If
            If lIndex < 1 Or lIndex > mp_aoCollection.Count() Then
                mp_oControl.mp_ErrorReport(v_lErr4, m_sObjectName & " object not found, Index out of bounds", "ActiveGanttVBECtl.cls" & m_sObjectName & "s.Remove")
                Return
            End If
            lRemovedIndex = lIndex
        End If
        mp_bIgnoreKeyChecks = True
        mp_aoCollection.RemoveAt(lRemovedIndex - 1)
        mp_bIgnoreKeyChecks = False
        m_ReconstKeys()
    End Sub

    Public Function m_lCopyAndMoveItems(ByVal v_lOriginIndex As Integer, ByVal v_lDestinationIndex As Integer) As Integer
        Dim Buffer As Object
        Dim lIndex As Integer
        mp_bIgnoreKeyChecks = True
        v_lOriginIndex = v_lOriginIndex - 1
        v_lDestinationIndex = v_lDestinationIndex - 1
        Buffer = mp_aoCollection(v_lOriginIndex)
        If v_lOriginIndex > v_lDestinationIndex Then
            For lIndex = v_lOriginIndex To v_lDestinationIndex + 1 Step -1
                mp_aoCollection(lIndex) = mp_aoCollection(lIndex - 1)
            Next lIndex
        Else
            For lIndex = v_lOriginIndex To v_lDestinationIndex - 1
                mp_aoCollection(lIndex) = mp_aoCollection(lIndex + 1)
            Next lIndex
        End If
        mp_aoCollection(v_lDestinationIndex) = Buffer
        mp_bIgnoreKeyChecks = False
        m_ReconstKeys()
        Return v_lDestinationIndex + 1
    End Function

    Public Sub m_ReconstKeys()
        Dim lIndex As Integer
        Dim lCount As Integer
        Dim sKey As String
        mp_oKeys = Nothing
        mp_oKeys = New clsDictionary()
        lCount = mp_aoCollection.Count()
        For lIndex = 1 To lCount
            sKey = mp_aoCollection(lIndex - 1).Key
            mp_aoCollection(lIndex - 1).Index = lIndex
            If sKey <> "" Then
                mp_oKeys.Add(lIndex, sKey)
            End If
        Next lIndex
    End Sub

    Public Function m_lFindIndexByKey(ByVal v_sKey As String) As Integer
        If mp_oKeys.Contains(v_sKey) = True Then
            Return mp_oKeys.Item(v_sKey)
        Else
            Return -1
        End If
    End Function

    Public Function m_bIsKeyUnique(ByVal v_sKey As String) As Boolean
        Return Not mp_oKeys.Contains(v_sKey)
    End Function

    Public Function m_bDoesKeyExist(ByVal v_sKey As String) As Boolean
        Return mp_oKeys.Contains(v_sKey)
    End Function

    Public Function m_oReturnArrayElement(ByVal r_lIndex As Integer) As Object
        Return mp_aoCollection(r_lIndex - 1)
    End Function

    Public Function m_oReturnArrayElementKey(ByVal v_sKey As String) As Object
        Dim lIndex As Integer
        If g_IsNumeric(v_sKey) Then
            Return mp_aoCollection(System.Convert.ToInt32(v_sKey) - 1)
        Else
            lIndex = m_lFindIndexByKey(v_sKey)
            If lIndex <> -1 Then
                Return mp_aoCollection(lIndex - 1)
            Else
                mp_oControl.mp_ErrorReport(SYS_ERRORS.ERR_RETARRELEMKEY_G, "Key not found", "ActiveGanttVBECtl.cls" & m_sObjectName & "s.m_oReturnArrayElementKey")
            End If
        End If
        Return Nothing
    End Function

    Public Sub m_Sort(ByVal sPropertyName As String, ByVal bDescending As Boolean, ByVal SortType As E_SORTTYPE, ByVal StartIndex As Integer, ByVal EndIndex As Integer)
        Dim aTempArray As New ArrayList
        Dim lIndex As Integer
        For lIndex = 1 To EndIndex
            aTempArray.Add(Nothing)
        Next
        mp_sPropertyName = sPropertyName
        mp_bDescending = bDescending
        mp_ySortType = SortType
        mp_bSortCells = False
        mp_Sort(mp_aoCollection, aTempArray, StartIndex - 1, EndIndex - 1)
        m_ReconstKeys()
    End Sub

    Public Sub m_SortCells(ByVal CellIndex As Integer, ByVal sPropertyName As String, ByVal bDescending As Boolean, ByVal SortType As E_SORTTYPE, ByVal StartIndex As Integer, ByVal EndIndex As Integer)
        Dim aTempArray As New ArrayList
        Dim lIndex As Integer
        For lIndex = 1 To EndIndex
            aTempArray.Add(Nothing)
        Next
        mp_sPropertyName = sPropertyName
        mp_bDescending = bDescending
        mp_ySortType = SortType
        mp_bSortCells = True
        mp_lCellIndex = CellIndex
        mp_Sort(mp_aoCollection, aTempArray, StartIndex - 1, EndIndex - 1)
        m_ReconstKeys()
    End Sub

    Private Sub mp_Sort(ByVal r_aSortArray As ArrayList, ByVal r_aTempArray As ArrayList, Optional ByVal first As Integer = -1, Optional ByVal last As Integer = -1)
        Dim lArrayMBound As Integer
        Dim lArrayLBound As Integer
        Dim lArrayUBound As Integer
        If first = -1 Then
            lArrayLBound = 0
        Else
            lArrayLBound = first
        End If
        If last = -1 Then
            lArrayUBound = r_aSortArray.Count
        Else
            lArrayUBound = last
        End If
        If lArrayUBound > lArrayLBound Then
            lArrayMBound = (lArrayUBound + lArrayLBound) \ 2
            mp_Sort(r_aSortArray, r_aTempArray, lArrayLBound, lArrayMBound)
            mp_Sort(r_aSortArray, r_aTempArray, lArrayMBound + 1, lArrayUBound)
            mp_Merge(r_aSortArray, r_aTempArray, lArrayLBound, lArrayMBound + 1, lArrayUBound)
        End If
    End Sub

    Private Sub mp_Merge(ByVal r_aSortArray As ArrayList, ByVal r_aTempArray As ArrayList, ByVal first As Integer, ByVal mid As Integer, ByVal last As Integer)
        Dim i As Integer
        Dim iLow As Integer
        Dim nNumElements As Integer
        Dim iTempPos As Integer
        iLow = mid - 1
        iTempPos = first
        nNumElements = last - first + 1
        Do While first <= iLow And mid <= last
            If mp_bDescending = False Then
                If mp_oGetProperty(r_aSortArray(first)) <= mp_oGetProperty(r_aSortArray(mid)) Then
                    r_aTempArray(iTempPos) = r_aSortArray(first)
                    iTempPos = iTempPos + 1
                    first = first + 1
                Else
                    r_aTempArray(iTempPos) = r_aSortArray(mid)
                    iTempPos = iTempPos + 1
                    mid = mid + 1
                End If
            Else
                If mp_oGetProperty(r_aSortArray(first)) >= mp_oGetProperty(r_aSortArray(mid)) Then
                    r_aTempArray(iTempPos) = r_aSortArray(first)
                    iTempPos = iTempPos + 1
                    first = first + 1
                Else
                    r_aTempArray(iTempPos) = r_aSortArray(mid)
                    iTempPos = iTempPos + 1
                    mid = mid + 1
                End If
            End If
        Loop
        Do While first <= iLow
            r_aTempArray(iTempPos) = r_aSortArray(first)
            first = first + 1
            iTempPos = iTempPos + 1
        Loop
        Do While mid <= last
            r_aTempArray(iTempPos) = r_aSortArray(mid)
            mid = mid + 1
            iTempPos = iTempPos + 1
        Loop
        For i = 0 To nNumElements - 1
            r_aSortArray(last) = r_aTempArray(last)
            last = last - 1
        Next i
    End Sub

    Private Function mp_oGetProperty(ByVal obj As Object) As Object
        If mp_bSortCells = False Then
            Select Case mp_ySortType
                Case E_SORTTYPE.ES_NUMERIC
                    Return System.Convert.ToInt32(CallByName(obj, mp_sPropertyName))
                Case E_SORTTYPE.ES_STRING
                    Return System.Convert.ToString(CallByName(obj, mp_sPropertyName))
                Case E_SORTTYPE.ES_DATE
                    Return CDate(CallByName(obj, mp_sPropertyName))
            End Select
        Else
            Select Case mp_ySortType
                Case E_SORTTYPE.ES_NUMERIC
                    Return System.Convert.ToInt32(CallByName(obj.Cell(mp_lCellIndex), mp_sPropertyName))
                Case E_SORTTYPE.ES_STRING
                    Return System.Convert.ToString(CallByName(obj.Cell(mp_lCellIndex), mp_sPropertyName))
                Case E_SORTTYPE.ES_DATE
                    Return CDate(CallByName(obj.Cell(mp_lCellIndex), mp_sPropertyName))
            End Select
        End If
        Return Nothing
    End Function

    Public Function m_lReturnIndex(ByVal v_sIndex As String, ByVal bIncludeDefault As Boolean) As Integer
        Dim lIndex As Integer
        Dim lReturn As Integer
        If (g_IsNumeric(v_sIndex)) Then
            lIndex = System.Convert.ToInt32(v_sIndex)
            If (bIncludeDefault = True) Then
                If (lIndex >= 0 And lIndex <= m_lCount()) Then
                    lReturn = lIndex
                Else
                    lReturn = -1
                End If
            Else
                If (lIndex >= 1 And lIndex <= m_lCount()) Then
                    lReturn = lIndex
                Else
                    lReturn = -1
                End If
            End If
        Else
            lReturn = m_lFindIndexByKey(v_sIndex)
        End If
        Return lReturn
    End Function

    Public Sub mp_SetKey(ByRef sCurrentKey As String, ByVal sNewKey As String, ByVal ErrNumber As Integer)
        If m_bIgnoreKeyChecks = False Then
            If g_IsNumeric(sNewKey) Or (sNewKey <> sCurrentKey And m_bIsKeyUnique(sNewKey) = False) Then
                mp_oControl.mp_ErrorReport(ErrNumber, "Numeric or duplicate " & m_sObjectName & " object key", "ActiveGanttVBECtl.cls" & m_sObjectName & ".Let Key")
                Return
            End If
        End If
        sCurrentKey = sNewKey
        If mp_bAddMode = False Then
            m_ReconstKeys()
        End If
    End Sub

    Public Property AddMode() As Boolean
        Get
            Return mp_bAddMode
        End Get
        Set(ByVal Value As Boolean)
            mp_bAddMode = Value
        End Set
    End Property

    Public Sub m_GetKeyAndIndex(ByVal sIndex As String, ByRef sKey As String, ByRef sReturnIndex As String)
        Dim oObject As Object
        oObject = m_oItem(sIndex, SYS_ERRORS.GETINDEXANDKEY_ITEM1, SYS_ERRORS.GETINDEXANDKEY_ITEM2, SYS_ERRORS.GETINDEXANDKEY_ITEM3, SYS_ERRORS.GETINDEXANDKEY_ITEM4)
        If oObject.Key <> "" Then
            sKey = oObject.Key
        Else
            sKey = System.Convert.ToString(oObject.Index)
        End If
        sReturnIndex = System.Convert.ToString(oObject.Index)
    End Sub

    Public Sub m_CollectionRemoveWhere(ByVal sPropertyName As String, ByVal sKey As String, ByVal sIndex As String)
        Dim lIndex As Integer
        Dim oObject As Object
        Dim sPropertyValue As String
        For lIndex = m_lCount() To 1 Step -1
            oObject = m_oReturnArrayElement(lIndex)
            sPropertyValue = CallByName(oObject, sPropertyName)
            If sPropertyValue = sKey Or sPropertyValue = sIndex Then
                m_Remove(lIndex.ToString(), SYS_ERRORS.ERR_COLLREMWHERE_1_G, SYS_ERRORS.ERR_COLLREMWHERE_2_G, SYS_ERRORS.ERR_COLLREMWHERE_3_G, SYS_ERRORS.ERR_COLLREMWHERE_4_G)
            End If
        Next lIndex
    End Sub

    Public Sub m_CollectionRemoveWhereNot(ByVal sPropertyName As String, ByVal sValue As String)
        Dim lIndex As Integer
        Dim oObject As Object
        Dim sPropertyValue As String
        Dim sIndex As String
        For lIndex = m_lCount() To 1 Step -1
            oObject = m_oReturnArrayElement(lIndex)
            sPropertyValue = CallByName(oObject, sPropertyName)
            If sPropertyValue <> sValue Then
                sIndex = System.Convert.ToString(lIndex)
                m_Remove(sIndex, SYS_ERRORS.ERR_COLLREMWHERENOT_1_G, SYS_ERRORS.ERR_COLLREMWHERENOT_2_G, SYS_ERRORS.ERR_COLLREMWHERENOT_3_G, SYS_ERRORS.ERR_COLLREMWHERENOT_4_G)
            End If
        Next lIndex
    End Sub

    Public Sub m_CollectionChange(ByVal sPropertyName As String, ByVal sKey As String, ByVal sIndex As String, ByVal sNewValue As String)
        Dim lIndex As Integer
        Dim oObject As Object
        Dim sPropertyValue As String
        For lIndex = 1 To m_lCount()
            oObject = m_oReturnArrayElement(lIndex)
            sPropertyValue = CallByName(oObject, sPropertyName)
            If sPropertyValue = sKey Or sPropertyValue = sIndex Then
                CallByName(oObject, sPropertyName, sNewValue)
            End If
        Next lIndex
    End Sub

    Public Sub m_CollectionChangeAll(ByVal sPropertyName As String, ByVal sNewValue As String)
        Dim lIndex As Integer
        Dim oObject As Object
        For lIndex = 1 To m_lCount()
            oObject = m_oReturnArrayElement(lIndex)
            CallByName(oObject, sPropertyName, sNewValue)
        Next lIndex
    End Sub

    Public Function CallByName(ByVal oObject As Object, ByVal sPropertyName As String) As String
        Dim oObjectType As System.Type = oObject.GetType()
        Dim oPropertyInfo As System.Reflection.PropertyInfo
        oPropertyInfo = oObjectType.GetProperty(sPropertyName)
        Return DirectCast(oPropertyInfo.GetValue(oObject, Nothing), System.String)
    End Function

    Public Sub CallByName(oObject As Object, sPropertyName As String, sNewValue As String)
        Dim oObjectType As System.Type = oObject.GetType()
        Dim oPropertyInfo As System.Reflection.PropertyInfo
        oPropertyInfo = oObjectType.GetProperty(sPropertyName)
        oPropertyInfo.SetValue(oObject, sNewValue, Nothing)
    End Sub


End Class
