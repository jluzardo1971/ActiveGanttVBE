Option Explicit On 

Friend Class clsCollectionBase

    Private m_sObjectName As String
    Friend mp_aoCollection As ArrayList
    Private mp_bAddMode As Boolean
    Private mp_bIgnoreKeyChecks As Boolean
    Friend mp_oKeys As clsDictionary

    Friend Sub New(ByVal sObjectName As String)
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
            If Not g_StrIsNumeric(v_lIndex) Then
                If v_lIndex = "" Then
                    g_ErrorReport(v_lErr1, "Invalid " & m_sObjectName & " object key, key cannot be an empty string", "ActiveGanttVBNCtl.cls" & m_sObjectName & "s.Get m_oItem")
                    Return Nothing
                End If
                lIndex = m_lFindIndexByKey(v_lIndex)
                If lIndex = -1 Then
                    g_ErrorReport(v_lErr2, m_sObjectName & " object not found, invalid key (""" & v_lIndex & """)", "ActiveGanttVBNCtl.cls" & m_sObjectName & "s.Get m_oItem")
                    Return Nothing
                End If
            Else
                lIndex = System.Convert.ToInt32(v_lIndex)
                If lIndex.ToString() <> v_lIndex Then
                    g_ErrorReport(v_lErr3, m_sObjectName & " object not found, invalid Index", "ActiveGanttVBNCtl.cls" & m_sObjectName & "s.Get m_oItem")
                    Return Nothing
                End If
                If lIndex < 1 Or lIndex > mp_aoCollection.Count() Then
                    g_ErrorReport(v_lErr4, m_sObjectName & " object not found, Index out of bounds", "ActiveGanttVBNCtl.cls" & m_sObjectName & "s.Get m_oItem")
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
        oItemBase = r_oObject
        Dim lUpperBounds As Integer
        If mp_bAddMode = False Then
            g_ErrorReport(SYS_ERRORS.ERR_ADDMODE_G, "AddMode must be set to true before executing oCollection.m_Add", "cls" & m_sObjectName & "s")
        End If
        If g_StrIsNumeric(v_sKey) Then
            g_ErrorReport(v_lErr1, "Invalid " & m_sObjectName & " object key, key cannot be numeric", "ActiveGanttVBNCtl.cls" & m_sObjectName & "s.Add")
            Return
        End If
        If v_sKey <> "" Then
            If m_bIsKeyUnique(v_sKey) = False Then
                g_ErrorReport(v_lErr2, "Key is not unique in " & m_sObjectName & "s collection", "ActiveGanttVBNCtl.cls" & m_sObjectName & "s.Add")
                Return
            End If
        End If
        If v_sKey = "" And v_bKeyRequired = True Then
            g_ErrorReport(v_lKeyError, "A Key is required for all objects in the " & m_sObjectName & "s collection", "ActiveGanttVBNCtl.cls" & m_sObjectName & "s.Add")
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
        If Not g_StrIsNumeric(v_sIndex) Then
            If v_sIndex = "" Then
                g_ErrorReport(v_lErr1, "Invalid " & m_sObjectName & " object key, key cannot be an empty string", "ActiveGanttVBNCtl.cls" & m_sObjectName & "s.Remove")
                Return
            End If
            lIndex = m_lFindIndexByKey(v_sIndex)
            If lIndex = -1 Then
                g_ErrorReport(v_lErr2, m_sObjectName & " object not found, invalid key (""" & v_sIndex & """)", "ActiveGanttVBNCtl.cls" & m_sObjectName & "s.Remove")
                Return
            End If
            lRemovedIndex = lIndex
        Else
            lIndex = System.Convert.ToInt32(v_sIndex)
            If lIndex.ToString() <> v_sIndex Then
                g_ErrorReport(v_lErr3, m_sObjectName & " object not found, invalid Index", "ActiveGanttVBNCtl.cls" & m_sObjectName & "s.Remove")
                Return
            End If
            If lIndex < 1 Or lIndex > mp_aoCollection.Count() Then
                g_ErrorReport(v_lErr4, m_sObjectName & " object not found, Index out of bounds", "ActiveGanttVBNCtl.cls" & m_sObjectName & "s.Remove")
                Return
            End If
            lRemovedIndex = lIndex
        End If
        mp_bIgnoreKeyChecks = True
        mp_aoCollection.RemoveAt(lRemovedIndex - 1)
        mp_bIgnoreKeyChecks = False
        m_ReconstKeys()
    End Sub

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

    Public Function m_oReturnArrayElement(ByVal r_lIndex As Integer) As Object
        Return mp_aoCollection(r_lIndex - 1)
    End Function

    Public Sub mp_SetKey(ByRef sCurrentKey As String, ByVal sNewKey As String, ByVal ErrNumber As Integer)
        If m_bIgnoreKeyChecks = False Then
            If g_StrIsNumeric(sNewKey) Or (sNewKey <> sCurrentKey And m_bIsKeyUnique(sNewKey) = False) Then
                g_ErrorReport(ErrNumber, "Numeric or duplicate " & m_sObjectName & " object key", "ActiveGanttVBNCtl.cls" & m_sObjectName & ".Let Key")
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








End Class
