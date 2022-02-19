Friend Class ArrayList

    Private mp_oObjects As Object()
    Private mp_lCount As Integer
    Private mp_lGrowFactor As Integer = 100

    Public Sub New()
        ReDim mp_oObjects(mp_lGrowFactor)
        mp_lCount = 0
    End Sub

    Public Function Add(ByVal value As Object) As Integer
        If mp_lCount > UBound(mp_oObjects, 1) Then
            ReDim Preserve mp_oObjects(mp_lCount + mp_lGrowFactor)
        End If
        mp_oObjects(mp_lCount) = value
        mp_lCount = mp_lCount + 1
        Return mp_lCount
    End Function

    'Private ReadOnly Property get_Item(ByVal Index As Integer) As Object
    '    Get
    '        Return 0
    '    End Get
    'End Property

    Default Public Property Item(ByVal index As Integer) As Object
        Get
            If index < 0 Or (index > mp_lCount - 1) Then
                Throw New ArgumentException("Index must be non-negative and less than the size of the ArrayList")
            End If
            Return mp_oObjects(index)
        End Get
        Set(ByVal value As Object)
            If index < 0 Or (index > mp_lCount - 1) Then
                Throw New ArgumentException("Index must be non-negative and less than the size of the ArrayList")
            End If
            mp_oObjects(index) = value
        End Set
    End Property

    Public Sub RemoveAt(ByVal index As Integer)
        Dim i As Integer
        If mp_lCount = 0 Then
            Throw New ArgumentException("ArrayList is empty")
        End If
        If index < 0 Or (index > mp_lCount - 1) Then
            Throw New ArgumentException("Index must be non-negative and less than the size of the ArrayList")
        End If
        For i = index To mp_lCount - 2
            mp_oObjects(i) = mp_oObjects(i + 1)
        Next
        mp_lCount = mp_lCount - 1
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Return mp_lCount
        End Get
    End Property

    Public Sub Clear()
        ReDim mp_oObjects(mp_lGrowFactor)
        mp_lCount = 0
    End Sub


End Class

