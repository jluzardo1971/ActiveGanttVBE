Public Interface IChecksum

    Property Value() As Long
    Sub Reset()
    Sub Update(value As Integer)
    Sub Update(buffer As Byte())
    Sub Update(buffer As Byte(), offset As Integer, count As Integer)

End Interface
