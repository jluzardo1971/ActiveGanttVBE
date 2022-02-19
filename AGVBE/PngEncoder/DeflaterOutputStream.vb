Imports System.IO


Public Class DeflaterOutputStream
    Inherits Stream

    Private buffer_ As Byte()
    Protected deflater_ As Deflater
    Protected baseOutputStream_ As Stream
    Private isClosed_ As Boolean
    Private isStreamOwner_ As Boolean = True
    Private m_password As String
    Private keys As UInteger()

    Public Sub New(baseOutputStream As Stream)
        Me.New(baseOutputStream, New Deflater(), 512)
    End Sub

    Public Sub New(baseOutputStream As Stream, deflater As Deflater)
        Me.New(baseOutputStream, deflater, 512)
    End Sub

    Public Sub New(baseOutputStream As Stream, deflater As Deflater, bufferSize As Integer)
        If baseOutputStream Is Nothing Then
            Throw New ArgumentNullException("baseOutputStream")
        End If

        If baseOutputStream.CanWrite = False Then
            Throw New ArgumentException("Must support writing", "baseOutputStream")
        End If

        If deflater Is Nothing Then
            Throw New ArgumentNullException("deflater")
        End If

        If bufferSize <= 0 Then
            Throw New ArgumentOutOfRangeException("bufferSize")
        End If

        baseOutputStream_ = baseOutputStream
        buffer_ = New Byte(bufferSize - 1) {}
        deflater_ = deflater
    End Sub

    Public Overridable Sub Finish()
        deflater_.Finish()
        While Not deflater_.IsFinished
            Dim len As Integer = deflater_.Deflate(buffer_, 0, buffer_.Length)
            If len <= 0 Then
                Exit While
            End If

            If keys IsNot Nothing Then
                EncryptBlock(buffer_, 0, len)
            End If

            baseOutputStream_.Write(buffer_, 0, len)
        End While

        If Not deflater_.IsFinished Then
            Throw New Exception("Can't deflate all input?")
        End If

        baseOutputStream_.Flush()


        If keys IsNot Nothing Then
            keys = Nothing
        End If

    End Sub

    Public Property IsStreamOwner() As Boolean
        Get
            Return isStreamOwner_
        End Get
        Set(value As Boolean)
            isStreamOwner_ = value
        End Set
    End Property

    Public ReadOnly Property CanPatchEntries() As Boolean
        Get
            Return baseOutputStream_.CanSeek
        End Get
    End Property

    Public Property Password() As String
        Get
            Return m_password
        End Get
        Set(value As String)
            If (value IsNot Nothing) AndAlso (value.Length = 0) Then
                m_password = Nothing
            Else
                m_password = value
            End If
        End Set
    End Property

    Protected Sub EncryptBlock(buffer As Byte(), offset As Integer, length As Integer)
        For i As Integer = offset To offset + (length - 1)
            Dim oldbyte As Byte = buffer(i)
            buffer(i) = buffer(i) Xor EncryptByte()
            UpdateKeys(oldbyte)
        Next
    End Sub

    Protected Function EncryptByte() As Byte
        Dim temp As UInteger = ((keys(2) And &HFFFF) Or 2)
        Return CByte((temp * (temp Xor 1)) >> 8)
    End Function

    Protected Sub UpdateKeys(ch As Byte)
        keys(0) = Crc32.ComputeCrc32(keys(0), ch)
        keys(1) = keys(1) + CByte(keys(0))
        keys(1) = keys(1) * 134775813 + 1
        keys(2) = Crc32.ComputeCrc32(keys(2), CByte(keys(1) >> 24))
    End Sub

    Protected Sub Deflate()
        While Not deflater_.IsNeedingInput
            Dim deflateCount As Integer = deflater_.Deflate(buffer_, 0, buffer_.Length)

            If deflateCount <= 0 Then
                Exit While
            End If

            If keys IsNot Nothing Then
                EncryptBlock(buffer_, 0, deflateCount)
            End If

            baseOutputStream_.Write(buffer_, 0, deflateCount)
        End While

        If Not deflater_.IsNeedingInput Then
            Throw New Exception("DeflaterOutputStream can't deflate all input?")
        End If
    End Sub

    Public Overrides ReadOnly Property CanRead() As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overrides ReadOnly Property CanSeek() As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overrides ReadOnly Property CanWrite() As Boolean
        Get
            Return baseOutputStream_.CanWrite
        End Get
    End Property

    Public Overrides ReadOnly Property Length() As Long
        Get
            Return baseOutputStream_.Length
        End Get
    End Property

    Public Overrides Property Position() As Long
        Get
            Return baseOutputStream_.Position
        End Get
        Set(value As Long)
            Throw New NotSupportedException("Position property not supported")
        End Set
    End Property

    Public Overrides Function Seek(offset As Long, origin As SeekOrigin) As Long
        Throw New NotSupportedException("DeflaterOutputStream Seek not supported")
    End Function

    Public Overrides Sub SetLength(value As Long)
        Throw New NotSupportedException("DeflaterOutputStream SetLength not supported")
    End Sub

    Public Overrides Function ReadByte() As Integer
        Throw New NotSupportedException("DeflaterOutputStream ReadByte not supported")
    End Function

    Public Overrides Function Read(buffer As Byte(), offset As Integer, count As Integer) As Integer
        Throw New NotSupportedException("DeflaterOutputStream Read not supported")
    End Function

    Public Overrides Function BeginRead(buffer As Byte(), offset As Integer, count As Integer, callback As AsyncCallback, state As Object) As IAsyncResult
        Throw New NotSupportedException("DeflaterOutputStream BeginRead not currently supported")
    End Function

    Public Overrides Function BeginWrite(buffer As Byte(), offset As Integer, count As Integer, callback As AsyncCallback, state As Object) As IAsyncResult
        Throw New NotSupportedException("BeginWrite is not supported")
    End Function

    Public Overrides Sub Flush()
        deflater_.Flush()
        Deflate()
        baseOutputStream_.Flush()
    End Sub

    Public Overrides Sub Close()
        If Not isClosed_ Then
            isClosed_ = True

            Try
                Finish()

                keys = Nothing
            Finally
                If isStreamOwner_ Then
                    baseOutputStream_.Close()
                End If
            End Try
        End If
    End Sub

    Public Overrides Sub WriteByte(value As Byte)
        Dim b As Byte() = New Byte(0) {}
        b(0) = value
        Write(b, 0, 1)
    End Sub

    Public Overrides Sub Write(buffer As Byte(), offset As Integer, count As Integer)
        deflater_.SetInput(buffer, offset, count)
        Deflate()
    End Sub

End Class

