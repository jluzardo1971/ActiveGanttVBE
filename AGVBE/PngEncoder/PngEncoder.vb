Imports System.IO
Imports System.Diagnostics

Public Class PngEncoder
    Public Const ENCODE_ALPHA As Boolean = True
    Public Const NO_ALPHA As Boolean = False
    Public Const FILTER_NONE As Integer = 0
    Public Const FILTER_SUB As Integer = 1
    Public Const FILTER_UP As Integer = 2
    Public Const FILTER_LAST As Integer = 2
    Protected Shared IHDR As Byte() = New Byte() {73, 72, 68, 82}
    Protected Shared IDAT As Byte() = New Byte() {73, 68, 65, 84}
    Protected Shared IEND As Byte() = New Byte() {73, 69, 78, 68}
    Protected pngBytes As Byte()
    Protected priorRow As Byte()
    Protected leftBytes As Byte()
    Protected width As Integer, height As Integer
    Protected bytePos As Integer, maxPos As Integer
    Protected crc As New Crc32()
    Protected crcValue As Long
    Protected encodeAlpha As Boolean
    Protected filter As Integer
    Protected bytesPerPixel As Integer
    Protected compressionLevel As Integer
    Protected pixelData As Integer()

    Public Sub New(pixel_data As Integer(), width As Integer, height As Integer, encodeAlpha As Boolean, whichFilter As Integer, compLevel As Integer)
        Me.pixelData = pixel_data
        Me.width = width
        Me.height = height
        Me.encodeAlpha = encodeAlpha

        Me.filter = FILTER_NONE
        If whichFilter <= FILTER_LAST Then
            Me.filter = whichFilter
        End If

        If compLevel >= 0 AndAlso compLevel <= 9 Then
            Me.compressionLevel = compLevel
        End If
    End Sub

    Public Function Encode(encodeAlpha As Boolean) As Byte()
        Dim pngIdBytes As Byte() = {&H89, &H50, &H4E, &H47, &HD, &HA, &H1A, &HA}
        pngBytes = New Byte(((width + 1) * height * 3) + 199) {}
        maxPos = 0
        bytePos = WriteBytes(pngIdBytes, 0)

        writeHeader()
        If WriteImageData() Then
            Debug.WriteLine("")
            writeEnd()
            pngBytes = ResizeByteArray(pngBytes, maxPos)
        Else
            pngBytes = Nothing
        End If
        Return pngBytes
    End Function

    Public Function pngEncode() As Byte()
        Return Encode(encodeAlpha)
    End Function


    Protected Function ResizeByteArray(array__1 As Byte(), newLength As Integer) As Byte()
        Dim newArray As Byte() = New Byte(newLength - 1) {}
        Dim oldLength As Integer = array__1.Length

        Array.Copy(array__1, 0, newArray, 0, Math.Min(oldLength, newLength))
        Return newArray
    End Function


    Protected Function WriteBytes(data As Byte(), offset As Integer) As Integer
        maxPos = Math.Max(maxPos, offset + data.Length)
        If data.Length + offset > pngBytes.Length Then
            pngBytes = ResizeByteArray(pngBytes, pngBytes.Length + Math.Max(1000, data.Length))
        End If

        Array.Copy(data, 0, pngBytes, offset, data.Length)
        Return offset + data.Length
    End Function

    Protected Function WriteBytes(data As Byte(), nBytes As Integer, offset As Integer) As Integer
        maxPos = Math.Max(maxPos, offset + nBytes)
        If nBytes + offset > pngBytes.Length Then
            pngBytes = ResizeByteArray(pngBytes, pngBytes.Length + Math.Max(1000, nBytes))
        End If

        Array.Copy(data, 0, pngBytes, offset, nBytes)
        Return offset + nBytes
    End Function


    Protected Function WriteInt2(n As Integer, offset As Integer) As Integer
        Dim temp As Byte() = {CByte((n >> 8) And &HFF), CByte(n And &HFF)}

        Return WriteBytes(temp, offset)
    End Function


    Protected Function WriteInt4(n As Integer, offset As Integer) As Integer
        Dim temp As Byte() = {CByte((n >> 24) And &HFF), CByte((n >> 16) And &HFF), CByte((n >> 8) And &HFF), CByte(n And &HFF)}

        Return WriteBytes(temp, offset)
    End Function


    Protected Function WriteByte(b As Integer, offset As Integer) As Integer
        Dim temp As Byte() = {CByte(b)}

        Return WriteBytes(temp, offset)
    End Function

    Protected Sub writeHeader()
        Dim startPos As Integer

        startPos = InlineAssignHelper(bytePos, WriteInt4(13, bytePos))

        bytePos = WriteBytes(IHDR, bytePos)
        bytePos = WriteInt4(width, bytePos)
        bytePos = WriteInt4(height, bytePos)
        bytePos = WriteByte(8, bytePos)
        bytePos = WriteByte(If((encodeAlpha), 6, 2), bytePos)
        bytePos = WriteByte(0, bytePos)
        bytePos = WriteByte(0, bytePos)
        bytePos = WriteByte(0, bytePos)

        crc.Reset()
        crc.Update(pngBytes, startPos, bytePos - startPos)
        crcValue = crc.Value

        bytePos = WriteInt4(CInt(crcValue), bytePos)
    End Sub

    Protected Sub FilterSub(pixels As Byte(), startPos As Integer, width As Integer)
        Dim i As Integer
        Dim offset As Integer = bytesPerPixel
        Dim actualStart As Integer = startPos + offset
        Dim nBytes As Integer = width * bytesPerPixel
        Dim leftInsert As Integer = offset
        Dim leftExtract As Integer = 0

        For i = actualStart To startPos + (nBytes - 1)
            leftBytes(leftInsert) = pixels(i)
            pixels(i) = CByte((pixels(i) - leftBytes(leftExtract)) Mod 256)
            leftInsert = (leftInsert + 1) Mod &HF
            leftExtract = (leftExtract + 1) Mod &HF
        Next
    End Sub


    Protected Sub FilterUp(pixels As Byte(), startPos As Integer, width As Integer)
        Dim i As Integer, nBytes As Integer
        Dim currentByte As Byte

        nBytes = width * bytesPerPixel

        For i = 0 To nBytes - 1
            currentByte = pixels(startPos + i)
            pixels(startPos + i) = CByte((pixels(startPos + i) - priorRow(i)) Mod 256)
            priorRow(i) = currentByte
        Next
    End Sub


    Protected Function WriteImageData() As Boolean
        Dim rowsLeft As Integer = height
        Dim startRow As Integer = 0
        Dim nRows As Integer

        Dim scanLines As Byte()
        Dim scanPos As Integer
        Dim startPos As Integer

        Dim compressedLines As Byte()
        Dim nCompressed As Integer


        bytesPerPixel = If((encodeAlpha), 4, 3)

        Dim scrunch As New Deflater(compressionLevel)
        Dim outBytes As New MemoryStream(1024)

        Dim compBytes As New DeflaterOutputStream(outBytes, scrunch)
        Try
            While rowsLeft > 0
                nRows = Math.Min(32767 \ (width * (bytesPerPixel + 1)), rowsLeft)
                nRows = Math.Max(nRows, 1)

                Dim pixels As Integer() = New Integer(width * nRows - 1) {}
                Array.Copy(Me.pixelData, width * startRow, pixels, 0, width * nRows)


                scanLines = New Byte(width * nRows * bytesPerPixel + (nRows - 1)) {}

                If filter = FILTER_SUB Then
                    leftBytes = New Byte(15) {}
                End If
                If filter = FILTER_UP Then
                    priorRow = New Byte(width * bytesPerPixel - 1) {}
                End If

                scanPos = 0
                startPos = 1

                For i As Integer = 0 To width * nRows - 1
                    If i Mod width = 0 Then
                        'scanLines[scanPos++] = (byte)filter;
                        scanLines(PosI(scanPos)) = CByte(filter)
                        startPos = scanPos
                    End If
                    'scanLines[scanPos++] = (byte)((pixels[i] >> 16) & 0xff);
                    scanLines(PosI(scanPos)) = CByte((pixels(i) >> 16) And &HFF)
                    'scanLines[scanPos++] = (byte)((pixels[i] >> 8) & 0xff);
                    scanLines(PosI(scanPos)) = CByte((pixels(i) >> 8) And &HFF)
                    'scanLines[scanPos++] = (byte)((pixels[i]) & 0xff);
                    scanLines(PosI(scanPos)) = CByte((pixels(i)) And &HFF)
                    If encodeAlpha Then
                        'scanLines[scanPos++] = (byte)((pixels[i] >> 24) & 0xff);
                        scanLines(PosI(scanPos)) = CByte((pixels(i) >> 24) And &HFF)
                    End If
                    If (i Mod width = width - 1) AndAlso (filter <> FILTER_NONE) Then
                        If filter = FILTER_SUB Then
                            FilterSub(scanLines, startPos, width)
                        End If
                        If filter = FILTER_UP Then
                            FilterUp(scanLines, startPos, width)
                        End If
                    End If
                Next


                compBytes.Write(scanLines, 0, scanPos)

                startRow += nRows
                rowsLeft -= nRows
            End While
            compBytes.Close()


            compressedLines = outBytes.ToArray()
            nCompressed = compressedLines.Length

            crc.Reset()
            bytePos = WriteInt4(nCompressed, bytePos)
            bytePos = WriteBytes(IDAT, bytePos)
            crc.Update(IDAT)
            bytePos = WriteBytes(compressedLines, nCompressed, bytePos)
            crc.Update(compressedLines, 0, nCompressed)

            crcValue = crc.Value
            bytePos = WriteInt4(ForceToInt32(crcValue), bytePos)
            scrunch.Finish()
            Return True
        Catch ex As Exception
            Debug.WriteLine(ex.StackTrace)
            Return False
        End Try
    End Function

    Protected Sub writeEnd()
        bytePos = WriteInt4(0, bytePos)
        bytePos = WriteBytes(IEND, bytePos)
        crc.Reset()
        crc.Update(IEND)
        crcValue = crc.Value
        bytePos = WriteInt4(ForceToInt32(crcValue), bytePos)
    End Sub

    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function
End Class

