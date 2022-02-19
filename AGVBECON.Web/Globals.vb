Imports System.IO

Module Globals

    Public Function g_ReadFile(ByVal sFullPath As String, Optional ByVal bDetectEncoding As Boolean = False, Optional ByVal oEncoding As System.Text.Encoding = Nothing) As String
        Dim oReader As StreamReader
        If bDetectEncoding = False Then
            If oEncoding Is Nothing Then
                oReader = New StreamReader(sFullPath)
            Else
                oReader = New StreamReader(sFullPath, oEncoding)
            End If
        Else
            oReader = New StreamReader(sFullPath, True)
        End If
        Dim sReturn As String = oReader.ReadToEnd()
        oReader.Close()
        Return sReturn
    End Function

    Public Sub g_WriteFile(ByVal sFullPath As String, ByVal sFileContents As String, Optional ByVal oEncoding As System.Text.Encoding = Nothing)
        Dim oWriter As StreamWriter
        If oEncoding Is Nothing Then
            oWriter = New StreamWriter(sFullPath)
        Else
            oWriter = New StreamWriter(sFullPath, False, oEncoding)
        End If
        oWriter.Write(sFileContents)
        oWriter.Close()
    End Sub

End Module
