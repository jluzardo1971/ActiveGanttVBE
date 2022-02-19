
Option Compare Binary
Option Infer On
Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.ComponentModel.DataAnnotations
Imports System.Linq
Imports System.ServiceModel.DomainServices.Hosting
Imports System.ServiceModel.DomainServices.Server
Imports System.IO

<EnableClientAccess()>  _
Public Class MSP2003Service
    Inherits DomainService

    Private Function mp_GetPath() As String
        Return HttpContext.Current.Request.MapPath(HttpContext.Current.Request.ApplicationPath) & "\MSP2003\"
    End Function

    <Invoke()> _
    Public Function GetXML(sFileName As String) As String
        Dim sReturn As String = g_ReadFile(mp_GetPath() & sFileName)
        Return sReturn
    End Function

    <Invoke()> _
    Public Function GetFileList() As String
        Dim oDirectory As DirectoryInfo = New DirectoryInfo(mp_GetPath())
        Dim oFile As FileInfo
        Dim sReturn As String = ""
        For Each oFile In oDirectory.GetFiles("*.xml")
            sReturn = sReturn & oFile.Name & "|"
        Next
        If sReturn.Length > 0 Then
            sReturn = sReturn.Substring(0, sReturn.Length() - 1)
        End If
        Return sReturn
    End Function

End Class

