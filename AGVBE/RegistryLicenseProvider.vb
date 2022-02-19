Option Explicit On 

Imports System
Imports System.ComponentModel
Imports Microsoft.Win32

Friend Class RegistryLicenseProvider
    'Inherits LicenseProvider

    'Public Sub New()

    'End Sub

    'Public Overloads Overrides Function GetLicense(ByVal context As LicenseContext, ByVal type As Type, ByVal instance As Object, ByVal allowExceptions As Boolean) As License
    '    If context.UsageMode = LicenseUsageMode.Runtime Then
    '        Return New RuntimeRegistryLicense(type)
    '    Else
    '        Try
    '            Dim licenseKey As RegistryKey = Registry.ClassesRoot.OpenSubKey("Licenses\\" & type.GUID.ToString())
    '            If Not licenseKey Is Nothing Then
    '                Dim strLic As String = CType(licenseKey.GetValue(""), String)
    '                If Not strLic Is Nothing Then
    '                    If String.Compare("HCMHCJPGGZTGRY", strLic, False) = 0 Then
    '                        Return New DesigntimeRegistryLicense(type)
    '                    End If
    '                End If
    '                If allowExceptions = True Then
    '                    Throw New LicenseException(type, instance, "You need a design time license to use this control in the design environment.")
    '                End If
    '                Return Nothing
    '            End If
    '        Catch
    '            Return Nothing
    '        End Try
    '    End If
    '    Return Nothing
    'End Function

End Class














































































