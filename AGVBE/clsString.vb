Option Explicit On 

Public Class clsString

    Private mp_oControl As ActiveGanttVBECtl

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
    End Sub

    Public Function StrFormat(ByVal Expression As Single, ByVal sFormat As String) As String
        Return Expression.ToString(sFormat, mp_oControl.Culture.NumberFormat)
    End Function

    Public Function StrFormat(ByVal Expression As Integer, ByVal sFormat As String) As String
        Return Expression.ToString(sFormat, mp_oControl.Culture.NumberFormat)
    End Function

    Public Function StrLeft(ByVal Expression As String, ByVal Length As Integer) As String
        If Length > StrLen(Expression) Then
            Return ""
        Else
            Return Expression.Substring(0, Length)
        End If
    End Function

    Public Function StrRight(ByVal Expression As String, ByVal Length As Integer) As String
        If Length > StrLen(Expression) Then
            Return ""
        Else
            Return Expression.Substring(Expression.Length - Length, Length)
        End If
    End Function

    Public Function StrMid(ByVal Expression As String, ByVal Start As Integer, ByVal Length As Integer) As String
        Return Expression.Substring(Start - 1, Length)
    End Function

    Public Function StrLowerCase(ByVal Expression As String) As String
        Return Expression.ToLower()
    End Function

    Public Function StrUpperCase(ByVal Expression As String) As String
        Return Expression.ToUpper()
    End Function

    Public Function StrIsNumeric(ByVal Expression As String) As Boolean
        Dim dDummy As Double
        Return Double.TryParse(Expression, dDummy)
    End Function

    Public Function StrCLng(ByVal Expression As String) As Integer
        Return System.Convert.ToInt32(Expression)
    End Function

    Public Function StrCStr(ByVal Expression As Integer) As String
        Return System.Convert.ToString(Expression)
    End Function

    Public Function StrCStr(ByVal Expression As Single) As String
        Return Expression.ToString()
    End Function

    Public Function StrCStr(ByVal Expression As String) As String
        Return Expression
    End Function
    
    Public Function StrTrim(ByVal Expression As String) As String
        Return Expression.Trim
    End Function

    Public Function StrReplace(ByVal Expression As String, ByVal sFind As String, ByVal sReplace As String) As String
        Return Expression.Replace(sFind, sReplace)
    End Function

    Friend Function GetDecimalSeparator() As String
        Return System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator()
    End Function
    
    Public Function StrLen(ByVal Expression As String) As Integer
        Return Expression.Length
    End Function

End Class
