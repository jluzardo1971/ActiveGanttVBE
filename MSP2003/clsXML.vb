Option Explicit On
Imports System.Xml
Imports System.Xml.Linq

Partial Friend Class clsXML

    Private xDoc As XDocument
    Private oControlElement As XElement
    Private mp_sObject As String
    Private mp_yLevel As PE_LEVEL
    Private mp_bSupportOptional As Boolean = False
    Private mp_bBoolsAreNumeric As Boolean = False

    Private Enum PE_LEVEL
        LVL_CONTROL = 0
    End Enum

    Friend Property SupportOptional() As Boolean
        Get
            Return mp_bSupportOptional
        End Get
        Set(ByVal value As Boolean)
            mp_bSupportOptional = value
        End Set
    End Property

    Friend Property BoolsAreNumeric() As Boolean
        Get
            Return mp_bBoolsAreNumeric
        End Get
        Set(ByVal value As Boolean)
            mp_bBoolsAreNumeric = value
        End Set
    End Property

    Friend Sub New(ByVal sObject As String)
        xDoc = New XDocument()
        mp_sObject = sObject
    End Sub

    Friend Sub InitializeWriter()
        xDoc = XDocument.Parse("<" & mp_sObject & "></" & mp_sObject & ">")
        oControlElement = GetDocumentElement(mp_sObject, 0)
        mp_yLevel = PE_LEVEL.LVL_CONTROL
    End Sub

    Friend Sub InitializeReader()
        oControlElement = GetDocumentElement(mp_sObject, 0)
        mp_yLevel = PE_LEVEL.LVL_CONTROL
    End Sub

    Private Function ParentElement() As XElement
        Select Case mp_yLevel
            Case PE_LEVEL.LVL_CONTROL
                Return oControlElement
        End Select
        Return Nothing
    End Function

    Private Function mp_oCreateEmptyDOMElement(ByVal sElementName As String) As XElement
        Dim oNodeBuff As XElement
        oNodeBuff = New XElement(sElementName)
        ParentElement().Add(oNodeBuff)
        Return oNodeBuff
    End Function

    Private Function GetDocumentElement(ByVal TagName As String, ByVal lIndex As Integer) As XElement
        Return xDoc.Elements(TagName)(lIndex)
    End Function

    Friend Function GetXML() As String
        Return xDoc.ToString()
    End Function

    Friend Sub SetXML(ByVal sXML As String)
        If mp_bSupportOptional = False Then
            xDoc = XDocument.Parse(sXML)
        Else
            If sXML.Length > 0 Then
                xDoc = XDocument.Parse(sXML)
            End If
        End If
    End Sub

#Region "Collections"

    Friend Function ReadCollectionObject(ByVal lIndex As Integer) As String
        If mp_bSupportOptional = False Then
            Return ParentElement().Elements()(lIndex - 1).ToString()
        Else
            If ParentElement() Is Nothing Or lIndex = 0 Then
                Return ""
            End If
            If ParentElement().Elements().Count > 0 Then
                Return ParentElement().Elements()(lIndex - 1).ToString()
            Else
                Return ""
            End If
        End If
    End Function

    Friend Function GetCollectionObjectName(ByVal lIndex As Integer) As String
        Dim sReturn As String = ""
        If ParentElement().Elements()(lIndex - 1).NodeType = XmlNodeType.Element Then
            Dim oElement As XElement = DirectCast(ParentElement().Elements()(lIndex - 1), XElement)
            sReturn = oElement.Name.LocalName
        Else
            sReturn = ""
        End If
        Return sReturn
    End Function

    Friend Function ReadCollectionCount() As Integer
        If mp_bSupportOptional = False Then
            Return ParentElement().Elements().Count
        Else
            If ParentElement() Is Nothing Then
                Return 0
            Else
                Return ParentElement().Elements().Count
            End If
        End If
    End Function

#End Region

#Region "XML Objects"

    Friend Function ReadObject(ByVal sObjectName As String) As String
        If mp_bSupportOptional = False Then
            Return ParentElement().Elements(sObjectName)(0).ToString()
        Else
            If ParentElement() Is Nothing Then
                Return ""
            End If
            If ParentElement().Elements(sObjectName).Count > 0 Then
                Return ParentElement().Elements(sObjectName)(0).ToString()
            Else
                Return ""
            End If
        End If
    End Function

    Friend Sub WriteObject(ByVal sObjectText As String)
        Dim xDoc1 As XDocument
        Dim oNodeBuff As XElement
        xDoc1 = New XDocument()
        xDoc1 = XDocument.Parse(sObjectText)
        oNodeBuff = New XElement(xDoc1.Elements()(0))
        ParentElement().Add(oNodeBuff)
    End Sub

#End Region

#Region "Enumerations"

    Public Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Object)
        Dim oNodeBuff As XElement
        oNodeBuff = New XElement(sElementName)
        oNodeBuff.Value = System.Convert.ToString(System.Convert.ToInt32(sElementValue))
        ParentElement().Add(oNodeBuff)
    End Sub

#End Region

#Region "Boolean"
    Friend Sub ReadProperty(ByVal sElementName As String, ByRef bElementValue As Boolean)
        If mp_bSupportOptional = False Then
            If ParentElement().Elements(sElementName)(0).Value = "false" Or ParentElement().Elements(sElementName)(0).Value = "0" Then
                bElementValue = False
            Else
                bElementValue = True
            End If
        Else
            If ParentElement() Is Nothing Then
                Return
            End If
            If ParentElement().Elements(sElementName).Count > 0 Then
                If ParentElement().Elements(sElementName)(0).Parent.Name = ParentElement().Name Then
                    If ParentElement().Elements(sElementName)(0).Value = "false" Or ParentElement().Elements(sElementName)(0).Value = "0" Then
                        bElementValue = False
                    Else
                        bElementValue = True
                    End If
                End If
            End If
        End If
    End Sub

    Friend Sub WriteProperty(ByVal sElementName As String, ByVal bElementValue As Boolean)
        Dim oNodeBuff As XElement
        oNodeBuff = New XElement(sElementName)
        If bElementValue = True Then
            If mp_bBoolsAreNumeric = False Then
                oNodeBuff.Value = "true"
            Else
                oNodeBuff.Value = "1"
            End If
        Else
            If mp_bBoolsAreNumeric = False Then
                oNodeBuff.Value = "false"
            Else
                oNodeBuff.Value = "0"
            End If
        End If
        ParentElement().Add(oNodeBuff)
    End Sub
#End Region

#Region "Byte"

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As Byte)
        sElementValue = System.Convert.ToByte(yReadProperty(sElementName, sElementValue))
    End Sub

    Public Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Byte)
        Dim oNodeBuff As XElement
        oNodeBuff = New XElement(sElementName)
        oNodeBuff.Value = System.Convert.ToString(sElementValue)
        ParentElement().Add(oNodeBuff)
    End Sub

#End Region

#Region "Int32"

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As Integer)
        sElementValue = yReadProperty(sElementName, sElementValue)
    End Sub

    Private Function yReadProperty(ByVal v_sNodeName As String, ByVal sElementValue As Integer) As Integer
        If mp_bSupportOptional = False Then
            Return System.Convert.ToInt32(ParentElement().Elements(v_sNodeName)(0).Value)
        Else
            If ParentElement() Is Nothing Then
                Return sElementValue
            End If
            If ParentElement().Elements(v_sNodeName).Count > 0 Then
                Return System.Convert.ToInt32(ParentElement().Elements(v_sNodeName)(0).Value)
            Else
                Return sElementValue
            End If
        End If
    End Function

    Public Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Integer)
        Dim oNodeBuff As XElement
        oNodeBuff = New XElement(sElementName)
        oNodeBuff.Value = System.Convert.ToString(sElementValue)
        ParentElement().Add(oNodeBuff)
    End Sub

#End Region

#Region "Int16"

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef iElementValue As Short)
        If mp_bSupportOptional = False Then
            iElementValue = System.Convert.ToInt16(ParentElement().Elements(sElementName)(0).Value)
        Else
            If ParentElement() Is Nothing Then
                Return
            End If
            If ParentElement().Elements(sElementName).Count > 0 Then
                iElementValue = System.Convert.ToInt16(ParentElement().Elements(sElementName)(0).Value)
            End If
        End If
    End Sub

    Public Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Short)
        Dim oNodeBuff As XElement
        oNodeBuff = New XElement(sElementName)
        oNodeBuff.Value = System.Convert.ToString(sElementValue)
        ParentElement().Add(oNodeBuff)
    End Sub

#End Region

#Region "Single"

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef fElementValue As Single)
        If mp_bSupportOptional = False Then
            fElementValue = System.Convert.ToSingle(ParentElement().Elements(sElementName)(0).Value)
        Else
            If ParentElement() Is Nothing Then
                Return
            End If
            If ParentElement().Elements(sElementName).Count > 0 Then
                fElementValue = System.Convert.ToSingle(ParentElement().Elements(sElementName)(0).Value)
            End If
        End If
    End Sub

    'Public Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Single)
    '    Dim oNodeBuff As XElement
    '    oNodeBuff = New XElement(sElementName)
    '    oNodeBuff.Value = System.Convert.ToString(sElementValue)
    '    ParentElement().Add(oNodeBuff)
    'End Sub

    Friend Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Single)
        Dim oNodeBuff As XElement
        oNodeBuff = New XElement(sElementName)
        Dim sReturn As String = sElementValue.ToString()
        If InStr(sReturn, "E") <> 0 Then
            Dim oDecimal As Decimal = System.Convert.ToDecimal(sElementValue)
            sReturn = oDecimal.ToString()
        End If
        oNodeBuff.Value = sReturn
        ParentElement.Add(oNodeBuff)
    End Sub

#End Region

#Region "String"

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As String)
        If mp_bSupportOptional = False Then
            sElementValue = ParentElement().Elements(sElementName)(0).Value
        Else
            If ParentElement() Is Nothing Then
                Return
            End If
            If ParentElement().Elements(sElementName).Count > 0 Then
                If ParentElement().Elements(sElementName)(0).Parent.Name = ParentElement().Name Then
                    sElementValue = ParentElement().Elements(sElementName)(0).Value
                End If
            End If
        End If
    End Sub

    Public Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As String)
        Dim oNodeBuff As XElement
        oNodeBuff = New XElement(sElementName)
        oNodeBuff.Value = sElementValue
        ParentElement().Add(oNodeBuff)
    End Sub

#End Region

#Region "System.DateTime"

    Private Sub ReadProperty(ByVal sElementName As String, ByRef dtElementValue As System.DateTime)
        If mp_bSupportOptional = False Then
            dtElementValue = mp_dtGetDateFromXML(ParentElement().Elements(sElementName)(0).Value)
        Else
            If ParentElement() Is Nothing Then
                Return
            End If
            If ParentElement().Elements(sElementName).Count > 0 Then
                dtElementValue = mp_dtGetDateFromXML(ParentElement().Elements(sElementName)(0).Value)
            End If
        End If
    End Sub

    Private Function mp_dtGetDateFromXML(ByVal sParam As String) As System.DateTime
        Dim dtReturn As System.DateTime
        Dim lYear As Integer = System.Convert.ToInt32(sParam.Substring(0, 4))
        Dim lMonth As Integer = System.Convert.ToInt32(sParam.Substring(5, 2))
        Dim lDay As Integer = System.Convert.ToInt32(sParam.Substring(8, 2))
        Dim lHours As Integer = System.Convert.ToInt32(sParam.Substring(11, 2))
        Dim lMinutes As Integer = System.Convert.ToInt32(sParam.Substring(14, 2))
        Dim lSeconds As Integer = System.Convert.ToInt32(sParam.Substring(17, 2))
        dtReturn = New System.DateTime(lYear, lMonth, lDay, lHours, lMinutes, lSeconds)
        Return dtReturn
    End Function

    Private Sub WriteProperty(ByVal sElementName As String, ByRef dtElementValue As System.DateTime)
        Dim oNodeBuff As XElement
        oNodeBuff = New XElement(sElementName)
        oNodeBuff.Value = mp_sGetXMLDateString(dtElementValue)
        ParentElement().Add(oNodeBuff)
    End Sub

    Private Function mp_sGetXMLDateString(ByVal dtParam As System.DateTime) As String
        Return dtParam.Year.ToString("0000") & "-" & dtParam.Month.ToString("00") & "-" & dtParam.Day.ToString("00") & "T" & dtParam.Hour.ToString("00") & ":" & dtParam.Minute.ToString("00") & ":" & dtParam.Second.ToString("00")
    End Function

#End Region

    '// Microsoft Project Integration

#Region "Attributes"

    Friend Sub AddAttribute(ByVal sName As String, ByVal sValue As String)
        Dim oAttribute As XAttribute = New XAttribute(sName, sValue)
        GetDocumentElement("Project", 0).Add(oAttribute)
    End Sub

#End Region

#Region "Time"

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As Time)
        If mp_bSupportOptional = False Then
            sElementValue.FromString(ParentElement().Elements(sElementName)(0).Value)
        Else
            If ParentElement() Is Nothing Then
                Return
            End If
            If ParentElement.Elements(sElementName).Count > 0 Then
                sElementValue.FromString(ParentElement.Elements(sElementName)(0).Value)
            End If
        End If
    End Sub

    Friend Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Time)
        Dim oNodeBuff As XElement
        oNodeBuff = New XElement(sElementName)
        oNodeBuff.Value = sElementValue.ToString()
        ParentElement.Add(oNodeBuff)
    End Sub

#End Region

#Region "Duration"

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As Duration)
        If mp_bSupportOptional = False Then
            sElementValue.FromString(ParentElement().Elements(sElementName)(0).Value)
        Else
            If ParentElement() Is Nothing Then
                Return
            End If
            If ParentElement.Elements(sElementName).Count > 0 Then
                sElementValue.FromString(ParentElement.Elements(sElementName)(0).Value)
            End If
        End If
    End Sub

    Friend Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Duration)
        Dim oNodeBuff As XElement
        oNodeBuff = New XElement(sElementName)
        oNodeBuff.Value = sElementValue.ToString()
        ParentElement.Add(oNodeBuff)
    End Sub

#End Region

#Region "Decimal"

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As Decimal)
        If mp_bSupportOptional = False Then
            sElementValue = System.Convert.ToDecimal(ParentElement().Elements(sElementName)(0).Value)
        Else
            If ParentElement() Is Nothing Then
                Return
            End If
            If ParentElement.Elements(sElementName).Count > 0 Then
                sElementValue = System.Convert.ToDecimal(ParentElement.Elements(sElementName)(0).Value)
            End If
        End If
    End Sub

    Friend Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Decimal)
        Dim oNodeBuff As XElement
        oNodeBuff = New XElement(sElementName)
        oNodeBuff.Value = sElementValue.ToString()
        ParentElement.Add(oNodeBuff)
    End Sub

#End Region

End Class
