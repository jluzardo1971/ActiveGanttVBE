Option Explicit On 
Imports System.Xml
Imports System.Xml.Linq
Imports System.IO
Imports System.Windows.Media.Imaging

Partial Friend Class clsXML

    Private mp_oControl As ActiveGanttVBECtl
    Private xDoc As XDocument
    Private oControlElement As XElement
    Private oFontElement As XElement
    Private oDateTimeElement As XElement
    Private mp_sObject As String
    Private mp_yLevel As PE_LEVEL
    Private mp_bSupportOptional As Boolean = False
    Private mp_bBoolsAreNumeric As Boolean = False

    Private Enum PE_LEVEL
        LVL_CONTROL = 0
        LVL_FONT = 5
        LVL_DATETIME = 6
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

    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal sObject As String)
        mp_oControl = Value
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
            Case PE_LEVEL.LVL_FONT
                Return oFontElement
            Case PE_LEVEL.LVL_DATETIME
                Return oDateTimeElement
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

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_BORDERSTYLE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_BORDERSTYLE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_TEXTPLACEMENT)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_TEXTPLACEMENT)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_PLACEMENT)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_PLACEMENT)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_REPORTERRORS)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_REPORTERRORS)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_SCROLLBEHAVIOUR)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_SCROLLBEHAVIOUR)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_STYLEAPPEARANCE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_STYLEAPPEARANCE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As GRE_BORDERSTYLE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), GRE_BORDERSTYLE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As GRE_PATTERN)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), GRE_PATTERN)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As GRE_BUTTONSTYLE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), GRE_BUTTONSTYLE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As GRE_FIGURETYPE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), GRE_FIGURETYPE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As GRE_GRADIENTFILLMODE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), GRE_GRADIENTFILLMODE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As GRE_HORIZONTALALIGNMENT)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), GRE_HORIZONTALALIGNMENT)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As GRE_LINEDRAWSTYLE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), GRE_LINEDRAWSTYLE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As GRE_VERTICALALIGNMENT)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), GRE_VERTICALALIGNMENT)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_ADDMODE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_ADDMODE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_LAYEROBJECTENABLE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_LAYEROBJECTENABLE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_TIMEBLOCKBEHAVIOUR)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_TIMEBLOCKBEHAVIOUR)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_MOVEMENTTYPE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_MOVEMENTTYPE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_CONSTRAINTTYPE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_CONSTRAINTTYPE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_PROGRESSLINELENGTH)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_PROGRESSLINELENGTH)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_PROGRESSLINETYPE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_PROGRESSLINETYPE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_TICKMARKTYPES)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_TICKMARKTYPES)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_TIERPOSITION)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_TIERPOSITION)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_TIERTYPE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_TIERTYPE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_CONTROLMODE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_CONTROLMODE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As GRE_BACKGROUNDMODE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), GRE_BACKGROUNDMODE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As GRE_HATCHSTYLE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), GRE_HATCHSTYLE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As GRE_FILLMODE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), GRE_FILLMODE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_WEEKDAY)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_WEEKDAY)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_TIMEBLOCKTYPE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_TIMEBLOCKTYPE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_RECURRINGTYPE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_RECURRINGTYPE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_INTERVAL)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_INTERVAL)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_SPLITTERTYPE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_SPLITTERTYPE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_TIERBACKGROUNDMODE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_TIERBACKGROUNDMODE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_TIERAPPEARANCESCOPE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_TIERAPPEARANCESCOPE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_TIERFORMATSCOPE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_TIERFORMATSCOPE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_SELECTIONRECTANGLEMODE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_SELECTIONRECTANGLEMODE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_PREDECESSORMODE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_PREDECESSORMODE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_TASKTYPE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_TASKTYPE)
    End Sub

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As E_TBINTERVALTYPE)
        sElementValue = DirectCast(yReadProperty(sElementName, sElementValue), E_TBINTERVALTYPE)
    End Sub

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

    Public Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As Single)
        Dim oNodeBuff As XElement
        oNodeBuff = New XElement(sElementName)
        oNodeBuff.Value = System.Convert.ToString(sElementValue)
        ParentElement().Add(oNodeBuff)
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

#Region "System.Windows.Media.Color"

    Public Sub ReadProperty(ByVal sElementName As String, ByRef sElementValue As System.Windows.Media.Color)
        Dim lResult As Long
        If mp_bSupportOptional = False Then
            lResult = Convert.ToInt32(ParentElement().Elements(sElementName)(0).Value)
        Else
            If ParentElement() Is Nothing Then
                Return
            End If
            If ParentElement().Elements(sElementName).Count > 0 Then
                lResult = Convert.ToInt32(ParentElement().Elements(sElementName)(0).Value)
            End If
        End If
        Dim lR As Byte
        Dim lG As Byte
        Dim lB As Byte
        lB = System.Convert.ToByte(System.Math.Floor(lResult / 65536))
        lResult = lResult - (lB * 65536)
        lG = System.Convert.ToByte(System.Math.Floor(lResult / 256))
        lResult = lResult - (lG * 256)
        lR = System.Convert.ToByte(lResult)
        sElementValue = Windows.Media.Color.FromArgb(255, lR, lG, lB)
    End Sub

    Public Sub WriteProperty(ByVal sElementName As String, ByVal sElementValue As System.Windows.Media.Color)
        Dim oNodeBuff As XElement
        oNodeBuff = New XElement(sElementName)
        Dim lResult As Long = (sElementValue.B() * 65536) + (sElementValue.G() * 256) + sElementValue.R()
        oNodeBuff.Value = lResult.ToString()
        ParentElement().Add(oNodeBuff)
    End Sub

#End Region

#Region "Font"

    Public Sub ReadProperty(ByVal sElementName As String, ByRef v_oFont As Font)
        Dim mp_yBackupLevel As PE_LEVEL
        Dim sName As String = ""
        Dim fSize As Single = 0
        Dim bDummy As Boolean
        oFontElement = ParentElement().Elements(sElementName)(0)
        mp_yBackupLevel = mp_yLevel
        mp_yLevel = PE_LEVEL.LVL_FONT
        ReadProperty("Name", sName)
        ReadProperty("Size", fSize)
        If sName = "MS Sans Serif" Then
            sName = "Microsoft Sans Serif"
        End If
        Dim oFont As New Font(sName, fSize, E_FONTSIZEUNITS.FSU_POINTS)
        ReadProperty("Bold", bDummy)
        If bDummy = True Then
            oFont.FontWeight = FontWeights.Bold
        End If
        ReadProperty("Italic", bDummy)
        If bDummy = True Then
            oFont.FontStyle = FontStyles.Italic
        End If
        ReadProperty("Underline", bDummy)
        oFont.Underline = bDummy
        mp_yLevel = mp_yBackupLevel
        v_oFont = oFont
    End Sub

    Public Sub WriteProperty(ByVal sElementName As String, ByRef oFont As Font)
        Dim mp_yBackupLevel As PE_LEVEL
        oFontElement = mp_oCreateEmptyDOMElement(sElementName)
        mp_yBackupLevel = mp_yLevel
        mp_yLevel = PE_LEVEL.LVL_FONT
        WriteProperty("Name", oFont.Name)
        WriteProperty("Size", oFont.GetSize(E_FONTSIZEUNITS.FSU_POINTS))
        WriteProperty("Bold", oFont.Bold)
        WriteProperty("Italic", oFont.Italic)
        WriteProperty("Underline", oFont.Underline)
        mp_yLevel = mp_yBackupLevel
    End Sub

#End Region

#Region "Image"

    Public Sub ReadProperty(ByVal sElementName As String, ByRef oImage As Image)
        If ParentElement().Elements(sElementName)(0).Value <> "" Then
            Dim data As String = ParentElement().Elements(sElementName)(0).Value
            Dim aImageBytes As Byte() = System.Convert.FromBase64String(data)
            Dim oMemoryStream As New MemoryStream(aImageBytes, 0, aImageBytes.Length)
            Dim oBitmapImage As BitmapImage = New BitmapImage()
            oBitmapImage.SetSource(oMemoryStream)
            If Not oBitmapImage Is Nothing Then
                oImage = New Image()
                oImage.Width = oBitmapImage.PixelWidth
                oImage.Height = oBitmapImage.PixelHeight
                oImage.Source = oBitmapImage
                If oImage.ActualWidth = 0 And oImage.ActualHeight = 0 Then
                    oImage = Nothing
                End If
                oMemoryStream.Close()
            Else
                oImage = Nothing
            End If
        End If
    End Sub

    Public Sub WriteProperty(ByVal sElementName As String, ByRef oImage As Image)
        Dim oNodeBuff As XElement
        If Not (oImage Is Nothing) Then
            Dim oWBitmap As System.Windows.Media.Imaging.WriteableBitmap = New System.Windows.Media.Imaging.WriteableBitmap(oImage, Nothing)
            Dim oEncoder1 As PngEncoder = New PngEncoder(oWBitmap.Pixels, System.Convert.ToInt32(oImage.ActualWidth), System.Convert.ToInt32(oImage.ActualHeight), True, 0, 0)
            Dim aEncodedArray As Byte() = oEncoder1.Encode(True)
            If Not aEncodedArray Is Nothing Then
                Dim sObjectText As String
                Dim sEncodedData As String
                Dim xDoc1 As XDocument
                Dim oMemoryStream As System.IO.MemoryStream
                xDoc1 = New XDocument()
                sObjectText = "<" & sElementName & " xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64""></" & sElementName & ">"
                xDoc1 = XDocument.Parse(sObjectText)
                oNodeBuff = New XElement(xDoc1.Elements()(0))
                oMemoryStream = New System.IO.MemoryStream(aEncodedArray)
                sEncodedData = Convert.ToBase64String(oMemoryStream.ToArray())
                oNodeBuff.Value = sEncodedData
            Else
                oNodeBuff = New XElement(sElementName)
            End If
        Else
            oNodeBuff = New XElement(sElementName)
        End If
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

#Region "AGVBE.DateTime"

    Friend Sub ReadProperty(ByVal sElementName As String, ByRef oDate As AGVBE.DateTime)
        Dim mp_yBackupLevel As PE_LEVEL
        oDateTimeElement = ParentElement().Elements(sElementName)(0)
        mp_yBackupLevel = mp_yLevel
        mp_yLevel = PE_LEVEL.LVL_DATETIME
        Dim dtDateTime As System.DateTime = New System.DateTime(0)
        Dim lSecondFraction As Integer = 0
        ReadProperty("DateTime", dtDateTime)
        ReadProperty("SecondFraction", lSecondFraction)
        oDate.DateTimePart = dtDateTime
        oDate.SecondFractionPart = lSecondFraction
        mp_yLevel = mp_yBackupLevel
    End Sub

    Friend Sub WriteProperty(ByVal sElementName As String, ByRef oDate As AGVBE.DateTime)
        Dim mp_yBackupLevel As PE_LEVEL
        mp_yBackupLevel = mp_yLevel
        oDateTimeElement = mp_oCreateEmptyDOMElement(sElementName)
        mp_yLevel = PE_LEVEL.LVL_DATETIME
        WriteProperty("DateTime", oDate.DateTimePart)
        WriteProperty("SecondFraction", oDate.SecondFractionPart)
        mp_yLevel = mp_yBackupLevel
    End Sub


#End Region

End Class


