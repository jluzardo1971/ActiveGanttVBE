Imports AGVBE

Partial Public Class fCarRentalVehicle
    Inherits ChildWindow

    Private mp_oParent As fCarRental
    Private mp_yDialogMode As PRG_DIALOGMODE
    Public mp_sRowID As String

    Public Property Mode() As PRG_DIALOGMODE
        Get
            Return mp_yDialogMode
        End Get
        Set(value As PRG_DIALOGMODE)
            mp_yDialogMode = value
        End Set
    End Property

    Public Sub New(ByVal oParent As fCarRental)
        InitializeComponent()
        mp_oParent = oParent
    End Sub

    Private Sub fCarRentalVehicle_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            '//TODO
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            '//TODO
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            drpCarTypeID.ItemsSource = mp_oParent.mp_o_AG_CR_Car_Types
            drpCarTypeID.DisplayMemberPath = "sDescription"
            drpCarTypeID.SelectedValuePath = "lCarTypeID"
            drpACRISS1.ItemsSource = mp_oParent.mp_o_AG_CR_ACRISS_Codes_1
            drpACRISS1.DisplayMemberPath = "sDescription"
            drpACRISS1.SelectedValuePath = "sLetter"
            drpACRISS2.ItemsSource = mp_oParent.mp_o_AG_CR_ACRISS_Codes_2
            drpACRISS2.DisplayMemberPath = "sDescription"
            drpACRISS2.SelectedValuePath = "sLetter"
            drpACRISS3.ItemsSource = mp_oParent.mp_o_AG_CR_ACRISS_Codes_3
            drpACRISS3.DisplayMemberPath = "sDescription"
            drpACRISS3.SelectedValuePath = "sLetter"
            drpACRISS4.ItemsSource = mp_oParent.mp_o_AG_CR_ACRISS_Codes_4
            drpACRISS4.DisplayMemberPath = "sDescription"
            drpACRISS4.SelectedValuePath = "sLetter"
        End If
        If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
            Me.Title = "Add New Vehicle"
            txtLicensePlates.Text = g_GenerateRandomLicense()
            drpCarTypeID.SelectedIndex = GetRnd(1, 48)
        Else
            Me.Title = "Edit Vehicle"

            Dim sACRISSCode As String = ""

            If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                '//TODO
            ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                '//TODO
            ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                Dim oCRRow As AG_CR_Row = Nothing
                For Each oCRRow In mp_oParent.mp_o_AG_CR_Rows
                    If oCRRow.lRowID = System.Convert.ToDouble(mp_sRowID) Then
                        Exit For
                    End If
                Next
                txtLicensePlates.Text = oCRRow.sLicensePlates
                drpCarTypeID.SelectedValue = oCRRow.lCarTypeID
                txtNotes.Text = oCRRow.sNotes
                txtRate.Text = oCRRow.dRate.ToString()
                sACRISSCode = oCRRow.sACRISSCode
            End If
            UpdatePicture()
            UpdateACRISSCode(sACRISSCode)
        End If
    End Sub

    Private Sub UpdateACRISSCode(ByVal sACRISSCode As String)
        drpACRISS1.SelectedValue = sACRISSCode.Substring(0, 1)
        drpACRISS2.SelectedValue = sACRISSCode.Substring(1, 1)
        drpACRISS3.SelectedValue = sACRISSCode.Substring(2, 1)
        drpACRISS4.SelectedValue = sACRISSCode.Substring(3, 1)
        lblACRISS1.Content = sACRISSCode.Substring(0, 1)
        lblACRISS2.Content = sACRISSCode.Substring(1, 1)
        lblACRISS3.Content = sACRISSCode.Substring(2, 1)
        lblACRISS4.Content = sACRISSCode.Substring(3, 1)
    End Sub

    Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles cmdOK.Click
        Dim oRow As clsRow = Nothing
        If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            '//TODO
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            '//TODO
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            Dim oCRRow As New AG_CR_Row
            If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
                Dim lRowID As Integer = 0
                For Each oCRRow In mp_oParent.mp_o_AG_CR_Rows
                    If oCRRow.lRowID > lRowID Then
                        lRowID = oCRRow.lRowID
                    End If
                Next
                lRowID = lRowID + 1
                oCRRow = New AG_CR_Row()
                oCRRow.lRowID = lRowID
                oRow = mp_oParent.ActiveGanttVBECtl1.Rows.Add("K" & lRowID.ToString())
                oRow.Node.Depth = 1
                mp_oParent.ActiveGanttVBECtl1.Rows.UpdateTree()
                mp_oParent.mp_o_AG_CR_Rows.Add(oCRRow)
            Else
                For Each oCRRow In mp_oParent.mp_o_AG_CR_Rows
                    If oCRRow.lRowID = System.Convert.ToDouble(mp_sRowID) Then
                        Exit For
                    End If
                Next
                oRow = mp_oParent.ActiveGanttVBECtl1.Rows.Item("K" & mp_sRowID)
            End If
            oCRRow.lDepth = 1
            oCRRow.sLicensePlates = txtLicensePlates.Text
            oCRRow.lCarTypeID = System.Convert.ToInt32(drpCarTypeID.SelectedValue)
            oCRRow.sNotes = txtNotes.Text
            oCRRow.dRate = System.Convert.ToDouble(txtRate.Text)
            oCRRow.sACRISSCode = lblACRISS1.Content.ToString() & lblACRISS2.Content.ToString() & lblACRISS3.Content.ToString() & lblACRISS4.Content.ToString()
        End If
        Dim oCarType As AG_CR_Car_Type = Nothing
        oRow.Cells.Item("1").Text = txtLicensePlates.Text
        oCarType = DirectCast(drpCarTypeID.SelectedItem, AGVBECON.Globals.AG_CR_Car_Type)
        oRow.Cells.Item("2").Image = mp_oParent.GetImage("CarRental/Small/" & oCarType.sDescription & ".jpg")
        oRow.Cells.Item("3").Text = oCarType.sDescription & vbCrLf & lblACRISS1.Content.ToString() & lblACRISS2.Content.ToString() & lblACRISS3.Content.ToString() & lblACRISS4.Content.ToString() & " - " & txtRate.Text & " USD"
        oRow.Node.Depth = 1
        oRow.Tag = lblACRISS1.Content.ToString() & lblACRISS2.Content.ToString() & lblACRISS3.Content.ToString() & lblACRISS4.Content.ToString() & "|" & txtRate.Text & "|" & drpCarTypeID.SelectedValue.ToString()
        If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
            Dim l As Integer
            l = System.Convert.ToInt32(System.Math.Floor(mp_oParent.ActiveGanttVBECtl1.CurrentViewObject.ClientArea.Height / 41))
            If ((mp_oParent.ActiveGanttVBECtl1.Rows.Count - l + 2) > 0) Then
                mp_oParent.ActiveGanttVBECtl1.VerticalScrollBar.Value = (mp_oParent.ActiveGanttVBECtl1.Rows.Count - l + 2)
            End If
        End If
        mp_oParent.ActiveGanttVBECtl1.Redraw()
        Me.DialogResult = True
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles cmdCancel.Click
        Me.DialogResult = False
    End Sub

    Private Sub drpCarTypeID_SelectionChanged(sender As Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles drpCarTypeID.SelectionChanged
        Dim sACRISSCode As String = ""
        UpdatePicture()
        Dim oCarType As AG_CR_Car_Type = Nothing
        oCarType = DirectCast(drpCarTypeID.SelectedItem, AGVBECON.Globals.AG_CR_Car_Type)
        sACRISSCode = oCarType.sACRISSCode
        UpdateACRISSCode(sACRISSCode)
        txtRate.Text = oCarType.dStdRate.ToString()
    End Sub

    Private Sub UpdatePicture()
        Dim oCarType As AG_CR_Car_Type
        oCarType = DirectCast(drpCarTypeID.SelectedItem, AGVBECON.Globals.AG_CR_Car_Type)
        If oCarType.sDescription = "" Then
            Return
        End If
        picMake.Source = GetImage("CarRental/Big/" & oCarType.sDescription & ".jpg").Source
    End Sub

    Friend Function GetImage(ByVal sImage As String) As Image
        Dim oReturnImage = New Image
        Dim oURI As System.Uri = Nothing
        If App.Current.Host.Source.ToString().Contains("file:///") Then
            Dim sSource = App.Current.Host.Source.ToString()
            sSource = sSource.Substring(0, sSource.IndexOf("AGVBECON")) & "AGVBECON.Web/" & sImage.Replace("\", "/")
            oURI = New System.Uri(sSource)
        Else
            oURI = New System.Uri(App.Current.Host.Source, "../" & sImage)
        End If
        Dim oBitmap As New System.Windows.Media.Imaging.BitmapImage()
        AddHandler oBitmap.ImageOpened, AddressOf mp_oBitmapOpened
        oBitmap.UriSource = oURI
        oReturnImage.Width = 442
        oReturnImage.Height = 186
        oReturnImage.Source = oBitmap
        Return oReturnImage
    End Function

    Private Sub mp_oBitmapOpened()
        'ActiveGanttVBECtl1.Redraw()
    End Sub

    Private Sub drpACRISS1_SelectionChanged(sender As Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles drpACRISS1.SelectionChanged
        Dim oACRISSCode As AG_CR_ACRISS_Code
        oACRISSCode = DirectCast(drpACRISS1.SelectedItem, AGVBECON.Globals.AG_CR_ACRISS_Code)
        lblACRISS1.Content = oACRISSCode.sLetter
    End Sub

    Private Sub drpACRISS2_SelectionChanged(sender As Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles drpACRISS2.SelectionChanged
        Dim oACRISSCode As AG_CR_ACRISS_Code
        oACRISSCode = DirectCast(drpACRISS2.SelectedItem, AGVBECON.Globals.AG_CR_ACRISS_Code)
        lblACRISS2.Content = oACRISSCode.sLetter
    End Sub

    Private Sub drpACRISS3_SelectionChanged(sender As Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles drpACRISS3.SelectionChanged
        Dim oACRISSCode As AG_CR_ACRISS_Code
        oACRISSCode = DirectCast(drpACRISS3.SelectedItem, AGVBECON.Globals.AG_CR_ACRISS_Code)
        lblACRISS3.Content = oACRISSCode.sLetter
    End Sub

    Private Sub drpACRISS4_SelectionChanged(sender As Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles drpACRISS4.SelectionChanged
        Dim oACRISSCode As AG_CR_ACRISS_Code
        oACRISSCode = DirectCast(drpACRISS4.SelectedItem, AGVBECON.Globals.AG_CR_ACRISS_Code)
        lblACRISS4.Content = oACRISSCode.sLetter
    End Sub
End Class
