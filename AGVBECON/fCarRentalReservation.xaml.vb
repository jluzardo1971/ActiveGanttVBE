Imports AGVBE

Partial Public Class fCarRentalReservation
    Inherits ChildWindow

    Private mp_oParent As fCarRental
    Private mp_yDialogMode As PRG_DIALOGMODE
    Public mp_sTaskID As String
    Private mp_oTask As clsTask

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

    Private Sub fCarRentalReservation_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        Dim sRowTag() As String
        If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
            Dim sCityName As String = ""
            Dim sStateName As String = ""
            Dim lID As Integer = 0
            If mp_oParent.Mode = fCarRental.HPE_ADDMODE.AM_RESERVATION Then
                Me.Title = "Add Reservation"
                lblMode.Content = "Add Reservation"
                lblMode.Background = New SolidColorBrush(Color.FromArgb(255, 153, 170, 194))
            Else
                Me.Title = "Add Rental"
                lblMode.Content = "Add Rental"
                lblMode.Background = New SolidColorBrush(Color.FromArgb(255, 162, 78, 50))
            End If
            g_GenerateRandomCity(sCityName, sStateName, lID, mp_oParent.mp_yDataSourceType)
            mp_oTask = mp_oParent.ActiveGanttVBECtl1.Tasks.Item(mp_oParent.ActiveGanttVBECtl1.Tasks.Count.ToString())
            txtCity.Text = sCityName
            txtName.Text = g_GenerateRandomName(False, mp_oParent.mp_yDataSourceType)
            txtState.Text = sStateName
            txtPhone.Text = g_GenerateRandomPhone("")
            txtMobile.Text = g_GenerateRandomPhone(txtPhone.Text.Substring(0, 5))
            txtAddress.Text = g_GenerateRandomAddress(mp_oParent.mp_yDataSourceType)
            txtZIP.Text = g_GenerateRandomZIP()
            SetStartDate(mp_oTask.StartDate.DateTimePart)
            SetEndDate(mp_oTask.EndDate.DateTimePart)
            GetStateTax()

            If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                '//TODO
            ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                '//TODO
            ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                Dim oTSO As AG_CR_Tax_Surcharge_Option
                For Each oTSO In mp_oParent.mp_o_AG_CR_Taxes_Surcharges_Options
                    If oTSO.sID = "GPS" Then
                        chkGPS.Content = oTSO.sDescription
                        chkGPS.Tag = oTSO.dRate
                    End If
                    If oTSO.sID = "LDW" Then
                        chkLDW.Content = oTSO.sDescription
                        chkLDW.Tag = oTSO.dRate
                    End If
                    If oTSO.sID = "PAI" Then
                        chkPAI.Content = oTSO.sDescription
                        chkPAI.Tag = oTSO.dRate
                    End If
                    If oTSO.sID = "PEP" Then
                        chkPEP.Content = oTSO.sDescription
                        chkPEP.Tag = oTSO.dRate
                    End If
                    If oTSO.sID = "ALI" Then
                        chkALI.Content = oTSO.sDescription
                        chkALI.Tag = oTSO.dRate
                    End If
                    If oTSO.sID = "ERF" Then
                        txtERF.Tag = oTSO.dRate
                    End If
                    If oTSO.sID = "CRF" Then
                        txtCRF.Tag = oTSO.dRate
                    End If
                    If oTSO.sID = "RCFC" Then
                        txtRCFC.Tag = oTSO.dRate
                    End If
                    If oTSO.sID = "WTB" Then
                        txtWTB.Tag = oTSO.dRate
                    End If
                    If oTSO.sID = "VLF" Then
                        txtVLF.Tag = oTSO.dRate
                    End If
                Next
            End If

        Else
            mp_oTask = mp_oParent.ActiveGanttVBECtl1.Tasks.Item("K" & mp_sTaskID)
            If CType(mp_oTask.Tag, System.Int32) = 0 Then
                Me.Title = "Edit Reservation"
                lblMode.Content = "Edit Reservation"
                lblMode.Background = New SolidColorBrush(Color.FromArgb(255, 153, 170, 194))
            Else
                Me.Title = "Edit Rental"
                lblMode.Content = "Edit Rental"
                lblMode.Background = New SolidColorBrush(Color.FromArgb(255, 162, 78, 50))
            End If

            If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                '//TODO
            ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                '//TODO
            ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                Dim oRental As AG_CR_Rental = Nothing
                For Each oRental In mp_oParent.mp_o_AG_CR_Rentals
                    If oRental.lTaskID = System.Convert.ToDouble(mp_sTaskID) Then
                        Exit For
                    End If
                Next
                txtCity.Text = oRental.sCity
                txtName.Text = oRental.sName
                txtState.Text = oRental.sState
                txtPhone.Text = oRental.sPhone
                txtMobile.Text = oRental.sMobile
                txtAddress.Text = oRental.sAddress
                txtZIP.Text = oRental.sZIP
                SetStartDate(oRental.dtPickUp)
                SetEndDate(oRental.dtReturn)
                txtRate.Text = oRental.dRate.ToString()
                chkGPS.Tag = oRental.dGPS
                chkLDW.Tag = oRental.dLDW
                chkPAI.Tag = oRental.dPAI
                chkPEP.Tag = oRental.dPEP
                chkALI.Tag = oRental.dALI
                txtERF.Tag = oRental.dERF
                txtCRF.Tag = oRental.dCRF
                txtRCFC.Tag = oRental.dRCFC
                txtWTB.Tag = oRental.dWTB
                txtVLF.Tag = oRental.dVLF
                lblTax.Tag = oRental.dTax
                txtEstimatedTotal.Tag = oRental.dEstimatedTotal
                chkGPS.IsChecked = oRental.bGPS
                chkFSO.IsChecked = oRental.bFSO
                chkLDW.IsChecked = oRental.bLDW
                chkPAI.IsChecked = oRental.bPAI
                chkPEP.IsChecked = oRental.bPEP
                chkALI.IsChecked = oRental.bALI
            End If
        End If
        GetStateTax()
        sRowTag = mp_oTask.Row.Tag.Split("|"c)
        txtDescription.Text = mp_oParent.GetDescription(System.Convert.ToInt32(sRowTag(2)))
        picCarType.Source = mp_oParent.GetImage("CarRental/Small/" & txtDescription.Text & ".jpg").Source
        txtRate.Text = g_Format(CType(sRowTag(1), System.Double), "0.00")
        txtACRISS1.Text = GetACRISSDescription(1, sRowTag(0).Substring(0, 1))
        txtACRISS2.Text = GetACRISSDescription(2, sRowTag(0).Substring(1, 1))
        txtACRISS3.Text = GetACRISSDescription(3, sRowTag(0).Substring(2, 1))
        txtACRISS4.Text = GetACRISSDescription(4, sRowTag(0).Substring(3, 1))
        CalculateRate()
    End Sub

    Public Function GetACRISSDescription(ByVal sPosition As Integer, ByVal sLetter As String) As String
        Dim sReturn As String = "GetACRISSDescription Error"
        If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            '//TODO
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            '//TODO
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            Dim oACRISSCode As AG_CR_ACRISS_Code = Nothing
            For Each oACRISSCode In mp_oParent.mp_o_AG_CR_ACRISS_Codes
                If sLetter = oACRISSCode.sLetter Then
                    Exit For
                End If
            Next
            Return oACRISSCode.sDescription
        End If
        Return sReturn
    End Function

    Private Sub GetStateTax()
        Dim sState As String = ""
        Dim dTax As Double = 0
        dTax = mp_oParent.GetStateTax(mp_oTask, sState)
        lblTax.Content = sState & " State Tax (" & g_Format(dTax * 100, "0") & "%):"
        lblTax.Tag = dTax
    End Sub

    Private Sub CalculateRate()
        Dim dFactor As Double = 0
        Dim sRowTag As String()
        Dim dRate As Double = 0
        Dim dSubTotal As Double = 0
        Dim dOptions As Double = 0
        dFactor = CType(mp_oParent.ActiveGanttVBECtl1.MathLib.DateTimeDiff(E_INTERVAL.IL_HOUR, FromDate(GetStartDate()), FromDate(GetEndDate())) / 24, System.Double)
        If chkGPS.IsChecked = True Then
            txtGPS.Text = g_Format(DirectCast(chkGPS.Tag, System.Double) * dFactor, "0.00")
            txtGPS.Tag = g_Format(DirectCast(chkGPS.Tag, System.Double) * dFactor, "0.00")
        Else
            txtGPS.Text = ""
            txtGPS.Tag = 0
        End If
        If chkLDW.IsChecked = True Then
            txtLDW.Text = g_Format(DirectCast(chkLDW.Tag, System.Double) * dFactor, "0.00")
            txtLDW.Tag = g_Format(DirectCast(chkLDW.Tag, System.Double) * dFactor, "0.00")
        Else
            txtLDW.Text = ""
            txtLDW.Tag = 0
        End If
        If chkPAI.IsChecked = True Then
            txtPAI.Text = g_Format(DirectCast(chkPAI.Tag, System.Double) * dFactor, "0.00")
            txtPAI.Tag = g_Format(DirectCast(chkPAI.Tag, System.Double) * dFactor, "0.00")
        Else
            txtPAI.Text = ""
            txtPAI.Tag = 0
        End If
        If chkPEP.IsChecked = True Then
            txtPEP.Text = g_Format(DirectCast(chkPEP.Tag, System.Double) * dFactor, "0.00")
            txtPEP.Tag = g_Format(DirectCast(chkPEP.Tag, System.Double) * dFactor, "0.00")
        Else
            txtPEP.Text = ""
            txtPEP.Tag = 0
        End If
        If chkALI.IsChecked = True Then
            txtALI.Text = g_Format(DirectCast(chkALI.Tag, System.Double) * dFactor, "0.00")
            txtALI.Tag = g_Format(DirectCast(chkALI.Tag, System.Double) * dFactor, "0.00")
        Else
            txtALI.Text = ""
            txtALI.Tag = 0
        End If
        sRowTag = mp_oTask.Row.Tag.Split("|"c)
        dRate = CType(sRowTag(1), System.Double)
        txtERF.Text = g_Format(CType(txtERF.Tag, System.Double) * dFactor, "0.00")
        txtWTB.Text = g_Format(CType(txtWTB.Tag, System.Double) * dFactor, "0.00")
        txtRCFC.Text = g_Format(CType(txtRCFC.Tag, System.Double) * dFactor, "0.00")
        txtVLF.Text = g_Format(CType(txtVLF.Tag, System.Double) * dFactor, "0.00")
        txtCRF.Text = g_Format(CType(txtCRF.Tag, System.Double) * dRate * dFactor, "0.00")
        txtSurcharge.Tag = (CType(txtERF.Tag, System.Double) * dFactor) + (CType(txtWTB.Tag, System.Double) * dFactor) + (CType(txtRCFC.Tag, System.Double) * dFactor) + (CType(txtVLF.Tag, System.Double) * dFactor) + (CType(txtCRF.Tag, System.Double) * dRate * dFactor)
        txtSurcharge.Text = g_Format(CType(txtSurcharge.Tag, System.Double), "0.00")

        dOptions = CType(txtGPS.Tag, System.Double) + CType(txtLDW.Tag, System.Double) + CType(txtPAI.Tag, System.Double) + CType(txtPEP.Tag, System.Double) + CType(txtALI.Tag, System.Double)
        dSubTotal = CType(txtSurcharge.Tag, System.Double) + (dRate * dFactor)

        txtTax.Tag = dSubTotal * DirectCast(lblTax.Tag, System.Double)
        txtTax.Text = g_Format(DirectCast(txtTax.Tag, System.Double), "0.00")

        txtEstimatedTotal.Tag = dSubTotal + CType(txtTax.Tag, System.Double) + dOptions
        txtEstimatedTotal.Text = g_Format(CType(txtEstimatedTotal.Tag, System.Double), "0.00")
    End Sub

    Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles cmdOK.Click
        If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            '//TODO
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            '//TODO
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            Dim oRental As AG_CR_Rental = Nothing
            If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
                Dim lTaskID As Integer = 0
                For Each oRental In mp_oParent.mp_o_AG_CR_Rentals
                    If oRental.lTaskID > lTaskID Then
                        lTaskID = oRental.lTaskID
                    End If
                Next
                lTaskID = lTaskID + 1
                oRental = New AG_CR_Rental()
                oRental.lTaskID = lTaskID
                mp_oTask.Key = "K" & lTaskID.ToString()
            Else
                For Each oRental In mp_oParent.mp_o_AG_CR_Rentals
                    If oRental.lTaskID = System.Convert.ToInt32(mp_oTask.Key.Replace("K", "")) Then
                        Exit For
                    End If
                Next
            End If
            oRental.lRowID = System.Convert.ToInt32(mp_oTask.RowKey.Replace("K", ""))
            oRental.yMode = System.Convert.ToInt32(mp_oParent.Mode)
            oRental.sCity = txtCity.Text
            oRental.sName = txtName.Text
            oRental.sState = txtState.Text
            oRental.sPhone = txtPhone.Text
            oRental.sMobile = txtMobile.Text
            oRental.sAddress = txtAddress.Text
            oRental.sZIP = txtZIP.Text
            oRental.dtPickUp = GetStartDate()
            oRental.dtReturn = GetEndDate()
            oRental.dRate = System.Convert.ToDouble(txtRate.Text)
            oRental.dGPS = System.Convert.ToDouble(chkGPS.Tag)
            oRental.dLDW = System.Convert.ToDouble(chkLDW.Tag)
            oRental.dPAI = System.Convert.ToDouble(chkPAI.Tag)
            oRental.dPEP = System.Convert.ToDouble(chkPEP.Tag)
            oRental.dALI = System.Convert.ToDouble(chkALI.Tag)
            oRental.dERF = System.Convert.ToDouble(txtERF.Tag)
            oRental.dCRF = System.Convert.ToDouble(txtCRF.Tag)
            oRental.dRCFC = System.Convert.ToDouble(txtRCFC.Tag)
            oRental.dWTB = System.Convert.ToDouble(txtWTB.Tag)
            oRental.dVLF = System.Convert.ToDouble(txtVLF.Tag)
            oRental.dTax = System.Convert.ToDouble(lblTax.Tag)
            oRental.dEstimatedTotal = System.Convert.ToDouble(txtEstimatedTotal.Tag)
            oRental.bGPS = System.Convert.ToBoolean(chkGPS.IsChecked)
            oRental.bFSO = System.Convert.ToBoolean(chkFSO.IsChecked)
            oRental.bLDW = System.Convert.ToBoolean(chkLDW.IsChecked)
            oRental.bPAI = System.Convert.ToBoolean(chkPAI.IsChecked)
            oRental.bPEP = System.Convert.ToBoolean(chkPEP.IsChecked)
            oRental.bALI = System.Convert.ToBoolean(chkALI.IsChecked)
            If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
                mp_oParent.mp_o_AG_CR_Rentals.Add(oRental)
            End If
        End If

        mp_oTask.Text = txtName.Text & vbCrLf & "Phone: " & txtPhone.Text & vbCrLf & "Estimated Total: " & txtEstimatedTotal.Text & " USD"
        mp_oTask.Tag = CType(CType(mp_oParent.Mode, System.Int32), System.String)
        mp_oParent.ActiveGanttVBECtl1.Redraw()
        Me.DialogResult = True
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles cmdCancel.Click
        If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
            mp_oParent.ActiveGanttVBECtl1.Tasks.Remove(mp_oParent.ActiveGanttVBECtl1.Tasks.Count.ToString)
            mp_oParent.ActiveGanttVBECtl1.Redraw()
        End If
        Me.DialogResult = False
    End Sub

    Private Sub SetStartDate(ByVal dtDate As System.DateTime)
        txtSDYear.Text = dtDate.Year.ToString()
        txtSDMonth.Text = dtDate.Month.ToString()
        txtSDDay.Text = dtDate.Day.ToString()
        txtSDHours.Text = dtDate.Hour.ToString()
        txtSDMinutes.Text = dtDate.Minute.ToString()
        txtSDSeconds.Text = dtDate.Second.ToString()
    End Sub

    Private Sub SetEndDate(ByVal dtDate As System.DateTime)
        txtEDYear.Text = dtDate.Year.ToString()
        txtEDMonth.Text = dtDate.Month.ToString()
        txtEDDay.Text = dtDate.Day.ToString()
        txtEDHours.Text = dtDate.Hour.ToString()
        txtEDMinutes.Text = dtDate.Minute.ToString()
        txtEDSeconds.Text = dtDate.Second.ToString()
    End Sub

    Private Function GetStartDate() As System.DateTime
        Dim dtDate As System.DateTime = New System.DateTime(System.Convert.ToInt32(txtSDYear.Text), System.Convert.ToInt32(txtSDMonth.Text), System.Convert.ToInt32(txtSDDay.Text), System.Convert.ToInt32(txtSDHours.Text), System.Convert.ToInt32(txtSDMinutes.Text), System.Convert.ToInt32(txtSDSeconds.Text))
        Return dtDate
    End Function

    Private Function GetEndDate() As System.DateTime
        Dim dtDate As System.DateTime = New System.DateTime(System.Convert.ToInt32(txtEDYear.Text), System.Convert.ToInt32(txtEDMonth.Text), System.Convert.ToInt32(txtEDDay.Text), System.Convert.ToInt32(txtEDHours.Text), System.Convert.ToInt32(txtEDMinutes.Text), System.Convert.ToInt32(txtEDSeconds.Text))
        Return dtDate
    End Function

    Private Sub chkGPS_Checked(sender As Object, e As System.Windows.RoutedEventArgs) Handles chkGPS.Checked
        CalculateRate()
    End Sub

    Private Sub chkGPS_Unchecked(sender As Object, e As System.Windows.RoutedEventArgs) Handles chkGPS.Unchecked
        CalculateRate()
    End Sub

    Private Sub chkLDW_Checked(sender As Object, e As System.Windows.RoutedEventArgs) Handles chkLDW.Checked
        CalculateRate()
    End Sub

    Private Sub chkLDW_Unchecked(sender As Object, e As System.Windows.RoutedEventArgs) Handles chkLDW.Unchecked
        CalculateRate()
    End Sub

    Private Sub chkPAI_Checked(sender As Object, e As System.Windows.RoutedEventArgs) Handles chkPAI.Checked
        CalculateRate()
    End Sub

    Private Sub chkPAI_Unchecked(sender As Object, e As System.Windows.RoutedEventArgs) Handles chkPAI.Unchecked
        CalculateRate()
    End Sub

    Private Sub chkPEP_Checked(sender As Object, e As System.Windows.RoutedEventArgs) Handles chkPEP.Checked
        CalculateRate()
    End Sub

    Private Sub chkPEP_Unchecked(sender As Object, e As System.Windows.RoutedEventArgs) Handles chkPEP.Unchecked
        CalculateRate()
    End Sub

    Private Sub chkALI_Checked(sender As Object, e As System.Windows.RoutedEventArgs) Handles chkALI.Checked
        CalculateRate()
    End Sub

    Private Sub chkALI_Unchecked(sender As Object, e As System.Windows.RoutedEventArgs) Handles chkALI.Unchecked
        CalculateRate()
    End Sub
End Class
