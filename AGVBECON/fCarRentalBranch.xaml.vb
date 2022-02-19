Imports AGVBE

Partial Public Class fCarRentalBranch
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

    Private Sub fCarRentalBranch_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
            Me.Title = "Add New Branch"
            Dim sCityName As String = ""
            Dim sStateName As String = ""
            Dim lID As Integer = 0
            g_GenerateRandomCity(sCityName, sStateName, lID, mp_oParent.mp_yDataSourceType)
            txtCity.Text = sCityName
            txtBranchName.Text = sCityName
            txtState.Text = sStateName
            txtPhone.Text = g_GenerateRandomPhone("")
            txtManagerName.Text = g_GenerateRandomName(False, mp_oParent.mp_yDataSourceType)
            txtManagerMobile.Text = g_GenerateRandomPhone(txtPhone.Text.Substring(0, 5))
            txtAddress.Text = g_GenerateRandomAddress(mp_oParent.mp_yDataSourceType)
            txtZIP.Text = g_GenerateRandomZIP()
        Else
            Me.Title = "Edit Branch"
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
                txtCity.Text = oCRRow.sCity
                txtBranchName.Text = oCRRow.sBranchName
                txtState.Text = oCRRow.sState
                txtPhone.Text = oCRRow.sPhone
                txtManagerName.Text = oCRRow.sManagerName
                txtManagerMobile.Text = oCRRow.sManagerMobile
                txtAddress.Text = oCRRow.sAddress
                txtZIP.Text = oCRRow.sZIP
            End If
        End If
    End Sub

    Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles cmdOK.Click
        Dim oRow As clsRow = Nothing
        If mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            '//TODO
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            '//TODO
        ElseIf mp_oParent.mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            Dim oCRRow As AG_CR_Row = Nothing
            If mp_yDialogMode = PRG_DIALOGMODE.DM_ADD Then
                Dim lRowID As Integer = 0
                For Each oCRRow In mp_oParent.mp_o_AG_CR_Rows
                    If oCRRow.lRowID > lRowID Then
                        lRowID = oCRRow.lRowID
                    End If
                Next
                lRowID = lRowID + 1
                oCRRow = New AG_CR_Row
                oCRRow.lRowID = lRowID
                oCRRow.lOrder = mp_oParent.ActiveGanttVBECtl1.Rows.Count() + 1
                oRow = mp_oParent.ActiveGanttVBECtl1.Rows.Add("K" & lRowID)
                oRow.Node.Depth = 0
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
            oCRRow.lDepth = 0
            oCRRow.sCity = txtCity.Text
            oCRRow.sBranchName = txtBranchName.Text
            oCRRow.sState = txtState.Text
            oCRRow.sPhone = txtPhone.Text
            oCRRow.sManagerName = txtManagerName.Text
            oCRRow.sManagerMobile = txtManagerMobile.Text
            oCRRow.sAddress = txtAddress.Text
            oCRRow.sZIP = txtZIP.Text
        End If
        oRow.Text = txtBranchName.Text & ", " & txtState.Text & vbCrLf & "Phone: " & txtPhone.Text
        oRow.MergeCells = True
        oRow.Container = False
        oRow.StyleIndex = "Branch"
        oRow.ClientAreaStyleIndex = "BranchCA"
        oRow.UseNodeImages = True
        oRow.Node.ExpandedImage = mp_oParent.GetImage("CarRental/minus.jpg")
        oRow.Node.Image = mp_oParent.GetImage("CarRental/plus.jpg")
        oRow.AllowMove = False
        oRow.AllowSize = False
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


End Class
