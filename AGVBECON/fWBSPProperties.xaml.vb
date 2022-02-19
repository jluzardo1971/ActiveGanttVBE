Partial Public Class fWBSPProperties
    Inherits ChildWindow

    Dim mp_oParent As fWBSProject

    Public Sub New(ByVal oParent As fWBSProject)
        InitializeComponent()
        mp_oParent = oParent
    End Sub

    Private Sub fWBSPProperties_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        chkCheckBoxes.IsChecked = mp_oParent.ActiveGanttVBECtl1.Treeview.CheckBoxes
        chkImages.IsChecked = mp_oParent.ActiveGanttVBECtl1.Treeview.Images
        chkPlusMinusSigns.IsChecked = mp_oParent.ActiveGanttVBECtl1.Treeview.PlusMinusSigns
        chkFullColumnSelect.IsChecked = mp_oParent.ActiveGanttVBECtl1.Treeview.FullColumnSelect
        chkTreeLines.IsChecked = mp_oParent.ActiveGanttVBECtl1.Treeview.TreeLines
        chkEnforcePredecessors.IsChecked = mp_oParent.ActiveGanttVBECtl1.EnforcePredecessors
        cboPredecessorMode.SelectedValue = System.Convert.ToInt32(mp_oParent.ActiveGanttVBECtl1.PredecessorMode).ToString()
    End Sub

    Private Sub OKButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles OKButton.Click
        mp_oParent.ActiveGanttVBECtl1.Treeview.CheckBoxes = System.Convert.ToBoolean(chkCheckBoxes.IsChecked)
        mp_oParent.ActiveGanttVBECtl1.Treeview.Images = System.Convert.ToBoolean(chkImages.IsChecked)
        mp_oParent.ActiveGanttVBECtl1.Treeview.PlusMinusSigns = System.Convert.ToBoolean(chkPlusMinusSigns.IsChecked)
        mp_oParent.ActiveGanttVBECtl1.Treeview.FullColumnSelect = System.Convert.ToBoolean(chkFullColumnSelect.IsChecked)
        mp_oParent.ActiveGanttVBECtl1.Treeview.TreeLines = System.Convert.ToBoolean(chkTreeLines.IsChecked)
        mp_oParent.ActiveGanttVBECtl1.EnforcePredecessors = System.Convert.ToBoolean(chkEnforcePredecessors.IsChecked)
        mp_oParent.ActiveGanttVBECtl1.PredecessorMode = DirectCast(System.Convert.ToInt32(cboPredecessorMode.SelectedValue), AGVBE.E_PREDECESSORMODE)
        mp_oParent.ActiveGanttVBECtl1.Redraw()
        Me.DialogResult = True
    End Sub


End Class
