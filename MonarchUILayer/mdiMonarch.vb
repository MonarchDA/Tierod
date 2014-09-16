Imports System.Windows.Forms
Imports IFLBaseDataLayer
Imports IFLCustomUILayer
Imports IFLCommonLayer

Public Class mdiWeldedCylinder

#Region "Default Subs"

    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs)
        ' Create a new instance of the child form.
        Dim ChildForm As New System.Windows.Forms.Form
        ' Make it a child of this MDI form before showing it.
        ChildForm.MdiParent = Me

        m_ChildFormNumber += 1
        ChildForm.Text = "Window " & m_ChildFormNumber

        ChildForm.Show()
    End Sub

    Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs)
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = OpenFileDialog.FileName
            ' TODO: Add code here to open the file.
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

        If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = SaveFileDialog.FileName
            ' TODO: Add code here to save the current contents of the form to a file.
        End If
    End Sub

    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Global.System.Windows.Forms.Application.Exit()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Use My.Computer.Clipboard.GetText() or My.Computer.Clipboard.GetData to retrieve information from the clipboard.
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticleToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer = 0

#End Region

#Region "SubProdeures"

    Private Sub mdiWeldedCylinder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetLoginDetails()
        ObjClsWeldedCylinder = New clsWeldedCylinder
        ObjClsWeldedCylinder.InitialiseAllChildFormObjects()
        Dim oFrmBaseEndDetails As New frmBaseEndDetails
        ObjClsWeldedCylinder.ObjCurrentForm = oFrmBaseEndDetails
        oFrmBaseEndDetails.TopLevel = False
        oFrmBaseEndDetails.Show()
        pnlChildFormArea.Controls.Add(oFrmBaseEndDetails)
        oFrmBaseEndDetails.Dock = DockStyle.Fill
        btnBack.Enabled = False
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click, btnNext.Click
        For Each oItem As Object In ObjClsWeldedCylinder.FormNavigationOrder
            If oItem(clsWeldedCylinder.EOrderOfFormNavigationArraylist.CurrentFormName).ToString.Equals(ObjClsWeldedCylinder.ObjCurrentForm.name) Then
                Dim oForm As Form = Nothing
                Dim oCurrentForm As Form = Nothing
                If sender.Equals(btnBack) Then
                    oForm = CType(oItem(clsWeldedCylinder.EOrderOfFormNavigationArraylist.PreviousFormObject), Form)
                    If Not IsNothing(oForm) Then
                        pnlChildFormArea.Controls.Clear()
                        ObjClsWeldedCylinder.ObjCurrentForm = oForm
                        oForm.TopLevel = False
                        oForm.Dock = DockStyle.Fill
                        oForm.Show()
                        pnlChildFormArea.Controls.Add(oForm)
                        Exit For
                    End If
                ElseIf sender.Equals(btnNext) Then
                    oForm = CType(oItem(clsWeldedCylinder.EOrderOfFormNavigationArraylist.NextFormObject), Form)
                    If Not IsNothing(oForm) Then
                        pnlChildFormArea.Controls.Clear()
                        ObjClsWeldedCylinder.ObjCurrentForm = oForm
                        oForm.TopLevel = False
                        oForm.Dock = DockStyle.Fill
                        oForm.Show()
                        pnlChildFormArea.Controls.Add(oForm)
                        Exit For
                    End If
                End If
            End If
        Next

        btnBack.Enabled = True
        btnNext.Enabled = True
        If ObjClsWeldedCylinder.ObjCurrentForm.name = "frmBaseEndDetails" Then
            btnBack.Enabled = False
        ElseIf ObjClsWeldedCylinder.ObjCurrentForm.name = "frmPortInTubeDetails" Then
            btnNext.Enabled = False
        End If
    End Sub

    Private Sub GetLoginDetails()
        Try
            IFLConnectionObject = IFLConnectionClass.GetConnectionObject("IEGHPDCWS106\SQLEXPRESS", "MIL", "System.Data.SqlClient", , , "sspi")
            lvwLoginDetails.Items.Clear()
            If Not IsNothing(IFLConnectionObject) Then
                lvwLoginDetails.Columns.Add("Property", 107, HorizontalAlignment.Center)
                lvwLoginDetails.Columns.Add("Value", 200, HorizontalAlignment.Center)

                Dim oListviewItem1 As ListViewItem
                oListviewItem1 = lvwLoginDetails.Items.Add("ServerName")
                lvwLoginDetails.Items(0).SubItems.Add(IFLConnectionObject.ServerName)

                Dim oListViewItem2 As ListViewItem
                oListViewItem2 = lvwLoginDetails.Items.Add("DataBase")
                lvwLoginDetails.Items(1).SubItems.Add(IFLConnectionObject.DataBaseName)

                Dim oListViewItem3 As ListViewItem
                oListViewItem3 = lvwLoginDetails.Items.Add("UserName")
                lvwLoginDetails.Items(2).SubItems.Add(My.User.Name)

                Dim oListViewItem4 As ListViewItem
                oListViewItem4 = lvwLoginDetails.Items.Add("ComputerName")
                lvwLoginDetails.Items(3).SubItems.Add(My.Computer.Name)

            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnHome_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHome.Click
        Application.Exit()
    End Sub

#End Region
   
End Class
