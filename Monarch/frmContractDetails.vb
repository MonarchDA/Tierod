Imports MonarchFunctionalLayer
Imports System.Drawing
Imports System.Drawing.Graphics
Public Class frmContractDetails

    Private _browseFileName As String
    Private oReadValuesFromExcel As New ReadValuesFromExcel
    Private oExcel As ExcelUtil
    Private _btnVisible As Boolean
    Private _IsSecondRowRead As Boolean = False
    Private excelRowNumber As Integer = 3
    ' Private _isBrowseBtnClicked As Boolean = False
    Private _strSearchedCustomer As String

    Dim oRowNumber As Integer = 3
    Dim isFirstRowGenerated As Boolean = False

    Private Property ReadValuesFromExcel() As ReadValuesFromExcel
        Get
            Return oReadValuesFromExcel
        End Get
        Set(ByVal value As ReadValuesFromExcel)
            oReadValuesFromExcel = value
        End Set
    End Property

    Public Property BrowseFileName() As String
        Get
            Return _browseFileName
        End Get
        Set(ByVal value As String)
            _browseFileName = value
        End Set
    End Property

    Public Property BtnVisible() As Boolean
        Get
            Return _btnVisible
        End Get
        Set(ByVal value As Boolean)
            _btnVisible = value
        End Set
    End Property

    'Public Property IsBrowseBtnClicked() As Boolean
    '    Get
    '        Return _isBrowseBtnClicked
    '    End Get
    '    Set(ByVal value As Boolean)
    '        _isBrowseBtnClicked = value
    '    End Set
    'End Property


    Public Property IsSecondRowRead() As Boolean
        Get
            Return _IsSecondRowRead
        End Get
        Set(ByVal value As Boolean)
            _IsSecondRowRead = value
        End Set
    End Property

    Private Sub txtlPartCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtlPartCode.TextChanged

        txtlPartCodeTextChanged()

    End Sub

    Public Sub txtlPartCodeTextChanged()

        PartCode1 = txtlPartCode.Text

    End Sub

    Private Sub cmbAssemblyType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                Handles cmbAssemblyType.SelectedIndexChanged

        cmbAssemblyTypeSelectedIndexChanged()

    End Sub

    Public Sub cmbAssemblyTypeSelectedIndexChanged()

        If cmbAssemblyType.Text <> "" Then
            AssemblyType = cmbAssemblyType.Text
        End If

    End Sub

    Private Sub chkManageCustomers_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                        Handles chkManageCustomers.CheckedChanged
        '22_02_2010    RAGAVA
        Try
            If chkManageCustomers.Checked = True Then
                GroupBox1.Visible = True
                LabelGradient1.Visible = True
                pnlManageCustomerDetails.Visible = True
                txtCustomerName_Add.Enabled = True
                cmbCustomerName_Delete.Enabled = True
            Else
                GroupBox1.Visible = False
                LabelGradient1.Visible = False
                pnlManageCustomerDetails.Visible = False
                txtCustomerName_Add.Enabled = False
                cmbCustomerName_Delete.Enabled = False
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtCustomerName_Add_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                      Handles txtCustomerName_Add.TextChanged
        '22_02_2010   RAGAVA
        Try
            Dim strCustomerName As String
            strCustomerName = Trim(txtCustomerName_Add.Text)
            If strCustomerName <> "" Then
                'If sender.Name = "txtCustomerName" Then
                '    ListBox1.Visible = True
                '    ListBox1.Location = New Point(ListBox1.Location.X, ListBox1.Location.Y + 25)
                'End If
                Dim strQuery As String = "Select * from CustomerManageDetails where CustomerName like '" & _
                                                    strCustomerName & "%' order by CustomerName Asc"
                Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                ListBox1.BringToFront()
                ListBox1.Items.Clear()
                If objDT.Rows.Count > 0 Then
                    ListBox1.Visible = True
                    For Each dr As DataRow In objDT.Rows
                        ListBox1.Items.Add(dr(0).ToString)
                    Next
                Else
                    ListBox1.Visible = False
                End If
            Else
                ListBox1.Visible = False
            End If
        Catch ex As Exception
        End Try
    End Sub

    '22_02_2010   RAGAVA
    Private Function GetPoint(ByVal textBoxControl As TextBox) As Point
        Try
            Dim coord As Point
            Dim size As SizeF
            Dim graphics As Graphics
            'graphics = graphics.FromHwnd(textBoxControl.Handle)
            graphics = Drawing.Graphics.FromHwnd(textBoxControl.Handle)
            size = graphics.MeasureString(textBoxControl.Text.Substring(0, textBoxControl.SelectionStart), _
                                                                textBoxControl.Font)
            coord = New Point(Convert.ToInt16(size.Width) + textBoxControl.Location.X, Convert.ToInt16(size.Height) _
                                                                            + textBoxControl.Location.Y)
            Return coord
        Catch ex As Exception
        End Try
    End Function

    '22_02_2010   RAGAVA
    Private Sub txtCustomerName_Add_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
                                                                        Handles txtCustomerName_Add.KeyUp
        Try
            If e.KeyCode = Keys.Down Or e.KeyCode = Keys.Up Then
                ListBox1.Focus()
            End If
        Catch ex As Exception
        End Try
    End Sub

    '22_02_2010   RAGAVA
    Private Sub ListBox1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.Leave
        Try
            txtCustomerName_Add.Text = ListBox1.Text.ToString()
            txtCustomerName_Add.SelectionStart = txtCustomerName_Add.Text.Length
            ListBox1.SendToBack()
            txtCustomerName_Add.Focus()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ListBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.LostFocus
        ListBox1.Visible = False
    End Sub

    '22_02_2010   RAGAVA
    Private Sub cmbCustomerName_Delete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCustomerName_Delete.Click
        Try
            Dim strQuery As String = "Select * from CustomerManageDetails order by CustomerName Asc"
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            'cmbCustomerName_Delete.DataSource = objDT
            cmbCustomerName_Delete.Items.Clear()
            cmbCustomerName_Delete.Items.Add(" ")
            For Each dr As DataRow In objDT.Rows
                cmbCustomerName_Delete.Items.Add(dr(0).ToString)
                'cmbCustomerName_Delete.DisplayMember = "CustomerName"
            Next
        Catch ex As Exception
        End Try
    End Sub

    '22_02_2010   RAGAVA
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Try
            If Trim(txtCustomerName_Add.Text) <> "" Then
                Dim strQuery As String = "Select * from CustomerManageDetails Where CustomerName = '" & Trim(txtCustomerName_Add.Text) & "'"
                Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery)
                If objDT1.Rows.Count = 0 Then
                    strQuery = "Insert into CustomerManageDetails Values ('" & UCase(Trim(txtCustomerName_Add.Text)) & "')"
                    Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    '22_02_2010   RAGAVA
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If Trim(cmbCustomerName_Delete.Text) <> "" Then
                Dim strQuery As String = "Delete from CustomerManageDetails Where CustomerName ='" & Trim(cmbCustomerName_Delete.Text) & "'"
                Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                cmbCustomerName_Delete_Click(sender, e)
            End If
        Catch ex As Exception
        End Try
    End Sub

    '22_02_2010   RAGAVA
    Private Sub ListBox1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) _
                                                                                    Handles ListBox1.MouseDoubleClick
        Try
            txtCustomerName_Add.Text = Trim(ListBox1.Text)
            txtCustomerName_Add.Focus()
        Catch ex As Exception
        End Try
    End Sub

    '22_02_2010   RAGAVA
    Private Sub cmbCustomerName_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCustomerName.Click
        AddCustomerNames()
    End Sub

    Private Sub AddCustomerNames()
        Try
            Dim strQuery As String = "Select * from CustomerManageDetails order by CustomerName Asc"
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            cmbCustomerName.Items.Clear()
            cmbCustomerName.Items.Add(" ")
            For Each dr As DataRow In objDT.Rows
                cmbCustomerName.Items.Add(dr(0).ToString)
            Next
            If blnRevision Then
                cmbCustomerName.Text = CustomerName
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub cmbCustomerName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                            Handles cmbCustomerName.SelectedIndexChanged

        cmbCustomerNameSelectedIndexChanged()

    End Sub

    Public Sub cmbCustomerNameSelectedIndexChanged()

        If Not cmbCustomerName.Text = "" Then
            CustomerName = cmbCustomerName.Text
        End If

    End Sub

    Private Sub frmContractDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            If blnRevision Then
                HideControls()
            End If
            cmbAssemblyType.SelectedIndex = 0 'ANUP 16-12-2010 
            ColorTheForm()
            AddCustomerNames()
            ShowingChangePartNumberButton()

            'Me.txtlPartCode.Text = PartCode1      '16_08_2011   RAGAVA
            'If Module1.ArraList2.Count > 0 Then
            '    cmbCustomerName.Text = Module1.ArraList2.Item(0).ToString()
            '    cmbAssemblyType.Text = "Tie Rod Cylinder Assembly"
            '    txtlPartCode.Text = Convert.ToInt32(Module1.ArraList2.Item(1).ToString())
            'End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ColorTheForm()
        FunctionalClassObject.LabelGradient_GreenBorder_ColoringTheScreens(LabelGradient3, LabelGradient5, LabelGradient4, LabelGradient2)
        FunctionalClassObject.LabelGradient_OrangeBorder_ColoringTheScreens(lblGradientContractDetails)
        FunctionalClassObject.LabelGradient_OrangeBorder_ColoringTheScreens(lblBackBrowse)
        FunctionalClassObject.LabelGradient_OrangeBorder_ColoringTheScreens(LabelGradient1)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(lblGradientPrimaryInformation)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(lblLogInformation)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient6)
    End Sub

    Private Sub ShowingChangePartNumberButton()
        Try
            If IsNew_Revision_Released = "Revision" OrElse IsNew_Revision_Released = "Released" Then
                btnChangePartNumber.Visible = True
            ElseIf IsNew_Revision_Released = "New" Then
                btnChangePartNumber.Visible = False
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnChangePartNumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChangePartNumber.Click
        Try
            If IsPartNumberUpdatedToDB() = False Then
                MessageBox.Show("Part Number is not updated", "Error in Updating", MessageBoxButtons.OK, _
                                                            MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Function IsPartNumberUpdatedToDB() As Boolean
        IsPartNumberUpdatedToDB = False
        Try
            Dim strContractNumber As String = String.Empty
            Dim strQuery As String = String.Empty
            For Each listviewItem As ListViewItem In ofrmMonarch.lvwContractDetails.SelectedItems
                strContractNumber = listviewItem.SubItems(0).Text
            Next
            If Not IsNothing(strContractNumber) Then
                If DoesContractExists(cmbCustomerName.Text, strContractNumber) Is Nothing Then
                    strQuery = "insert into dbo.ContractDetails_Revision values('" & strContractNumber & "','" & _
                                                    cmbCustomerName.Text & "','" & txtlPartCode.Text & "')"
                Else
                    strQuery = "update dbo.ContractDetails_Revision set CustomerPartNUmber ='" & txtlPartCode.Text & _
                        "' where CustomerName ='" & cmbCustomerName.Text & "' and ContractNumber ='" & strContractNumber & "'"
                End If
                Dim strQuery2 As String = "update dbo.ContractMaster set CustomerPartCode ='" & _
                                    txtlPartCode.Text & "' where ContractNumber ='" & strContractNumber & "'"
                Dim Query2Updation As Boolean = IFLConnectionObject.ExecuteQuery(strQuery2)
                IsPartNumberUpdatedToDB = IFLConnectionObject.ExecuteQuery(strQuery)
                If Query2Updation AndAlso IsPartNumberUpdatedToDB Then
                    IsPartNumberUpdatedToDB = True
                Else
                    IsPartNumberUpdatedToDB = False
                End If
            Else
                IsPartNumberUpdatedToDB = False
                MessageBox.Show("Contract Number is not available")
            End If
        Catch ex As Exception
            IsPartNumberUpdatedToDB = False
        End Try

    End Function

    Private Function DoesContractExists(ByVal strCustomerName As String, ByVal strContractNumber As String) As DataRow

        DoesContractExists = Nothing
        Try
            Dim strQuery As String = "select * from ContractDetails_Revision where ContractNumber ='" & strContractNumber _
                                                        & "' and CustomerName ='" & strCustomerName & "'"
            DoesContractExists = IFLConnectionObject.GetDataRow(strQuery)
            If DoesContractExists Is Nothing Then
                DoesContractExists = Nothing
            End If
        Catch ex As Exception
            DoesContractExists = Nothing
        End Try

    End Function

    Private Function OpenSheet(ByVal strFileName As String) As ExcelUtil
        Dim oExcel As New ExcelUtil

        If Not oExcel.OpenWorkBook(strFileName, False) Then
            Return Nothing
        End If

        If Not oExcel.OpenWorksheet("Sheet1") Then
            Return Nothing
        End If

        Return oExcel
    End Function

    Private Function GetFileLocation() As String

        Dim strFileName As String = Nothing
        Dim oOpenFileWindow As New OpenFileDialog
        oOpenFileWindow.Multiselect = False
        oOpenFileWindow.InitialDirectory = "D:\Raju_Home\SDGT\HookUps\Testing"
        oOpenFileWindow.Filter = " Excel Files 2003 (*.xls)|*.xls|Excel Files 2007 (*.xlsx)|*.xlsx|Excel Files 2007 (*.xlsm)|*.xlsm"
        oOpenFileWindow.RestoreDirectory = True
        oOpenFileWindow.Title = "Select HookUp"
        If oOpenFileWindow.ShowDialog = Windows.Forms.DialogResult.OK Then
            strFileName = oOpenFileWindow.FileName
            'Else
            '    txtFileLocation.Text = ""
        End If
        Return strFileName

    End Function

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click         ' sugandhi

        IsNewBtnClicked = False
        lVLogInformation.Items.Clear()
        BrowseFileName = GetFileLocation()
        If BrowseFileName Is Nothing Then
            _btnVisible = False
            Return
        End If
        'LVCustomer.Items.Clear()
        'lvwContractDetails.Items.Clear()

        oExcel = OpenSheet(BrowseFileName)

        If oExcel Is Nothing Then
            _btnVisible = False
            Module1.LogInfo.Add("File not found.")
            ErrorMessages()
            mdiMonarch.BtnsVisibleFalse()
            Exit Sub
        Else
            Me.Cursor = Cursors.WaitCursor
            excelRowNumber = 3
            Module1.LogInfo.Clear()
            mdiMonarch.BtnsVisibleFalse()
            Module1.BtnBrowseClicked = True
            CheckExcelRowValues()

            If Module1.LogInfo.Count > 0 Then
                Exit Sub
            End If

            If Module1.ArraList1.Count = 0 Then
                Me.Cursor = Cursors.Default
                Module1.LogInfo.Add("    Row not found.")
                ErrorMessages()
                Exit Sub
            Else
                SettingValuesFromExcel(Module1.ArraList1)
            End If

            For i As Integer = 0 To Module1.RowCount - 1

                If Not excelRowNumber = Module1.RowCount + 3 Then

                    ofrmTieRod1.ClearAllFielsTieRod1()
                    ofrmTieRod2.ClearAllFielsTieRod2()

                    mdiMonarch.CheckAllFormsValues(excelRowNumber)
                    excelRowNumber = excelRowNumber + 1
                    If Not excelRowNumber = Module1.RowCount + 3 Then
                        SettingSelectedArrayListValues(excelRowNumber)
                    End If
                End If

            Next
            If Module1.LogInfo.Count > 0 Then
                mdiMonarch.ShowFrmContactDetails()
                ErrorMessages()
                btnBrowse.Enabled = False
                mdiMonarch.BtnsVisibleFalse()
                Me.Cursor = Cursors.Default
            Else
                _btnVisible = True
                btnBrowse.Enabled = False
            End If

        End If

        ofrmTieRod1.ClearAllFielsTieRod1()
        ofrmTieRod2.ClearAllFielsTieRod2()

        SettingValuesFromExcel(Module1.ArraList1)

        cmbCustomerName.Text = oReadValuesFromExcel.CustomerName
        cmbAssemblyType.Text = oReadValuesFromExcel.Type
        txtlPartCode.Text = oReadValuesFromExcel.CustomerPortCode

        mdiMonarch.GetExcelFile()

        Me.Cursor = Cursors.Default

        oExcel.Close()

    End Sub

    Public Sub CheckExcelRowValues()

        Dim j As Integer = 0
        For excelRow As Integer = 3 To 102
            If Not oExcel.Read(excelRow, 1) Is Nothing Then
                Module1.RowCount = j + 1
                j = j + 1
                If Not CheckExcelValues(excelRow) Then
                    ErrorMessages()
                    Exit Sub
                End If

                If Module1.RowCount = 1 Then
                    Module1.ArraList1 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 2 Then
                    Module1.ArraList2 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 3 Then
                    Module1.ArraList3 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 4 Then
                    Module1.ArraList4 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 5 Then
                    Module1.ArraList5 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 6 Then
                    Module1.ArraList6 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 7 Then
                    Module1.ArraList7 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 8 Then
                    Module1.ArraList8 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 9 Then
                    Module1.ArraList9 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 10 Then
                    Module1.ArraList10 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 11 Then
                    Module1.ArraList11 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 12 Then
                    Module1.ArraList12 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 13 Then
                    Module1.ArraList13 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 14 Then
                    Module1.ArraList14 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 15 Then
                    Module1.ArraList15 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 16 Then
                    Module1.ArraList16 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 17 Then
                    Module1.ArraList17 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 18 Then
                    Module1.ArraList18 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 19 Then
                    Module1.ArraList19 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 20 Then
                    Module1.ArraList20 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 21 Then
                    Module1.ArraList21 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 22 Then
                    Module1.ArraList22 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 23 Then
                    Module1.ArraList23 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 24 Then
                    Module1.ArraList24 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 25 Then
                    Module1.ArraList25 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 26 Then
                    Module1.ArraList26 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 27 Then
                    Module1.ArraList27 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 28 Then
                    Module1.ArraList28 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 29 Then
                    Module1.ArraList29 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 30 Then
                    Module1.ArraList30 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 31 Then
                    Module1.ArraList31 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 32 Then
                    Module1.ArraList32 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 33 Then
                    Module1.ArraList33 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 34 Then
                    Module1.ArraList34 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 35 Then
                    Module1.ArraList35 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 36 Then
                    Module1.ArraList36 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 37 Then
                    Module1.ArraList37 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 38 Then
                    Module1.ArraList38 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 39 Then
                    Module1.ArraList39 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 40 Then
                    Module1.ArraList40 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 41 Then
                    Module1.ArraList41 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 42 Then
                    Module1.ArraList42 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 43 Then
                    Module1.ArraList43 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 44 Then
                    Module1.ArraList44 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 45 Then
                    Module1.ArraList45 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 46 Then
                    Module1.ArraList46 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 47 Then
                    Module1.ArraList47 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 48 Then
                    Module1.ArraList48 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 49 Then
                    Module1.ArraList49 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 50 Then
                    Module1.ArraList50 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 51 Then
                    Module1.ArraList51 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 52 Then
                    Module1.ArraList52 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 53 Then
                    Module1.ArraList53 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 54 Then
                    Module1.ArraList54 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 55 Then
                    Module1.ArraList55 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 56 Then
                    Module1.ArraList56 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 57 Then
                    Module1.ArraList57 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 58 Then
                    Module1.ArraList58 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 59 Then
                    Module1.ArraList59 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 60 Then
                    Module1.ArraList60 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 61 Then
                    Module1.ArraList61 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 62 Then
                    Module1.ArraList62 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 63 Then
                    Module1.ArraList63 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 64 Then
                    Module1.ArraList64 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 65 Then
                    Module1.ArraList65 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 66 Then
                    Module1.ArraList66 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 67 Then
                    Module1.ArraList67 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 68 Then
                    Module1.ArraList68 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 69 Then
                    Module1.ArraList69 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 70 Then
                    Module1.ArraList70 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 71 Then
                    Module1.ArraList71 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 72 Then
                    Module1.ArraList72 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 73 Then
                    Module1.ArraList73 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 74 Then
                    Module1.ArraList74 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 75 Then
                    Module1.ArraList75 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 76 Then
                    Module1.ArraList76 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 77 Then
                    Module1.ArraList77 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 78 Then
                    Module1.ArraList78 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 79 Then
                    Module1.ArraList79 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 80 Then
                    Module1.ArraList80 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 81 Then
                    Module1.ArraList81 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 82 Then
                    Module1.ArraList82 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 83 Then
                    Module1.ArraList83 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 84 Then
                    Module1.ArraList84 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 85 Then
                    Module1.ArraList85 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 86 Then
                    Module1.ArraList86 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 87 Then
                    Module1.ArraList87 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 88 Then
                    Module1.ArraList88 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 89 Then
                    Module1.ArraList89 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 90 Then
                    Module1.ArraList90 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 91 Then
                    Module1.ArraList91 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 92 Then
                    Module1.ArraList92 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 93 Then
                    Module1.ArraList93 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 94 Then
                    Module1.ArraList94 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 95 Then
                    Module1.ArraList95 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 96 Then
                    Module1.ArraList96 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 97 Then
                    Module1.ArraList97 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 98 Then
                    Module1.ArraList98 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 99 Then
                    Module1.ArraList99 = SettingValuesFromExcelToArrayList(excelRow)
                ElseIf Module1.RowCount = 100 Then
                    Module1.ArraList100 = SettingValuesFromExcelToArrayList(excelRow)
                End If

            Else
                Exit For
            End If
        Next

    End Sub

    Public Sub ReadingExcelRowvalues()

        CylinderCodeNumber = ofrmTieRod1.GetCylinderCodeNumber()
        'System.Threading.Thread.Sleep(1000)
        mdiMonarch.GenerateBtnFuctionality(mdiMonarch.GenerateBtnSender)

    End Sub

    Private Function CheckExcelValues(ByVal rowNumber As Integer) As Boolean       'SUGANDHI

        Dim bool As Boolean = False

        For j As Integer = 1 To 42

            If oExcel.Read(rowNumber, j) Is Nothing Then
                If j = 4 Then
                    If Not oExcel.Read(rowNumber, 3) = "TP-High" Or oExcel.Read(rowNumber, 3) = "TP-Low" Then
                    Else
                        Module1.LogInfo.Add(oExcel.Read(2, j) + "  :  " + "Please fill the Cell at (" + rowNumber.ToString() _
                                                                + " , " + j.ToString() + ")")
                        bool = True
                    End If
                ElseIf j = 15 Then
                    If oExcel.Read(rowNumber, j) Is Nothing Then

                    End If
                ElseIf j = 41 Then
                    If oExcel.Read(rowNumber, j) Is Nothing Then

                    End If
                ElseIf j = 42 Then
                    If oExcel.Read(rowNumber, j) Is Nothing Then

                    End If
                Else
                    Module1.LogInfo.Add(oExcel.Read(2, j) + "  :  " + "Please fill the Cell at (" + rowNumber.ToString() + _
                                                                                " , " + j.ToString() + ")")
                    bool = True
                End If
            End If
        Next

        If bool Then
            Return False
        Else
            Return True
            isFirstRowGenerated = True
        End If

    End Function

    Public Sub ErrorMessages()                 'SUGANDHI

        Dim listView As ListViewItem

        For Each oMessage As String In Module1.LogInfo

            listView = lVLogInformation.Items.Add(oMessage)
            'listView.SubItems.Add(oMessage.getMessage)
            listView.ForeColor = Color.Red
        Next

    End Sub

    Public Function SettingValuesFromExcelToArrayList(ByVal RowNumber As Integer) As ArrayList         ' sugandhi

        Dim arrList As New ArrayList

        arrList.Add(oExcel.Read("B" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("C" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("D" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("E" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("F" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("G" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("H" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("I" + RowNumber.ToString()))

        If oExcel.Read("J" + RowNumber.ToString()).ToString() = "Yes" Then
            arrList.Add("True")
        Else
            arrList.Add("False")
        End If

        arrList.Add(oExcel.Read("K" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("L" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("M" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("N" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("O" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("P" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("Q" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("R" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("S" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("T" + RowNumber.ToString()))

        If oExcel.Read("U" + RowNumber.ToString()).ToString() = "Yes" Then
            arrList.Add("True")
        Else
            arrList.Add("False")
        End If

        arrList.Add(oExcel.Read("V" + RowNumber.ToString()))

        If oExcel.Read("W" + RowNumber.ToString()).ToString() = "Yes" Then
            arrList.Add("True")
        Else
            arrList.Add("False")
        End If

        If oExcel.Read("X" + RowNumber.ToString()).ToString() = "Yes" Then
            arrList.Add("True")
        Else
            arrList.Add("False")
        End If

        arrList.Add(oExcel.Read("Y" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("Z" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("AA" + RowNumber.ToString()))

        If oExcel.Read("AB" + RowNumber.ToString()).ToString() = "Yes" Then
            arrList.Add("True")
        Else
            arrList.Add("False")
        End If

        ' arrList.Add(oExcel.Read("AB" + excelRowNoTable1.ToString()))
        arrList.Add(oExcel.Read("AC" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("AD" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("AE" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("AF" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("AG" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("AH" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("AI" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("AJ" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("AK" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("AL" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("AM" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("AN" + RowNumber.ToString()))
        arrList.Add(oExcel.Read("AO" + RowNumber.ToString()))

        If Not oExcel.Read("AP" + RowNumber.ToString()) Is Nothing Then
            If oExcel.Read("AP" + RowNumber.ToString()).ToString() = "Yes" Then
                arrList.Add("True")
            Else
                arrList.Add("False")
            End If
        Else
            arrList.Add("True")
        End If
       
        If Not oExcel.Read("AQ" + RowNumber.ToString()) Is Nothing Then
            If oExcel.Read("AQ" + RowNumber.ToString()).ToString() = "Yes" Then
                arrList.Add("True")
            Else
                arrList.Add("False")
            End If
        Else
            arrList.Add("False")
        End If
      
        'arrList.Add(oExcel.Read("AR" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("AS" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("AT" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("AU" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("AV" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("AW" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("AX" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("AY" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("AZ" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("BA" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("BB" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("BC" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("BD" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("BE" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("BF" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("BG" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("BH" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("BI" + RowNumber.ToString()))
        'arrList.Add(oExcel.Read("BJ" + RowNumber.ToString()))

        Return arrList

    End Function

    Public Sub SettingValuesFromExcel(ByVal arrList As ArrayList)          'SUGANDHI

        oReadValuesFromExcel.CustomerName = arrList.Item(0).ToString()
        oReadValuesFromExcel.CustomerPortCode = arrList.Item(1).ToString()
        oReadValuesFromExcel.Series = arrList.Item(2).ToString()

        If oReadValuesFromExcel.Series = "TP-High" Or oReadValuesFromExcel.Series = "TP-Low" Then
            oReadValuesFromExcel.RephasingPortPosition = arrList.Item(3).ToString()
        End If

        oReadValuesFromExcel.Style = arrList.Item(4).ToString()
        oReadValuesFromExcel.Bore = arrList.Item(5).ToString()
        oReadValuesFromExcel.StrokeLength = Convert.ToInt32(arrList.Item(6).ToString())
        oReadValuesFromExcel.RodAdder = Convert.ToInt32(arrList.Item(7).ToString())

        oReadValuesFromExcel.StopTube = arrList.Item(8).ToString()

        oReadValuesFromExcel.ClevisCapPinHole = arrList.Item(9).ToString()
        oReadValuesFromExcel.RodClevisPinHole = arrList.Item(10).ToString()
        oReadValuesFromExcel.StandardRunQty = Convert.ToInt32(arrList.Item(11).ToString())
        oReadValuesFromExcel.RodMaterials = arrList.Item(12).ToString()

        'Dim s As String = Convert.ToDouble(arrList.Item(13).ToString())
        'Dim words As String() = s.Split(New Char() {"."c})

        'Dim word As String = words(1)
        'If Not word.Length = 2 Then
        '    word = word + "0"
        '    word = words(0) + "." + word
        '    oReadValuesFromExcel.RodDiameter = Convert.ToDouble(word)
        'Else
        oReadValuesFromExcel.RodDiameter = Convert.ToDouble(arrList.Item(13).ToString())
        ' End If
        If Not arrList.Item(14) Is Nothing Then
            oReadValuesFromExcel.RodDeratedPressureAtmaximumExtension = Convert.ToDouble(arrList.Item(14).ToString())
        End If

        oReadValuesFromExcel.PortOrientationForClevisCap = arrList.Item(15).ToString()
        oReadValuesFromExcel.PortOrientationForRodCap = arrList.Item(16).ToString()
        oReadValuesFromExcel.ClevisCapPort = arrList.Item(17).ToString()
        oReadValuesFromExcel.RodCapPort = arrList.Item(18).ToString()

        oReadValuesFromExcel.StrokeControl = arrList.Item(19).ToString()

        oReadValuesFromExcel.StrokeControlStages = Convert.ToInt32(arrList.Item(20).ToString())

        oReadValuesFromExcel.ClevisCapPins = arrList.Item(21).ToString()

        oReadValuesFromExcel.RodClevisPins = arrList.Item(22).ToString()

        oReadValuesFromExcel.PinMaterial = arrList.Item(23).ToString()
        oReadValuesFromExcel.ClevisCapPinClips = arrList.Item(24).ToString()
        oReadValuesFromExcel.ThreadProtected = arrList.Item(25).ToString()

        oReadValuesFromExcel.RodClevisCheck = arrList.Item(26).ToString()

        oReadValuesFromExcel.RodEndThreadSize = Convert.ToDouble(arrList.Item(27).ToString())
        oReadValuesFromExcel.RodClevisPinClips = arrList.Item(28).ToString()
        oReadValuesFromExcel.PistonStealPackage = arrList.Item(29).ToString()
        oReadValuesFromExcel.Paint = arrList.Item(30).ToString()
        oReadValuesFromExcel.GenerationType = arrList.Item(31).ToString()
        oReadValuesFromExcel.RetractedLength = arrList.Item(32).ToString()
        oReadValuesFromExcel.ExtendedLength = arrList.Item(33).ToString()
        oReadValuesFromExcel.PinSizeDetails = arrList.Item(34).ToString()
        oReadValuesFromExcel.RodClevis = arrList.Item(35).ToString()
        oReadValuesFromExcel.RodWiper = arrList.Item(36).ToString()
        oReadValuesFromExcel.StopTubeLength = arrList.Item(37).ToString()
        oReadValuesFromExcel.NutSize = arrList.Item(38).ToString()
        oReadValuesFromExcel.RodSealPackage = arrList.Item(39).ToString()
        oReadValuesFromExcel.NOLABELONCYLINDER = arrList.Item(40).ToString()
        oReadValuesFromExcel.BagRequired = arrList.Item(41).ToString()
        

        'oReadValuesFromExcel.PinsAreInLineWithPort = Convert.ToInt32(arrList.Item(32).ToString())
        'oReadValuesFromExcel.RetractedLengthTieRod3 = Convert.ToInt32(arrList.Item(33).ToString())
        'oReadValuesFromExcel.ExtendedLengthTieRod3 = Convert.ToInt32(arrList.Item(34).ToString())
        'oReadValuesFromExcel.RodDiameterTieRod3 = Convert.ToInt32(arrList.Item(35).ToString())
        'oReadValuesFromExcel.Ports = Convert.ToInt32(arrList.Item(36).ToString())
        'oReadValuesFromExcel.PercentAirTest = Convert.ToInt32(arrList.Item(37).ToString())
        'oReadValuesFromExcel.PercentOilTest = Convert.ToInt32(arrList.Item(38).ToString())
        'oReadValuesFromExcel.RephaseOnExtension = Convert.ToInt32(arrList.Item(39).ToString())
        'oReadValuesFromExcel.RephaseOnRetraction = Convert.ToInt32(arrList.Item(40).ToString())
        'oReadValuesFromExcel.InstallStrokeControl = Convert.ToInt32(arrList.Item(41).ToString())
        'oReadValuesFromExcel.StampCustomerPartAndDateCodeOnTube = Convert.ToInt32(arrList.Item(42).ToString())
        'oReadValuesFromExcel.StampCustomerPartOnTube = Convert.ToInt32(arrList.Item(43).ToString())
        'oReadValuesFromExcel.StampCountryOfOriginOnTube = Convert.ToInt32(arrList.Item(44).ToString())
        'oReadValuesFromExcel.RodMaterialsNitroSteel = Convert.ToInt32(arrList.Item(45).ToString())
        'oReadValuesFromExcel.InstallSteelPlugsInAllPorts = Convert.ToInt32(arrList.Item(46).ToString())
        'oReadValuesFromExcel.InstallHardenedBushingsAndRodClevisEnd = Convert.ToInt32(arrList.Item(47).ToString())
        'oReadValuesFromExcel.InstallHardenedBushingsAndClevisCapEnd = Convert.ToInt32(arrList.Item(48).ToString())
        'oReadValuesFromExcel.AssemblyStopTubeToCylinder = Convert.ToInt32(arrList.Item(49).ToString())
        'oReadValuesFromExcel.MaskPerBOMAndSOP = Convert.ToInt32(arrList.Item(50).ToString())
        'oReadValuesFromExcel.MaskBushingsBeforePainting = Convert.ToInt32(arrList.Item(51).ToString())
        'oReadValuesFromExcel.MaskExposedThreadsAfterWashing = Convert.ToInt32(arrList.Item(52).ToString())
        'oReadValuesFromExcel.MaskPinHoles = Convert.ToInt32(arrList.Item(53).ToString())
        'oReadValuesFromExcel.Prime = Convert.ToInt32(arrList.Item(54).ToString())
        'oReadValuesFromExcel.PaintTieRod3 = Convert.ToInt32(arrList.Item(55).ToString())
        'oReadValuesFromExcel.AffixLabelPerSOP = Convert.ToInt32(arrList.Item(56).ToString())
        'oReadValuesFromExcel.IncludePinKitPerBOM = Convert.ToInt32(arrList.Item(57).ToString())
        'oReadValuesFromExcel.PackCylinderInPlasticBag = Convert.ToInt32(arrList.Item(58).ToString())
        'oReadValuesFromExcel.AffixLabelToBag = Convert.ToInt32(arrList.Item(59).ToString())
        'oReadValuesFromExcel.PackagePerSOP = Convert.ToInt32(arrList.Item(60).ToString())

        Module1.ReadValuesFromExcel = oReadValuesFromExcel

    End Sub

    Public Sub SettingSelectedArrayListValues(ByVal rowno As Integer)

        If rowno = 4 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList2)
            End If
        ElseIf rowno = 5 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList3)
            End If

        ElseIf rowno = 6 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList4)
            End If

        ElseIf rowno = 7 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList5)
            End If

        ElseIf rowno = 8 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList6)
            End If

        ElseIf rowno = 9 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList7)
            End If

        ElseIf rowno = 10 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList8)
            End If

        ElseIf rowno = 11 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList9)
            End If

        ElseIf rowno = 12 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList10)
            End If

        ElseIf rowno = 13 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList11)
            End If

        ElseIf rowno = 14 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList12)
            End If

        ElseIf rowno = 15 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList13)
            End If

        ElseIf rowno = 16 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList14)
            End If

        ElseIf rowno = 17 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList15)
            End If

        ElseIf rowno = 18 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList16)
            End If

        ElseIf rowno = 19 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList17)
            End If

        ElseIf rowno = 20 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList18)
            End If

        ElseIf rowno = 21 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList19)
            End If

        ElseIf rowno = 22 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList20)
            End If

        ElseIf rowno = 23 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList21)
            End If

        ElseIf rowno = 24 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList22)
            End If

        ElseIf rowno = 25 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList23)
            End If

        ElseIf rowno = 26 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList24)
            End If

        ElseIf rowno = 27 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList25)
            End If

        ElseIf rowno = 28 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList26)
            End If

        ElseIf rowno = 29 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList27)
            End If

        ElseIf rowno = 30 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList28)
            End If

        ElseIf rowno = 31 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList29)
            End If

        ElseIf rowno = 32 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList30)
            End If

        ElseIf rowno = 33 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList31)
            End If

        ElseIf rowno = 34 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList32)
            End If

        ElseIf rowno = 35 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList33)
            End If

        ElseIf rowno = 36 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList34)
            End If

        ElseIf rowno = 37 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList35)
            End If

        ElseIf rowno = 38 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList36)
            End If

        ElseIf rowno = 39 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList37)
            End If

        ElseIf rowno = 40 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList38)
            End If

        ElseIf rowno = 41 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList39)
            End If

        ElseIf rowno = 42 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList40)
            End If

        ElseIf rowno = 43 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList41)
            End If

        ElseIf rowno = 44 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList42)
            End If

        ElseIf rowno = 45 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList43)
            End If

        ElseIf rowno = 46 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList44)
            End If

        ElseIf rowno = 47 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList45)
            End If

        ElseIf rowno = 48 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList46)
            End If

        ElseIf rowno = 49 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList47)
            End If

        ElseIf rowno = 50 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList48)
            End If

        ElseIf rowno = 51 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList49)
            End If

        ElseIf rowno = 52 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList50)
            End If

        ElseIf rowno = 53 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList51)
            End If

        ElseIf rowno = 54 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList52)
            End If

        ElseIf rowno = 55 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList53)
            End If

        ElseIf rowno = 56 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList54)
            End If

        ElseIf rowno = 57 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList55)
            End If

        ElseIf rowno = 58 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList56)
            End If

        ElseIf rowno = 59 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList57)
            End If

        ElseIf rowno = 60 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList58)
            End If

        ElseIf rowno = 61 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList59)
            End If

        ElseIf rowno = 62 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList60)
            End If

        ElseIf rowno = 63 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList61)
            End If

        ElseIf rowno = 64 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList62)
            End If

        ElseIf rowno = 65 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList63)
            End If

        ElseIf rowno = 66 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList64)
            End If

        ElseIf rowno = 67 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList65)
            End If

        ElseIf rowno = 68 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList66)
            End If

        ElseIf rowno = 69 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList67)
            End If

        ElseIf rowno = 70 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList68)
            End If

        ElseIf rowno = 71 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList69)
            End If

        ElseIf rowno = 72 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList70)
            End If

        ElseIf rowno = 73 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList71)
            End If

        ElseIf rowno = 74 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList72)
            End If

        ElseIf rowno = 75 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList73)
            End If

        ElseIf rowno = 76 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList74)
            End If

        ElseIf rowno = 77 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList75)
            End If

        ElseIf rowno = 78 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList76)
            End If

        ElseIf rowno = 79 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList77)
            End If

        ElseIf rowno = 80 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList78)
            End If

        ElseIf rowno = 81 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList79)
            End If

        ElseIf rowno = 82 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList80)
            End If

        ElseIf rowno = 83 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList81)
            End If

        ElseIf rowno = 84 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList82)
            End If

        ElseIf rowno = 85 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList83)
            End If

        ElseIf rowno = 86 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList84)
            End If

        ElseIf rowno = 87 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList85)
            End If

        ElseIf rowno = 88 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList86)
            End If

        ElseIf rowno = 89 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList87)
            End If

        ElseIf rowno = 90 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList88)
            End If

        ElseIf rowno = 91 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList89)
            End If

        ElseIf rowno = 92 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList90)
            End If

        ElseIf rowno = 93 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList91)
            End If

        ElseIf rowno = 94 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList92)
            End If

        ElseIf rowno = 95 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList93)
            End If

        ElseIf rowno = 96 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList94)
            End If

        ElseIf rowno = 97 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList95)
            End If

        ElseIf rowno = 98 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList96)
            End If

        ElseIf rowno = 99 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList97)
            End If

        ElseIf rowno = 100 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList98)
            End If

        ElseIf rowno = 101 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList99)
            End If

        ElseIf rowno = 102 Then
            If Not Module1.ArraList2.Count = 0 Then
                SettingValuesFromExcel(Module1.ArraList100)
            End If

        End If

    End Sub

    Public Sub HideControls()     'sugandhi

        lblBackBrowse.Visible = False
        GroupBox3.Visible = False
        GroupBox2.Visible = False
        LabelGradient8.Visible = False

    End Sub

    Private Sub cmbCustomerName_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
                                                Handles cmbCustomerName.KeyUp      'sugandhi

        Try
            If (e.KeyValue >= 65 AndAlso e.KeyValue <= 90) OrElse (e.KeyValue >= 48 AndAlso e.KeyValue <= 57) _
            OrElse (e.KeyValue >= 186 AndAlso e.KeyValue <= 192) OrElse (e.KeyValue >= 219 AndAlso e.KeyValue <= 222) _
            OrElse e.KeyValue = 106 OrElse e.KeyValue = 109 OrElse e.KeyValue = 110 OrElse e.KeyValue = 111 OrElse _
                                                                        e.KeyCode = Keys.Space Then


                If e.Shift Then
                    If (e.KeyValue >= 48 AndAlso e.KeyValue <= 57) OrElse e.KeyValue = 190 Then
                        _strSearchedCustomer += HTReturnDKeyCharacters.Item(e.KeyValue)(0).ToString
                    End If
                Else
                    If (e.KeyValue >= 48 AndAlso e.KeyValue <= 57) OrElse e.KeyValue = 190 Then
                        _strSearchedCustomer += HTReturnDKeyCharacters.Item(e.KeyValue)(1).ToString
                    ElseIf e.KeyCode = Keys.Space Then
                        _strSearchedCustomer += " "
                    Else
                        _strSearchedCustomer += e.KeyData.ToString
                    End If
                End If

                Dim strQuery As String = "select CustomerName from CustomerManageDetails where CustomerName like '" _
                                                            & _strSearchedCustomer & "%' order by CustomerName Asc"
                Dim CustomerNameDataTable As DataTable = oDataClass.GetDataTable(strQuery)
                If Not IsNothing(CustomerNameDataTable) AndAlso CustomerNameDataTable.Rows.Count > 0 Then
                    ' cmbCustomerName.Items.Clear()
                    For Each oCustomerNameDataRow As DataRow In CustomerNameDataTable.Rows
                        If Not IsDBNull(oCustomerNameDataRow(0)) Then
                            cmbCustomerName.Text = oCustomerNameDataRow(0).ToString
                            Exit For
                        End If
                    Next
                Else
                    _strSearchedCustomer = String.Empty
                    _strSearchedCustomer = e.KeyData.ToString
                End If
            End If

        Catch ex As Exception
        End Try

    End Sub

    Private Sub cmbCustomerName_DropDownClosed(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                            Handles cmbCustomerName.DropDownClosed

        Try
            _strSearchedCustomer = String.Empty
        Catch ex As Exception
        End Try

    End Sub

    Private ReadOnly Property HTReturnDKeyCharacters() As Hashtable        'sugandhi
        Get
            Dim htReturnDKeyChar As New Hashtable
            htReturnDKeyChar.Add(48, New Object(1) {")", 0})
            htReturnDKeyChar.Add(49, New Object(1) {"!", 1})
            htReturnDKeyChar.Add(50, New Object(1) {"@", 2})
            htReturnDKeyChar.Add(51, New Object(1) {"#", 3})
            htReturnDKeyChar.Add(52, New Object(1) {"$", 4})
            htReturnDKeyChar.Add(53, New Object(1) {"%", 5})
            htReturnDKeyChar.Add(54, New Object(1) {"^", 6})
            htReturnDKeyChar.Add(55, New Object(1) {"&", 7})
            htReturnDKeyChar.Add(56, New Object(1) {"*", 8})
            htReturnDKeyChar.Add(57, New Object(1) {"(", 9})
            htReturnDKeyChar.Add(190, New Object(1) {">", "."})
            Return htReturnDKeyChar
        End Get
    End Property

    Public Sub ClearErrorLogInfo()     ' sugandhi
        lVLogInformation.Items.Clear()
    End Sub

End Class