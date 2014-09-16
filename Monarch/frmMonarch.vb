Imports System.Drawing.Drawing2D
Imports IFLBaseDataLayer
Imports IFLCustomUILayer
Imports IFLCommonLayer
Imports MonarchFunctionalLayer
Imports IFLBusinessLayer
Imports LabelGradient
Imports System.IO
Imports System.Drawing.Imaging
Imports ExcelModule


''' <summary>
''' This class is the primary class of the project.
''' </summary>
''' <remarks>
''' This class starts the application, this class checks for LICENSE and if the LICENSE KEY is valid then it starts 
''' the application else it prompts a error message.
''' </remarks>
Public Class frmMonarch

#Region "Class Level Variables"
    Private oExcel As ExcelUtil
    Private m_Middle As Single = 0
    Private m_Delta As Single = 0.1
    Private _status As Integer
    Dim _oIFLCustomUILayer As IFLCustomUILayer.IFLListView
    Dim _IsLoadExecuted As Boolean = False
    Dim oIFLWindowsRegistryClass As New IFLCommonLayer.IFLWindowsRegistryClass
    Private _IsItemSelected As Boolean = False
    Private _sProjectNumber As String = ""
    Private strCustomerName As String
    Private CodeNumberTable As DataTable
    Private list As New ArrayList
    Private _btnVisible As Boolean = False

    Private _browseFileName As String
    Private oReadValuesFromExcel As New ReadValuesFromExcel

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

    Public ReadOnly Property GetSelectedRow() As String
        Get
            Return _sProjectNumber
        End Get
    End Property

    Public ReadOnly Property IsItemSelected() As Boolean
        Get
            Return _IsItemSelected
        End Get
    End Property
#End Region

#Region "SubProcedures"
    ''' <summary>
    ''' This function overrides the close function of system object.
    ''' </summary>
    ''' <param name="m">
    ''' parameter is "m" of type message.
    ''' </param>
    ''' <remarks>
    ''' This function overrides the close functionality of the system object and disables the close option on the 
    ''' control box.
    ''' </remarks>
    Protected Overrides Sub WndProc(ByRef m As Message)

        Dim SC_Close As Integer = &HF060
        Dim WM_SysCommand As Integer = &H112
        Select Case m.Msg
            Case &H112 'WM_SYSCOMMAND
                ' The WM_ACTIVATEAPP message occurs when the application
                ' becomes the active application or becomes inactive.
                Select Case m.WParam.ToInt32
                    Case &HF060 'SC_Close 'User clicked on "X"
                        'Do something if you want then exit sub without 
                        'passing the Message on to MyBase...this will stop the form from firing the 
                        'close event
                        Exit Sub
                        'End If
                End Select
        End Select
        MyBase.WndProc(m)

    End Sub

    Private Function getaRegistryKeys() As ArrayList

        Dim aRegistryKeys As New ArrayList()
        aRegistryKeys.Add("SOFTWARE")
        aRegistryKeys.Add("BBL")
        Return aRegistryKeys

    End Function

    Private Function getaRegistryColumns() As ArrayList

        Dim aRegistryColumns As New ArrayList()
        aRegistryColumns.Add("Server")
        aRegistryColumns.Add("Database")
        aRegistryColumns.Add("UserID")
        aRegistryColumns.Add("Password")
        aRegistryColumns.Add("IntegratedSecurity")
        Return aRegistryColumns
    End Function

#End Region

   

#Region "Events"
    ''' <summary>
    ''' This event occurs when user clicks button "NewValidationProject".
    ''' </summary>
    ''' <param name="sender">
    ''' parameter is "sender" of type System.Object
    ''' </param>
    ''' <param name="e">
    ''' parametern is "e" of type System.EventArgs
    ''' </param>
    ''' <remarks>
    ''' This event calls a subprocedure "NewValidationFunctionality" to perform some task.
    ''' </remarks>
    Private Sub btnNewValidationProject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '  NewValidationFunctionality()
    End Sub

    ''' <summary>
    ''' This event occures when button "ExitApplication" is clicked.
    ''' </summary>
    ''' <param name="sender">
    ''' parameter is "sender" of type System.EventArgs
    ''' </param>
    ''' <param name="e">
    ''' Parameter is "e" of type System.EventArgs
    ''' </param>
    ''' <remarks>
    ''' This subprocedure stops the application.
    ''' </remarks>
    Private Sub btnExitApplication_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub frmDesignValidationAutomation_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        If _IsLoadExecuted Then
            ' ExecuteLoadFunction()

            '28_04_2011  RAGAVA
            Try
                'Me.AutoScrollMargin = New Drawing.Size(400, 250)
                'Me.AutoScrollPosition = New Point(160, 115)
                Me.AutoScrollMargin = ofrmMdiMonarch.pnlChildFormArea.AutoScrollMargin
                Me.AutoScrollMargin = New Drawing.Size(250, 180)
                Me.AutoScrollPosition = New Point(150, 95)
            Catch ex As Exception
            End Try
            'Till  Here
        End If

    End Sub

    ''' <summary>
    ''' This event is called when form is loaded first.
    ''' </summary>
    ''' <param name="sender">
    ''' parameter is "sender" of type Sysytem.Object
    ''' </param>
    ''' <param name="e">
    ''' parameter is "e" of type System.EventArgs
    ''' </param>
    ''' <remarks>
    ''' This event loads images into the picture boxes when the form is loaded first and checks the license validity.
    ''' </remarks>
    Private Sub frmDesignValidationAutomation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                                        Handles MyBase.Load
        ColorTheForm()
        Dim Execution_Path1 As String = Application.StartupPath
        Dim Execution_Path As String = "X:\Master_Library"
        _IsLoadExecuted = True
        SetListViewColumnProperties("CustomerMaster", LVCustomer)

        '28_04_2011  RAGAVA
        Try
            'Me.AutoScrollMargin = New Drawing.Size(400, 250)
            'Me.AutoScrollPosition = New Point(160, 115)
            Me.AutoScrollMargin = New Drawing.Size(250, 180)
            ''Me.AutoScrollOffset = New Point(-150, -95)
            Me.AutoScrollPosition = New Point(150, 110)
        Catch ex As Exception
        End Try

    End Sub
#End Region

    Private Function GetTableRelatedForm(ByVal strFormName As String) As Form

        GetTableRelatedForm = Nothing
        For Each oForm As Form In GetProjectFormSequenceInstances
            If strFormName.ToUpper.Equals(oForm.Name.ToUpper) Then
                GetTableRelatedForm = oForm
            End If
        Next

    End Function
#Region "UnUsed"
    Private Sub ExecuteLoadFunction()
        Try
            If Not IsNothing(IFLConnectionObject) Then
                Me.lvwContractDetails.SourceTable = GetDesignValidationTableData("ContractMaster")
                checkTableProperties(lvwContractDetails, "ContractMaster")
                Me.LVCustomer.SourceTable = GetDesignValidationTableData("CustomerMaster")
                checkTableProperties(LVCustomer, "CustomerMaster")
            End If
        Catch oException As Exception
            MessageBox.Show(oException.Message + vbCrLf + vbCrLf + oException.StackTrace.ToString)
        End Try
    End Sub

    Private Sub checkTableProperties(ByVal lvTable As IFLListView, ByVal TableName As String)
        If Not lvTable.SourceTable Is Nothing Then
            If lvTable.Columns.Count = 0 Then
                SetListViewColumnProperties(TableName, lvwContractDetails)
            End If
            lvTable.Items.Clear()
            If lvTable.Columns.Count = 0 Then
                lvTable.Populate()
            Else
                lvTable.PopulateData()
            End If
        End If
    End Sub

    Private Function GetDesignValidationTableData(ByVal TableName As String) As DataTable
        Dim strQuery As String = Nothing
        Select Case TableName
            Case "ContractMaster"
                strQuery = "select ContractNumber as ProjectNumber,Description,ContractRevision as Revision,AssemblyType,CustomerPartCode as Customer_Part_Code " + vbCrLf
                strQuery += "from ContractMaster " + vbCrLf
            Case "CustomerMaster"
                strQuery = "select distinct CustomerName " + vbCrLf      '14_10_2009  ragava  Distinct Added
                strQuery += "from CustomerMaster " + vbCrLf
        End Select
        GetDesignValidationTableData = IFLConnectionObject.GetTable(strQuery)
    End Function
#End Region

    Private Sub SetListViewColumnProperties(ByVal TableName As String, ByVal lvName As IFLListView)

        lvName.Columns.Clear()
        Dim aColumns As New ArrayList
        Select Case TableName
            Case "ContractMaster"
                aColumns.Add(New Object(2) {"ProjectNumber", "Contract Number", True})      '04_11_2009   Ragava
                aColumns.Add(New Object(2) {"Description", "Description", True})
                aColumns.Add(New Object(2) {"Revision", "Contract Revision", True})      '04_11_2009   Ragava
                aColumns.Add(New Object(2) {"AssemblyType", "Assembly Type", True})
                aColumns.Add(New Object(2) {"Customer_Part_Code", "Customer PartCode", True})    '04_11_2009   Ragava
                aColumns.Add(New Object(2) {"IsCompleteModelGeneration", "IsCompleteModelGeneration", True})    '30-06-10-10am
            Case "CustomerMaster"
                aColumns.Add(New Object(2) {"CustomerName", "Customer Name", True})
        End Select
        lvName.DisplayHeaders = aColumns
        lvName.FullRowSelect = True
        lvName.IsTypeSearchEnable = True
        lvName.SearchObject = Me.txtListViewSearchObject
        lvName.IsFilterOptionEnabled = True
        lvName.SourceTable = GetWorkOrderMasterTableData(TableName)
        lvName.Populate()

    End Sub

    Public Function GetWorkOrderMasterTableData(ByVal TableName As String) As DataTable

        Dim strQuery As String = Nothing
        Select Case TableName
            Case "ContractMaster"
                strQuery = "select ContractNumber as ProjectNumber,Description,ContractRevision as Revision,AssemblyType,CustomerPartCode as Customer_Part_Code, IsCompleteModelGeneration " + vbCrLf
                strQuery += "from ContractMaster " + vbCrLf
            Case "CustomerMaster"
                strQuery = "select distinct CustomerName " + vbCrLf
                strQuery += "from CustomerMaster " + vbCrLf
        End Select
        GetWorkOrderMasterTableData = IFLConnectionObject.GetTable(strQuery)

    End Function

#Region "Property"

    Private ReadOnly Property GetProjectFormSequenceInstances() As ArrayList
        Get
            Dim aFormInstance As New ArrayList
            Return aFormInstance
        End Get
    End Property

#End Region

    Private Function DataFetchingForEditAndCopy() As Boolean

        DataFetchingForEditAndCopy = False
        _oIFLCustomUILayer = New IFLCustomUILayer.IFLListView
        Me._IsItemSelected = (Me.lvwContractDetails.SelectedItems.Count > 0)

    End Function

    Public Sub GetContract_CustomerDetails()

        Try
            If LVCustomer.GetCurrentIndex <> -1 Then
                ' Dim strCustomerName As String
                Dim oSelectedListviewItem As ListViewItem = Me.LVCustomer.SelectedItems(0)
                strCustomerName = oSelectedListviewItem.SubItems(0).Text
                'Dim strQuery As String = "select ContractNumber,Description,ContractRevision,AssemblyType,CustomerPartCode from ContractMaster Where CustomerMasterID in (Select IFL_ID from CustomerMaster where CustomerName = '" & strCustomerName & "') order by DateAndTime Desc"
                CustomerDetails()
                CustomerName = strCustomerName
            End If
            '04_11_2009  Ragava
            For Each listviewItem As ListViewItem In LVCustomer.Items
                Dim index As Integer = LVCustomer.Items.IndexOf(listviewItem)
                LVCustomer.Items(index).BackColor = Color.Ivory
                LVCustomer.Items(index).ForeColor = Color.Black
            Next
            For Each listviewItem As ListViewItem In LVCustomer.SelectedItems
                Dim index As Integer = LVCustomer.Items.IndexOf(listviewItem)
                LVCustomer.Items(index).BackColor = Color.CornflowerBlue
                LVCustomer.Items(index).ForeColor = Color.White
            Next
            '04_11_2009  Ragava   Till  Here
        Catch ex As Exception
        End Try

    End Sub

    Public Sub CustomerDetails()

        Dim strQuery As String = "select ContractNumber as ProjectNumber,Description,ContractRevision as Revision,AssemblyType,CustomerPartCode as Customer_Part_Code, "
        strQuery += " case when IsCompleteModelGeneration = 1 then 'Complete Model Generated' else 'Only Costing Report Generated' end as Comment "
        strQuery += "from ContractMaster Where CustomerMasterID in (Select IFL_ID from CustomerMaster where CustomerName = '" & strCustomerName & "') order by DateAndTime Desc"      '04_11_2009  Ragava
        Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
        objDT = IsReleasedOrNotValidation(objDT, strCustomerName)
        Dim codeNumberQuery As String = "select CustomerPartCode as Customer_Part_Code from ContractMaster Where CustomerMasterID in (Select IFL_ID from CustomerMaster where CustomerName = '" & strCustomerName & "') order by DateAndTime Desc"
        CodeNumberTable = oDataClass.GetDataTable(codeNumberQuery)
        list.Clear()
        For i As Integer = 0 To CodeNumberTable.Rows.Count - 1
            list.Add(CodeNumberTable.Rows(i).Item(0).ToString())
        Next

        'ANUP 28-10-2010 TILL HERE

        '' select case when IsCompleteModelGeneration=1 then 'ModelGenerated' else 'CostGenerated' end as comments from ContractMaster
        'objDT.Columns.Add("Comment")
        'For Each oRow As DataRow In objDT.Rows
        '    If oRow("IsCompleteModelGeneration") Then
        '        oRow("Comment") = "Complete Model Generated"
        '    Else
        '        oRow("Comment") = "Only Costing Report Generated"
        '    End If
        'Next
        'objDT.Columns.Remove("IsCompleteModelGeneration")
        lvwContractDetails.FlushListViewData()
        lvwContractDetails.SourceTable = objDT
        lvwContractDetails.Populate()

    End Sub

    Private Sub LVCustomer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                Handles LVCustomer.SelectedIndexChanged

        lvwContractDetails.Clear()
        If LVCustomer.SelectedItems.Count > 0 Then
            oCustomerListviewItem = Me.LVCustomer.SelectedItems(0)
            GetContract_CustomerDetails()
        Else
            oCustomerListviewItem = Nothing
            oContractListviewItem = Nothing
        End If

    End Sub

    Private Sub lvwContractDetails_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                    Handles lvwContractDetails.SelectedIndexChanged
        Try

            If lvwContractDetails.SelectedItems.Count > 0 Then
                oContractListviewItem = Me.lvwContractDetails.SelectedItems(0)
                Try
                    PartCode1 = Me.lvwContractDetails.SelectedItems(0).SubItems.Item(4).Text       '16_08_2011   RAGAVA
                Catch ex As Exception

                End Try
                '12_07_2011  RAGAVA
                If IsNew_Revision_Released = "Released" Then
                    btnRelease.Visible = True
                Else
                    btnRelease.Visible = False
                End If
            Else
                btnRelease.Visible = False
                'Till   Here
            End If
            '04_11_2009  Ragava
            For Each listviewItem As ListViewItem In lvwContractDetails.Items
                Dim index As Integer = lvwContractDetails.Items.IndexOf(listviewItem)
                lvwContractDetails.Items(index).BackColor = Color.Ivory
                lvwContractDetails.Items(index).ForeColor = Color.Black
            Next
            For Each listviewItem As ListViewItem In lvwContractDetails.SelectedItems
                Dim index As Integer = lvwContractDetails.Items.IndexOf(listviewItem)
                lvwContractDetails.Items(index).BackColor = Color.CornflowerBlue
                lvwContractDetails.Items(index).ForeColor = Color.White
            Next
            '04_11_2009  Ragava   Till  Here

        Catch ex As Exception
        End Try

    End Sub

    Private Sub ColorTheForm()

        FunctionalClassObject.LabelGradient_GreenBorder_ColoringTheScreens(LabelGradient6, LabelGradient2, _
                                                                        LabelGradient3, LabelGradient7)
        FunctionalClassObject.LabelGradient_OrangeBorder_ColoringTheScreens(LabelGradient1)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient5)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient4)

    End Sub

    'ANUP 28-10-2010 START
    Private Function IsReleasedOrNotValidation(ByVal dtFinalDataTable As DataTable, ByVal strCustomerName As String) As DataTable

        IsReleasedOrNotValidation = Nothing
        Try

            Dim strQuery As String = "select ContractNumber as ProjectNumber,Description,ContractRevision as Revision,AssemblyType,"
            strQuery += " CustomerPartCode as Customer_Part_Code, case when IsCompleteModelGeneration = 1 then "
            strQuery += " 'Complete Model Generated' else 'Only Costing Report Generated' end as Comment from "
            strQuery += " ContractMaster CM,ReleasedCylinderCodes RCC Where CM.ContractNumber = RCC.ReleasedCylinderCodeNumber and CustomerMasterID in (Select IFL_ID from CustomerMaster where CustomerName = '" & strCustomerName & "') order by DateAndTime Desc"
            Dim oTable As DataTable = IFLConnectionObject.GetTable(strQuery)
            If IsNew_Revision_Released = "Revision" Then
                dtFinalDataTable.Columns.Add("IsItReleased")
            End If

            If Not IsNothing(oTable) Then
                For Each oDataRow As DataRow In oTable.Rows
                    For Each oFinalDataRow As DataRow In dtFinalDataTable.Rows
                        If oFinalDataRow("ProjectNumber") = oDataRow("ProjectNumber") Then
                            If IsNew_Revision_Released = "Revision" Then
                                oFinalDataRow("IsItReleased") = "Released"
                            ElseIf IsNew_Revision_Released = "Released" Then
                                oFinalDataRow.Delete()
                                dtFinalDataTable.AcceptChanges()
                            End If
                            Exit For
                        End If
                    Next
                Next
            End If
            IsReleasedOrNotValidation = dtFinalDataTable
        Catch ex As Exception
        End Try

    End Function
    'ANUP 28-10-2010 TILL HERE

    '12_07_2011   RAGAVA
    Private Sub btnRelease_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRelease.Click

        Try
            Dim oClsReleaseCylinderFunctionality As New clsReleaseCylinderFunctionality
            If IsNew_Revision_Released = "Released" Then
                EditProjectFunctionality()
                oClsReleaseCylinderFunctionality.RevisionCounterValidation(ContractNumber)
            End If
            If lvwContractDetails.SelectedItems.Count > 0 Then
                oContractListviewItem = Me.lvwContractDetails.SelectedItems(0)
                oDataClass.UpdateRevision_Details(oContractListviewItem.Text, intContractRevisionNumber + 1, "Release")
            End If
            'C:\MONARCH_TESTING\CMS_TEMP\
            If Not oClsReleaseCylinderFunctionality.MainFunctionality() Then
                MessageBox.Show("Release Cylinder Validation Failed", "ERROR", MessageBoxButtons.OK, _
                                                            MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
            End If
            If Directory.Exists("C:\MONARCH_TESTING\CMS_TEMP\" & ContractNumber & "_CMS") Then
                My.Computer.FileSystem.MoveDirectory("C:\MONARCH_TESTING\CMS_TEMP\" & ContractNumber & _
                                                        "_CMS", "W:\TIEROD\CMS\" & ContractNumber & "_CMS")
                My.Computer.FileSystem.MoveFile("C:\MONARCH_TESTING\CNC_TEMP\0" & strRodCodeNumber & _
                                                        "1.MIN", "W:\TIEROD\CNC\0" & strRodCodeNumber & "1.MIN")
                'Directory.Move("C:\MONARCH_TESTING\CMS_TEMP\" & ContractNumber & "_CMS", "W:\TIEROD\CMS\" & ContractNumber & "_CMS")
            End If
            oClsReleaseCylinderFunctionality.DropRod_Tube_Stoptube_TieRodCodesToDB(strRodCodeNumber, _
                                strBoreCodeNumber, StopTubeCodeNumber, strTieRodCodeNumber, oContractListviewItem.Text)
            btnRelease.Enabled = False
            MsgBox("Cylinder Released Successfully")
        Catch ex As Exception
        End Try

    End Sub

  
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

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        BrowseFileName = GetFileLocation()
        If BrowseFileName Is Nothing Then
            Return
        End If
        LVCustomer.Items.Clear()
        lvwContractDetails.Items.Clear()

        oExcel = OpenSheet(BrowseFileName)
        Dim i As Integer = 3
        If oExcel Is Nothing Then
            Throw New Exception("File not found exception.")
        Else
            ' SettingValuesFromExcel(i)
        End If

        strCustomerName = ReadValuesFromExcel.CustomerName
        Dim listView As ListViewItem
        listView = LVCustomer.Items.Add(strCustomerName)
        LVCustomer.Items(0).Selected = True

        Dim boolValuePresent As Boolean = False
        'CustomerDetails()

        For j As Integer = 0 To list.Count - 1
            If list.Item(j).ToString() = ReadValuesFromExcel.CustomerPortCode.ToString() Then
                lvwContractDetails.Items(j).Selected = True
                boolValuePresent = True
                Exit For
            End If
        Next
        If Not boolValuePresent Then
            MessageBox.Show("Please Enter valid Code Number")
            Return
        End If
        _btnVisible = True

        mdiMonarch.GetExcelFile()

        oExcel.Close()

    End Sub

End Class
