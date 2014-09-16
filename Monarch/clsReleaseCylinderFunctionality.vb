Imports Microsoft.Win32.Registry
Imports Microsoft.Win32.RegistryKey
Imports System.Diagnostics.Process
Imports Microsoft.Office.Interop
Imports Microsoft.Win32
Imports System.IO
Imports MonarchFunctionalLayer

Public Class clsReleaseCylinderFunctionality

#Region "Variables"

    Private _strCurrentWorkingDirectory As String = System.Environment.CurrentDirectory + "\ECR"
    Private _strSourceExcelPath As String = _strCurrentWorkingDirectory + "\ECR_Codes.xls"

    'anup 04-02-2011 start
    '    Private _strDestinationExcelPath As String = "W:\ECR_TieRod\ECR_Codes.xls"
    Private _strDestinationExcelPath As String = "W:\ECR\ECR_Codes.xls"
    'anup 04-032-2011 till here

    Private _oExApplication As Excel.Application
    Private _oExWorkbook As Excel.Workbook
    Private _oExcelSheet_MainAssembly As Excel.Worksheet

    Private _strCylinderCodeNumber As String = CylinderCodeNumber

    Private _blnIsNewCylinder As Boolean
    Private _blnIsNewTube As Boolean
    Private _blnIsNewRod As Boolean
    Private _blnIsNewTierod As Boolean
    Private _blnIsNewStopTube As Boolean
    Private _strECRNumber As String = String.Empty
    Private _strNewExcelPath As String = String.Empty

    Private _htCodeNumbers As New Hashtable
    Private _strTableDrawingNumber As String = String.Empty
    Private _IsDrawingRevisedOrCreated As Boolean

#End Region

    Public Function MainFunctionality() As Boolean
        MainFunctionality = False
        Try
            If ReleaseExcelFunctionality() Then
                If CreateExcelBasedOnNewPartsGenerated() Then
                    If DropReleasedCodeNumbersToDB() Then
                        'anup 23-12-2010 start
                        '    If IsNew_Revision_Released = "Released" Then
                        'RevisionCounterValidation()
                        '  End If
                        'anup 23-12-2010 till here
                        MainFunctionality = True
                    End If
                End If
            End If
        Catch ex As Exception
            MainFunctionality = False
        End Try
    End Function

#Region "ECR Code Generation Excel Functionality"

    Private Function ReleaseExcelFunctionality() As Boolean
        ReleaseExcelFunctionality = False
        Try
            If CreateExcelObjects() Then
                If DropDataToExistingSheet() Then
                    If SaveExcel() Then
                        ReleaseExcelFunctionality = True
                    End If
                End If
            End If
        Catch ex As Exception
            ReleaseExcelFunctionality = False
        End Try
    End Function

    Private Function CreateExcelObjects() As Boolean
        CreateExcelObjects = False
        Try
            If CheckForExcel() Then
                If DoesExcelExists() Then
                    If CreateExcel() Then
                        CreateExcelObjects = True
                    End If
                End If
            End If
        Catch ex As Exception
            CreateExcelObjects = False
        End Try
    End Function

    Private Function CheckForExcel() As Boolean
        CheckForExcel = True
        Dim strSubKey As String = "Excel.Application"
        Dim oKey As RegistryKey = Registry.ClassesRoot
        Dim oSubKey As RegistryKey = oKey.OpenSubKey("Word.Application")
        If Not IsNothing(oSubKey) Then
            oKey.Close()
            Return True
        Else
            MessageBox.Show("Error with Excel" + vbCrLf + "Kindly check whether the Excel is installed" + vbCrLf + _
             "You can proceed with application but, Excel report will not be generated", "Error with Excel", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2)
            Return False
        End If
    End Function

    Private Function DoesExcelExists() As Boolean
        DoesExcelExists = False

        Try
            'anup 04-02-2011 start
            'If Not Directory.Exists("W:\ECR_TieRod\") Then
            '    Directory.CreateDirectory("W:\ECR_TieRod\")
            'End If

            If Not Directory.Exists("W:\ECR\") Then
                Directory.CreateDirectory("W:\ECR\")
            End If

            'anup 04-02-2011 till here
            If Not File.Exists(_strDestinationExcelPath) Then
                File.Copy(_strSourceExcelPath, _strDestinationExcelPath, True)
            End If

            If Not IsNothing(_oExApplication) Then
                _oExApplication = Nothing
            End If
            DoesExcelExists = True

        Catch ex As Exception
            DoesExcelExists = False
        End Try

    End Function

    Private Function CreateExcel() As Boolean
        CreateExcel = True
        Try
            _oExApplication = New Excel.Application
            _oExApplication.Visible = False
            _oExWorkbook = _oExApplication.Workbooks.Open(_strDestinationExcelPath)
            _oExcelSheet_MainAssembly = _oExApplication.Sheets(1)

        Catch ex As Exception
            CreateExcel = False
            MessageBox.Show("Unable to open Excel sheet", "Information", MessageBoxButtons.OK, _
            MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Function

    Private Function SaveExcel() As Boolean
        SaveExcel = False
        Try
            _oExWorkbook.Save()
            For Each oProcess As Process In Process.GetProcessesByName("Excel")
                oProcess.Kill()
                GC.Collect()
                System.Threading.Thread.Sleep(1000)
            Next

            SaveExcel = True
        Catch ex As Exception
            SaveExcel = False
        End Try
    End Function

    Private Function DropDataToExistingSheet() As Boolean

        DropDataToExistingSheet = False
        Dim intTotalCostExcelRange As Integer = 2

        Do
            If IsNothing(_oExcelSheet_MainAssembly.Range("A" + intTotalCostExcelRange.ToString).Value) Then
                intTotalCostExcelRange = intTotalCostExcelRange
                Exit Do
            Else
                intTotalCostExcelRange += 1
            End If
        Loop

        Dim intrSNo As Integer = intTotalCostExcelRange - 1

        Try
            '19_07_2011   RAGAVA
            If _strCylinderCodeNumber = "" Then
                _strCylinderCodeNumber = ContractNumber
            End If
            'Till  Here


            Dim strDescription As String = "Release " & _strCylinderCodeNumber & " Cylinder"

            'anup 04-02-2011 start
            '_oExcelSheet_MainAssembly.Range("A" + intTotalCostExcelRange.ToString).Value = "10IFL"
            _oExcelSheet_MainAssembly.Range("A" + intTotalCostExcelRange.ToString).Value = "11IFL"
            'anup 04-02-2011 till here

            _oExcelSheet_MainAssembly.Range("C" + intTotalCostExcelRange.ToString).Value = intrSNo
            _oExcelSheet_MainAssembly.Range("D" + intTotalCostExcelRange.ToString).Value = strDescription
            _oExcelSheet_MainAssembly.Range("E" + intTotalCostExcelRange.ToString).Value = Date.Today
            _oExcelSheet_MainAssembly.Range("F" + intTotalCostExcelRange.ToString).Value = 1 'IT MAY CHANGE IN FUTURE

            'anup 04-02-2011 start
            '_strECRNumber = "10IFL-" & intrSNo.ToString
            _strECRNumber = "11IFL-" & intrSNo.ToString
            'anup 04-02-2011 till here

            _oExWorkbook.Save()

            DropDataToExistingSheet = True

            '19_07_2011   RAGAVA
            Dim strquery As String = "Update RevisionTable Set ECR_Number = '" & _strECRNumber & "' where ContractNumber = '" & _strCylinderCodeNumber.ToString & "'"
            Dim blnSuccess As Boolean = IFLConnectionObject.ExecuteQuery(strquery)
        Catch ex As Exception
            DropDataToExistingSheet = False
        End Try
    End Function

#End Region

#Region " Excel For Each ECR Code Generated Functionality"

    Private Function CreateExcelBasedOnNewPartsGenerated() As Boolean
        CreateExcelBasedOnNewPartsGenerated = False
        Dim drECR_Details As DataRow = Nothing
        Try
            '14_07_2011    RAGAVA
            If blnGenerateClicked = False Then
                Try
                    Dim _strQuery As String = String.Empty
                    _strQuery = "select * from StoreECR_PartsDetails_ReleaseOnClick where CylinderCodeNumber = '" + ContractNumber + "'"
                    drECR_Details = IFLConnectionObject.GetDataRow(_strQuery)
                    _strCylinderCodeNumber = drECR_Details("CylinderCodeNumber")
                    strBoreCodeNumber = drECR_Details("BoreCode")
                    strTieRodCodeNumber = drECR_Details("TieRodCode")
                    strRodCodeNumber = drECR_Details("RodCode")
                    StopTubeCodeNumber = drECR_Details("StopTubeCode")

                Catch ex As Exception
                End Try
            End If
            'Till  Here
            If CreateDirectoryForNewExcel() Then
                If _strCylinderCodeNumber <> "" Then
                    _htCodeNumbers.Add("CYLINDER CODE", _strCylinderCodeNumber)
                End If
                If strBoreCodeNumber <> "" Then
                    _htCodeNumbers.Add("TUBE", strBoreCodeNumber)
                End If
                If strTieRodCodeNumber <> "" Then
                    _htCodeNumbers.Add("TIEROD", strTieRodCodeNumber)
                End If
                If strRodCodeNumber <> "" Then
                    _htCodeNumbers.Add("ROD", strRodCodeNumber)
                End If
                If StopTubeCodeNumber <> "" Then
                    _htCodeNumbers.Add("STOPTUBE", StopTubeCodeNumber)
                End If

                'CheckForNewOrExisting()
                CheckForNewOrExisting(drECR_Details)       '14_07_2011  RAGAVA
                If DropDataToNewExcelSheet() Then
                    CreateExcelBasedOnNewPartsGenerated = True
                End If
            End If
        Catch ex As Exception
            CreateExcelBasedOnNewPartsGenerated = False
        End Try
    End Function

    Private Sub CheckForNewOrExisting(Optional ByVal drECR_Details As DataRow = Nothing)           '14_07_2011   RAGAVA  Optional parameter added
        If Not IsNothing(_strCylinderCodeNumber) Then
            'If _strCylinderCodeNumber.StartsWith("7") Then
            _blnIsNewCylinder = True
            'Else
            '_blnIsNewCylinder = False
            'End If
        End If

        If Not IsNothing(strBoreCodeNumber) Then
            If drECR_Details Is Nothing Then            '14_07_2011   RAGAVA
                If strCodeNumber_BeforeApplicationStart > strBoreCodeNumber Then           '21_01_2011    RAGAVA
                    _blnIsNewTube = False
                Else
                    _blnIsNewTube = True
                End If
            Else
                _blnIsNewTube = drECR_Details("IsNewBore")
            End If
        End If

        If Not IsNothing(strRodCodeNumber) Then
            If drECR_Details Is Nothing Then            '14_07_2011   RAGAVA
                If strCodeNumber_BeforeApplicationStart > strRodCodeNumber Then           '21_01_2011    RAGAVA
                    _blnIsNewRod = False
                Else
                    _blnIsNewRod = True
                End If
            Else
                _blnIsNewRod = drECR_Details("IsNewRod")
            End If
        End If

        If Not IsNothing(strTieRodCodeNumber) Then
            If drECR_Details Is Nothing Then            '14_07_2011   RAGAVA
                If strCodeNumber_BeforeApplicationStart > strTieRodCodeNumber Then           '21_01_2011    RAGAVA
                    _blnIsNewTierod = False
                Else
                    _blnIsNewTierod = True
                End If
            Else
                _blnIsNewTierod = drECR_Details("IsNewTieRod")
            End If
        End If

        If Not IsNothing(StopTubeCodeNumber) Then
            If drECR_Details Is Nothing Then            '14_07_2011   RAGAVA
                If strCodeNumber_BeforeApplicationStart > StopTubeCodeNumber Then           '21_01_2011    RAGAVA
                    _blnIsNewStopTube = False
                Else
                    _blnIsNewStopTube = True
                End If
            Else
                _blnIsNewStopTube = drECR_Details("IsNewStopTube")
            End If
        End If
    End Sub

    Private Function CreateDirectoryForNewExcel() As Boolean
        CreateDirectoryForNewExcel = False
        Try

            'anup 04-02-2011 start
            'If Not Directory.Exists("W:\ECR_TieRod\ECR_NewExcels\") Then
            '    Directory.CreateDirectory("W:\ECR_TieRod\ECR_NewExcels\")
            'End If
            If Not Directory.Exists("W:\ECR\ECR_NewExcels\") Then
                Directory.CreateDirectory("W:\ECR\ECR_NewExcels\")
            End If
            'anup 04-02-2011 till here

            Dim strNewReleasedExcel As String = _strCurrentWorkingDirectory + "\MasterNewPartsExcel.xls\"


            If File.Exists(_strCurrentWorkingDirectory & "\MasterNewPartsExcel.xls") Then

                'anup 04-02-2011 start
                'File.Copy(_strCurrentWorkingDirectory & "\MasterNewPartsExcel.xls", "W:\ECR_TieRod\ECR_NewExcels\" & _strECRNumber & ".xls", True)
                '_strNewExcelPath = "W:\ECR_TieRod\ECR_NewExcels\" & _strECRNumber & ".xls"
                File.Copy(_strCurrentWorkingDirectory & "\MasterNewPartsExcel.xls", "W:\ECR\ECR_NewExcels\" & _strECRNumber & ".xls", True)
                _strNewExcelPath = "W:\ECR\ECR_NewExcels\" & _strECRNumber & ".xls"
                'anup 04-02-2011 till here

                CreateDirectoryForNewExcel = True
            Else
                MessageBox.Show("MasterNewPartsExcel.xls dosn't exist", "Check the path for Excel", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
            End If

        Catch ex As Exception
            CreateDirectoryForNewExcel = False
        End Try
    End Function

    Private Function DropDataToNewExcelSheet() As Boolean
        DropDataToNewExcelSheet = False
        Try
            _oExApplication = New Excel.Application
            _oExApplication.Visible = False
            _oExWorkbook = _oExApplication.Workbooks.Open(_strNewExcelPath)
            _oExcelSheet_MainAssembly = _oExApplication.Sheets(1)

            Dim IsNewOrExisting As Boolean
            Dim intTotalCostExcelRange As Integer = 2
            Dim intItemCounter As Integer = 0

            Dim blnRouting As Boolean 'ANUP 26-11-2010
            Try

                For Each oItem As DictionaryEntry In _htCodeNumbers


                    If oItem.Key = "TUBE" Then
                        IsNewOrExisting = _blnIsNewTube
                        IsDrawingTableCreated(IsRowInserted_Tube, BoreDrawingNumber)
                        blnRouting = IsNewOrExisting 'ANUP 26-11-2010
                    ElseIf oItem.Key = "TIEROD" Then
                        IsNewOrExisting = _blnIsNewTierod
                        IsDrawingTableCreated(IsRowInserted_Tierod, TieRodDrawingNumber)
                        blnRouting = False 'ANUP 26-11-2010
                    ElseIf oItem.Key = "ROD" Then
                        IsNewOrExisting = _blnIsNewRod
                        IsDrawingTableCreated(IsRowInserted_Rod, RodDrawingNumber)
                        blnRouting = IsNewOrExisting 'ANUP 26-11-2010
                    ElseIf oItem.Key = "STOPTUBE" Then
                        IsNewOrExisting = _blnIsNewStopTube
                        IsDrawingTableCreated(IsRowInserted_StopTube, StopTubeDrawingNumber)
                        blnRouting = False 'ANUP 26-11-2010
                    ElseIf oItem.Key = "CYLINDER CODE" Then
                        IsNewOrExisting = _blnIsNewCylinder
                        _IsDrawingRevisedOrCreated = True
                        _strTableDrawingNumber = String.Empty
                        blnRouting = IsNewOrExisting 'ANUP 26-11-2010
                    End If

                    _oExcelSheet_MainAssembly.Range("A" + intTotalCostExcelRange.ToString).Value = oItem.Value
                    _oExcelSheet_MainAssembly.Range("B" + intTotalCostExcelRange.ToString).Value = IsNewOrExisting
                    _oExcelSheet_MainAssembly.Range("C" + intTotalCostExcelRange.ToString).Value = _IsDrawingRevisedOrCreated
                    _oExcelSheet_MainAssembly.Range("D" + intTotalCostExcelRange.ToString).Value = _strTableDrawingNumber
                    _oExcelSheet_MainAssembly.Range("E" + intTotalCostExcelRange.ToString).Value = blnRouting   'ANUP 26-11-2010

                    intTotalCostExcelRange += 1
                    intItemCounter += 1
                Next

                _oExWorkbook.Save()
                DropDataToNewExcelSheet = True
                SaveExcel()
            Catch ex As Exception
                DropDataToNewExcelSheet = False
            End Try
        Catch ex As Exception
            DropDataToNewExcelSheet = False
        End Try
    End Function

    Private Sub IsDrawingTableCreated(ByVal blnIsTableCreated As Boolean, ByVal TableDrawingNumber As String)
        Try
            If blnIsTableCreated Then
                _IsDrawingRevisedOrCreated = True
                _strTableDrawingNumber = TableDrawingNumber
            Else
                _IsDrawingRevisedOrCreated = False
                _strTableDrawingNumber = String.Empty
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Function DropReleasedCodeNumbersToDB() As Boolean
        DropReleasedCodeNumbersToDB = False
        Try
            'anup 21-03-2011 start
            If IsNew_Revision_Released = "Released" OrElse IsNew_Revision_Released = "Revision" Then
                Dim strQuery1 As String = String.Empty
                strQuery1 = "DELETE FROM dbo.ReleasedCylinderCodes WHERE ReleasedCylinderCodeNumber = '" & _strCylinderCodeNumber & "'"
                IFLConnectionObject.ExecuteQuery(strQuery1)
                'anup 21-03-2011 till here

                Dim strQuery As String = String.Empty
                strQuery = "INSERT INTO dbo.ReleasedCylinderCodes(ReleasedCylinderCodeNumber) VALUES(" & _strCylinderCodeNumber & ")"
                DropReleasedCodeNumbersToDB = IFLConnectionObject.ExecuteQuery(strQuery)
                If DropReleasedCodeNumbersToDB = False Then
                    MessageBox.Show("Error in updating Released Cylinder Code to Data Table", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            DropReleasedCodeNumbersToDB = False
        End Try
    End Function

    'anup 17-02-2011 start

    Public Function DropRod_Tube_Stoptube_TieRodCodesToDB(ByVal strRod As String, ByVal strTube As String, ByVal strStoptube As String, ByVal strTierod As String, ByVal strReleasedCylinderCode As String) As Boolean
        DropRod_Tube_Stoptube_TieRodCodesToDB = False
        Try
            Dim strQuery As String = String.Empty
            strQuery = "UPDATE dbo.ReleasedCylinderCodes   SET RodCode = '" & strRod & "',TubeCode ='" & strTube & "',StopTubeCode = '" & strStoptube & "',TieRodCode = '" & strTierod & "' WHERE ReleasedCylinderCodeNumber = '" & strReleasedCylinderCode & "'"
            DropRod_Tube_Stoptube_TieRodCodesToDB = IFLConnectionObject.ExecuteQuery(strQuery)
            If DropRod_Tube_Stoptube_TieRodCodesToDB = False Then
                MessageBox.Show("Error in updating Released Codes to Data Table", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
            End If
        Catch ex As Exception
            DropRod_Tube_Stoptube_TieRodCodesToDB = False
        End Try
    End Function

    Private ReadOnly Property ColumnName() As Hashtable
        Get
            Dim htColumnName As New Hashtable
            htColumnName.Add("ROD", "RodCode")
            htColumnName.Add("TUBE", "TubeCode")
            htColumnName.Add("STOPTUBE", "StopTubeCode")
            htColumnName.Add("TIEROD", "TieRodCode")
            Return htColumnName
        End Get
    End Property


    Public Function DoesCodeExistInDB(ByVal strCode As String, Optional ByVal strCodeName As String = "") As Boolean
        DoesCodeExistInDB = False
        Try
            If strCode.StartsWith("7") Then 'anup 16-03-2011
                Dim strColoumnName As String = String.Empty
                strColoumnName = ColumnName.Item(strCodeName)

                If Not String.IsNullOrEmpty(strColoumnName) Then
                    Dim strQuery As String = String.Empty
                    strQuery = "select * from ReleasedCylinderCodes where " & strColoumnName & " ='" & strCode & "'"
                    Dim dtDoesCodeExistInDB As DataTable = Nothing
                    dtDoesCodeExistInDB = IFLConnectionObject.GetTable(strQuery)
                    If IsNothing(dtDoesCodeExistInDB) OrElse dtDoesCodeExistInDB.Rows.Count < 1 Then
                        DoesCodeExistInDB = True
                    End If
                End If
            End If
        Catch ex As Exception
            DoesCodeExistInDB = False
        End Try
    End Function

   
    'anup 17-02-2011 till here

#End Region

#Region "Setting Revision Counter To Zero"

    Private Sub SettingRevisionCounterToZero(Optional ByVal CylinderCodeNumber As String = "")      '12_07_2011  RAGAVA  optional parameter added
        Try
            Dim strQuery As String = String.Empty
            strQuery = "UPDATE dbo.ContractMaster SET ContractRevision = 0 where ContractNumber = " & CylinderCodeNumber
            IFLConnectionObject.ExecuteQuery(strQuery)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DeletingExistingRevisionDetails(Optional ByVal CylinderCodeNumber As String = "")      '12_07_2011  RAGAVA  optional parameter added
        Try
            Dim strQuery As String = String.Empty
            strQuery = "DELETE FROM dbo.RevisionTable WHERE ContractNumber =" & CylinderCodeNumber
            IFLConnectionObject.ExecuteQuery(strQuery)
        Catch ex As Exception

        End Try
    End Sub

    Public Sub RevisionCounterValidation(Optional ByVal strContractNumber As String = "")      '14_07_2011  RAGAVA  optional parameter added
        Try
            DeletingExistingRevisionDetails(strContractNumber)
            SettingRevisionCounterToZero(strContractNumber)
        Catch ex As Exception

        End Try
    End Sub

#End Region

End Class

