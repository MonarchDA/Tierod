Imports MonarchFunctionalLayer
Imports MonarchAPILayer
Imports System.io


Public Class frmTieRod3

    Dim fso As New Scripting.FileSystemObject

#Region "Procedures"

    Public Sub GenerateModel()
        blnGenerateClicked = True    '14_07_2011   RAGAVA
        NumberButtonClick()
        Me.Cursor = Cursors.WaitCursor
        'If Not FunctionalClassObject.validateForm(Me) Is Nothing Then
        '    MessageBox.Show(FunctionalClassObject.ErrorMessage)
        '    Me.Cursor = Cursors.Arrow      '20_04_2010   RAGAVA
        '    FunctionalClassObject.validateForm(Me).Focus()
        'Else
        Try
            mdiMonarch.BtnsVisibleFalse()
            FunctionalClassObject.PopulateFormscontrolsData(Me)

            'anup 23-12-2010 start
            Try
                Dim oClsReleaseCylinderFunctionality As New clsReleaseCylinderFunctionality
                If IsNew_Revision_Released = "Released" Then
                    oClsReleaseCylinderFunctionality.RevisionCounterValidation()
                End If
            Catch ex As Exception
            End Try
            'anup 23-12-2010 till here


            If blnRevision = True Then
                Try
                    'anup 23-12-2010 start
                    If IsNew_Revision_Released = "Released" Then
                        oDataClass.UpdateRevision_Details(Trim(txtCylinderCodeNumber.Text), intContractRevisionNumber + 1, "Release")
                    Else
                        oDataClass.UpdateRevision_Details(Trim(txtCylinderCodeNumber.Text), intContractRevisionNumber + 1)
                    End If
                    'anup 23-12-2010 till here
                Catch ex As Exception

                End Try
                Dim ofrmRevision As New frmRevisionTable
                ofrmRevision.ShowDialog()
            End If

            If ApplicationStop = True Then
                Dim strMsg As String
                strMsg = "Invalid Code Numbers are encounter asper MonarchIndustries !" & vbNewLine
                strMsg = strMsg + "Please check the Code Numbers" & vbNewLine
                strMsg = strMsg + "Application will not proceed for the model generation"
                MessageBox.Show(strMsg, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ofrmMdiMonarch.btnGenerate.Visible = False
            Else
                ofrmMdiMonarch.btnGenerate.Visible = True
            End If

            Dim strMessage As String = ""
            'If IsCompleteModelGeneration Then   s
            '    strMessage = "Do you want to generate Complete Model"   s
            'Else   s
            '    strMessage = "Do you want to generate only Costing Report"    s
            'End If       s
            ' If MessageBox.Show(strMessage, "Confirmation", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.OK Then         s
            StopWatchAndProgressBar = "Start"
            'Try   s
            KillExcel()
            oExcelClass.objApp = Nothing
            KillAllSolidWorksServices()
            'Catch ex As Exception     s

            'End Try
            Execution_Path1 = Application.StartupPath
            Execution_Path = "X:\Master_Library"
            ' CylinderCodeNumber = txtCylinderCodeNumber.Text
            strCylinderDescription = txtCylinderDesc.Text
            Try
                If fso.FolderExists(Execution_Path1 & "\Reports") = False Then
                    fso.CreateFolder(Execution_Path1 & "\Reports")
                End If
                ReportFile = Execution_Path1 & "\Reports\" & CylinderCodeNumber & ".xls"
                updateDesignTables()
                'Rename(ReportFile & "\GUI_PARAMETERS_report.xls", ReportFile & "\" & CylinderCodeNumber & ".xls")

            Catch ex As Exception

            End Try

            Try
                SaveModelFolder()
            Catch ex As Exception

            End Try
            Try
                updateMainDesignTables()
            Catch ex As Exception

            End Try
            Try
                strPinKitId = Me.txtInstallPinandClips.Text.ToString      '23_09_2011   RAGAVA
                GenerateNotes()
                'Saveas_detached()        '31_08_2012   RAGAVA
            Catch ex As Exception

            End Try

            If IsCompleteModelGeneration Then
                oTieRodCylinder.PerformTieRodCylinderFunctionalities(CylinderCodeNumber)
            End If

            Try
                Try
                    StoreContractDetails()
                    InsertData_PartsDetails_ReleaseOnClick(CylinderCodeNumber)           '14_07_2011   RAGAVA
                Catch ex As Exception

                End Try
            Catch ex As Exception

            End Try
            'Else    s
            '    Me.Cursor = Cursors.Default   s
            '    Exit Sub   s
            'End If   s

        Catch oException As Exception
        End Try
        captureImages(Me)

        'ANUP 26-10-2010 START
        'Put Release Cylinder Data into Excel
        If IsReleaseCylinderChecked Then
            Dim oClsReleaseCylinderFunctionality As New clsReleaseCylinderFunctionality
            If Not oClsReleaseCylinderFunctionality.MainFunctionality() Then
                MessageBox.Show("Release Cylinder Validation Failed", "ERROR", MessageBoxButtons.OK, _
                                            MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
            End If
        End If
        'ANUP 26-10-2010 TILL HERE


        'Sandeep 04-03-10-4pm 
        AddCodeNumbersToCostingExcelRetrivedFromDB()

        'Sandeep 18-03-10-10am
        ObjClsCostingDetails.GetPaint_Package_LabelDetails()
        '**************

        '20_01_2011   RAGAVA
        Try
            If ht_CodeNumbers("TUBE") <> "" Then
                strBoreCodeNumber = ht_CodeNumbers("TUBE")
            End If
            If ht_CodeNumbers("TIEROD") <> "" Then
                strTieRodCodeNumber = ht_CodeNumbers("TIEROD")
            End If
            If ht_CodeNumbers("ROD") <> "" Then
                strRodCodeNumber = ht_CodeNumbers("ROD")
            End If
            If ht_CodeNumbers("STOPTUBE") <> "" Then
                StopTubeCodeNumber = ht_CodeNumbers("STOPTUBE")
            End If
            For Each oCodeNumber_AfterAscendingItem As DataRow In ObjClsCostingDetails.CodeNumber_BeforeAscendimg.Rows
                If oCodeNumber_AfterAscendingItem(1) = "Rod Code Number" Then
                    oCodeNumber_AfterAscendingItem(0) = strRodCodeNumber
                    If strCodeNumber_BeforeApplicationStart > strRodCodeNumber Then
                        oCodeNumber_AfterAscendingItem(2) = "Existing"
                    End If
                ElseIf oCodeNumber_AfterAscendingItem(1) = "Tube Code Number" Then
                    oCodeNumber_AfterAscendingItem(0) = strBoreCodeNumber
                    If strCodeNumber_BeforeApplicationStart > strBoreCodeNumber Then
                        oCodeNumber_AfterAscendingItem(2) = "Existing"
                    End If
                ElseIf oCodeNumber_AfterAscendingItem(1) = "Tie Rod Code Number" Then
                    oCodeNumber_AfterAscendingItem(0) = strTieRodCodeNumber
                    If strCodeNumber_BeforeApplicationStart > strTieRodCodeNumber Then
                        oCodeNumber_AfterAscendingItem(2) = "Existing"
                    End If
                ElseIf oCodeNumber_AfterAscendingItem(1) = "Stop Tube Code Number" Then
                    oCodeNumber_AfterAscendingItem(0) = StopTubeCodeNumber
                    If strCodeNumber_BeforeApplicationStart > StopTubeCodeNumber Then
                        oCodeNumber_AfterAscendingItem(2) = "Existing"
                    End If
                End If
            Next
        Catch ex As Exception
        End Try
        'Till   Here

        ObjClsCostingDetails.Costingfunctionality()
        '**************

        'ANUP 26-10-2010 START
        'If IsReleaseCylinderChecked Then
        Dim objClsCMSIntegration As New clsCMSIntegration
        objClsCMSIntegration.CMSIntegrationfunctionality()
        CNC_Code()
        'End If

        Me.Cursor = Cursors.Default
        StopWatchAndProgressBar = "Stop"

        If blnRevision Then
            Application.Exit()
        End If
        If IsGenerateBtnClicked Then
            Application.Exit()
        End If

    End Sub

    '14_07_2011   RAGAVA
    Public Function InsertData_PartsDetails_ReleaseOnClick(ByVal ContractNumber As String) As Boolean

        Try
            Dim strQuery As String = String.Empty
            '21_07_2011   RAGAVA
            If ht_CodeNumbers.ContainsKey("STOPTUBE") = True Then
                StopTubeCodeNumber = ht_CodeNumbers("STOPTUBE")
            End If
            'Till   Here
            strQuery = "Insert into StoreECR_PartsDetails_ReleaseOnClick values ('" & ContractNumber.ToString _
                & "','" & strBoreCodeNumber & "','" & strTieRodCodeNumber & "','" & strRodCodeNumber & "','" _
                & StopTubeCodeNumber & "'," & IsNewPart(strBoreCodeNumber) & "," & _
                IsNewPart(strTieRodCodeNumber) & "," & IsNewPart(strRodCodeNumber) & "," & IsNewPart(StopTubeCodeNumber) & ")"
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
        Catch ex As Exception
        End Try

    End Function
    '14_07_2011   RAGAVA
    Public Function IsNewPart(ByVal strCode As String) As Integer

        Try
            If strCodeNumber_BeforeApplicationStart > strCode Then
                IsNewPart = 0
            Else
                IsNewPart = 1
            End If
            Return IsNewPart
        Catch ex As Exception
        End Try

    End Function

    Private Sub CNC_Code()

        Dim oCyl As New clsCNCUtil
        If oCyl.DoCNCCodeGeneration() Then
            'MsgBox("CNC Code generated ")
        Else
            MessageBox.Show(oCyl.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Private Function StoreContractDetails() As Boolean

        Dim strContractNumber As String = txtCylinderCodeNumber.Text
        Dim strCustomerName As String = CustomerName
        Dim iRevisionNumber As Integer = 0
        Dim oSelectedListviewItem As ListViewItem
        Try
            If Not oContractListviewItem Is Nothing Then
                strCustomerName = Trim(oCustomerListviewItem.SubItems(0).Text)
                oSelectedListviewItem = oContractListviewItem
                If IsCompleteModelGeneration Then
                    iRevisionNumber = Val(oSelectedListviewItem.SubItems(2).Text) + 1
                Else
                    iRevisionNumber = Val(oSelectedListviewItem.SubItems(2).Text)
                End If
            End If
            Dim objDT As DataTable
            Dim strQuery As String
            If iRevisionNumber = 0 Then
                strQuery = "Insert into CustomerMaster(IFL_ID,CustomerName) values (" & _
                                            Val(strContractNumber) & ",'" & strCustomerName & "')"
                objDT = oDataClass.GetDataTable(strQuery)
                strQuery = ""
                objDT.Clear()
            Else
                Try
                    strQuery = "Delete from ContractMaster where ContractNumber = '" & strContractNumber _
                                                & "' and ContractRevision <= " & iRevisionNumber.ToString
                    Dim objdt1 As DataTable = oDataClass.GetDataTable(strQuery)
                    strQuery = ""
                Catch ex As Exception
                End Try
            End If

            'anup 23-12-2010 start
            If IsNew_Revision_Released = "Released" Then
                iRevisionNumber = 0
            End If
            'anup 23-12-2010 till here


            Dim aData As Byte() = GetDataToSave(ofrmMdiMonarch.btnGenerate)
            Dim oRow As DataRow
            Dim oDataBaseClass As New IFLBaseDataLayer.IFLBaseDataClass(IFLConnectionObject)
            Dim IsNewRecord As Boolean = False
            oDataBaseClass.LoadTable("ContractMaster", , "IFL_ID", "IFLID")

            oRow = oDataBaseClass.GetNewRecord("ContractMaster")
            oRow("ContractNumber") = Trim(txtCylinderCodeNumber.Text)
            oRow("ContractRevision") = iRevisionNumber.ToString
            oRow("Description") = Trim(txtCylinderDesc.Text)
            oRow("AssemblyType") = Trim(AssemblyType)
            oRow("Project_XML") = aData
            oRow("CustomerPartCode") = Trim(PartCode1)
            oRow("CustomerMasterID") = strContractNumber
            oRow("IsCompleteModelGeneration") = IsCompleteModelGeneration

            oDataBaseClass.AddNewRecord("ContractMaster", oRow)
            StoreContractDetails = oDataBaseClass.SaveData

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    Private Sub btnReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        ShowSaveDialog()

    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If MessageBox.Show("Are you sure to close the application?", "Confirm", MessageBoxButtons.YesNo, _
            MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
            Me.Close()
            Application.Exit()
        End If

    End Sub

    Private Sub chkAssemblyNotes_CheckedChanged(ByVal sender As System.Object, ByVal e As _
                        System.EventArgs) Handles chkAssemblyNotes.CheckedChanged

        If sender.Checked = True Then
            RichTextBox1.Enabled = True
        Else
            RichTextBox1.Enabled = False
        End If

    End Sub

    Private Sub chkPaintingNote_CheckedChanged(ByVal sender As System.Object, ByVal e As _
                    System.EventArgs) Handles chkPaintingNote.CheckedChanged

        If sender.Checked = True Then
            RichTextBox2.Enabled = True
        Else
            RichTextBox2.Enabled = False
        End If

    End Sub

    Public Sub ActivatedCodeTieRod3()
        '16_06_2011   RAGAVA
        Try
            Validate_PinandClipsNotes()
        Catch ex As Exception
        End Try
        'Till   Here
        '10_05_2011   RAGAVA 
        Try
            'ASSEMBLY NOTES
            If (SeriesForCosting.ToString).StartsWith("TX") = True OrElse WorkingPressure >= 2500 Then
                'chk100AirTest.Checked = False  VAMSI 10-09-2014
                'chk100AirTest.Enabled = False
                'txtAirTest.Enabled = False

                chk100OilTest.Checked = False
                chk100OilTest.Enabled = False
                txtOilTest.Enabled = False

                ChkRephaseExtension.Checked = False
                ChkRephaseExtension.Enabled = False
                txtRephaseOnExtension.Enabled = False

                ChkRephaseRetraction.Checked = False
                ChkRephaseRetraction.Enabled = False
                txtRephaseOnRetraction.Enabled = False
            Else
                'chk100AirTest.Enabled = True VAMSI 10-09-2014
                'txtAirTest.Enabled = True
                chk100OilTest.Enabled = True
                txtOilTest.Enabled = True
                ChkRephaseExtension.Enabled = True
                txtRephaseOnExtension.Enabled = True
                ChkRephaseRetraction.Enabled = True
                txtRephaseOnRetraction.Enabled = True
            End If
            If (SeriesForCosting.ToString).StartsWith("TX") = True Then
                chkInstallStrokeControl.Checked = False
                chkInstallStrokeControl.Enabled = False
                txtInstallStrokeLength.Enabled = False

                ChkStampCountryOfOrigin.Checked = False
                ChkStampCountryOfOrigin.Enabled = False
                txtStampCountry.Enabled = False

                ChkInstallSteelPlugs.Checked = False
                ChkInstallSteelPlugs.Enabled = False
                txtInstallSteelPlugs.Enabled = False

                ChkHardenedBushingsRodClevisEnd.Checked = False
                ChkHardenedBushingsRodClevisEnd.Enabled = False
                txtInstallHardenedBushingsRodClevis.Enabled = False

                ChkHardenedBushingsClevisCapEnd.Checked = False
                ChkHardenedBushingsClevisCapEnd.Enabled = False
                txtInstallHardenedBushingsClevisCap.Enabled = False

                ChkAssemblyStopTube.Checked = False
                ChkAssemblyStopTube.Enabled = False
                txtAssemblyStopTube.Enabled = False
            Else
                chkInstallStrokeControl.Enabled = True
                txtInstallStrokeLength.Enabled = True
                ChkStampCountryOfOrigin.Enabled = True
                txtStampCountry.Visible = True
                ChkInstallSteelPlugs.Enabled = True
                txtInstallSteelPlugs.Enabled = True
                ChkHardenedBushingsRodClevisEnd.Enabled = True
                txtInstallHardenedBushingsRodClevis.Enabled = True
                ChkHardenedBushingsClevisCapEnd.Enabled = True
                txtInstallHardenedBushingsClevisCap.Enabled = True
                ChkAssemblyStopTube.Enabled = True
                txtAssemblyStopTube.Enabled = True
            End If

            'PAINT NOTES
            If (SeriesForCosting.ToString).StartsWith("TX") Then
                chkMaskPerBOM.Checked = False
                chkMaskPerBOM.Enabled = False
                txtMaskPerBOM.Enabled = False

            Else
                chkMaskPerBOM.Enabled = True
                txtMaskPerBOM.Enabled = True

            End If

            'If ofrmTieRod2.rdbRodClevisYes.Checked = True OrElse Trim(ofrmTieRod1.cmbRodClevisPinHole.Text).IndexOf("Bushing") _
            '                <> -1 OrElse Trim(ofrmTieRod1.cmbClevisCapPinHole.Text).IndexOf("Bushing") <> -1 Then  'vamsi 16-09-2014
            If Trim(ofrmTieRod1.cmbRodClevisPinHole.Text).IndexOf("Bushing") _
                            <> -1 OrElse Trim(ofrmTieRod1.cmbClevisCapPinHole.Text).IndexOf("Bushing") <> -1 Then
                chkMaskPerBOM.Enabled = False
                chkMaskPerBOM.Checked = True
            Else
                'chkMaskPerBOM.Enabled = False  'vamsi 12-09-14
                'chkMaskPerBOM.Checked = False

                chkMaskPerBOM.Enabled = True  'vamsi 12-09-14
                chkMaskPerBOM.Checked = False
                txtMaskPerBOM.Clear()
            End If
            'TILL   HERE
        Catch ex As Exception

        End Try

        txtColumnLoad.Text = ColumnLoad
        txtWorkingPressure.Text = WorkingPressure
        txtCylinderDesc.Text = ""
        txtCylinderDesc.Text = SetCodeDesciption1()

        '10_02_2010   RAGAVA   moved to 1st Screen
        If Trim(txtCylinderCodeNumber.Text) = "" Or Trim(txtCylinderCodeNumber.Text) = _
                                (Convert.ToInt32(CylinderCodeNumber) - 1).ToString() Then
            txtCylinderCodeNumber.Text = CylinderCodeNumber()
        End If

        PartCode = Trim(txtCylinderCodeNumber.Text)
        PartCode1 = Trim(ofrmContractDetails.txtlPartCode.Text)
        CylinderCodeNumber = txtCylinderCodeNumber.Text
        LoadInformation()
        If chkAssemblyNotes.Checked = True Then
            RichTextBox1.Enabled = True
        Else
            RichTextBox1.Enabled = False
        End If
        If chkPaintingNote.Checked = True Then
            RichTextBox2.Enabled = True
        Else
            RichTextBox2.Enabled = False
        End If
        'If Trim(ofrmTieRod1.cmbPortOrientation.Text) = Trim(ofrmTieRod1.cmbPortOrientationForRodCap.Text) Then  VAMSI 10-09-2014
        '    If Trim(ofrmTieRod1.cmbPortOrientation.Text).IndexOf("90") <> -1 Then
        '        ChkPins.Text = "Pins Are 90 Degrees To Ports"
        '    Else
        '        ChkPins.Text = "Pins Are In Line With Ports"
        '    End If
        'Else
        '    ChkPins.Text = "Rod Cap Port " & Trim(ofrmTieRod1.cmbPortOrientationForRodCap.Text) _
        '                        & ", Clevis Cap Port " & Trim(ofrmTieRod1.cmbPortOrientation.Text)
        'End If

        If Trim(ofrmTieRod1.cmbPortOrientation.Text) = Trim(ofrmTieRod1.cmbPortOrientationForRodCap.Text) Then
            If Trim(ofrmTieRod1.cmbPortOrientation.Text).IndexOf("90") <> -1 Then
                ChkPins.Text = "Pins Are 90 Degrees To Ports"
            Else
                ChkPins.Text = "Pins Are In Line With Ports"
            End If
        Else
            ChkPins.Text = "PINS ARE AS SHOWN"
        End If



        ChkPins.Checked = True
        ChkPins.Enabled = False
        If Trim(ofrmTieRod1.cmbClevisCapPort.Text) = Trim(ofrmTieRod1.cmbRodCapPort.Text) Then
            ChkPorts.Text = "Ports      ____________"
        Else
            ChkPorts.Text = "Rod Cap Port ____________" & ", Clevis Cap Port ____________"
        End If
        ChkPorts.Checked = True
        ChkPorts.Enabled = False
        ChkRetractedLength.Checked = True
        ChkRetractedLength.Enabled = False
        ChkExtendedLength.Checked = True
        ChkExtendedLength.Enabled = False
        ChkRodDiameter.Checked = True
        ChkRodDiameter.Enabled = False

        '16_08_2012   RAGAVA
        'chk100AirTest.Enabled = False VAMSI 10-09-2014
        'chk100AirTest.Checked = True
        'Till   Here

        If Trim(SeriesForCosting).IndexOf("TP") <> -1 Then
            'chk100AirTest.Enabled = False  VAMSI 10-09-2014
            'chk100AirTest.Checked = True
            chk100OilTest.Checked = True
            chk100OilTest.Enabled = False
        Else
            If blnRevision = False Then                   '04_11_2009    Ragava
                ' txtAirTest.Clear() VAMSI 10-09-2014
                txtOilTest.Clear()
            End If
        End If
        If Trim(ofrmTieRod1.cmbRephasingPortPosition.Text).IndexOf("At Extension") <> -1 Or _
                                    Trim(ofrmTieRod1.cmbRephasingPortPosition.Text).IndexOf("At Both") <> -1 Then
            ChkRephaseExtension.Checked = True
            ChkRephaseExtension.Enabled = False
            txtRephaseOnExtension.Clear()
            txtRephaseOnExtension.Enabled = True
        ElseIf Trim(ofrmTieRod1.cmbRephasingPortPosition.Text).IndexOf("At Retraction") <> -1 Or _
                                Trim(ofrmTieRod1.cmbRephasingPortPosition.Text).IndexOf("At Both") <> -1 Then
            ChkRephaseRetraction.Checked = True
            ChkRephaseRetraction.Enabled = False
            txtRephaseOnRetraction.Clear()
            txtRephaseOnRetraction.Enabled = True
        Else
            ChkRephaseExtension.Checked = False
            ChkRephaseExtension.Enabled = False
            txtRephaseOnExtension.Clear()
            txtRephaseOnExtension.Enabled = False

            ChkRephaseRetraction.Checked = False
            ChkRephaseRetraction.Enabled = False
            txtRephaseOnRetraction.Clear()
            txtRephaseOnRetraction.Enabled = False
        End If
        If ofrmTieRod1.optStrokeControlYes.Checked = True Then
            chkInstallStrokeControl.Checked = True       '23_11_2009   Ragava
            chkInstallStrokeControl.Enabled = False
        Else
            chkInstallStrokeControl.Checked = False
            chkInstallStrokeControl.Enabled = False
            txtInstallStrokeLength.Enabled = False
            txtInstallStrokeLength.Clear()
        End If
        If Trim(ofrmTieRod1.cmbRodMaterial.Text).IndexOf("Nitro") <> -1 Then
            ChkRodMaterial.Checked = True      '23_11_2009   Ragava
            ChkRodMaterial.Enabled = False
        Else
            ChkRodMaterial.Checked = False
            ChkRodMaterial.Enabled = False
            txtRodMaterialNitroSteel.Clear()
            txtRodMaterialNitroSteel.Enabled = False
        End If
        If Trim(ofrmTieRod2.cmbThreadProtected.Text).IndexOf("All") <> -1 Then
            ChkInstallSteelPlugs.Enabled = False
            ChkInstallSteelPlugs.Checked = True       '23_11_2009   Ragava
        Else
            ChkInstallSteelPlugs.Enabled = False
            ChkInstallSteelPlugs.Checked = False
            txtInstallSteelPlugs.Clear()
            txtInstallSteelPlugs.Enabled = False
        End If
        If ofrmTieRod1.rdbStopTubeYes.Checked = True Then
            ChkAssemblyStopTube.Enabled = False
            ChkAssemblyStopTube.Checked = True       '23_11_2009   Ragava
        Else
            ChkAssemblyStopTube.Checked = False
            ChkAssemblyStopTube.Enabled = False
            txtAssemblyStopTube.Clear()
            txtAssemblyStopTube.Enabled = False
        End If

        'Painting Notes
        If ofrmTieRod2.optPinsYes.Checked = True Then
            chkInstallPinAndClips.Checked = True
        End If
        If ofrmTieRod2.optPinsNo.Checked = True Then
            chkInstallPinAndClips.Checked = False
        End If
        If blnRevision = False Then             '04_11_2009   Ragava
            ' chkAffixLabel.Checked = True       'Sugandhi_20120601
        End If
        If Trim(ofrmTieRod1.cmbClevisCapPinHole.Text).IndexOf("Bushing") <> -1 Then
            ChkMaskBushings.Enabled = False
            ChkMaskBushings.Checked = True      '23_11_2009   Ragava
            ChkHardenedBushingsClevisCapEnd.Checked = True       '25_11_2009   Ragava
            ChkHardenedBushingsClevisCapEnd.Enabled = False      '25_11_2009   Ragava
        Else
            ChkMaskBushings.Checked = False
            ChkMaskBushings.Enabled = False
            txtMaskBushings.Clear()
            txtMaskBushings.Enabled = False
            ChkHardenedBushingsClevisCapEnd.Checked = False
            ChkHardenedBushingsClevisCapEnd.Enabled = False
            txtInstallHardenedBushingsClevisCap.Clear()
            txtInstallHardenedBushingsClevisCap.Enabled = False
        End If
        If Trim(ofrmTieRod1.cmbRodClevisPinHole.Text).IndexOf("Bushing") <> -1 Then
            ChkHardenedBushingsRodClevisEnd.Checked = True
            ChkHardenedBushingsRodClevisEnd.Enabled = False
        Else
            ChkHardenedBushingsRodClevisEnd.Checked = False
            ChkHardenedBushingsRodClevisEnd.Enabled = False
            txtInstallHardenedBushingsRodClevis.Clear()
            txtInstallHardenedBushingsRodClevis.Enabled = False
        End If
        If ofrmTieRod2.rdbRodClevisNo.Checked = True Then
            ChkMaskExposedThreads.Checked = True     '23_11_2009   Ragava
            ChkMaskExposedThreads.Enabled = False
        Else
            ChkMaskExposedThreads.Enabled = False
            ChkMaskExposedThreads.Checked = False
            txtMaskExposedThreads.Clear()
            txtMaskExposedThreads.Enabled = False
        End If

        If Trim(ofrmTieRod2.cmbPaint.Text).IndexOf("Prime") <> -1 Then
            ChkPrime.Enabled = False
            ChkPrime.Checked = True     '23_11_2009   Ragava
        Else
            ChkPrime.Enabled = False
            ChkPrime.Checked = False
            txtPrime.Clear()
            txtPrime.Enabled = False
        End If
        If blnRevision = False Then             '04_11_2009   Ragava
            ChkPaint.Checked = True

        End If
        Try
            If Trim(ofrmTieRod2.cmbPaint.Text) <> "" Then    '01_12_2009    Ragava
                'Dim strQuery As String = "Select Color,PrimerCode,PaintCodeNumber from TieRodPaintDetails where Color = '" _
                '                                    & Trim(ofrmTieRod2.cmbPaint.Text) & "'" 'vamsi 10-09-2014
                Dim strQuery As String = "Select paintColor,PrimerCode,PaintCode from PaintDetails where paintColor = '" _
                                                   & Trim(ofrmTieRod2.cmbPaint.Text) & "'"


                Dim objdt As DataTable = oDataClass.GetDataTable(strQuery)
                If Not objdt.Rows(0).Item("PrimerCode") Is Nothing And Trim(objdt.Rows(0).Item("PrimerCode")) <> "" Then
                    ChkPrime.Checked = True
                    txtPrime.Enabled = True
                Else
                    ChkPrime.Checked = False
                    txtPrime.Clear()
                    txtPrime.Enabled = False
                End If
                ChkPrime.Enabled = False
                If Not objdt.Rows(0).Item("PaintCode") Is Nothing And Trim(objdt.Rows(0).Item("PaintCode")) <> "" Then
                    ChkPaint.Checked = True
                    txtPaint.Enabled = True
                Else
                    ChkPaint.Checked = False
                    txtPaint.Clear()
                    txtPaint.Enabled = False
                End If
                ChkPaint.Enabled = False
            End If
        Catch ex As Exception
        End Try
        If ofrmTieRod2.optPinsYes.Checked = True Then
            chkInstallPinAndClips.Enabled = True
        Else
            chkInstallPinAndClips.Checked = False
            chkInstallPinAndClips.Enabled = False
            txtInstallPinandClips.Clear()
            txtInstallPinandClips.Enabled = False
        End If


        '23_09_2011  RAGAVA  Commented as per Jinka Logic
        ''06_06_2011   RAGAVA
        'If blnInstallPinsandClips = False Then
        '    chkInstallPinAndClips.Checked = False
        '    chkInstallPinAndClips.Enabled = False
        'Else
        '    chkInstallPinAndClips.Enabled = True
        'End If
        ''TILL   HERE

    End Sub

#End Region

#Region "Functions"
    Private Function GetCylinderCodeNumber() As String

        GetCylinderCodeNumber = ""
        Try
            If oContractListviewItem Is Nothing Then
                Try
                    Dim iContractNumber As Long = 800001
                    Dim strQuery As String
                    strQuery = "select top 1 ContractNumber from ContractNumberDetails order by ContractNumber Desc"
                    Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                    If objDT.Rows.Count > 0 Then
                        iContractNumber = Val(objDT.Rows(0).Item("ContractNumber").ToString)
                        iContractNumber = iContractNumber + 1
                    End If
                    Try
                        strQuery = "Insert into ContractNumberDetails Values ('" & iContractNumber.ToString & "')"
                        objDT = oDataClass.GetDataTable(strQuery)
                    Catch ex As Exception
                    End Try
                    GetCylinderCodeNumber = iContractNumber.ToString
                Catch ex As Exception
                End Try
            Else
                GetCylinderCodeNumber = oContractListviewItem.SubItems(0).Text
            End If
        Catch ex As Exception
        End Try
        Return GetCylinderCodeNumber

    End Function

    Public Function SetCodeDesciption1() As String

        SetCodeDesciption1 = ""
        '05_07_2011   RAGAVA
        Dim strDesc1 As String = String.Empty
        If CodeDesc = "LN" Then
            strDesc1 = "THC"
        Else
            strDesc1 = CodeDesc
        End If
        'TILL   HERE
        Dim s As String = StrokeLength
        Dim words As String() = s.Split(New Char() {"."c})

        If RodDiameter.ToString = "1.12" Then

            If words.Length = 1 Then
                SetCodeDesciption1 = Math.Floor((Val(ofrmTieRod1.cmbBore.Text) * 10)) & strDesc1 & _
                                    Format(StrokeLength, "00").ToString & "-" & (1.13 * 100)
            Else
                SetCodeDesciption1 = Math.Floor((Val(ofrmTieRod1.cmbBore.Text) * 10)) & strDesc1 & _
                        Format(StrokeLength, "00.00").ToString & "-" & (1.13 * 100)    '22_02_2010    RAGAVA
            End If
        Else
            If words.Length = 1 Then
                SetCodeDesciption1 = Math.Floor((Val(ofrmTieRod1.cmbBore.Text) * 10)) & strDesc1 & _
                    Format(StrokeLength, "00").ToString & "-" & (Math.Round(RodDiameter, 2) * 100)    '22_02_2010    RAGAVA
            Else
                SetCodeDesciption1 = Math.Floor((Val(ofrmTieRod1.cmbBore.Text) * 10)) & strDesc1 & _
                        Format(StrokeLength, "00.00").ToString & "-" & (Math.Round(RodDiameter, 2) * 100)
            End If
        End If
        SetCodeDesciption = SetCodeDesciption1

    End Function
#End Region

    '19_10_2009   ragava
    Private Sub btnNumberNotes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                        Handles btnNumberNotes.Click
        NumberButtonClick()

    End Sub

    Public Sub NumberButtonClick()           'SUGANDHI
        Try
            Dim iCount As Integer = 1
            If ChkPins.Checked = True Then
                txtPins.Text = iCount.ToString
                iCount += 1
            Else
                txtPins.Clear()
            End If
            txtRetractedLength.Text = iCount.ToString
            iCount += 1
            txtExtenedLength.Text = iCount.ToString
            iCount += 1
            txtRodDiameter.Text = iCount.ToString
            iCount += 1
            If ChkPorts.Checked = True Then
                txtPorts.Text = iCount.ToString
                iCount += 1
            Else
                txtPorts.Clear()
            End If
            'vamsi 12-09-14
            'If chk100AirTest.Checked = True Then  
            '    txtAirTest.Text = iCount.ToString
            '    iCount += 1
            'Else
            '    txtAirTest.Clear()
            'End If
            'vamsi 12-09-14
            If chk100OilTest.Checked = True Then
                txtOilTest.Text = iCount.ToString
                iCount += 1
            Else
                txtOilTest.Clear()
            End If
            If ChkRephaseExtension.Checked = True Then
                txtRephaseOnExtension.Text = iCount.ToString
                iCount += 1
            Else
                txtRephaseOnExtension.Clear()
            End If
            If ChkRephaseRetraction.Checked = True Then
                txtRephaseOnRetraction.Text = iCount.ToString
                iCount += 1
            Else
                txtRephaseOnRetraction.Clear()
            End If
            If chkInstallStrokeControl.Checked = True Then
                txtInstallStrokeLength.Text = iCount.ToString
                iCount += 1
            Else
                txtInstallStrokeLength.Clear()
            End If
            If chkStampCustomerPartandDate.Checked = True Then
                txtStampCustomerPartandDate.Text = iCount.ToString
                iCount += 1
            Else
                txtStampCustomerPartandDate.Clear()
            End If
            If chkStampCustomerPartOnTube.Checked = True Then
                txtStampCustomerPart.Text = iCount.ToString
                iCount += 1
            Else
                txtStampCustomerPart.Clear()
            End If
            If ChkStampCountryOfOrigin.Checked = True Then
                txtStampCountry.Text = iCount.ToString
                iCount += 1
            Else
                txtStampCountry.Clear()
            End If
            If ChkRodMaterial.Checked = True Then
                txtRodMaterialNitroSteel.Text = iCount.ToString
                iCount += 1
            Else
                txtRodMaterialNitroSteel.Clear()
            End If
            If ChkInstallSteelPlugs.Checked = True Then
                txtInstallSteelPlugs.Text = iCount.ToString
                iCount += 1
            Else
                txtInstallSteelPlugs.Clear()
            End If
            'If ChkInstallHardenedBushings.Checked = True Then
            '    txtInstallHardenedBushings.Text = iCount.ToString
            '    iCount += 1
            'Else
            '    txtInstallHardenedBushings.Clear()
            'End If
            If ChkHardenedBushingsRodClevisEnd.Checked = True Then
                txtInstallHardenedBushingsRodClevis.Text = iCount.ToString
                iCount += 1
            Else
                txtInstallHardenedBushingsRodClevis.Clear()
            End If
            If ChkHardenedBushingsClevisCapEnd.Checked = True Then
                txtInstallHardenedBushingsClevisCap.Text = iCount.ToString
                iCount += 1
            Else
                txtInstallHardenedBushingsClevisCap.Clear()
            End If
            '02_11_2009   Ragava
            'If chkAffixLabelToBag.Checked = True Then
            '    txtAffixLabeltoBag.Text = iCount.ToString
            '    iCount += 1
            'Else
            '    txtAffixLabeltoBag.Clear()
            'End If
            '02_11_2009   Ragava  Till  Here
            If ChkAssemblyStopTube.Checked = True Then
                txtAssemblyStopTube.Text = iCount.ToString
                iCount += 1
            Else
                txtAssemblyStopTube.Clear()
            End If
            '07_09_2012   RAGAVA
            If ChkFluidFilmInternal.Checked = True Then
                txtFluidFilmInternal.Text = iCount.ToString
                iCount += 1
            Else
                txtFluidFilmInternal.Clear()
            End If

            '20_10_2009   ragava
            If chkAssemblyNotes.Checked = True Then
                Dim strRichTextBox1Text As String = String.Empty
                For Each str As String In RichTextBox1.Lines
                    If Trim(strRichTextBox1Text) <> "" Then
                        strRichTextBox1Text = strRichTextBox1Text & vbNewLine
                    End If
                    If str.IndexOf("}") <> -1 Then
                        If str.IndexOf("}") > 0 Then                   '01_12_2009  Ragava
                            str = str.Replace(str.Substring(0, str.IndexOf("}")), iCount.ToString)
                        Else
                            str = str.Insert(0, iCount.ToString)       '01_12_2009  Ragava
                        End If
                    Else
                        str = iCount.ToString & "} " & str
                    End If
                    iCount += 1
                    strRichTextBox1Text = strRichTextBox1Text & str
                Next
                RichTextBox1.Text = strRichTextBox1Text
            End If
            '20_10_2009   ragava   Till   Here

            'Paint Notes Numbering
            iCount = 1

            '29_04_2011  RAGAVA
            If chkMaskPerBOM.Checked = True Then
                txtMaskPerBOM.Text = iCount.ToString
                iCount += 1
            End If
            'Till  Here

            If ChkMaskBushings.Checked = True Then
                txtMaskBushings.Text = iCount.ToString
                iCount += 1
            Else
                txtMaskBushings.Clear()
            End If
            If ChkMaskExposedThreads.Checked = True Then
                txtMaskExposedThreads.Text = iCount.ToString
                iCount += 1
            Else
                txtMaskExposedThreads.Clear()
            End If
            If ChkMaskPinHoles.Checked = True Then
                txtMaskPinholes.Text = iCount.ToString
                iCount += 1
            Else
                txtMaskPinholes.Clear()
            End If
            If ChkPrime.Checked = True Then
                txtPrime.Text = iCount.ToString
                iCount += 1
            Else
                txtPrime.Clear()
            End If
            If ChkPaint.Checked = True Then
                txtPaint.Text = iCount.ToString
                iCount += 1
            Else
                txtPaint.Clear()
            End If
            If chkAffixLabel.Checked = True Then
                txtAffixLabel.Text = iCount.ToString
                iCount += 1
            Else
                txtAffixLabel.Clear()
            End If
            If chkInstallPinAndClips.Checked = True Then
                txtInstallPinandClips.Text = iCount.ToString
                iCount += 1
            Else
                txtInstallPinandClips.Clear()
            End If

            txtPackagePerSOP.Text = iCount.ToString
            iCount += 1

            '07_09_2012   RAGAVA
            If chkNoLabelOnCylinder.Checked = True Then
                txtNoLabelOnCylinder.Text = iCount.ToString
                iCount += 1
            Else
                txtNoLabelOnCylinder.Clear()
            End If

            'Till  Here

            '20_10_2009   ragava
            If chkPaintingNote.Checked = True Then
                Dim strRichTextBox2Text As String = String.Empty
                For Each str As String In RichTextBox2.Lines
                    If Trim(strRichTextBox2Text) <> "" Then
                        strRichTextBox2Text = strRichTextBox2Text & vbNewLine
                    End If
                    If str.IndexOf("}") <> -1 Then
                        If str.IndexOf("}") > 0 Then                   '01_12_2009  Ragava
                            str = str.Replace(str.Substring(0, str.IndexOf("}")), iCount.ToString)
                        Else
                            str = str.Insert(0, iCount.ToString)       '01_12_2009  Ragava
                        End If
                    Else
                        str = iCount.ToString & "} " & str
                    End If
                    iCount += 1
                    strRichTextBox2Text = strRichTextBox2Text & str
                Next
                RichTextBox2.Text = strRichTextBox2Text
            End If
        Catch ex As Exception
        End Try

    End Sub

    Private Sub chkPackPinsAndClipsInPlasticBag_CheckedChanged(ByVal sender As System.Object, _
                                                    ByVal e As System.EventArgs)
        blnPinsPlasticBag = False            '07_04_2010   RAGAVA
        If sender.Checked = True Then
            blnPinsPlasticBag = True         '07_04_2010   RAGAVA
            chkInstallPinAndClips.Checked = False
            txtInstallPinandClips.Clear()
        End If

    End Sub
    '19_10_2009  ragava
    Private Sub chkInstallPinAndClips_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                As System.EventArgs) Handles chkInstallPinAndClips.CheckedChanged

        If sender.Checked = True Then
            txtInstallPinandClips.Enabled = True
            blnInstallPinsandClips_Checked = True           '16_06_2011   RAGAVA
        Else
            txtInstallPinandClips.Enabled = False
            blnInstallPinsandClips_Checked = False           '16_06_2011   RAGAVA
            txtInstallPinandClips.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkPins_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                    As System.EventArgs) Handles ChkPins.CheckedChanged
        If sender.Checked = True Then
            txtPins.Enabled = True
        Else
            txtPins.Enabled = False
            txtPins.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkRetractedLength_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                As System.EventArgs) Handles ChkRetractedLength.CheckedChanged

        If sender.Checked = True Then
            txtRetractedLength.Enabled = True
        Else
            txtRetractedLength.Enabled = False
            txtRetractedLength.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkExtendedLength_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                    As System.EventArgs) Handles ChkExtendedLength.CheckedChanged

        If sender.Checked = True Then
            txtExtenedLength.Enabled = True
        Else
            txtExtenedLength.Enabled = False
            txtExtenedLength.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkRodDiameter_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                        As System.EventArgs) Handles ChkRodDiameter.CheckedChanged

        If sender.Checked = True Then
            txtRodDiameter.Enabled = True
        Else
            txtRodDiameter.Enabled = False
            txtRodDiameter.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkPorts_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                    As System.EventArgs) Handles ChkPorts.CheckedChanged

        If sender.Checked = True Then
            txtPorts.Enabled = True
        Else
            txtPorts.Enabled = False
            txtPorts.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub
    ' REMOVED  10-09-2014
    'Private Sub chk100AirTest_CheckedChanged(ByVal sender As System.Object, ByVal e _
    '                            As System.EventArgs) Handles chk100AirTest.CheckedChanged

    '    If sender.Checked = True Then
    '        txtAirTest.Enabled = True
    '    Else
    '        txtAirTest.Enabled = False
    '        txtAirTest.Clear()           '10_02_2010   RAGAVA
    '    End If

    'End Sub

    Private Sub chk100OilTest_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                As System.EventArgs) Handles chk100OilTest.CheckedChanged

        If sender.Checked = True Then
            txtOilTest.Enabled = True
        Else
            txtOilTest.Enabled = False
            txtOilTest.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkRephaseExtension_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                    As System.EventArgs) Handles ChkRephaseExtension.CheckedChanged

        If sender.Checked = True Then
            txtRephaseOnExtension.Enabled = True
        Else
            txtRephaseOnExtension.Enabled = False
            txtRephaseOnExtension.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkRephaseRetraction_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                As System.EventArgs) Handles ChkRephaseRetraction.CheckedChanged

        If sender.Checked = True Then
            txtRephaseOnRetraction.Enabled = True
        Else
            txtRephaseOnRetraction.Enabled = False
            txtRephaseOnRetraction.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub chkInstallStrokeControl_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                As System.EventArgs) Handles chkInstallStrokeControl.CheckedChanged

        If sender.Checked = True Then
            txtInstallStrokeLength.Enabled = True
        Else
            txtInstallStrokeLength.Enabled = False
            txtInstallStrokeLength.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub chkStampCustomerPartandDate_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                            As System.EventArgs) Handles chkStampCustomerPartandDate.CheckedChanged

        If sender.Checked = True Then
            txtStampCustomerPartandDate.Enabled = True
            chkStampCustomerPartOnTube.Checked = False            '06_04_2010    RAGAVA
        Else
            txtStampCustomerPartandDate.Enabled = False
            txtStampCustomerPartandDate.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub chkStampCustomerPartOnTube_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                As System.EventArgs) Handles chkStampCustomerPartOnTube.CheckedChanged

        If sender.Checked = True Then
            txtStampCustomerPart.Enabled = True
            chkStampCustomerPartandDate.Checked = False           '06_04_2010    RAGAVA
        Else
            txtStampCustomerPart.Enabled = False
            txtStampCustomerPart.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkStampCountryOfOrigin_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                As System.EventArgs) Handles ChkStampCountryOfOrigin.CheckedChanged

        If sender.Checked = True Then
            txtStampCountry.Enabled = True
        Else
            txtStampCountry.Enabled = False
            txtStampCountry.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkRodMaterial_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                    As System.EventArgs) Handles ChkRodMaterial.CheckedChanged

        If sender.Checked = True Then
            txtRodMaterialNitroSteel.Enabled = True
        Else
            txtRodMaterialNitroSteel.Enabled = False
            txtRodMaterialNitroSteel.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkInstallSteelPlugs_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                    As System.EventArgs) Handles ChkInstallSteelPlugs.CheckedChanged

        If sender.Checked = True Then
            txtInstallSteelPlugs.Enabled = True
        Else
            txtInstallSteelPlugs.Enabled = False
            txtInstallSteelPlugs.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    'Private Sub ChkInstallHardenedBushings_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If sender.Checked = True Then
    '         txtInstallHardenedBushings.Enabled = True
    '    Else
    '         txtInstallHardenedBushings.Enabled = False
    '    End If
    'End Sub

    Private Sub ChkHardenedBushingsRodClevisEnd_CheckedChanged(ByVal sender As System.Object, ByVal _
                                e As System.EventArgs) Handles ChkHardenedBushingsRodClevisEnd.CheckedChanged

        If sender.Checked = True Then
            txtInstallHardenedBushingsRodClevis.Enabled = True
        Else
            txtInstallHardenedBushingsRodClevis.Enabled = False
            txtInstallHardenedBushingsRodClevis.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkHardenedBushingsClevisCapEnd_CheckedChanged(ByVal sender As System.Object, _
                    ByVal e As System.EventArgs) Handles ChkHardenedBushingsClevisCapEnd.CheckedChanged

        If sender.Checked = True Then
            txtInstallHardenedBushingsClevisCap.Enabled = True
        Else
            txtInstallHardenedBushingsClevisCap.Enabled = False
            txtInstallHardenedBushingsClevisCap.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkAssemblyStopTube_CheckedChanged(ByVal sender As System.Object, _
                            ByVal e As System.EventArgs) Handles ChkAssemblyStopTube.CheckedChanged

        If sender.Checked = True Then
            txtAssemblyStopTube.Enabled = True
        Else
            txtAssemblyStopTube.Enabled = False
            txtAssemblyStopTube.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkMaskBushings_CheckedChanged(ByVal sender As System.Object, ByVal _
                                        e As System.EventArgs) Handles ChkMaskBushings.CheckedChanged

        If sender.Checked = True Then
            txtMaskBushings.Enabled = True
        Else
            txtMaskBushings.Enabled = False
            txtMaskBushings.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkInstallVinylCap_CheckedChanged(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles ChkMaskExposedThreads.CheckedChanged

        If sender.Checked = True Then
            txtMaskExposedThreads.Enabled = True
        Else
            txtMaskExposedThreads.Enabled = False
            txtMaskExposedThreads.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkNoPaintInPinHoles_CheckedChanged(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles ChkMaskPinHoles.CheckedChanged

        If sender.Checked = True Then
            txtMaskPinholes.Enabled = True
        Else
            txtMaskPinholes.Enabled = False
            txtMaskPinholes.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkPrime_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                                As System.EventArgs) Handles ChkPrime.CheckedChanged

        If sender.Checked = True Then
            txtPrime.Enabled = True
        Else
            txtPrime.Enabled = False
            txtPrime.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub ChkPaint_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                                As System.EventArgs) Handles ChkPaint.CheckedChanged

        If sender.Checked = True Then
            txtPaint.Enabled = True
        Else
            txtPaint.Enabled = False
            txtPaint.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub

    Private Sub chkAffixValueLineLabel_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                                    As System.EventArgs) Handles chkAffixLabel.CheckedChanged

        If sender.Checked = True Then
            txtAffixLabel.Enabled = True
        Else
            txtAffixLabel.Enabled = False
            txtAffixLabel.Clear()           '10_02_2010   RAGAVA
        End If

    End Sub


    Private Sub chkShipLabelLooseWithCylinder_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                        As System.EventArgs) Handles chkPackagePerSOP.CheckedChanged

        If sender.Checked = True Then
            txtAffixLabel.Clear()
            txtAffixLabel.Enabled = False
            'chkAffixLabel.Checked = False        'sugandhi
        End If

    End Sub

    Private Sub rdbCompleteGeneration_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
                                                 rdbCompleteGeneration.CheckedChanged, rdbOnlyCosting.CheckedChanged
        If rdbCompleteGeneration.Checked Then
            IsCompleteModelGeneration = True
        Else
            IsCompleteModelGeneration = False
        End If

    End Sub

    Private Sub ColorTheForm()

        FunctionalClassObject.LabelGradient_GreenBorder_ColoringTheScreens(LabelGradient3, LabelGradient1, LabelGradient4, LabelGradient2)
        FunctionalClassObject.LabelGradient_OrangeBorder_ColoringTheScreens(LabelGradient5)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient15)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient16)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient17)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient18)

    End Sub

    Private Sub frmTieRod3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ColorTheForm()
        chkAffixLabel.Enabled = False         'sugandhi

    End Sub

    Private Sub chkMaskPerBOM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                Handles chkMaskPerBOM.CheckedChanged

        If sender.Checked = True Then
            txtMaskPerBOM.Enabled = True
        Else
            'txtMaskPerBOM.Enabled = False  'vamsi 12-09-14
            txtMaskPerBOM.Enabled = True

            txtMaskPerBOM.Clear()
        End If

    End Sub

    Public Sub LoadingDataFromExcelTieRod3()     'SUGANDHI

        'If ChkPins.Checked = True Then
        '    txtPins.Text = Module1.ReadValuesFromExcel.PinsAreInLineWithPort.ToString()
        'End If
        'If ChkRetractedLength.Checked = True Then
        '    txtRetractedLength.Text = Module1.ReadValuesFromExcel.RetractedLengthTieRod3.ToString()
        'End If
        'If ChkExtendedLength.Checked = True Then
        '    txtExtenedLength.Text = Module1.ReadValuesFromExcel.ExtendedLengthTieRod3.ToString()
        'End If
        'If ChkRodDiameter.Checked = True Then
        '    txtRodDiameter.Text = Module1.ReadValuesFromExcel.RodDiameterTieRod3.ToString()
        'End If
        'If ChkPorts.Checked = True Then
        '    txtPorts.Text = Module1.ReadValuesFromExcel.Ports.ToString()
        'End If
        'If chk100AirTest.Checked = True Then
        '    txtAirTest.Text = Module1.ReadValuesFromExcel.PercentAirTest.ToString()
        'End If
        'If chk100OilTest.Checked = True Then
        '    txtOilTest.Text = Module1.ReadValuesFromExcel.PercentOilTest.ToString()
        'End If
        'If ChkRephaseExtension.Checked = True Then
        '    txtRephaseOnExtension.Text = Module1.ReadValuesFromExcel.RephaseOnExtension.ToString()
        'End If
        'If ChkRephaseRetraction.Checked = True Then
        '    txtRephaseOnRetraction.Text = Module1.ReadValuesFromExcel.RephaseOnRetraction.ToString()
        'End If
        'If chkInstallStrokeControl.Checked = True Then
        '    txtInstallStrokeLength.Text = Module1.ReadValuesFromExcel.InstallStrokeControl.ToString()
        'End If
        'If chkStampCustomerPartandDate.Checked = True Then
        '    txtStampCustomerPartandDate.Text = Module1.ReadValuesFromExcel.StampCustomerPartAndDateCodeOnTube.ToString()
        'End If
        'If chkStampCustomerPartOnTube.Checked = True Then
        '    txtStampCustomerPart.Text = Module1.ReadValuesFromExcel.StampCustomerPartOnTube.ToString()
        'End If
        'If ChkStampCountryOfOrigin.Checked = True Then
        '    txtStampCountry.Text = Module1.ReadValuesFromExcel.StampCountryOfOriginOnTube.ToString()
        'End If
        'If ChkRodMaterial.Checked = True Then
        '    txtRodMaterialNitroSteel.Text = Module1.ReadValuesFromExcel.RodMaterialsNitroSteel.ToString()
        'End If
        'If ChkInstallSteelPlugs.Checked = True Then
        '    txtInstallSteelPlugs.Text = Module1.ReadValuesFromExcel.InstallSteelPlugsInAllPorts.ToString()
        'End If
        'If ChkHardenedBushingsRodClevisEnd.Checked = True Then
        '    txtInstallHardenedBushingsRodClevis.Text = Module1.ReadValuesFromExcel.InstallHardenedBushingsAndRodClevisEnd.ToString()
        'End If
        'If ChkHardenedBushingsClevisCapEnd.Checked = True Then
        '    txtInstallHardenedBushingsClevisCap.Text = Module1.ReadValuesFromExcel.InstallHardenedBushingsAndClevisCapEnd.ToString()
        'End If
        'If ChkAssemblyStopTube.Checked = True Then
        '    txtAssemblyStopTube.Text = Module1.ReadValuesFromExcel.AssemblyStopTubeToCylinder.ToString()
        'End If
        'If chkMaskPerBOM.Checked = True Then
        '    txtMaskPerBOM.Text = Module1.ReadValuesFromExcel.MaskPerBOMAndSOP.ToString()
        'End If
        'If ChkMaskBushings.Checked = True Then
        '    txtMaskBushings.Text = Module1.ReadValuesFromExcel.MaskBushingsBeforePainting.ToString()
        'End If
        'If ChkMaskExposedThreads.Checked = True Then
        '    txtMaskExposedThreads.Text = Module1.ReadValuesFromExcel.MaskExposedThreadsAfterWashing.ToString()
        'End If
        'If ChkMaskPinHoles.Checked = True Then
        '    txtMaskPinholes.Text = Module1.ReadValuesFromExcel.MaskPinHoles.ToString()
        'End If
        'If ChkPrime.Checked = True Then
        '    txtPrime.Text = Module1.ReadValuesFromExcel.Prime.ToString()
        'End If
        'If ChkPaint.Checked = True Then
        '    txtPaint.Text = Module1.ReadValuesFromExcel.PaintTieRod3.ToString()
        'End If
        'If chkAffixLabel.Checked = True Then
        '    txtAffixLabel.Text = Module1.ReadValuesFromExcel.AffixLabelPerSOP.ToString()
        'End If
        'If chkInstallPinAndClips.Checked = True Then
        '    txtInstallPinandClips.Text = Module1.ReadValuesFromExcel.IncludePinKitPerBOM.ToString()
        'End If
        'If chkPackCylinderInPlasticBag.Checked = True Then
        '    txtPackCylinder.Text = Module1.ReadValuesFromExcel.PackCylinderInPlasticBag.ToString()
        'End If
        'If chkAffixLabelToBag.Checked = True Then
        '    txtAffixLabeltoBag.Text = Module1.ReadValuesFromExcel.AffixLabelToBag.ToString()
        'End If
        'If chkPackagePerSOP.Checked = True Then
        '    txtAffixLabeltoBag.Text = Module1.ReadValuesFromExcel.AffixLabelToBag.ToString()
        'End If
        'If chkPackagePerSOP.Checked = True Then
        '    txtPackagePerSOP.Text = Module1.ReadValuesFromExcel.PackagePerSOP.ToString()
        'End If

        If Module1.ReadValuesFromExcel.GenerationType = "Complete Model Generation" Then
            rdbCompleteGeneration.Checked = True
        ElseIf Module1.ReadValuesFromExcel.GenerationType = "Only Costing Generation" Then
            rdbOnlyCosting.Checked = True
        End If

        If Module1.ReadValuesFromExcel.NOLABELONCYLINDER Then
            rbYesLabelRequired.Checked = True
        Else
            rbNoLabelRequired.Checked = False
        End If

        If Module1.ReadValuesFromExcel.BagRequired Then
            rbYesBagRequired.Checked = True
        Else
            rbNoBagRequired.Checked = False
        End If

    End Sub

    Private Sub rbYesLabelRequired_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbYesLabelRequired.CheckedChanged

        If ChkPrime.Checked Or ChkPaint.Checked And rbYesBagRequired.Checked Then
            If chkNoLabelOnCylinder.Checked Then
                txtAffixLabel.Enabled = False
                txtAffixLabel.Text = ""
                chkAffixLabel.Enabled = False
                chkAffixLabel.Checked = False
                chkNoLabelOnCylinder.Enabled = True
                txtNoLabelOnCylinder.Enabled = True
                txtNoLabelOnCylinder.Text = ""
            End If
        Else
            rbNoLabelRequired.Checked = True
        End If

    End Sub

    Private Sub rbNoLabelRequired_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNoLabelRequired.CheckedChanged

        chkNoLabelOnCylinder.Enabled = True
        chkNoLabelOnCylinder.Checked = True
        txtNoLabelOnCylinder.Enabled = True
        txtNoLabelOnCylinder.Text = ""

    End Sub

    Private Sub rbYesBagRequired_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbYesBagRequired.CheckedChanged

    End Sub

    Private Sub ChkFluidFilmInternal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkFluidFilmInternal.CheckedChanged
        If sender.Checked = True Then
            txtFluidFilmInternal.Enabled = True
        Else
            txtFluidFilmInternal.Enabled = False
            txtFluidFilmInternal.Clear()
        End If
    End Sub

    Private Sub chkNoLabelOnCylinder_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNoLabelOnCylinder.CheckedChanged
        If sender.Checked = True Then
            txtNoLabelOnCylinder.Enabled = True
        Else
            txtNoLabelOnCylinder.Enabled = False
            txtNoLabelOnCylinder.Clear()
        End If
    End Sub

    Private Sub txtAirTest_TextChanged(sender As System.Object, e As System.EventArgs)

    End Sub
End Class