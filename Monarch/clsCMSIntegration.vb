Imports Microsoft.Win32.Registry
Imports Microsoft.Win32.RegistryKey
Imports System.Diagnostics.Process
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Win32
Imports System.IO
Imports IFLBaseDataLayer
Imports MonarchFunctionalLayer
Imports System.Data
Public Class clsCMSIntegration

#Region "Private Variables"

    Private _strCurrentWorkingDirectory As String = System.Environment.CurrentDirectory

    Private _strCMSInterationMasterFilePath As String = _strCurrentWorkingDirectory + "\CMSIntegration_Master.xls"

    Private _strCMSInterationChildFilePath As String = _strCurrentWorkingDirectory + "\Reports\" + CylinderCodeNumber + "_CMS\" + CylinderCodeNumber + "_CMS.xls"

    Private _strCMSInterationChildFilePath_IFL As String = _strCurrentWorkingDirectory + "\Reports\" + CylinderCodeNumber + "_CMS\" + CylinderCodeNumber + "_CMS_IFL.xls"

    Private _strDirectoryName As String = _strCurrentWorkingDirectory + "\Reports\" + CylinderCodeNumber + "_CMS"

    Private _oExApplication As Excel.Application

    Private _oExWorkbook As Excel.Workbook

    Private _oExcelSheet_STKMM_TieRodCylinder As Excel.Worksheet

    Private _oExcelSheet_STKMM_Tube As Excel.Worksheet

    Private _oExcelSheet_STKMM_Rod As Excel.Worksheet

    Private _oExcelSheet_STKMP_TieRod As Excel.Worksheet

    Private _oExcelSheet_STKMP_StopTube As Excel.Worksheet

    Private _oExcelSheet_STKA_TieRodCylinder As Excel.Worksheet

    Private _oExcelSheet_STKA_Tube As Excel.Worksheet

    Private _oExcelSheet_STKA_Rod As Excel.Worksheet

    Private _oExcelSheet_STKA_TieRod As Excel.Worksheet

    Private _oExcelSheet_STKA_StopTube As Excel.Worksheet

    Private _oExcelSheet_METHDM_TieRodCylinder As Excel.Worksheet

    Private _oExcelSheet_MainAssembly As Excel.Worksheet

    Private _oExcelSheet_METHDM_Tube As Excel.Worksheet

    Private _oExcelSheet_METHDM_Rod As Excel.Worksheet

    Private _oExcelSheet_METHDR_TieRodCylinder As Excel.Worksheet

    Private _oExcelSheet_METHDR_Tube As Excel.Worksheet

    Private _oExcelSheet_METHDR_Rod As Excel.Worksheet

    Private _oExcelSheet_MTHL_TieRodCylinder As Excel.Worksheet

    Private _oExcelSheet_MTHL_Tube As Excel.Worksheet

    Private _oExcelSheet_MTHL_Rod As Excel.Worksheet

    Private _blnIsNewTube As Boolean

    Private _blnIsNewRod As Boolean

    Private _blnIsNewTierod As Boolean

    Private _blnIsNewStopTube As Boolean

    'anup 17-02-2011 start
    Private _IsExistingCodeButNotReleased_Rod As Boolean
    Private _IsExistingCodeButNotReleased_Tube As Boolean
    Private _IsExistingCodeButNotReleased_StopTube As Boolean
    Private _IsExistingCodeButNotReleased_TieRod As Boolean
    'anup 17-02-2011 till here

#End Region

#Region "Public Properties"



#End Region

#Region "Sub Procedures"

    Private Function StartLogic() As Boolean


        Try
            If ht_CodeNumbers.Count > 0 Then
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
            End If

            '14_07_2011  RAGAVA
            If IsNew_Revision_Released <> "Released" Then
                'If Directory.Exists("C:\MONARCH_TESTING\CMS_TEMP") = False Then
                '    Directory.CreateDirectory("C:\MONARCH_TESTING\CMS_TEMP")

                If Directory.Exists("K:\USR\_CYLINDER\CYLOEM\IFL DWG NR\TIEROD\CMS\") = False Then
                    Directory.CreateDirectory("K:\USR\_CYLINDER\CYLOEM\IFL DWG NR\TIEROD\CMS\")
                End If
            End If
            'Till  Here
        Catch ex As Exception

        End Try

        Try
            CheckForNewOrExisting() '14-07-10-Sunny

            STKMM_TieRodCylinder_Functionality() 'STKMM_STKMP

            STKA_TieRodCylinder_Functionality() 'STKA

            METHDM_TieRodCylinder_Functionality() 'METHDM

            METHDR_TieRodCylinder_Functionality() 'METHDR

            MTHL_TieRodCylinder_Functionality() 'MTHL

            File.Copy(_strCMSInterationChildFilePath, _strCMSInterationChildFilePath_IFL) 'Copy of child Excel for reference

            CSVConversionFunctionality()

            MoveDirectoryToW()

            METHE_TieRodCylinder_Functionality()     '15_09_2010    RAGAVA  'METHE 

            ShuffleCMSfiles()        '14_12_2010   RAGAVA
        Catch ex As Exception

        End Try
    End Function

    '14_12_2010       RAGAVA
    Private Sub ShuffleCMSfiles()
        Try
            'anup 10-03-2011 sart
            Dim strReleasedRodCodeNumber As String = String.Empty
            Dim strReleasedBoreCodeNumber As String = String.Empty
            Dim strReleasedStopTubeCodeNumber As String = String.Empty
            Dim strReleasedTieRodCodeNumber As String = String.Empty
            'anup 10-03-2011 till here

            KillExcel()

            '14_07_2011   RAGAVA
            Dim strCMSLocation As String = String.Empty
            If IsNew_Revision_Released = "Released" Then
                strCMSLocation = "W:\TIEROD\CMS\" + CylinderCodeNumber + "_CMS"
            Else
                'strCMSLocation = "C:\MONARCH_TESTING\CMS_TEMP\" + CylinderCodeNumber + "_CMS"
                strCMSLocation = "K:\USR\_CYLINDER\CYLOEM\IFL DWG NR\TIEROD\CMS\" + CylinderCodeNumber + "_CMS"
            End If
            'Dim strCMSLocation As String = "W:\TIEROD\CMS\" + CylinderCodeNumber + "_CMS"
            'Till   Here


            Dim strFileList() As String = Directory.GetFiles(strCMSLocation)
            Directory.CreateDirectory(strCMSLocation & "\CYLINDER" & CylinderCodeNumber.ToString)

            For Each strfile As String In strFileList
                If strfile.EndsWith(CylinderCodeNumber & ".csv") = True Then
                    strfile = strfile.Substring(strfile.LastIndexOf("\") + 1, strfile.Length - (strfile.LastIndexOf("\") + 1))
                    File.Move(strCMSLocation & "\" & strfile, strCMSLocation & "\CYLINDER" & CylinderCodeNumber.ToString & "\" & strfile.Replace(CylinderCodeNumber, "RO"))
                End If
            Next

            Try
                ReModify_STKAfiles(strCMSLocation & "\CYLINDER" & CylinderCodeNumber.ToString)            '10_01_2011   RAGAVA
            Catch ex As Exception
            End Try

            If _blnIsNewTube OrElse _IsExistingCodeButNotReleased_Tube Then  'anup 17-02-2011 'anup 10-03-2011
                strReleasedBoreCodeNumber = strBoreCodeNumber  'anup 10-03-2011
                Directory.CreateDirectory(strCMSLocation & "\TUBE" & strBoreCodeNumber.ToString)
                For Each strfile As String In strFileList
                    If strfile.EndsWith(strBoreCodeNumber & ".csv") = True Then
                        strfile = strfile.Substring(strfile.LastIndexOf("\") + 1, strfile.Length - (strfile.LastIndexOf("\") + 1))
                        File.Move(strCMSLocation & "\" & strfile, strCMSLocation & "\TUBE" & strBoreCodeNumber.ToString & "\" & strfile.Replace(strBoreCodeNumber, "RO"))
                    End If
                Next
                Try
                    ReModify_STKAfiles(strCMSLocation & "\TUBE" & strBoreCodeNumber.ToString)            '10_01_2011   RAGAVA
                Catch ex As Exception
                End Try
            End If
            If _blnIsNewRod OrElse _IsExistingCodeButNotReleased_Rod Then  'anup 17-02-2011'anup 10-03-2011
                strReleasedRodCodeNumber = strRodCodeNumber  'anup 10-03-2011
                Directory.CreateDirectory(strCMSLocation & "\ROD" & strRodCodeNumber.ToString)
                For Each strfile As String In strFileList
                    If strfile.EndsWith(strRodCodeNumber & ".csv") = True Then
                        strfile = strfile.Substring(strfile.LastIndexOf("\") + 1, strfile.Length - (strfile.LastIndexOf("\") + 1))
                        File.Move(strCMSLocation & "\" & strfile, strCMSLocation & "\ROD" & strRodCodeNumber.ToString & "\" & strfile.Replace(strRodCodeNumber, "RO"))
                    End If
                Next
                Try
                    ReModify_STKAfiles(strCMSLocation & "\ROD" & strRodCodeNumber.ToString)            '10_01_2011   RAGAVA
                Catch ex As Exception
                End Try
            End If
            If _blnIsNewTierod OrElse _IsExistingCodeButNotReleased_TieRod Then  'anup 17-02-2011'anup 10-03-2011
                strReleasedTieRodCodeNumber = strTieRodCodeNumber  'anup 10-03-2011
                Directory.CreateDirectory(strCMSLocation & "\TIEROD" & strTieRodCodeNumber.ToString)
                For Each strfile As String In strFileList
                    If strfile.EndsWith(strTieRodCodeNumber & ".csv") = True Then
                        strfile = strfile.Substring(strfile.LastIndexOf("\") + 1, strfile.Length - (strfile.LastIndexOf("\") + 1))
                        File.Move(strCMSLocation & "\" & strfile, strCMSLocation & "\TIEROD" & strTieRodCodeNumber.ToString & "\" & strfile.Replace(strTieRodCodeNumber, "RO"))
                    End If
                Next
                Try
                    ReModify_STKAfiles(strCMSLocation & "\TIEROD" & strTieRodCodeNumber.ToString)            '10_01_2011   RAGAVA
                Catch ex As Exception
                End Try
            End If
            If IsStopTubeSelected Then
                strReleasedStopTubeCodeNumber = StopTubeCodeNumber   'anup 10-03-2011
                Directory.CreateDirectory(strCMSLocation & "\STOPTUBE" & StopTubeCodeNumber.ToString)
                For Each strfile As String In strFileList
                    If strfile.EndsWith(StopTubeCodeNumber & ".csv") = True Then
                        strfile = strfile.Substring(strfile.LastIndexOf("\") + 1, strfile.Length - (strfile.LastIndexOf("\") + 1))
                        File.Move(strCMSLocation & "\" & strfile, strCMSLocation & "\STOPTUBE" & StopTubeCodeNumber.ToString & "\" & strfile.Replace(StopTubeCodeNumber, "RO"))
                    End If
                Next
                Try
                    ReModify_STKAfiles(strCMSLocation & "\STOPTUBE" & StopTubeCodeNumber.ToString)            '10_01_2011   RAGAVA
                Catch ex As Exception
                End Try
            End If


            'anup 17-02-2011 start
            If IsNew_Revision_Released = "Released" Then       '19_07_2011  RAGAVA  
                Dim oClsReleaseCylinderFunctionality As New clsReleaseCylinderFunctionality
                oClsReleaseCylinderFunctionality.DropRod_Tube_Stoptube_TieRodCodesToDB(strReleasedRodCodeNumber, strReleasedBoreCodeNumber, strReleasedStopTubeCodeNumber, strReleasedTieRodCodeNumber, CylinderCodeNumber) 'anup 10-03-2011
            End If
            'anup 17-02-2011 till here

        Catch ex As Exception

        End Try
    End Sub

    '10_01_2011    RAGAVA
    Private Sub ReModify_STKAfiles(ByVal strFolderLocation As String)
        Try
            Dim strFiles As String() = Directory.GetFiles(strFolderLocation)
            For Each strFile As String In strFiles
                If strFile.IndexOf("STKA") <> -1 Then
                    Dim sr As New StreamReader(strFile)
                    Dim fs1 As New FileStream(strFolderLocation & "\dummy.txt", FileMode.Create)
                    Dim sw As New StreamWriter(fs1)
                    sw.Write(sr.ReadLine)
                    sw.Close()
                    sr.Close()
                    GC.Collect()
                    File.Delete(strFile)
                    File.Move(strFolderLocation & "\dummy.txt", strFolderLocation & "\STKARO.CSV")
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub

    '14-07-10-Sunny
    Private Sub CheckForNewOrExisting()
        Try
            If Not IsNothing(strBoreCodeNumber) Then
                '21_01_2011    RAGAVA
                'If strBoreCodeNumber.StartsWith("7") Then
                '    _blnIsNewTube = True
                'Else
                '    _blnIsNewTube = False
                'End If
                If strCodeNumber_BeforeApplicationStart > strBoreCodeNumber Then
                    _blnIsNewTube = False
                Else
                    _blnIsNewTube = True
                End If
                'If strCodeNumber_BeforeApplicationStart > strRodCodeNumber Then
                '    oCodeNumber_AfterAscendingItem(2) = "Existing"
                'End If
            End If

            If Not IsNothing(strRodCodeNumber) Then
                If strCodeNumber_BeforeApplicationStart > strRodCodeNumber Then        '21_01_2011    RAGAVA
                    _blnIsNewRod = False
                Else
                    _blnIsNewRod = True
                End If
            End If

            If Not IsNothing(strTieRodCodeNumber) Then
                If strCodeNumber_BeforeApplicationStart > strTieRodCodeNumber Then        '21_01_2011    RAGAVA
                    _blnIsNewTierod = False
                Else
                    _blnIsNewTierod = True
                End If
            End If

            If Not IsNothing(StopTubeCodeNumber) Then
                If strCodeNumber_BeforeApplicationStart > StopTubeCodeNumber Then        '21_01_2011    RAGAVA
                    _blnIsNewStopTube = False
                Else
                    _blnIsNewStopTube = True
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

#Region "STKMM_STKMP"

    Private Sub STKMM_TieRodCylinder_Functionality()
        Try
            Dim objclsCMS_STKMM_STKMP_CylinderSheet As New clsCMS_STKMM_STKMP
            objclsCMS_STKMM_STKMP_CylinderSheet.SetCommenPropertyValue()
            STKMM_STKMP_CylinderSheetLogics(objclsCMS_STKMM_STKMP_CylinderSheet)
            objclsCMS_STKMM_STKMP_CylinderSheet.SetDataToExcel(_oExcelSheet_STKMM_TieRodCylinder)
            _oExWorkbook.Save()

            Dim objclsCMS_STKMM_STKMP_TubeSheet As New clsCMS_STKMM_STKMP
            objclsCMS_STKMM_STKMP_TubeSheet.SetCommenPropertyValue()
            STKMM_STKMP_TubeSheetLogics(objclsCMS_STKMM_STKMP_TubeSheet)
            objclsCMS_STKMM_STKMP_TubeSheet.SetDataToExcel(_oExcelSheet_STKMM_Tube)
            _oExWorkbook.Save()

            Dim objclsCMS_STKMM_STKMP_RodSheet As New clsCMS_STKMM_STKMP
            objclsCMS_STKMM_STKMP_RodSheet.SetCommenPropertyValue()
            STKMM_STKMP_RodSheetLogics(objclsCMS_STKMM_STKMP_RodSheet)
            objclsCMS_STKMM_STKMP_RodSheet.SetDataToExcel(_oExcelSheet_STKMM_Rod)
            _oExWorkbook.Save()

            Dim objclsCMS_STKMM_STKMP_TieRodSheet As New clsCMS_STKMM_STKMP
            objclsCMS_STKMM_STKMP_TieRodSheet.SetCommenPropertyValue()
            STKMM_STKMP_TieRodSheetLogics(objclsCMS_STKMM_STKMP_TieRodSheet)
            objclsCMS_STKMM_STKMP_TieRodSheet.SetDataToExcel(_oExcelSheet_STKMP_TieRod)
            _oExWorkbook.Save()

            If IsStopTubeSelected Then
                Dim objclsCMS_STKMM_STKMP_StopTubeSheet As New clsCMS_STKMM_STKMP
                objclsCMS_STKMM_STKMP_StopTubeSheet.SetCommenPropertyValue()
                STKMM_STKMP_SheetStopTubeLogics(objclsCMS_STKMM_STKMP_StopTubeSheet)
                objclsCMS_STKMM_STKMP_StopTubeSheet.SetDataToExcel(_oExcelSheet_STKMP_StopTube)
                _oExWorkbook.Save()
            End If
        Catch ex As Exception

        End Try
        
    End Sub

    Private Sub STKMM_STKMP_CylinderSheetLogics(ByVal objclsCMS_STKMM_STKMP_CylinderSheet As clsCMS_STKMM_STKMP) 'STKMM_TieRodCylinder
        Try

      
            objclsCMS_STKMM_STKMP_CylinderSheet.InternalPartNumber = CylinderCodeNumber
            objclsCMS_STKMM_STKMP_CylinderSheet.PartType = 1

            '12-07-10-sunny
            'anup 08-03-2011 start
            'If PartCode1 <> "tba" OrElse PartCode1 <> "" Then
            '    'objclsCMS_STKMM_STKMP_CylinderSheet.PartDescriptionLine1 = "CYL " + SetCodeDesciption + CustomerName + "#" + PartCode1
            '    If CustomerName.Length < 4 Then
            '        objclsCMS_STKMM_STKMP_CylinderSheet.PartDescriptionLine1 = "CYL " + SetCodeDesciption + " " + IIf(CustomerName = "", "", CustomerName.Substring(0, 3)) + "#" + PartCode1           '17_11_2010   RAGAVA     '04_10_2010   RAGAVA  Space added before customer name        '06_09_2010    RAGAVA        
            '    Else
            '        objclsCMS_STKMM_STKMP_CylinderSheet.PartDescriptionLine1 = "CYL " + SetCodeDesciption + " " + IIf(CustomerName = "", "", CustomerName.Substring(0, 4)) + "#" + PartCode1    '17_11_2010   RAGAVA      '04_10_2010   RAGAVA  Space added before customer name    '06_09_2010    RAGAVA
            '    End If
            'Else
            '    'objclsCMS_STKMM_STKMP_CylinderSheet.PartDescriptionLine1 = "CYL " + SetCodeDesciption + CustomerName
            '    If CustomerName.Length < 4 Then
            '        objclsCMS_STKMM_STKMP_CylinderSheet.PartDescriptionLine1 = "CYL " + SetCodeDesciption + " " + IIf(CustomerName = "", "", CustomerName.Substring(0, 3))      '17_11_2010   RAGAVA     '04_10_2010   RAGAVA  Space added before customer name    '06_09_2010   RAGAVA
            '    Else
            '        objclsCMS_STKMM_STKMP_CylinderSheet.PartDescriptionLine1 = "CYL " + SetCodeDesciption + " " + IIf(CustomerName = "", "", CustomerName.Substring(0, 4))      '17_11_2010   RAGAVA      '04_10_2010   RAGAVA  Space added before customer name  '06_09_2010   RAGAVA
            '    End If
            'End If
            Dim intLength As Integer
            If CustomerName.Length < 4 Then
                objclsCMS_STKMM_STKMP_CylinderSheet.PartDescriptionLine1 = "CYL " + SetCodeDesciption + " " + IIf(CustomerName = "", "", CustomerName.Substring(0, 3))
                intLength = 3
            Else
                objclsCMS_STKMM_STKMP_CylinderSheet.PartDescriptionLine1 = "CYL " + SetCodeDesciption + " " + IIf(CustomerName = "", "", CustomerName.Substring(0, 4))
                intLength = 4
            End If

            If PartCode1 <> "tba" AndAlso PartCode1 <> "" Then
                Dim strTempDescription As String = "CYL " + SetCodeDesciption + " " + IIf(CustomerName = "", "", CustomerName.Substring(0, intLength)) + "#" + PartCode1
                If strTempDescription.Length <= 30 Then
                    objclsCMS_STKMM_STKMP_CylinderSheet.PartDescriptionLine1 += "#" + PartCode1
                End If
            End If
            'anup 08-03-2011 till here
        Catch ex As Exception

        End Try
        Try
            objclsCMS_STKMM_STKMP_CylinderSheet.PartDescriptionLine2 = ""
            objclsCMS_STKMM_STKMP_CylinderSheet.GLExpenseCode = "FGB"
            objclsCMS_STKMM_STKMP_CylinderSheet.MajorGroupCode = "'" & Format(30, "000")     '06_09_2010    RAGAVA
            Dim strSearchString As String = ""
            If SeriesForCosting.Equals("TL (TC)") Then
                strSearchString = "TL(TC)"
                objclsCMS_STKMM_STKMP_CylinderSheet.MinorGroupCode = "T23"         '06_09_2010   RAGAVA
            ElseIf SeriesForCosting.Equals("TH (TD)") Then
                strSearchString = "TH(TD)"
                objclsCMS_STKMM_STKMP_CylinderSheet.MinorGroupCode = "T24"         '06_09_2010   RAGAVA
            ElseIf SeriesForCosting.Equals("TP-High") OrElse SeriesForCosting.Equals("TP-Low") Then
                strSearchString = "TP"
                objclsCMS_STKMM_STKMP_CylinderSheet.MinorGroupCode = "T25"         '06_09_2010   RAGAVA
            ElseIf SeriesForCosting.Equals("TX (TXC)") Then
                If strStyleModified.Equals("ASAE") Then
                    strSearchString = "TX - ASAE"
                    objclsCMS_STKMM_STKMP_CylinderSheet.MinorGroupCode = "T40"         '06_09_2010   RAGAVA
                ElseIf strStyleModified.Equals("NON ASAE") Then
                    strSearchString = "TX - Non ASAE"
                    objclsCMS_STKMM_STKMP_CylinderSheet.MinorGroupCode = "T41"         '06_09_2010   RAGAVA
                End If
            ElseIf SeriesForCosting = "LN" Then         '20_01_2011   RAGAVA
                strSearchString = "LN"        '20_01_2011   RAGAVA
                objclsCMS_STKMM_STKMP_CylinderSheet.MinorGroupCode = "T47"        '20_01_2011   RAGAVA
            End If

            Try
                objclsCMS_STKMM_STKMP_CylinderSheet.MinorGroupCode = IFLConnectionObject.GetValue("select MinorGroupCode_STKMM_STKMP from CMS_OtherFields where Details = '" + strSearchString + "'")
            Catch ex As Exception
            End Try

            objclsCMS_STKMM_STKMP_CylinderSheet.MajorSalesCode = objclsCMS_STKMM_STKMP_CylinderSheet.MajorGroupCode
            objclsCMS_STKMM_STKMP_CylinderSheet.MinorSalesCode = objclsCMS_STKMM_STKMP_CylinderSheet.MinorGroupCode

            '13-07-10-sunny
            If UCase(PartCode1).IndexOf("TBA") = -1 AndAlso PartCode1 <> "" Then
                objclsCMS_STKMM_STKMP_CylinderSheet.CustomerPartNumber = PartCode1 'Sunny 13-07-10
            End If

            Try
                objclsCMS_STKMM_STKMP_CylinderSheet.CatalogId = IFLConnectionObject. _
                GetValue("select CagalogID_STKMM_STKMP from CMS_OtherFields where ItemType = 'Tie Rod Cylinder' and MinorGroupCode_STKMM_STKMP = '" _
                                   + objclsCMS_STKMM_STKMP_CylinderSheet.MinorGroupCode + "'")
            Catch ex As Exception
            End Try
            objclsCMS_STKMM_STKMP_CylinderSheet.NetWeight = Math.Round(Volume_Assembly * 0.27812, 4)   '06_09_2010    RAGAVA 'Density of steel
            objclsCMS_STKMM_STKMP_CylinderSheet.HarmonizationCode = "8412.21.0015"
            objclsCMS_STKMM_STKMP_CylinderSheet.UserVerificationTemplateCode = "Customs"   'Sunny 24-05-10
            objclsCMS_STKMM_STKMP_CylinderSheet.AntiDumpingTracking = ""
            objclsCMS_STKMM_STKMP_CylinderSheet.DumpingSubjectIndicator = ""
            objclsCMS_STKMM_STKMP_CylinderSheet.ServiceChargePart = 2
        Catch ex As Exception

        End Try
    End Sub

    Private Sub STKMM_STKMP_TubeSheetLogics(ByVal objclsCMS_STKMM_STKMP_TubeSheet As clsCMS_STKMM_STKMP) 'STKMM_Tube
        objclsCMS_STKMM_STKMP_TubeSheet.InternalPartNumber = strBoreCodeNumber
        objclsCMS_STKMM_STKMP_TubeSheet.PartType = 1
        Try
            objclsCMS_STKMM_STKMP_TubeSheet.PartDescriptionLine1 = IFLConnectionObject.GetValue("select Description from CostingDetails where PartCode = '" + strBoreCodeNumber + "'")
            If IsNothing(objclsCMS_STKMM_STKMP_TubeSheet.PartDescriptionLine1) Then
                objclsCMS_STKMM_STKMP_TubeSheet.PartDescriptionLine1 = "TUBE CYL " + BoreDiameter.ToString + "-" + StrokeLength.ToString
            End If
        Catch ex As Exception
        End Try
        objclsCMS_STKMM_STKMP_TubeSheet.PartDescriptionLine2 = ""
        objclsCMS_STKMM_STKMP_TubeSheet.GLExpenseCode = "SFB"
        objclsCMS_STKMM_STKMP_TubeSheet.MajorGroupCode = "'" & Format(35, "000")     '06_09_2010   RAGAVA
        objclsCMS_STKMM_STKMP_TubeSheet.MinorGroupCode = "V22"
        objclsCMS_STKMM_STKMP_TubeSheet.MajorSalesCode = objclsCMS_STKMM_STKMP_TubeSheet.MajorGroupCode
        objclsCMS_STKMM_STKMP_TubeSheet.MinorSalesCode = objclsCMS_STKMM_STKMP_TubeSheet.MinorGroupCode
        objclsCMS_STKMM_STKMP_TubeSheet.CustomerPartNumber = ""
        Try
            objclsCMS_STKMM_STKMP_TubeSheet.CatalogId = IFLConnectionObject.GetValue("select CagalogID_STKMM_STKMP from CMS_OtherFields where ItemType = 'Tube'")
        Catch ex As Exception
        End Try
        objclsCMS_STKMM_STKMP_TubeSheet.NetWeight = Math.Round(Volume_Bore * 0.27812, 4)       '06_09_2010   RAGAVA 'Density of steel
        objclsCMS_STKMM_STKMP_TubeSheet.HarmonizationCode = "8412.90.9005"
        objclsCMS_STKMM_STKMP_TubeSheet.UserVerificationTemplateCode = "Customs" 'Sunny 24-05-10
        objclsCMS_STKMM_STKMP_TubeSheet.AntiDumpingTracking = ""
        objclsCMS_STKMM_STKMP_TubeSheet.DumpingSubjectIndicator = ""
        objclsCMS_STKMM_STKMP_TubeSheet.ServiceChargePart = 2
    End Sub

    Private Sub STKMM_STKMP_RodSheetLogics(ByVal objclsCMS_STKMM_STKMP_RodSheet As clsCMS_STKMM_STKMP) 'STKMM_Rod
        objclsCMS_STKMM_STKMP_RodSheet.InternalPartNumber = strRodCodeNumber
        objclsCMS_STKMM_STKMP_RodSheet.PartType = 1

        Try
            objclsCMS_STKMM_STKMP_RodSheet.PartDescriptionLine1 = IFLConnectionObject.GetValue("select Purchased_Manfractured, Description from CostingDetails where PartCode = '" + strRodCodeNumber + "'")
            If IsNothing(objclsCMS_STKMM_STKMP_RodSheet.PartDescriptionLine1) Then
                objclsCMS_STKMM_STKMP_RodSheet.PartDescriptionLine1 = "ROD CYL " + RodDiameter.ToString + "-" + StrokeLength.ToString + "-"
                If IsNothing(PistonThreadSize) Then
                    objclsCMS_STKMM_STKMP_RodSheet.PartDescriptionLine1 += dblRodThreadSize.ToString
                Else
                    objclsCMS_STKMM_STKMP_RodSheet.PartDescriptionLine1 += PistonThreadSize.ToString + "_" + dblRodThreadSize.ToString
                End If
            End If
        Catch ex As Exception
        End Try

        If RodMaterialForCosting = "Chrome" Then
            objclsCMS_STKMM_STKMP_RodSheet.PartDescriptionLine2 = "CHROME"
        ElseIf RodMaterialForCosting = "Nitro Steel" Then
            objclsCMS_STKMM_STKMP_RodSheet.PartDescriptionLine2 = "NITRO STEEL"
        End If
        objclsCMS_STKMM_STKMP_RodSheet.GLExpenseCode = "SFB"
        objclsCMS_STKMM_STKMP_RodSheet.MajorGroupCode = "'" & Format(35, "000")
        objclsCMS_STKMM_STKMP_RodSheet.MinorGroupCode = "V22"
        objclsCMS_STKMM_STKMP_RodSheet.MajorSalesCode = objclsCMS_STKMM_STKMP_RodSheet.MajorGroupCode
        objclsCMS_STKMM_STKMP_RodSheet.MinorSalesCode = objclsCMS_STKMM_STKMP_RodSheet.MinorGroupCode
        objclsCMS_STKMM_STKMP_RodSheet.CustomerPartNumber = ""
        Try
            objclsCMS_STKMM_STKMP_RodSheet.CatalogId = IFLConnectionObject.GetValue("select CagalogID_STKMM_STKMP from CMS_OtherFields where ItemType = 'Rod'")
        Catch ex As Exception
        End Try
        objclsCMS_STKMM_STKMP_RodSheet.NetWeight = Math.Round(Volume_Rod * 0.27812, 4)        '14_09_2010   RAGAVA 'Density of steel
        objclsCMS_STKMM_STKMP_RodSheet.HarmonizationCode = "8412.90.9005"
        objclsCMS_STKMM_STKMP_RodSheet.UserVerificationTemplateCode = "Customs" 'Sunny 24-05-10
        objclsCMS_STKMM_STKMP_RodSheet.AntiDumpingTracking = ""
        objclsCMS_STKMM_STKMP_RodSheet.DumpingSubjectIndicator = ""
        objclsCMS_STKMM_STKMP_RodSheet.ServiceChargePart = 2
    End Sub

    Private Sub STKMM_STKMP_TieRodSheetLogics(ByVal objclsCMS_STKMM_STKMP_TieRodSheet As clsCMS_STKMM_STKMP) 'STKMP_TieRod
        Dim strPurc_Manu As String = ""
        Dim strDesc As String = ""
        objclsCMS_STKMM_STKMP_TieRodSheet.InternalPartNumber = strTieRodCodeNumber
        objclsCMS_STKMM_STKMP_TieRodSheet.PartType = 2

        Try
            objclsCMS_STKMM_STKMP_TieRodSheet.PartDescriptionLine1 = IFLConnectionObject.GetValue("select Purchased_Manfractured, Description from CostingDetails where PartCode = '" + strTieRodCodeNumber + "'")
            If IsNothing(objclsCMS_STKMM_STKMP_TieRodSheet.PartDescriptionLine1) Then
                objclsCMS_STKMM_STKMP_TieRodSheet.PartDescriptionLine1 = "TIE ROD " + TieRodSize.ToString + "-" + StrokeLength.ToString + "-"
                If IsNothing(PistonThreadSize) Then
                    objclsCMS_STKMM_STKMP_TieRodSheet.PartDescriptionLine1 += dblRodThreadSize.ToString
                Else
                    objclsCMS_STKMM_STKMP_TieRodSheet.PartDescriptionLine1 += PistonThreadSize.ToString + "_" + dblRodThreadSize.ToString
                End If
            End If
        Catch ex As Exception
        End Try

        objclsCMS_STKMM_STKMP_TieRodSheet.PartDescriptionLine2 = ""
        objclsCMS_STKMM_STKMP_TieRodSheet.GLExpenseCode = "RAB"
        objclsCMS_STKMM_STKMP_TieRodSheet.MajorGroupCode = "'" & "035"
        objclsCMS_STKMM_STKMP_TieRodSheet.MinorGroupCode = "V01"
        objclsCMS_STKMM_STKMP_TieRodSheet.MajorSalesCode = objclsCMS_STKMM_STKMP_TieRodSheet.MajorGroupCode
        objclsCMS_STKMM_STKMP_TieRodSheet.MinorSalesCode = objclsCMS_STKMM_STKMP_TieRodSheet.MinorGroupCode
        objclsCMS_STKMM_STKMP_TieRodSheet.CustomerPartNumber = ""
        Try
            objclsCMS_STKMM_STKMP_TieRodSheet.CatalogId = IFLConnectionObject.GetValue("select CagalogID_STKMM_STKMP from CMS_OtherFields where ItemType = 'Tie Rod'")
        Catch ex As Exception
        End Try
        objclsCMS_STKMM_STKMP_TieRodSheet.NetWeight = Math.Round(Volume_TieRod * 0.27812, 4)        '14_09_2010   RAGAVA 'Density of steel
        objclsCMS_STKMM_STKMP_TieRodSheet.HarmonizationCode = "9801.00.1095"
        objclsCMS_STKMM_STKMP_TieRodSheet.UserVerificationTemplateCode = "Customs"  'Sunny 24-05-10
        objclsCMS_STKMM_STKMP_TieRodSheet.AntiDumpingTracking = 2
        objclsCMS_STKMM_STKMP_TieRodSheet.DumpingSubjectIndicator = 2
        objclsCMS_STKMM_STKMP_TieRodSheet.ServiceChargePart = ""
    End Sub

    Private Sub STKMM_STKMP_SheetStopTubeLogics(ByVal objclsCMS_STKMM_STKMP_StopTubeSheet As clsCMS_STKMM_STKMP) 'STKMP_StopTube
        Dim strPurc_Manu As String = ""
        Dim strDesc As String = ""
        objclsCMS_STKMM_STKMP_StopTubeSheet.InternalPartNumber = StopTubeCodeNumber
        objclsCMS_STKMM_STKMP_StopTubeSheet.PartType = 2

        Try
            objclsCMS_STKMM_STKMP_StopTubeSheet.PartDescriptionLine1 = IFLConnectionObject.GetValue("select Purchased_Manfractured, Description from CostingDetails where PartCode = '" + StopTubeCodeNumber + "'")
            If IsNothing(objclsCMS_STKMM_STKMP_StopTubeSheet.PartDescriptionLine1) Then
                objclsCMS_STKMM_STKMP_StopTubeSheet.PartDescriptionLine1 = "STOP TUBE " + StopTubeID.ToString + "-" + StopTubeOD.ToString + "-" + StrokeLength.ToString
            End If
        Catch ex As Exception
        End Try

        objclsCMS_STKMM_STKMP_StopTubeSheet.PartDescriptionLine2 = ""
        objclsCMS_STKMM_STKMP_StopTubeSheet.GLExpenseCode = "RAB"
        objclsCMS_STKMM_STKMP_StopTubeSheet.MajorGroupCode = "'" & "035"
        objclsCMS_STKMM_STKMP_StopTubeSheet.MinorGroupCode = "V01"
        objclsCMS_STKMM_STKMP_StopTubeSheet.MajorSalesCode = objclsCMS_STKMM_STKMP_StopTubeSheet.MajorGroupCode
        objclsCMS_STKMM_STKMP_StopTubeSheet.MinorSalesCode = objclsCMS_STKMM_STKMP_StopTubeSheet.MinorGroupCode
        objclsCMS_STKMM_STKMP_StopTubeSheet.CustomerPartNumber = ""
        Try
            objclsCMS_STKMM_STKMP_StopTubeSheet.CatalogId = IFLConnectionObject.GetValue("select CagalogID_STKMM_STKMP from CMS_OtherFields where ItemType = 'Stop Tube'")
        Catch ex As Exception
        End Try
        objclsCMS_STKMM_STKMP_StopTubeSheet.NetWeight = Math.Round(Volume_StopTube * 0.27812, 4)        '14_09_2010   RAGAVA 'Density of steel
        objclsCMS_STKMM_STKMP_StopTubeSheet.HarmonizationCode = "8412.90.9005"
        objclsCMS_STKMM_STKMP_StopTubeSheet.UserVerificationTemplateCode = ""
        objclsCMS_STKMM_STKMP_StopTubeSheet.AntiDumpingTracking = 2
        objclsCMS_STKMM_STKMP_StopTubeSheet.DumpingSubjectIndicator = 2
        objclsCMS_STKMM_STKMP_StopTubeSheet.ServiceChargePart = ""
    End Sub

#End Region

#Region "STKA"

    Private Sub STKA_TieRodCylinder_Functionality()
        Try
            Dim objClsCMS_STKA_CylinderSheet As New clsCMS_STKA
            objClsCMS_STKA_CylinderSheet.SetCommenPropertyValue()
            STKA_CylinderSheetLogics(objClsCMS_STKA_CylinderSheet)
            objClsCMS_STKA_CylinderSheet.SetDataToExcel(_oExcelSheet_STKA_TieRodCylinder)
            _oExWorkbook.Save()

            Dim objClsCMS_STKA_TubeSheet As New clsCMS_STKA
            objClsCMS_STKA_TubeSheet.SetCommenPropertyValue()
            STKA_TubeSheetLogics(objClsCMS_STKA_TubeSheet)
            objClsCMS_STKA_TubeSheet.SetDataToExcel(_oExcelSheet_STKA_Tube)
            _oExWorkbook.Save()

            Dim objClsCMS_STKA_RodSheet As New clsCMS_STKA
            objClsCMS_STKA_RodSheet.SetCommenPropertyValue()
            STKA_RodSheetLogics(objClsCMS_STKA_RodSheet)
            objClsCMS_STKA_RodSheet.SetDataToExcel(_oExcelSheet_STKA_Rod)
            _oExWorkbook.Save()

            Dim objClsCMS_STKA_TieRodSheet As New clsCMS_STKA
            objClsCMS_STKA_TieRodSheet.SetCommenPropertyValue()
            STKA_TieRodSheetLogics(objClsCMS_STKA_TieRodSheet)
            objClsCMS_STKA_TieRodSheet.SetDataToExcel(_oExcelSheet_STKA_TieRod)
            _oExWorkbook.Save()

            If IsStopTubeSelected Then
                Dim objClsCMS_STKA_StopTubeSheet As New clsCMS_STKA
                objClsCMS_STKA_StopTubeSheet.SetCommenPropertyValue()
                STKA_StopTubeSheetLogics(objClsCMS_STKA_StopTubeSheet)
                objClsCMS_STKA_StopTubeSheet.SetDataToExcel(_oExcelSheet_STKA_StopTube)
                _oExWorkbook.Save()
            End If
        Catch ex As Exception

        End Try
        
    End Sub

    Private Sub STKA_CylinderSheetLogics(ByVal objClsCMS_STKA_CylinderSheet As clsCMS_STKA) 'STKA_TieRodCylinder
        objClsCMS_STKA_CylinderSheet.InternalPartNumber = CylinderCodeNumber
        objClsCMS_STKA_CylinderSheet.ReplenishmentType = 1
        objClsCMS_STKA_CylinderSheet.VendorLeadTime_TransferLeadTime_inDays = ""
        objClsCMS_STKA_CylinderSheet.MinimumOrderQuantity_inUnitofIssue = 1
        objClsCMS_STKA_CylinderSheet.BuyerCode = ""
        objClsCMS_STKA_CylinderSheet.PlannerCode = "TC"

        'Sunny 22-06-10
        objClsCMS_STKA_CylinderSheet.ReceiveToLocation = "BPC"     '06_09_2010    RAGAVA "C01BPC"

        '16_09_2010  RAGAVA
        If UCase(CustomerName).IndexOf("CNH") <> -1 Then
            objClsCMS_STKA_CylinderSheet.InspectionProcedure = "ISIR"
        Else
            objClsCMS_STKA_CylinderSheet.InspectionProcedure = "PROTO"  '19_11_2010   ANUP
            '    objClsCMS_STKA_CylinderSheet.InspectionProcedure = ""               '18_11_2010   ANUP
        End If
        'Till  Here
        '************

        objClsCMS_STKA_CylinderSheet.MaterialPreparationLeadTime_inDays = 12 'vamsi 06-01-13
        'objClsCMS_STKA_CylinderSheet.MaterialPreparationLeadTime_inDays = 7                'sugandhi 08_05_2012
        objClsCMS_STKA_CylinderSheet.CountryOfOrigin = "CA"
        objClsCMS_STKA_CylinderSheet.ScheduleType = ""
        objClsCMS_STKA_CylinderSheet.OptimumRun_PurchaseSize = SRQ
        objClsCMS_STKA_CylinderSheet.MinimumRun_PurchaseSize = 1
        objClsCMS_STKA_CylinderSheet.SalesForecastTimeFence_inDays = 70
        objClsCMS_STKA_CylinderSheet.RepetitiveControl = "N"
    End Sub

    Private Sub STKA_TubeSheetLogics(ByVal objClsCMS_STKA_TubeSheet As clsCMS_STKA) 'STKA_Tube
        objClsCMS_STKA_TubeSheet.InternalPartNumber = strBoreCodeNumber
        objClsCMS_STKA_TubeSheet.ReplenishmentType = 1
        objClsCMS_STKA_TubeSheet.VendorLeadTime_TransferLeadTime_inDays = ""
        objClsCMS_STKA_TubeSheet.MinimumOrderQuantity_inUnitofIssue = 1
        objClsCMS_STKA_TubeSheet.BuyerCode = ""
        objClsCMS_STKA_TubeSheet.PlannerCode = "TC"

        'Sunny 22-06-10
        objClsCMS_STKA_TubeSheet.ReceiveToLocation = "BT5"    '06_09_2010   RAGAVA  "C01BT5"
        'objClsCMS_STKA_TubeSheet.InspectionProcedure = "PROTO"    '19_11_2010   ANUP

        'ANUP 19-11-2010 
        Try

            Dim strPartType_Purc As String = IFLConnectionObject.GetValue("Select Purchased_Manfractured from CostingDetails where PartCode = '" + strBoreCodeNumber + "'")
            If IsNothing(strPartType_Purc) OrElse strPartType_Purc = "M" Then
                objClsCMS_STKA_TubeSheet.InspectionProcedure = "PROTO"
            Else
                objClsCMS_STKA_TubeSheet.InspectionProcedure = ""
            End If

        Catch ex As Exception

        End Try
        'objClsCMS_STKA_TubeSheet.InspectionProcedure = ""  '18_11_2010   ANUPx
        '************

        objClsCMS_STKA_TubeSheet.MaterialPreparationLeadTime_inDays = 2
        objClsCMS_STKA_TubeSheet.CountryOfOrigin = "CA"
        objClsCMS_STKA_TubeSheet.ScheduleType = "B"
        objClsCMS_STKA_TubeSheet.OptimumRun_PurchaseSize = SRQ
        objClsCMS_STKA_TubeSheet.MinimumRun_PurchaseSize = 1
        objClsCMS_STKA_TubeSheet.SalesForecastTimeFence_inDays = 70
        objClsCMS_STKA_TubeSheet.RepetitiveControl = "Y"
    End Sub

    Private Sub STKA_RodSheetLogics(ByVal objClsCMS_STKA_RodSheet As clsCMS_STKA) 'STKA_Rod
        objClsCMS_STKA_RodSheet.InternalPartNumber = strRodCodeNumber
        objClsCMS_STKA_RodSheet.ReplenishmentType = 1
        objClsCMS_STKA_RodSheet.VendorLeadTime_TransferLeadTime_inDays = ""
        objClsCMS_STKA_RodSheet.MinimumOrderQuantity_inUnitofIssue = 1
        objClsCMS_STKA_RodSheet.BuyerCode = ""
        objClsCMS_STKA_RodSheet.PlannerCode = "TC"


        'ANUP 19-11-2010 
        Try

            Dim strPartType_Purc As String = IFLConnectionObject.GetValue("Select Purchased_Manfractured from CostingDetails where PartCode = '" + strRodCodeNumber + "'")
            If IsNothing(strPartType_Purc) OrElse strPartType_Purc = "M" Then
                objClsCMS_STKA_RodSheet.InspectionProcedure = "PROTO"
            Else
                objClsCMS_STKA_RodSheet.InspectionProcedure = ""
            End If

        Catch ex As Exception

        End Try



        'Sunny 22-06-10
        objClsCMS_STKA_RodSheet.ReceiveToLocation = "BT5"     '06_09_2010   RAGAVA    "C01BT5"


        'objClsCMS_STKA_RodSheet.InspectionProcedure = ""  '18_11_2010   ANUPx
        '************

        objClsCMS_STKA_RodSheet.MaterialPreparationLeadTime_inDays = 2
        objClsCMS_STKA_RodSheet.CountryOfOrigin = "CA"
        objClsCMS_STKA_RodSheet.ScheduleType = "B"
        objClsCMS_STKA_RodSheet.OptimumRun_PurchaseSize = SRQ
        objClsCMS_STKA_RodSheet.MinimumRun_PurchaseSize = 1
        objClsCMS_STKA_RodSheet.SalesForecastTimeFence_inDays = 70
        objClsCMS_STKA_RodSheet.RepetitiveControl = "Y"
    End Sub

    Private Sub STKA_TieRodSheetLogics(ByVal objClsCMS_STKA_TieRodSheet As clsCMS_STKA) 'STKA_TieRod
        objClsCMS_STKA_TieRodSheet.InternalPartNumber = strTieRodCodeNumber
        objClsCMS_STKA_TieRodSheet.ReplenishmentType = 2
        objClsCMS_STKA_TieRodSheet.VendorLeadTime_TransferLeadTime_inDays = ""
        objClsCMS_STKA_TieRodSheet.MinimumOrderQuantity_inUnitofIssue = ""
        'objClsCMS_STKA_TieRodSheet.BuyerCode = "TTP"
        objClsCMS_STKA_TieRodSheet.BuyerCode = "TPP"          '10_01_2011       RAGAVA
        objClsCMS_STKA_TieRodSheet.PlannerCode = ""

        'Sunny 22-06-10
        objClsCMS_STKA_TieRodSheet.ReceiveToLocation = "BT5"     '06_09_2010   RAGAVA   "C01BT5"

        '16_09_2010  RAGAVA
        If UCase(CustomerName).IndexOf("CNH") <> -1 Then
            objClsCMS_STKA_TieRodSheet.InspectionProcedure = "ISIR"
        Else
            objClsCMS_STKA_TieRodSheet.InspectionProcedure = "PROTO"   '19_11_2010   ANUP
            ' objClsCMS_STKA_TieRodSheet.InspectionProcedure = ""    '18_11_2010   ANUPx
        End If
        ' objClsCMS_STKA_TieRodSheet.InspectionProcedure = "ISIR"
        'Till  Here


        '************

        objClsCMS_STKA_TieRodSheet.MaterialPreparationLeadTime_inDays = ""
        objClsCMS_STKA_TieRodSheet.CountryOfOrigin = "US"
        objClsCMS_STKA_TieRodSheet.ScheduleType = "B"
        objClsCMS_STKA_TieRodSheet.OptimumRun_PurchaseSize = 1
        objClsCMS_STKA_TieRodSheet.MinimumRun_PurchaseSize = 0
        objClsCMS_STKA_TieRodSheet.SalesForecastTimeFence_inDays = 70
        objClsCMS_STKA_TieRodSheet.RepetitiveControl = "N"
    End Sub

    Private Sub STKA_StopTubeSheetLogics(ByVal objClsCMS_STKA_StopTubeSheet As clsCMS_STKA) 'STKA_StopTube
        objClsCMS_STKA_StopTubeSheet.InternalPartNumber = StopTubeCodeNumber
        objClsCMS_STKA_StopTubeSheet.ReplenishmentType = 2
        objClsCMS_STKA_StopTubeSheet.VendorLeadTime_TransferLeadTime_inDays = 120
        objClsCMS_STKA_StopTubeSheet.MinimumOrderQuantity_inUnitofIssue = ""
        objClsCMS_STKA_StopTubeSheet.BuyerCode = "NWP"
        objClsCMS_STKA_StopTubeSheet.PlannerCode = ""

        'Sunny 22-06-10
        objClsCMS_STKA_StopTubeSheet.ReceiveToLocation = "BS1"       '06_09_2010   RAGAVA    "C01BS1"

        '16_09_2010  RAGAVA
        If UCase(CustomerName).IndexOf("CNH") <> -1 Then
            objClsCMS_STKA_StopTubeSheet.InspectionProcedure = "ISIR"
        Else
            objClsCMS_STKA_StopTubeSheet.InspectionProcedure = "PROTO"  '19_11_2010   ANUP
            '  objClsCMS_STKA_StopTubeSheet.InspectionProcedure = ""    '18_11_2010   ANUPx
        End If
        'objClsCMS_STKA_StopTubeSheet.InspectionProcedure = "ISIR"
        'Till  Here

        '************

        objClsCMS_STKA_StopTubeSheet.MaterialPreparationLeadTime_inDays = ""
        objClsCMS_STKA_StopTubeSheet.CountryOfOrigin = "CN"
        objClsCMS_STKA_StopTubeSheet.ScheduleType = "B"
        objClsCMS_STKA_StopTubeSheet.OptimumRun_PurchaseSize = 1
        objClsCMS_STKA_StopTubeSheet.MinimumRun_PurchaseSize = 0
        objClsCMS_STKA_StopTubeSheet.SalesForecastTimeFence_inDays = 90
        objClsCMS_STKA_StopTubeSheet.RepetitiveControl = "N"
    End Sub

#End Region

#Region "METHDM"

    Private Sub METHDM_TieRodCylinder_Functionality()
        Try
            Dim objClsCMS_METHDM_TieRodCylinder As New clsCMS_METHDM
            objClsCMS_METHDM_TieRodCylinder.SetCommenPropertyValue()
            METHDM_TieRodCylinderSheetLogics(objClsCMS_METHDM_TieRodCylinder)
            _oExWorkbook.Save()

            Dim objClsCMS_METHDM_Tube As New clsCMS_METHDM
            objClsCMS_METHDM_Tube.SetCommenPropertyValue()
            METHDM_TubeSheetLogics(objClsCMS_METHDM_Tube)
            _oExWorkbook.Save()

            Dim objClsCMS_METHDM_Rod As New clsCMS_METHDM
            objClsCMS_METHDM_Rod.SetCommenPropertyValue()
            METHDM_RodSheetLogics(objClsCMS_METHDM_Rod)
            objClsCMS_METHDM_Rod.SetDataToExcel(_oExcelSheet_METHDM_Rod, 2)
            _oExWorkbook.Save()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub METHDM_TieRodCylinderSheetLogics(ByVal objClsCMS_METHDM_TieRodCylinder As clsCMS_METHDM) 'METHDM_Cylinder
        objClsCMS_METHDM_TieRodCylinder.PartNumber = CylinderCodeNumber

        Dim intCount As Integer = 2
        Dim intMETHDM_LineNumberCount As Integer = 10
        For Each oDataRow As DataRow In CostDetails_Costing.Rows

            If oDataRow("CodeNumber").Equals("469832") Then
                Continue For
            End If

            objClsCMS_METHDM_TieRodCylinder.LineNumber = intMETHDM_LineNumberCount

            Try
                Dim strChangedPartCode As String = GetPurchasedCode(oDataRow("CodeNumber"))
                If Not IsNothing(strChangedPartCode) Then
                    objClsCMS_METHDM_TieRodCylinder.MaterialPartNumber = strChangedPartCode
                Else
                    objClsCMS_METHDM_TieRodCylinder.MaterialPartNumber = oDataRow("CodeNumber")
                End If
            Catch ex As Exception
            End Try

            Try
                objClsCMS_METHDM_TieRodCylinder.StockType = IFLConnectionObject.GetValue("Select Purchased_Manfractured from CostingDetails where PartCode = '" + oDataRow("CodeNumber") + "'")
                If IsNothing(objClsCMS_METHDM_TieRodCylinder.StockType) OrElse objClsCMS_METHDM_TieRodCylinder.StockType = "M" Then
                    objClsCMS_METHDM_TieRodCylinder.StockType = "M"
                Else
                    objClsCMS_METHDM_TieRodCylinder.StockType = "R"
                End If
                '17_06_2011  RAGAVA
                If oDataRow("CodeNumber").ToString.StartsWith("7") = False Then
                    objClsCMS_METHDM_TieRodCylinder.StockType = "R"
                End If
                'TILL   HERE
            Catch ex As Exception
            End Try

            Try
                objClsCMS_METHDM_TieRodCylinder.QuantityPer = oDataRow("Quantity")
            Catch ex As Exception
            End Try

            Try
                objClsCMS_METHDM_TieRodCylinder.UnitofMeansure_forQuantityPer = oDataRow("Units")
            Catch ex As Exception
            End Try

            Try
                If oDataRow("Description").ToString.Contains("BAG PLASTIC") Then
                    objClsCMS_METHDM_TieRodCylinder.BlowThroughPart = 1
                Else
                    objClsCMS_METHDM_TieRodCylinder.BlowThroughPart = ""
                End If
            Catch ex As Exception
            End Try

            Try
                If oDataRow("Description").ToString.Contains("PAINT") OrElse oDataRow("CodeNumber") = "174040" OrElse oDataRow("Description").ToString.Contains("MASK") Then
                    objClsCMS_METHDM_TieRodCylinder.StockLocation = "C01BPC"
                    objClsCMS_METHDM_TieRodCylinder.SequenceNumber = "30"      '06_09_2010   RAGAVA
                Else
                    Dim strLableCodeNumber As String = IFLConnectionObject.GetValue("Select IFLID from LableDetails where PartCode = '" + oDataRow("CodeNumber") + "'")
                    If oDataRow("Description").ToString.Contains("PIN ") OrElse oDataRow("Description").ToString.Contains("LABEL") OrElse Not IsNothing(strLableCodeNumber) OrElse oDataRow("Description").ToString.Contains("PLASTIC") OrElse oDataRow("Description").ToString.Contains("DECAL ") Then    '26_11_2010   RAGAVA   "PIN " Condition added     '08_10_2010   RAGAVA
                        objClsCMS_METHDM_TieRodCylinder.StockLocation = "C01BPC"
                    Else
                        '14_12_2010   RAGAVA
                        'objClsCMS_METHDM_TieRodCylinder.StockLocation = "C01BT5"
                        If oDataRow("CodeNumber").ToString.StartsWith("19") = True Then
                            objClsCMS_METHDM_TieRodCylinder.StockLocation = "C01BH5"
                        Else
                            objClsCMS_METHDM_TieRodCylinder.StockLocation = "C01BT5"
                        End If
                        'Till   Here
                    End If

                    '06_09_2010   RAGAVA
                    If oDataRow("Description").ToString.Contains("BAG PLASTIC") OrElse oDataRow("Description").ToString.Contains("END KIT") OrElse oDataRow("Description").ToString.Contains("PIN ") OrElse oDataRow("Description").ToString.Contains("DECAL ") OrElse oDataRow("Description").ToString.Contains("LABEL") Then        '05_07_2011  RAGAVA  oDataRow("Description").ToString.Contains("END KIT")
                        objClsCMS_METHDM_TieRodCylinder.SequenceNumber = "30"
                    Else
                        objClsCMS_METHDM_TieRodCylinder.SequenceNumber = "10"
                    End If
                    'Till  Here
                End If
                objClsCMS_METHDM_TieRodCylinder.MaterialDescription = "" '13_09_2012   RAGAVA  oDataRow("Description").ToString          '17_08_2012   RAGAVA
            Catch ex As Exception
            End Try

            Try
                objClsCMS_METHDM_TieRodCylinder.ItemNumber = IFLConnectionObject.GetValue("Select ReferenceNumber from CostingDetails where PartCode = '" + oDataRow("CodeNumber") + "'")

                '19_07_2011   RAGAVA
                If objClsCMS_METHDM_TieRodCylinder.MaterialPartNumber = "235004" OrElse objClsCMS_METHDM_TieRodCylinder.MaterialPartNumber = "235012" _
                OrElse objClsCMS_METHDM_TieRodCylinder.MaterialPartNumber = "235005" OrElse objClsCMS_METHDM_TieRodCylinder.MaterialPartNumber = "235006" _
                OrElse objClsCMS_METHDM_TieRodCylinder.MaterialPartNumber = "235007" Then
                    objClsCMS_METHDM_TieRodCylinder.ItemNumber = "8"
                End If
                'Till  Here


                If objClsCMS_METHDM_TieRodCylinder.ItemNumber = "" Then
                    If oDataRow("Description").ToString.Contains("Tube") Then
                        objClsCMS_METHDM_TieRodCylinder.ItemNumber = 5
                    ElseIf oDataRow("Description").ToString.Contains("Tie Rod") Then
                        objClsCMS_METHDM_TieRodCylinder.ItemNumber = 6
                    ElseIf oDataRow("Description").ToString.Contains("Rod") Then
                        objClsCMS_METHDM_TieRodCylinder.ItemNumber = 4
                    Else
                        objClsCMS_METHDM_TieRodCylinder.ItemNumber = 0
                    End If
                End If
            Catch ex As Exception
            End Try

            objClsCMS_METHDM_TieRodCylinder.SetDataToExcel(_oExcelSheet_METHDM_TieRodCylinder, intCount)
            _oExWorkbook.Save()
            intCount += 1
            intMETHDM_LineNumberCount += 10
        Next


    End Sub

    Private Sub METHDM_TubeSheetLogics(ByVal objClsCMS_METHDM_Tube As clsCMS_METHDM) 'METHDM_Tube
        objClsCMS_METHDM_Tube.PartNumber = strBoreCodeNumber
        objClsCMS_METHDM_Tube.LineNumber = 10

        Try
            Dim strChangedPartCode As String = GetPurchasedCode(TubeMaterialCode_Costing)
            If Not IsNothing(strChangedPartCode) Then
                objClsCMS_METHDM_Tube.MaterialPartNumber = strChangedPartCode
            Else
                objClsCMS_METHDM_Tube.MaterialPartNumber = TubeMaterialCode_Costing
            End If
        Catch ex As Exception
        End Try

        Try
            objClsCMS_METHDM_Tube.StockType = IFLConnectionObject.GetValue("Select Purchased_Manfractured from CostingDetails where PartCode = '" + TubeMaterialCode_Costing + "'")
            If IsNothing(objClsCMS_METHDM_Tube.StockType) OrElse objClsCMS_METHDM_Tube.StockType = "M" Then
                objClsCMS_METHDM_Tube.StockType = "M"
            Else
                objClsCMS_METHDM_Tube.StockType = "R"
            End If
        Catch ex As Exception
        End Try

        Dim dblTubeLength As Double = Math.Ceiling(TubeLength)
        objClsCMS_METHDM_Tube.QuantityPer = Math.Round((dblTubeLength + (3 / 8)) / 12, 4)        '14_09_2010    RAGAVA
        objClsCMS_METHDM_Tube.UnitofMeansure_forQuantityPer = "FT"
        If TubeMaterial1.StartsWith("@") = True Then
            objClsCMS_METHDM_Tube.StockLocation = "C01BSB"
        Else
            objClsCMS_METHDM_Tube.StockLocation = "C01BT5"
        End If
        objClsCMS_METHDM_Tube.ItemNumber = 0
        objClsCMS_METHDM_Tube.SetDataToExcel(_oExcelSheet_METHDM_Tube, 2)



        If SeriesForCosting.Contains("TP") Then
            objClsCMS_METHDM_Tube.LineNumber = 20
            objClsCMS_METHDM_Tube.MaterialPartNumber = 469832
            Try
                objClsCMS_METHDM_Tube.StockType = IFLConnectionObject.GetValue("Select Purchased_Manfractured from CostingDetails where PartCode = '469832'")
                If IsNothing(objClsCMS_METHDM_Tube.StockType) OrElse objClsCMS_METHDM_Tube.StockType = "M" Then
                    objClsCMS_METHDM_Tube.StockType = "M"
                Else
                    objClsCMS_METHDM_Tube.StockType = "R"
                End If
            Catch ex As Exception
            End Try

            If strRephasing.Contains("Both") Then
                objClsCMS_METHDM_Tube.QuantityPer = 2
            Else
                objClsCMS_METHDM_Tube.QuantityPer = 1
            End If

            objClsCMS_METHDM_Tube.UnitofMeansure_forQuantityPer = "EA"
            'objClsCMS_METHDM_Tube.StockLocation = "C01BSB"
            objClsCMS_METHDM_Tube.StockLocation = "C01BT5"       '08_10_2010  RAGAVA
            objClsCMS_METHDM_Tube.ItemNumber = 0
            objClsCMS_METHDM_Tube.SetDataToExcel(_oExcelSheet_METHDM_Tube, 3)
        End If

    End Sub

    Private Sub METHDM_RodSheetLogics(ByVal objClsCMS_METHDM_Rod As clsCMS_METHDM) 'METHDM_Rod
        objClsCMS_METHDM_Rod.PartNumber = strRodCodeNumber
        objClsCMS_METHDM_Rod.LineNumber = 10

        Try
            Dim strChangedPartCode As String = GetPurchasedCode(RodMaterialCode_Costing)
            If Not IsNothing(strChangedPartCode) Then
                objClsCMS_METHDM_Rod.MaterialPartNumber = strChangedPartCode
            Else
                objClsCMS_METHDM_Rod.MaterialPartNumber = RodMaterialCode_Costing
            End If
        Catch ex As Exception
        End Try

        Try
            objClsCMS_METHDM_Rod.StockType = IFLConnectionObject.GetValue("Select Purchased_Manfractured from CostingDetails where PartCode = '" + RodMaterialCode_Costing + "'")
            If IsNothing(objClsCMS_METHDM_Rod.StockType) OrElse objClsCMS_METHDM_Rod.StockType = "M" Then
                objClsCMS_METHDM_Rod.StockType = "M"
            Else
                objClsCMS_METHDM_Rod.StockType = "R"
            End If
        Catch ex As Exception
        End Try

        Try
            Dim dblRodWeightPerFoot As Double = IFLConnectionObject.GetValue("select WeightPerFoot from RodWeightDetails where RodDiameter = " + RodDiameter.ToString)
            objClsCMS_METHDM_Rod.QuantityPer = Math.Round((RodLength + 0.25) * (dblRodWeightPerFoot / 12), 4)     '14_09_2010   RAGAVA
        Catch ex As Exception
        End Try

        objClsCMS_METHDM_Rod.UnitofMeansure_forQuantityPer = "LB"
        objClsCMS_METHDM_Rod.StockLocation = "C01BSB"
        objClsCMS_METHDM_Rod.ItemNumber = 0
    End Sub

    Private Function GetRodWeight() As Double
        Try

        Catch ex As Exception
            GetRodWeight = 0
        End Try
    End Function

    Private Function GetPurchasedCode(ByVal strPartCode As String) As String
        GetPurchasedCode = Nothing
        Try
            GetPurchasedCode = IFLConnectionObject.GetValue("Select PurchasePartCode from CostingDetails where PartCode = '" _
                               + strPartCode + "' and Purchased_Manfractured = 'P' and PurchasePartCode <> ''") 'When part is changed from Manu to Pur
        Catch ex As Exception
            GetPurchasedCode = Nothing
        End Try
    End Function

#End Region

#Region "METHDR"

    Private Sub METHDR_TieRodCylinder_Functionality()

        Try
            Dim objClsCMS_METHDR_TieRodCylinder As New clsCMS_METHDR
            objClsCMS_METHDR_TieRodCylinder.SetCommenPropertyValue()
            METHDR_TieRodCylinderSheetLogics(objClsCMS_METHDR_TieRodCylinder)
        Catch ex As Exception
        End Try

        Try
            Dim objClsCMS_METHDR_Tube As New clsCMS_METHDR
            objClsCMS_METHDR_Tube.SetCommenPropertyValue()
            METHDR_TubeSheetLogics(objClsCMS_METHDR_Tube)
        Catch ex As Exception
        End Try

        Try
            Dim objClsCMS_METHDR_Rod As New clsCMS_METHDR
            objClsCMS_METHDR_Rod.SetCommenPropertyValue()
            METHDR_RodSheetLogics(objClsCMS_METHDR_Rod)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub METHDR_TieRodCylinderSheetLogics(ByVal objClsCMS_METHDR_TieRodCylinder As clsCMS_METHDR) 'METHDR CYLINDER
        objClsCMS_METHDR_TieRodCylinder.PartNumber = CylinderCodeNumber
        objClsCMS_METHDR_TieRodCylinder.SetupStandard = 0

        Dim intCylinderSeq As Integer = 10
        For intCount As Integer = 2 To 4

            '07_09_2010   RAGAVA
            If intCount = 3 AndAlso _ofrmTieRod3.chk100OilTest.Checked = False Then
                Continue For
            End If
            'Till   Here

            If intCount = 4 Then
                objClsCMS_METHDR_TieRodCylinder.Seq = 30
                objClsCMS_METHDR_TieRodCylinder.Department = objClsCMS_METHDR_TieRodCylinder.PlantCode + "06"
                objClsCMS_METHDR_TieRodCylinder.Resource = "WC631"
                objClsCMS_METHDR_TieRodCylinder.Operation = "P02"
                objClsCMS_METHDR_TieRodCylinder.ScheduleRunStandard = METHDRPaintRunStandard
            Else
                objClsCMS_METHDR_TieRodCylinder.Seq = intCylinderSeq
                objClsCMS_METHDR_TieRodCylinder.Department = objClsCMS_METHDR_TieRodCylinder.PlantCode + "18"
                If intCount = 2 Then
                    objClsCMS_METHDR_TieRodCylinder.Resource = METHDRAssemblyResource
                    objClsCMS_METHDR_TieRodCylinder.Operation = "A01"
                    objClsCMS_METHDR_TieRodCylinder.ScheduleRunStandard = METHDRAssemblyRunStandard
                Else
                    objClsCMS_METHDR_TieRodCylinder.Resource = "WC626"     '05_10_2010   RAGAVA   'WC625
                    objClsCMS_METHDR_TieRodCylinder.Operation = "T01"
                    '06_10_2010   RAGAVA
                    If UCase(CustomerName).IndexOf("CNH") <> -1 Then
                        If BoreDiameter <= 3 Then
                            objClsCMS_METHDR_TieRodCylinder.ScheduleRunStandard = 12.63
                        ElseIf BoreDiameter > 3 AndAlso BoreDiameter <= 4 Then
                            objClsCMS_METHDR_TieRodCylinder.ScheduleRunStandard = 11.34
                        ElseIf BoreDiameter > 4 Then
                            objClsCMS_METHDR_TieRodCylinder.ScheduleRunStandard = 9.3
                        End If
                    Else
                        If BoreDiameter <= 3 Then
                            objClsCMS_METHDR_TieRodCylinder.ScheduleRunStandard = 16
                        ElseIf BoreDiameter > 3 AndAlso BoreDiameter <= 4 Then
                            objClsCMS_METHDR_TieRodCylinder.ScheduleRunStandard = 14
                        ElseIf BoreDiameter > 4 Then
                            objClsCMS_METHDR_TieRodCylinder.ScheduleRunStandard = 11
                        End If
                    End If
                    'objClsCMS_METHDR_TieRodCylinder.ScheduleRunStandard = 16
                    'Till   Here
                End If
            End If


            '18_02_2011   RAGAVA
            Dim iMen As Integer = 1
            'If intCount = 4 Then      '01_05_2012  RAGAVA Commented    '31_10_2011   RAGAVA
            Try
                Dim strSql As String = "Select NumberOfMen from MIL_WELDED.dbo.WorkCenter_MenMachineDetails where WorkCenter = '" & objClsCMS_METHDR_TieRodCylinder.Resource.ToString.Substring(2, 3) & "'"
                iMen = IFLConnectionObject.GetValue(strSql)
                '16_06_2011   RAGAVA
                If iMen <= 0 Then
                    iMen = 1
                End If
                'TILL   HERE
                objClsCMS_METHDR_TieRodCylinder.OfMen = iMen 'anup 16-03-2011
                objClsCMS_METHDR_TieRodCylinder.ScheduleRunStandard = objClsCMS_METHDR_TieRodCylinder.ScheduleRunStandard * iMen
            Catch ex As Exception
            End Try
            'End If
            'Till   Here


            objClsCMS_METHDR_TieRodCylinder.CostingRunStandard = objClsCMS_METHDR_TieRodCylinder.ScheduleRunStandard
            '16_09_2010    RAGAVA
            If intCount = 4 AndAlso _ofrmTieRod3.chk100OilTest.Checked = False Then
                objClsCMS_METHDR_TieRodCylinder.SetDataToExcel(_oExcelSheet_METHDR_TieRodCylinder, intCount - 1)
            Else
                objClsCMS_METHDR_TieRodCylinder.SetDataToExcel(_oExcelSheet_METHDR_TieRodCylinder, intCount)
            End If
            'objClsCMS_METHDR_TieRodCylinder.SetDataToExcel(_oExcelSheet_METHDR_TieRodCylinder, intCount)
            'Till   Here
            _oExWorkbook.Save()
            intCylinderSeq += 10
        Next
    End Sub

    Private Sub METHDR_TubeSheetLogics(ByVal objClsCMS_METHDR_Tube As clsCMS_METHDR) 'METHDR TUBE
        objClsCMS_METHDR_Tube.PartNumber = strBoreCodeNumber
        objClsCMS_METHDR_Tube.Department = objClsCMS_METHDR_Tube.PlantCode + "17"
        objClsCMS_METHDR_Tube.Operation = "M01"
        objClsCMS_METHDR_Tube.SetupStandard = 0.5        '01_11_2010   ANUP          ' 0.25

        Dim dblWC099_RunStandard As Double = 0
        Dim dblWC087_RunStandard As Double = 0
        GetTubeRunStandardValues(dblWC099_RunStandard, dblWC087_RunStandard)
        Dim inTubeSeq As Integer = 10
        For intCount As Integer = 2 To 5

            'anup 17-03-2011 start
            If intCount > 3 AndAlso Not SeriesForCosting.IndexOf("TP-") <> -1 Then
                Exit For
            End If
            'anup 17-03-2011 till here

            objClsCMS_METHDR_Tube.Seq = inTubeSeq

            If intCount = 2 Then
                objClsCMS_METHDR_Tube.Resource = "WC099"
                objClsCMS_METHDR_Tube.ScheduleRunStandard = dblWC099_RunStandard
            ElseIf intCount = 3 Then
                objClsCMS_METHDR_Tube.Resource = "WC087"
                objClsCMS_METHDR_Tube.ScheduleRunStandard = dblWC087_RunStandard
            ElseIf SeriesForCosting.IndexOf("TP-") <> -1 Then
                If intCount = 4 Then
                    objClsCMS_METHDR_Tube.Resource = "WC136"
                    objClsCMS_METHDR_Tube.ScheduleRunStandard = "125"
                ElseIf intCount = 5 Then
                    objClsCMS_METHDR_Tube.Resource = "WC136"
                    objClsCMS_METHDR_Tube.ScheduleRunStandard = "30"
                End If
            End If

            'anup 16-03-2011 start
            Dim iMen As Integer = 1
            Try
                Dim strSql As String = "Select NumberOfMen from MIL_WELDED.dbo.WorkCenter_MenMachineDetails where WorkCenter = '" & objClsCMS_METHDR_Tube.Resource.ToString.Substring(2, 3) & "'"
                iMen = IFLConnectionObject.GetValue(strSql)
                If iMen = 0 Then
                    iMen = 1
                End If
                objClsCMS_METHDR_Tube.OfMen = iMen
                objClsCMS_METHDR_Tube.ScheduleRunStandard = objClsCMS_METHDR_Tube.ScheduleRunStandard * iMen
            Catch ex As Exception
            End Try
            'anup 16-03-2011 till here

            objClsCMS_METHDR_Tube.CostingRunStandard = objClsCMS_METHDR_Tube.ScheduleRunStandard

            objClsCMS_METHDR_Tube.SetDataToExcel(_oExcelSheet_METHDR_Tube, intCount)
            _oExWorkbook.Save()
            inTubeSeq += 10
        Next
    End Sub

    Private Sub METHDR_RodSheetLogics(ByVal objClsCMS_METHDR_Rod As clsCMS_METHDR) 'METHDR ROD
        objClsCMS_METHDR_Rod.PartNumber = strRodCodeNumber
        objClsCMS_METHDR_Rod.Department = objClsCMS_METHDR_Rod.PlantCode + "17"
        objClsCMS_METHDR_Rod.Operation = "M01"

        Dim dblWC083_RunStandard As Double = 0
        Dim dblWCNumberValue As Double = 0
        GetRodRunStandardValues(dblWC083_RunStandard, dblWCNumberValue)

        Dim inRodSeq As Integer = 10
        For intCount As Integer = 2 To 3
            objClsCMS_METHDR_Rod.Seq = inRodSeq

            If intCount = 2 Then
                objClsCMS_METHDR_Rod.Resource = "WC083"
                objClsCMS_METHDR_Rod.SetupStandard = 0.3
                objClsCMS_METHDR_Rod.ScheduleRunStandard = dblWC083_RunStandard
            Else
                objClsCMS_METHDR_Rod.Resource = GetRodMachiningWorkCenter()
                objClsCMS_METHDR_Rod.SetupStandard = 0.25
                objClsCMS_METHDR_Rod.ScheduleRunStandard = dblWCNumberValue
            End If

            'anup 16-03-2011 start
            Dim iMen As Integer = 1
            Try
                Dim strSql As String = "Select NumberOfMen from MIL_WELDED.dbo.WorkCenter_MenMachineDetails where WorkCenter = '" & objClsCMS_METHDR_Rod.Resource.ToString.Substring(2, 3) & "'"
                iMen = IFLConnectionObject.GetValue(strSql)
                If iMen = 0 Then
                    iMen = 1
                End If
                objClsCMS_METHDR_Rod.OfMen = iMen
                objClsCMS_METHDR_Rod.ScheduleRunStandard = objClsCMS_METHDR_Rod.ScheduleRunStandard * iMen
            Catch ex As Exception
            End Try
            'anup 16-03-2011 till here

            objClsCMS_METHDR_Rod.CostingRunStandard = objClsCMS_METHDR_Rod.ScheduleRunStandard

            objClsCMS_METHDR_Rod.SetDataToExcel(_oExcelSheet_METHDR_Rod, intCount)
            _oExWorkbook.Save()
            inRodSeq += 10
        Next
    End Sub

    Private Function GetRodMachiningWorkCenter() As String
        Dim strRodDiameterColumn As String = ""
        Dim strWCNumber As String = ""
        Dim dblCostRodLength As Double = Math.Ceiling(RodLength)

        If RodDiameter = 1.12 Then
            strRodDiameterColumn = "BoreDiameter_1_12"
        ElseIf RodDiameter = 1.25 Then
            strRodDiameterColumn = "BoreDiameter_1_25"
        ElseIf RodDiameter = 1.38 Then
            strRodDiameterColumn = "BoreDiameter_1_38"
        ElseIf RodDiameter = 1.5 Then
            strRodDiameterColumn = "BoreDiameter_1_5"
        ElseIf RodDiameter = 1.75 Then
            strRodDiameterColumn = "BoreDiameter_1_75"
        ElseIf RodDiameter = 2 Then
            strRodDiameterColumn = "BoreDiameter_2"
        End If

        If RodMaterialForCosting = "Chrome" Or UCase(RodMaterialForCosting).IndexOf("LION") <> -1 Then    '06_09_2010   RAGAVA   Lion Condition Added
            Dim strQuery3 As String = "Select " + strRodDiameterColumn + " from TRChromeRodMachiningWCDetails where TubeLength =" + dblCostRodLength.ToString
            Try
                strWCNumber = IFLConnectionObject.GetValue(strQuery3)
            Catch ex As Exception
                strWCNumber = ""
            End Try
        ElseIf RodMaterialForCosting = "Nitro Steel" Then
            Dim strQuery3 As String = "Select " + strRodDiameterColumn + " from TRNitroRodMachiningWCDetails where TubeLength =" + dblCostRodLength.ToString
            Try
                strWCNumber = IFLConnectionObject.GetValue(strQuery3)
            Catch ex As Exception
                strWCNumber = ""
            End Try
        End If

        Return strWCNumber
    End Function

    Private Sub GetTubeRunStandardValues(ByRef dblWC099_RunStandard As Double, ByRef dblWC087_RunStandard As Double)
        Dim dblTubeLength As Double = Math.Ceiling(TubeLength)
        Dim strBoreDiameterColumn_TubeCode As String = ""

        Dim dblBoreDiamteter_TubeCode As Double = Val(ofrmTieRod1.cmbBore.Text)
        If dblBoreDiamteter_TubeCode = 2 Then
            strBoreDiameterColumn_TubeCode = "BoreDiameter_2"
        ElseIf dblBoreDiamteter_TubeCode = 2.25 Then
            strBoreDiameterColumn_TubeCode = "BoreDiameter_2_25"
        ElseIf dblBoreDiamteter_TubeCode = 2.5 Then
            strBoreDiameterColumn_TubeCode = "BoreDiameter_2_5"
        ElseIf dblBoreDiamteter_TubeCode = 2.75 Then
            strBoreDiameterColumn_TubeCode = "BoreDiameter_2_75"
        ElseIf dblBoreDiamteter_TubeCode = 3 Then
            strBoreDiameterColumn_TubeCode = "BoreDiameter_3"
        ElseIf dblBoreDiamteter_TubeCode = 3.25 Then
            strBoreDiameterColumn_TubeCode = "BoreDiameter_3_25"
        ElseIf dblBoreDiamteter_TubeCode = 3.5 Then
            strBoreDiameterColumn_TubeCode = "BoreDiameter_3_5"
        ElseIf dblBoreDiamteter_TubeCode = 3.75 Then
            strBoreDiameterColumn_TubeCode = "BoreDiameter_3_75"
        ElseIf dblBoreDiamteter_TubeCode = 4 Then
            strBoreDiameterColumn_TubeCode = "BoreDiameter_4"
        ElseIf dblBoreDiamteter_TubeCode = 4.25 Then
            strBoreDiameterColumn_TubeCode = "BoreDiameter_4_25"
        ElseIf dblBoreDiamteter_TubeCode = 4.5 Then
            strBoreDiameterColumn_TubeCode = "BoreDiameter_4_5"
        ElseIf dblBoreDiamteter_TubeCode = 4.75 Then
            strBoreDiameterColumn_TubeCode = "BoreDiameter_4_75"
        ElseIf dblBoreDiamteter_TubeCode = 5 Then
            strBoreDiameterColumn_TubeCode = "BoreDiameter_5"
        End If

        Dim strQuery1 As String = "Select " + strBoreDiameterColumn_TubeCode + " from TubeCutDetails where TubeLength =" + dblTubeLength.ToString
        Try
            dblWC099_RunStandard = IFLConnectionObject.GetValue(strQuery1)
        Catch ex As Exception
            dblWC099_RunStandard = 0
        End Try

        Dim strQuery2 As String = "Select " + strBoreDiameterColumn_TubeCode + " from TubeSkiveDetails where TubeLength =" + dblTubeLength.ToString
        Try
            dblWC087_RunStandard = IFLConnectionObject.GetValue(strQuery2)
        Catch ex As Exception
            dblWC087_RunStandard = 0
        End Try
    End Sub

    Private Sub GetRodRunStandardValues(ByRef dblWC083_RunStandard As Double, ByRef dblWCNumberValue1 As Double)

        'ANUP 15-12-2010 START
        Dim oCostingAndCMSCommon As New clsCostingAnsCMSCommon
        'ANUP 15-12-2010 TILL HERE

        Dim strRodDiameterColumn As String = ""
        If RodDiameter = 1.12 Then
            strRodDiameterColumn = "BoreDiameter_1_12"
        ElseIf RodDiameter = 1.25 Then
            strRodDiameterColumn = "BoreDiameter_1_25"
        ElseIf RodDiameter = 1.38 Then
            strRodDiameterColumn = "BoreDiameter_1_38"
        ElseIf RodDiameter = 1.5 Then
            strRodDiameterColumn = "BoreDiameter_1_5"
        ElseIf RodDiameter = 1.75 Then
            strRodDiameterColumn = "BoreDiameter_1_75"
        ElseIf RodDiameter = 2 Then
            strRodDiameterColumn = "BoreDiameter_2"
        End If

        Dim dblCostRodLength As Double = Math.Ceiling(RodLength)
        If RodMaterialForCosting = "Chrome" Or UCase(RodMaterialForCosting).IndexOf("LION") <> -1 Then    '06_09_2010   RAGAVA   Lion Condition Added

            Dim strQuery1 As String = "Select " + strRodDiameterColumn + " from TRChromeRodCuttingDetails where TubeLength =" + dblCostRodLength.ToString
            Try
                dblWC083_RunStandard = IFLConnectionObject.GetValue(strQuery1)
            Catch ex As Exception
                dblWC083_RunStandard = 0
            End Try

            'ANUP 15-12-2010 START
            'Chrome Machining Details
            Dim strQuery2 As String = "Select " + strRodDiameterColumn + " from " + oCostingAndCMSCommon.GetChromeMachiningTableName() + " where TubeLength =" + dblCostRodLength.ToString
            'ANUP 15-12-2010 TILL HERE
            Try
                dblWCNumberValue1 = IFLConnectionObject.GetValue(strQuery2)
            Catch ex As Exception
                dblWCNumberValue1 = 0
            End Try

        ElseIf RodMaterialForCosting = "Nitro Steel" Then

            Dim strQuery1 As String = "Select " + strRodDiameterColumn + " from TRNitroRodCuttingDetails where TubeLength =" + dblCostRodLength.ToString
            Try
                dblWC083_RunStandard = IFLConnectionObject.GetValue(strQuery1)
            Catch ex As Exception
                dblWC083_RunStandard = 0
            End Try

            'ANUP 15-12-2010 START
            Dim strQuery2 As String = "Select " + strRodDiameterColumn + " from " + oCostingAndCMSCommon.GetNitroRodMachiningTableName() + " where TubeLength =" + dblCostRodLength.ToString
            'ANUP 15-12-2010 TILL HERE

            Try
                dblWCNumberValue1 = IFLConnectionObject.GetValue(strQuery2)
            Catch ex As Exception
                dblWCNumberValue1 = 0
            End Try

        ElseIf RodMaterialForCosting = "Induction Hardened" Then
            dblWC083_RunStandard = 0
            'ANUP 15-12-2010 START
            '  dblWCNumberValue1 = 0

            Dim strQuery2 As String = "Select " + strRodDiameterColumn + " from " + oCostingAndCMSCommon.GetInductionHBMachiningTableName() + " where TubeLength =" + dblCostRodLength.ToString
            Try
                dblWCNumberValue1 = IFLConnectionObject.GetValue(strQuery2)
            Catch ex As Exception
                dblWCNumberValue1 = 0
            End Try
            'ANUP 15-12-2010 TILL HERE
        End If

    End Sub

#End Region

#Region "MTHL"

    Private Sub MTHL_TieRodCylinder_Functionality()

        Try
            Dim objClsCMS_MTHL_Tube As New clsCMS_MTHL
            objClsCMS_MTHL_Tube.SetCommenPropertyValue()
            MTHL_TubeSheetLogics(objClsCMS_MTHL_Tube)
        Catch ex As Exception

        End Try

        Try
            Dim objClsCMS_MTHL_Rod As New clsCMS_MTHL
            objClsCMS_MTHL_Rod.SetCommenPropertyValue()
            MTHL_RodSheetLogics(objClsCMS_MTHL_Rod)
        Catch ex As Exception

        End Try

        '16_09_2010   RAGAVA
        Try
            'If _ofrmTieRod3.chk100OilTest.Checked = True Then           '26_11_2010   RAGAVA  Commented
            Dim objClsCMS_MTHL_TieRod As New clsCMS_MTHL
            objClsCMS_MTHL_TieRod.SetCommenPropertyValue()
            MTHL_TieRodCylinderSheetLogics(objClsCMS_MTHL_TieRod)
            'End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MTHL_TubeSheetLogics(ByVal objClsCMS_MTHL_Tube As clsCMS_MTHL) 'METHDR TUBE
        Try
            objClsCMS_MTHL_Tube.Part = strBoreCodeNumber
            Dim intLineCount As Integer = 1
            '04_10_2010   RAGAVA
            Dim iCount As Integer = 9
            Dim iweightcount As Integer = 0
            If SeriesForCosting.IndexOf("TP-") <> -1 Then
                iCount = 13
            End If
            For intCount As Integer = 2 To iCount
                objClsCMS_MTHL_Tube.Tool = ""
                If intCount = 7 Then
                    intLineCount = 0
                End If
                If intCount = 8 Then
                    objClsCMS_MTHL_Tube.Seq = 20
                    objClsCMS_MTHL_Tube.Line = intLineCount + 1
                    'objClsCMS_MTHL_Tube.Tool = "WI10-E-24"
                    objClsCMS_MTHL_Tube.Tool = "WI09-E-78"        '10_12_2010   RAGAVA
                ElseIf intCount = 9 Then
                    If Weight_Bore >= 40 Then
                        objClsCMS_MTHL_Tube.Seq = 20
                        objClsCMS_MTHL_Tube.Line = intLineCount + 1
                        objClsCMS_MTHL_Tube.Tool = "CAUTION WEIGHT"
                    Else
                        iweightcount = iweightcount + 1
                    End If
                ElseIf intCount > 9 Then
                    objClsCMS_MTHL_Tube.Seq = 30
                    objClsCMS_MTHL_Tube.Line = intLineCount + 1
                    If intCount = 10 Then
                        objClsCMS_MTHL_Tube.Tool = "WI09-E-95"
                    ElseIf intCount = 11 Then
                        objClsCMS_MTHL_Tube.Tool = "299776"
                    ElseIf intCount = 12 Then
                        objClsCMS_MTHL_Tube.Tool = IFLConnectionObject.GetValue("Select CamNumber from CMS_Gauge_Cam_Details where BoreDiameter = " + BoreDiameter.ToString)
                    ElseIf intCount = 13 Then
                        objClsCMS_MTHL_Tube.Tool = IFLConnectionObject.GetValue("Select GaugeNumber from CMS_Gauge_Cam_Details where BoreDiameter = " + BoreDiameter.ToString)
                    End If
                Else
                    objClsCMS_MTHL_Tube.Seq = 10
                    objClsCMS_MTHL_Tube.Line = intLineCount
                    intLineCount += 1

                    If intCount = 2 Then
                        objClsCMS_MTHL_Tube.Tool = "WI09-E-99"
                    ElseIf intCount = 3 Then
                        objClsCMS_MTHL_Tube.Tool = "275127"
                    ElseIf intCount = 4 Then
                        objClsCMS_MTHL_Tube.Tool = "275128"
                    ElseIf intCount = 5 Then
                        objClsCMS_MTHL_Tube.Tool = "275129"
                    ElseIf intCount = 6 Then
                        Dim strQuery As String
                        If SeriesForCosting.IndexOf("TX") <> -1 Then
                            strQuery = "select top 1 TX_ProgramCode from TubeTooling_ProgramList where BoreDiameter >= " & BoreDiameter.ToString
                        Else
                            strQuery = "select top 1 TR_ProgramCode from TubeTooling_ProgramList where BoreDiameter >= " & BoreDiameter.ToString
                        End If
                        objClsCMS_MTHL_Tube.Tool = IFLConnectionObject.GetValue(strQuery)
                    ElseIf intCount = 7 Then
                        If Weight_Bore >= 40 Then
                            objClsCMS_MTHL_Tube.Tool = "CAUTION WEIGHT"
                        Else
                            iweightcount = iweightcount + 1
                        End If
                    End If
                End If
                objClsCMS_MTHL_Tube.SetDataToExcel(_oExcelSheet_MTHL_Tube, intCount - iweightcount)
                _oExWorkbook.Save()
            Next
        Catch ex As Exception

        End Try
    End Sub

    '16_09_2010   RAGAVA
    Private Sub MTHL_TieRodCylinderSheetLogics(ByVal objClsCMS_MTHL_TieRodCylinder As clsCMS_MTHL) 'MTHL

        Dim iweightcount As Integer = 0
        objClsCMS_MTHL_TieRodCylinder.Part = CylinderCodeNumber

        Dim intLineCount As Integer = 1
        Dim iColCount As Integer = 13          '06_07_2011   RAGAVA
        If _ofrmTieRod3.chk100OilTest.Checked = True AndAlso (SeriesForCosting.Equals("TP-High") OrElse SeriesForCosting.Equals("TP-Low")) Then
            iColCount = 16        '28_08_2012  RAGAVA
        ElseIf _ofrmTieRod3.chk100OilTest.Checked = True Then
            iColCount = 16        '28_08_2012  RAGAVA
        End If
        objClsCMS_MTHL_TieRodCylinder.Seq = 10

        For intCount As Integer = 2 To iColCount
            objClsCMS_MTHL_TieRodCylinder.Tool = ""       '04_10_2010   RAGAVA
            If intCount = 6 Then
                objClsCMS_MTHL_TieRodCylinder.Seq = 20
                intLineCount = 1
            ElseIf intCount = 12 Then
                objClsCMS_MTHL_TieRodCylinder.Seq = 30
                intLineCount = 1
            End If
            'vamsi commented 28th August 2013 start

            'If intCount > 5 AndAlso intCount < 12 Then 
            '    If _ofrmTieRod3.chk100OilTest.Checked = False Then        '04_10_2010   RAGAVA
            '        Continue For
            '    ElseIf Not (SeriesForCosting.Equals("TP-High") OrElse SeriesForCosting.Equals("TP-Low")) Then
            '        If intCount > 7 AndAlso intCount < 10 Then
            '            Continue For
            '        End If
            '    End If
            'End If 
            'till here
            objClsCMS_MTHL_TieRodCylinder.Line = intLineCount
            intLineCount += 1
            '06_10_2010   RAGAVA
            If intCount = 2 Then
                objClsCMS_MTHL_TieRodCylinder.Tool = "WI09-E-36"
            ElseIf intCount = 3 Then
                If Weight_Assembly > 40 Then
                    objClsCMS_MTHL_TieRodCylinder.Tool = "CAUTION WEIGHT"
                Else
                    iweightcount = iweightcount + 1
                End If
                'Till  Here
                '06_09_2012   RAGAVA
                '275139 and 275141
            ElseIf intCount = 4 Then
                If SeriesForCosting.Equals("TP-High") OrElse SeriesForCosting.Equals("TP-Low") Then
                    objClsCMS_MTHL_TieRodCylinder.Tool = "275139"
                ElseIf SeriesForCosting.IndexOf("TX") <> -1 Then
                    objClsCMS_MTHL_TieRodCylinder.Tool = "275172"
                Else
                    objClsCMS_MTHL_TieRodCylinder.Tool = "275139"
                End If
            ElseIf intCount = 5 Then
                If SeriesForCosting.Equals("TP-High") OrElse SeriesForCosting.Equals("TP-Low") Then
                    objClsCMS_MTHL_TieRodCylinder.Tool = "275141"
                Else
                    iweightcount = iweightcount + 1
                End If
            ElseIf intCount = 6 Then
                objClsCMS_MTHL_TieRodCylinder.Tool = "WI10-E-03"
            ElseIf intCount = 7 Then
                objClsCMS_MTHL_TieRodCylinder.Tool = "CYL TEST PRESSURE 2"
            ElseIf intCount = 8 Then
                If _ofrmTieRod3.chk100OilTest.Checked = True Then
                    If Trim(ofrmTieRod1.cmbRephasingPortPosition.Text).IndexOf("At Both") <> -1 Then
                        objClsCMS_MTHL_TieRodCylinder.Tool = "CYLINDER STYLE 4"
                        '04_10_2010   RAGAVA
                    ElseIf Trim(ofrmTieRod1.cmbRephasingPortPosition.Text).IndexOf("At Extension") <> -1 OrElse Trim(ofrmTieRod1.cmbRephasingPortPosition.Text).IndexOf("At Retraction") <> -1 Then
                        objClsCMS_MTHL_TieRodCylinder.Tool = "CYLINDER STYLE 3"
                    Else
                        iweightcount = iweightcount + 1
                        intLineCount = intLineCount - 1 'vamsi 28th August 2013

                    End If
                Else
                    iweightcount = iweightcount + 1
                    intLineCount = intLineCount - 1 'vamsi 28th August 2013
                End If
            ElseIf intCount = 9 Then
                If SeriesForCosting.Equals("TP-High") Then
                    objClsCMS_MTHL_TieRodCylinder.Tool = "REPHASING FLOW HIGH"
                ElseIf SeriesForCosting.Equals("TP-Low") Then        '04_10_2010   RAGAVA
                    objClsCMS_MTHL_TieRodCylinder.Tool = "REPHASING FLOW LOW"
                Else
                    iweightcount = iweightcount + 1
                    intLineCount = intLineCount - 1 'vamsi 28th August 2013
                End If
            ElseIf intCount = 10 Then
                If Trim(ofrmTieRod2.cmbPaint.Text).IndexOf("Prime") <> -1 Then
                    objClsCMS_MTHL_TieRodCylinder.Tool = "275063"
                Else
                    iweightcount = iweightcount + 1
                    intLineCount = intLineCount - 1 'vamsi 28th August 2013
                End If
            ElseIf intCount = 11 Then
                If Weight_Assembly > 40 Then
                    objClsCMS_MTHL_TieRodCylinder.Tool = "CAUTION WEIGHT"
                Else
                    iweightcount = iweightcount + 1
                    intLineCount = intLineCount - 1 'vamsi 28th August 2013
                End If
            ElseIf intCount = 12 Then
                'objClsCMS_MTHL_TieRodCylinder.Tool = "WI09-E-09"
                objClsCMS_MTHL_TieRodCylinder.Tool = "WI09-E-11"            '26_11_2010   RAGAVA
                '06_07_2011   RAGAVA
            ElseIf intCount = 13 Then   'Sugandhi_20120607

                If ofrmTieRod3.rbYesBagRequired.Checked Then
                    objClsCMS_MTHL_TieRodCylinder.Tool = "275010"
                Else
                    objClsCMS_MTHL_TieRodCylinder.Tool = "275254"
                End If
                'If ofrmTieRod3.rbYesLabelRequired.Checked Then    'Sugandhi_20120614
                '    objClsCMS_MTHL_TieRodCylinder.Tool = "275009"
                'End If

            ElseIf intCount = 14 Then
                '06_09_2012   RAGAVA Commented
                'If ofrmTieRod3.rbYesLabelRequired.Checked Then    'Sugandhi_20120614
                '    objClsCMS_MTHL_TieRodCylinder.Tool = "275009"
                'End If
                iweightcount = iweightcount + 1
            ElseIf intCount = 15 Then
                objClsCMS_MTHL_TieRodCylinder.Tool = "275009"

            ElseIf intCount = 16 Then
                If Weight_Assembly > 40 Then
                    objClsCMS_MTHL_TieRodCylinder.Tool = "CAUTION WEIGHT"
                Else
                    iweightcount = iweightcount + 1
                End If
            End If
            objClsCMS_MTHL_TieRodCylinder.SetDataToExcel(_oExcelSheet_MTHL_TieRodCylinder, intCount - iweightcount)
            _oExWorkbook.Save()
        Next








        'Dim iweightcount As Integer = 0
        'objClsCMS_MTHL_TieRodCylinder.Part = CylinderCodeNumber

        'Dim intLineCount As Integer = 1
        ''Dim iColCount As Integer = 0
        ''Dim iColCount As Integer = 11          '26_11_2010   RAGAVA
        'Dim iColCount As Integer = 13          '06_07_2011   RAGAVA
        'If _ofrmTieRod3.chk100OilTest.Checked = True AndAlso (SeriesForCosting.Equals("TP-High") OrElse SeriesForCosting.Equals("TP-Low")) Then
        '    'iColCount = 11
        '    'iColCount = 13         '06_07_2011   RAGAVA
        '    iColCount = 14        '28_08_2012  RAGAVA
        'ElseIf _ofrmTieRod3.chk100OilTest.Checked = True Then
        '    'iColCount = 2
        '    'iColCount = 11       '04_10_2010   RAGAVA
        '    'iColCount = 13         '06_07_2011   RAGAVA
        '    iColCount = 14        '28_08_2012  RAGAVA
        'End If
        'objClsCMS_MTHL_TieRodCylinder.Seq = 10

        'For intCount As Integer = 2 To iColCount
        '    objClsCMS_MTHL_TieRodCylinder.Tool = ""       '04_10_2010   RAGAVA
        '    If intCount = 4 Then
        '        objClsCMS_MTHL_TieRodCylinder.Seq = 20
        '        intLineCount = 1
        '    ElseIf intCount = 10 Then
        '        objClsCMS_MTHL_TieRodCylinder.Seq = 30
        '        intLineCount = 1
        '    End If

        '    If intCount > 3 AndAlso intCount < 10 Then
        '        If _ofrmTieRod3.chk100OilTest.Checked = False Then        '04_10_2010   RAGAVA
        '            Continue For
        '        ElseIf Not (SeriesForCosting.Equals("TP-High") OrElse SeriesForCosting.Equals("TP-Low")) Then
        '            If intCount > 5 AndAlso intCount < 8 Then
        '                Continue For
        '            End If
        '        End If
        '    End If
        '    objClsCMS_MTHL_TieRodCylinder.Line = intLineCount
        '    intLineCount += 1
        '    '06_10_2010   RAGAVA
        '    If intCount = 2 Then
        '        objClsCMS_MTHL_TieRodCylinder.Tool = "WI09-E-36"
        '    ElseIf intCount = 3 Then
        '        If Weight_Assembly > 40 Then
        '            objClsCMS_MTHL_TieRodCylinder.Tool = "CAUTION WEIGHT"
        '        Else
        '            iweightcount = iweightcount + 1
        '        End If
        '        'Till  Here
        '    ElseIf intCount = 4 Then
        '        objClsCMS_MTHL_TieRodCylinder.Tool = "WI10-E-03"
        '    ElseIf intCount = 5 Then
        '        objClsCMS_MTHL_TieRodCylinder.Tool = "CYL TEST PRESSURE 2"
        '    ElseIf intCount = 6 Then
        '        If _ofrmTieRod3.chk100OilTest.Checked = True Then
        '            If Trim(ofrmTieRod1.cmbRephasingPortPosition.Text).IndexOf("At Both") <> -1 Then
        '                objClsCMS_MTHL_TieRodCylinder.Tool = "CYLINDER STYLE 4"
        '                '04_10_2010   RAGAVA
        '            ElseIf Trim(ofrmTieRod1.cmbRephasingPortPosition.Text).IndexOf("At Extension") <> -1 OrElse Trim(ofrmTieRod1.cmbRephasingPortPosition.Text).IndexOf("At Retraction") <> -1 Then
        '                objClsCMS_MTHL_TieRodCylinder.Tool = "CYLINDER STYLE 3"
        '            Else
        '                iweightcount = iweightcount + 1
        '            End If
        '        End If
        '    ElseIf intCount = 7 Then
        '        If SeriesForCosting.Equals("TP-High") Then
        '            objClsCMS_MTHL_TieRodCylinder.Tool = "REPHASING FLOW HIGH"
        '        ElseIf SeriesForCosting.Equals("TP-Low") Then        '04_10_2010   RAGAVA
        '            objClsCMS_MTHL_TieRodCylinder.Tool = "REPHASING FLOW LOW"
        '        Else
        '            iweightcount = iweightcount + 1
        '        End If
        '    ElseIf intCount = 8 Then
        '        If Trim(ofrmTieRod2.cmbPaint.Text).IndexOf("Prime") <> -1 Then
        '            objClsCMS_MTHL_TieRodCylinder.Tool = "275063"
        '        Else
        '            iweightcount = iweightcount + 1
        '        End If
        '    ElseIf intCount = 9 Then
        '        If Weight_Assembly > 40 Then
        '            objClsCMS_MTHL_TieRodCylinder.Tool = "CAUTION WEIGHT"
        '        Else
        '            iweightcount = iweightcount + 1
        '        End If
        '    ElseIf intCount = 10 Then
        '        'objClsCMS_MTHL_TieRodCylinder.Tool = "WI09-E-09"
        '        objClsCMS_MTHL_TieRodCylinder.Tool = "WI09-E-11"            '26_11_2010   RAGAVA
        '        '06_07_2011   RAGAVA
        '    ElseIf intCount = 11 Then   'Sugandhi_20120607

        '        If ofrmTieRod3.rbYesBagRequired.Checked Then
        '            objClsCMS_MTHL_TieRodCylinder.Tool = "275010"
        '        Else
        '            objClsCMS_MTHL_TieRodCylinder.Tool = "275254"
        '        End If
        '        'If ofrmTieRod3.rbYesLabelRequired.Checked Then    'Sugandhi_20120614
        '        '    objClsCMS_MTHL_TieRodCylinder.Tool = "275009"
        '        'End If

        '    ElseIf intCount = 12 Then
        '        If ofrmTieRod3.rbYesLabelRequired.Checked Then    'Sugandhi_20120614
        '            objClsCMS_MTHL_TieRodCylinder.Tool = "275009"
        '        End If

        '    ElseIf intCount = 13 Then
        '        'objClsCMS_MTHL_TieRodCylinder.Tool = "299801"        '28_08_2012  RAGAVA

        '    ElseIf intCount = 14 Then
        '        If Weight_Assembly > 40 Then
        '            objClsCMS_MTHL_TieRodCylinder.Tool = "CAUTION WEIGHT"
        '        End If
        '    End If
        '    objClsCMS_MTHL_TieRodCylinder.SetDataToExcel(_oExcelSheet_MTHL_TieRodCylinder, intCount - iweightcount)
        '    _oExWorkbook.Save()
        'Next

    End Sub

    Private Sub MTHL_RodSheetLogics(ByVal objClsCMS_MTHL_Rod As clsCMS_MTHL) 'METHDR ROD

        Try

            objClsCMS_MTHL_Rod.Part = strRodCodeNumber

            Dim intRodSeq As Integer = 10
            objClsCMS_MTHL_Rod.Line = 0
            Dim blnUnderCut As Boolean = False
            Dim iweightcount As Integer = 0
            '13_09_2010    RAGAVA
            For intCount As Integer = 2 To 14
                objClsCMS_MTHL_Rod.Tool = ""       '04_10_2010   RAGAVA
                objClsCMS_MTHL_Rod.Seq = intRodSeq
                If intCount = 2 OrElse intCount = 5 Then
                    objClsCMS_MTHL_Rod.Tool = "WI10-E-25"
                End If

                '04_10_2010   RAGAVA
                If intCount = 4 Then
                    If Weight_Rod >= 40 Then
                        objClsCMS_MTHL_Rod.Tool = "CAUTION WEIGHT"
                    Else
                        iweightcount = iweightcount + 1
                    End If
                End If
                'Till   Here

                If intCount = 3 OrElse intCount = 6 Then
                    '04_10_2010   RAGAVA
                    If intCount = 6 Then
                        If (SeriesForCosting.IndexOf("TX ") = -1 AndAlso (RodMaterialForCosting = "Chrome" OrElse UCase(RodMaterialForCosting).IndexOf("LION") <> -1)) OrElse (RodMaterialForCosting = "Nitro Steel" AndAlso (strStyleModified.IndexOf("NON ASAE") <> -1)) Then ' AndAlso strStyleModified.Equals("ASAE") Then
                            objClsCMS_MTHL_Rod.Tool = "WI09-E-79"
                        End If
                    Else
                        objClsCMS_MTHL_Rod.Tool = "WI09-E-79"
                    End If
                End If
                If intCount = 7 Then
                    objClsCMS_MTHL_Rod.Tool = "C0011"
                End If
                If intCount > 7 Then
                    If intCount = 8 Then
                        If RodLength >= 8 AndAlso RodLength < 24 Then
                            objClsCMS_MTHL_Rod.Tool = "C0047"
                        ElseIf RodLength >= 24 AndAlso RodLength <= 40 Then
                            objClsCMS_MTHL_Rod.Tool = "C0184"      '"C00184"
                        ElseIf RodLength > 40 Then
                            objClsCMS_MTHL_Rod.Tool = "500463"       '"CONTACT PROCESS ENGINEER"
                        End If
                    ElseIf intCount = 9 Then
                        If PistonThreadSize < 1 Then
                            objClsCMS_MTHL_Rod.Tool = "M001"
                        ElseIf PistonThreadSize >= 1 AndAlso PistonThreadSize < 2 Then
                            objClsCMS_MTHL_Rod.Tool = "M002"
                        End If
                    ElseIf intCount = 10 Then
                        If strStyleModified.Equals("ASAE") OrElse RodDiameter > dblRodThreadSize Then           '28_08_2012  RAGAVA
                            If dblRodThreadSize < 1 Then
                                objClsCMS_MTHL_Rod.Tool = "M003"
                            ElseIf dblRodThreadSize >= 1 AndAlso dblRodThreadSize < 2 Then
                                objClsCMS_MTHL_Rod.Tool = "M004"
                            End If
                        Else
                            blnUnderCut = True
                            Continue For
                        End If
                        'If PistonThreadSize <> dblRodThreadSize Then
                        '    If dblRodThreadSize < 1 Then
                        '        objClsCMS_MTHL_Rod.Tool = "M003"
                        '    ElseIf dblRodThreadSize >= 1 AndAlso dblRodThreadSize < 2 Then
                        '        objClsCMS_MTHL_Rod.Tool = "M004"
                        '    End If
                        'Else
                        '    blnUnderCut = True
                        '    Continue For
                        'End If
                    ElseIf intCount = 11 Then
                        If PistonThreadSize = 0.5 Then
                            objClsCMS_MTHL_Rod.Tool = "V5"
                        ElseIf PistonThreadSize = 0.63 Then
                            objClsCMS_MTHL_Rod.Tool = "V7"
                        ElseIf PistonThreadSize = 0.75 Then
                            objClsCMS_MTHL_Rod.Tool = "V8"
                        ElseIf PistonThreadSize = 0.87 Then
                            objClsCMS_MTHL_Rod.Tool = "V9"
                        ElseIf PistonThreadSize = 1.0 Then
                            'objClsCMS_MTHL_Rod.Tool = "V50"
                            objClsCMS_MTHL_Rod.Tool = "V11"      '05_01_2011    RAGAVA
                        ElseIf PistonThreadSize = 1.12 Then
                            objClsCMS_MTHL_Rod.Tool = "V12"
                        ElseIf PistonThreadSize = 1.26 Then
                            objClsCMS_MTHL_Rod.Tool = "V13"
                        ElseIf PistonThreadSize = 1.5 Then
                            objClsCMS_MTHL_Rod.Tool = "V14"
                        End If
                    ElseIf intCount = 12 Then
                        If PistonThreadSize = 0.5 Then
                            objClsCMS_MTHL_Rod.Tool = "RG32"
                        ElseIf PistonThreadSize = 0.63 Then
                            objClsCMS_MTHL_Rod.Tool = "RG34"
                        ElseIf PistonThreadSize = 0.75 Then
                            objClsCMS_MTHL_Rod.Tool = "RG26"
                        ElseIf PistonThreadSize = 0.87 Then
                            objClsCMS_MTHL_Rod.Tool = "RG33"
                        ElseIf PistonThreadSize = 1.0 Then
                            objClsCMS_MTHL_Rod.Tool = "RG27"
                        ElseIf PistonThreadSize = 1.12 Then
                            objClsCMS_MTHL_Rod.Tool = "RG28"
                        ElseIf PistonThreadSize = 1.26 Then
                            objClsCMS_MTHL_Rod.Tool = "RG29"
                        ElseIf PistonThreadSize = 1.5 Then
                            objClsCMS_MTHL_Rod.Tool = "RG30"
                        End If
                    ElseIf intCount = 13 Then
                        If dblRodThreadSize = 1.12 Then
                            objClsCMS_MTHL_Rod.Tool = "V31"
                        ElseIf dblRodThreadSize = 1.26 Then
                            objClsCMS_MTHL_Rod.Tool = "V29"
                        ElseIf dblRodThreadSize = 1.38 Then
                            objClsCMS_MTHL_Rod.Tool = "V30"
                        ElseIf dblRodThreadSize = 1.5 Then
                            objClsCMS_MTHL_Rod.Tool = "V28"
                        End If
                    ElseIf intCount = 14 Then
                        If Weight_Rod >= 40 Then
                            objClsCMS_MTHL_Rod.Tool = "CAUTION WEIGHT"
                        End If
                    End If
                    End If
                    If objClsCMS_MTHL_Rod.Tool <> "" Then     '04_10_2010   RAGAVA
                        If intCount = 5 Then
                            objClsCMS_MTHL_Rod.Line = 1
                        Else
                            objClsCMS_MTHL_Rod.Line = objClsCMS_MTHL_Rod.Line + 1
                        End If
                        If intCount > 4 Then
                            intRodSeq = 20
                            objClsCMS_MTHL_Rod.Seq = intRodSeq
                        End If
                        If blnUnderCut = True Then
                            objClsCMS_MTHL_Rod.SetDataToExcel(_oExcelSheet_MTHL_Rod, intCount - 1 - iweightcount)
                        Else
                            objClsCMS_MTHL_Rod.SetDataToExcel(_oExcelSheet_MTHL_Rod, intCount - iweightcount)
                        End If
                        _oExWorkbook.Save()
                    End If
            Next
            'Till    Here
            'For intCount As Integer = 2 To 3
            '    objClsCMS_MTHL_Rod.Seq = intRodSeq
            '    objClsCMS_MTHL_Rod.Tool = "WI10-E-25"
            '    objClsCMS_MTHL_Rod.Line = 1
            '    intRodSeq += 10

            '    objClsCMS_MTHL_Rod.SetDataToExcel(_oExcelSheet_MTHL_Rod, intCount)
            '    _oExWorkbook.Save()
            'Next
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "CSV Conversion Functionality"

    Private Sub CSVConversionFunctionality()
        DeleteIndexRow()
        _oExWorkbook.Save()
        KillExcel()

        WriteToCSV()
        KillExcel()

        DeleteEndCharFromCVS()
        File.Delete(_strCMSInterationChildFilePath) 'Actual worked excel
    End Sub

    Private Sub DeleteIndexRow()

        'STKMM_STKMP----------
        DeleteFirstRow(_oExcelSheet_STKMM_TieRodCylinder)
        DeleteFirstRow(_oExcelSheet_STKMM_Tube)
        DeleteFirstRow(_oExcelSheet_STKMM_Rod)
        DeleteFirstRow(_oExcelSheet_STKMP_TieRod)
        If IsStopTubeSelected Then
            DeleteFirstRow(_oExcelSheet_STKMP_StopTube)
        End If
        '---------------------


        'STKA-----------------
        DeleteFirstRow(_oExcelSheet_STKA_TieRodCylinder)
        DeleteFirstRow(_oExcelSheet_STKA_Tube)
        DeleteFirstRow(_oExcelSheet_STKA_Rod)
        DeleteFirstRow(_oExcelSheet_STKA_TieRod)
        If IsStopTubeSelected Then
            DeleteFirstRow(_oExcelSheet_STKA_StopTube)
        End If
        '---------------------

        'METHDM----------------
        DeleteFirstRow(_oExcelSheet_METHDM_TieRodCylinder)
        DeleteFirstRow(_oExcelSheet_METHDM_Tube)
        DeleteFirstRow(_oExcelSheet_METHDM_Rod)
        '---------------------

        'METHDR----------------
        DeleteFirstRow(_oExcelSheet_METHDR_TieRodCylinder)
        DeleteFirstRow(_oExcelSheet_METHDR_Tube)
        DeleteFirstRow(_oExcelSheet_METHDR_Rod)
        '---------------------

        'MTHL----------------
        DeleteFirstRow(_oExcelSheet_MTHL_TieRodCylinder)       '16_09_2010   RAGAVA
        DeleteFirstRow(_oExcelSheet_MTHL_Tube)
        DeleteFirstRow(_oExcelSheet_MTHL_Rod)
        '---------------------

    End Sub

    Private Sub WriteToCSV()
        Try
            'anup 17-02-2011 start
            Dim oClsReleaseCylinderFunctionality As New clsReleaseCylinderFunctionality

            If oClsReleaseCylinderFunctionality.DoesCodeExistInDB(strBoreCodeNumber, "TUBE") Then
                _IsExistingCodeButNotReleased_Tube = True
            End If
            If oClsReleaseCylinderFunctionality.DoesCodeExistInDB(strRodCodeNumber, "ROD") Then
                _IsExistingCodeButNotReleased_Rod = True
            End If
            If oClsReleaseCylinderFunctionality.DoesCodeExistInDB(StopTubeCodeNumber, "STOPTUBE") Then
                _IsExistingCodeButNotReleased_StopTube = True
            End If
            If oClsReleaseCylinderFunctionality.DoesCodeExistInDB(strTieRodCodeNumber, "TIEROD") Then
                _IsExistingCodeButNotReleased_TieRod = True
            End If
            'anup 17-02-2011 till here

            Dim oExApp As New Excel.Application
            oExApp.Visible = False

            Dim oExWB As Excel.Workbook = oExApp.Workbooks.Open(_strCMSInterationChildFilePath)

            'STKMM_STKMP----------
            Dim oSTKMM_STKMP_CylinderSheet As Excel.Worksheet = oExApp.Sheets(1) 'STKMM_STKMP Cylinder
            SaveasCSVFile(oSTKMM_STKMP_CylinderSheet, CylinderCodeNumber)

            Dim oSTKMM_STKMP_TubeSheet As Excel.Worksheet = oExApp.Sheets(2) 'STKMM_STKMP Tube
            If _blnIsNewTube OrElse _IsExistingCodeButNotReleased_Tube Then  'anup 17-02-2011 'anup 10-03-2011
                SaveasCSVFile(oSTKMM_STKMP_TubeSheet, strBoreCodeNumber)
            End If

            Dim oSTKMM_STKMP_RodSheet As Excel.Worksheet = oExApp.Sheets(3) 'STKMM_STKMP Rod
            If _blnIsNewRod OrElse _IsExistingCodeButNotReleased_Rod Then 'anup 17-02-2011  'anup 10-03-2011
                SaveasCSVFile(oSTKMM_STKMP_RodSheet, strRodCodeNumber)
            End If

            Dim oSTKMM_STKMP_TieRodSheet As Excel.Worksheet = oExApp.Sheets(4) 'STKMM_STKMP TieRod
            If _blnIsNewTierod OrElse _IsExistingCodeButNotReleased_TieRod Then 'anup 17-02-2011  'anup 10-03-2011
                'SaveasCSVFile(oSTKMM_STKMP_TieRodSheet, CylinderCodeNumber)
                SaveasCSVFile(oSTKMM_STKMP_TieRodSheet, strTieRodCodeNumber)      '06_01_2011    RAGAVA
            End If

            If IsStopTubeSelected Then
                Dim oSTKMM_STKMP_StopTubeSheet As Excel.Worksheet = oExApp.Sheets(5) 'STKMM_STKMP StopTube
                If _blnIsNewStopTube OrElse _IsExistingCodeButNotReleased_StopTube Then 'anup 17-02-2011  'anup 10-03-2011
                    SaveasCSVFile(oSTKMM_STKMP_StopTubeSheet, StopTubeCodeNumber)
                End If
            End If
            '---------------------


            'STKA-----------------
            Dim oSTKA_CylinderSheet As Excel.Worksheet = oExApp.Sheets(6) 'STKA Cylinder
            SaveasCSVFile(oSTKA_CylinderSheet, CylinderCodeNumber)

            Dim oSTKA_TubeSheet As Excel.Worksheet = oExApp.Sheets(7) 'STKA Tube
            If _blnIsNewTube OrElse _IsExistingCodeButNotReleased_Tube Then 'anup 17-02-2011 'anup 10-03-2011
                SaveasCSVFile(oSTKA_TubeSheet, strBoreCodeNumber)
            End If

            Dim oSTKA_RodSheet As Excel.Worksheet = oExApp.Sheets(8) 'STKA Rod
            If _blnIsNewRod OrElse _IsExistingCodeButNotReleased_Rod Then 'anup 17-02-2011 'anup 10-03-2011
                SaveasCSVFile(oSTKA_RodSheet, strRodCodeNumber)
            End If

            Dim oSTKA_TieRodSheet As Excel.Worksheet = oExApp.Sheets(9) 'STKA TieRod
            If _blnIsNewTierod OrElse _IsExistingCodeButNotReleased_TieRod Then 'anup 17-02-2011 'anup 10-03-2011
                SaveasCSVFile(oSTKA_TieRodSheet, strTieRodCodeNumber)
            End If

            If IsStopTubeSelected Then
                Dim oSTKA_StopTubeSheet As Excel.Worksheet = oExApp.Sheets(10) 'STKA StopTube
                If _blnIsNewStopTube OrElse _IsExistingCodeButNotReleased_StopTube Then 'anup 17-02-2011 'anup 10-03-2011
                    SaveasCSVFile(oSTKA_StopTubeSheet, StopTubeCodeNumber)
                End If
            End If
            '---------------------


            'METHDM----------------
            Dim oMETHDM_TieRodCylinderSheet As Excel.Worksheet = oExApp.Sheets(11) 'METHDM TieRodCylinder
            SaveasCSVFile(oMETHDM_TieRodCylinderSheet, CylinderCodeNumber)

            Dim oMETHDM_TubeSheet As Excel.Worksheet = oExApp.Sheets(12) 'METHDM Tube
            If _blnIsNewTube OrElse _IsExistingCodeButNotReleased_Tube Then 'anup 17-02-2011 'anup 10-03-2011
                SaveasCSVFile(oMETHDM_TubeSheet, strBoreCodeNumber)
            End If

            Dim oMETHDM_RodSheet As Excel.Worksheet = oExApp.Sheets(13) 'METHDM Rod
            If _blnIsNewRod OrElse _IsExistingCodeButNotReleased_Rod Then 'anup 17-02-2011 'anup 10-03-2011
                SaveasCSVFile(oMETHDM_RodSheet, strRodCodeNumber)
            End If
            '---------------------

            'METHDR----------------
            Dim oMETHDR_TieRodCylinderSheet As Excel.Worksheet = oExApp.Sheets(14) 'METHDR TieRodCylinder
            SaveasCSVFile(oMETHDR_TieRodCylinderSheet, CylinderCodeNumber)

            Dim oMETHDR_TubeSheet As Excel.Worksheet = oExApp.Sheets(15) 'METHDR Tube
            If _blnIsNewTube OrElse _IsExistingCodeButNotReleased_Tube Then 'anup 17-02-2011 'anup 10-03-2011
                SaveasCSVFile(oMETHDR_TubeSheet, strBoreCodeNumber)
            End If

            Dim oMETHDR_RodSheet As Excel.Worksheet = oExApp.Sheets(16) 'METHDR Rod
            If _blnIsNewRod OrElse _IsExistingCodeButNotReleased_Rod Then 'anup 17-02-2011 'anup 10-03-2011
                SaveasCSVFile(oMETHDR_RodSheet, strRodCodeNumber)
            End If
            '---------------------


            'MTHL----------------
            '05_10_2010   RAGAVA  uncommented
            'If _ofrmTieRod3.chk100OilTest.Checked = True Then             '26_11_2010   RAGAVA
            Dim oMTHL_TieRodCylinderSheet As Excel.Worksheet = oExApp.Sheets(17) 'MTHL TieRodCylinder
            SaveasCSVFile(oMTHL_TieRodCylinderSheet, CylinderCodeNumber)
            'End If
            Dim oMTHL_TubeSheet As Excel.Worksheet = oExApp.Sheets(18) 'MTHL Tube
            If _blnIsNewTube OrElse _IsExistingCodeButNotReleased_Tube Then 'anup 17-02-2011 'anup 10-03-2011
                SaveasCSVFile(oMTHL_TubeSheet, strBoreCodeNumber)
            End If

            Dim oMTHL_RodSheet As Excel.Worksheet = oExApp.Sheets(19) 'MTHL Rod
            If _blnIsNewRod OrElse _IsExistingCodeButNotReleased_Rod Then       '14_09_2010   RAGAVA    'anup 17-02-2011  'anup 10-03-2011
                SaveasCSVFile(oMTHL_RodSheet, strRodCodeNumber)
            End If
            '---------------------
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DeleteFirstRow(ByVal objExcelSheet As Excel.Worksheet)
        Try
            objExcelSheet.Range("A1").EntireRow.Delete()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SaveasCSVFile(ByVal objExcelSheet As Excel.Worksheet, Optional ByVal codenumber As String = "0")
        Try
            If File.Exists(_strDirectoryName + "\" + objExcelSheet.Name.Substring(0, objExcelSheet.Name.IndexOf("_")) & codenumber & ".csv") Then
                File.Delete(_strDirectoryName + "\" + objExcelSheet.Name.Substring(0, objExcelSheet.Name.IndexOf("_")) & codenumber & ".csv")
            End If
        Catch ex As Exception

        End Try
        Try
            If codenumber = "0" Then
                objExcelSheet.SaveAs(_strDirectoryName + "\" + objExcelSheet.Name, Excel.XlFileFormat.xlCSVWindows, , Excel.XlFileAccess.xlReadWrite)
            Else
                objExcelSheet.SaveAs(_strDirectoryName + "\" + objExcelSheet.Name.Substring(0, objExcelSheet.Name.IndexOf("_")) & codenumber, Excel.XlFileFormat.xlCSVWindows, , Excel.XlFileAccess.xlReadWrite)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DeleteEndCharFromCVS()
        Try
            Dim strFilePaths As String() = Directory.GetFiles(_strDirectoryName)
            For Each strFilePath As String In strFilePaths
                If strFilePath.Contains(".csv") Then
                    'Dim strTextFromFile As String = File.ReadAllText(strFilePath)
                    'Dim intLastIndexofComma As Integer = strTextFromFile.LastIndexOf(",")
                    'Dim strTextWithoutEnd As String = strTextFromFile.Substring(0, intLastIndexofComma)
                    'File.Delete(strFilePath)

                    'Dim fs As New FileStream(strFilePath, FileMode.Create, FileAccess.ReadWrite)
                    'Dim sw As New StreamWriter(fs)
                    'sw.BaseStream.Seek(0, SeekOrigin.End)
                    'sw.WriteLine(strTextWithoutEnd)
                    'sw.Close()
                    'GC.Collect()

                    Dim strTextFromFile As String() = File.ReadAllLines(strFilePath)

                    Dim strTextWithoutEnd As String = ""
                    For Each strLine As String In strTextFromFile
                        If strLine <> "" Then         '23_11_2010   RAGAVA
                            Dim intLastIndexofComma As Integer = strLine.LastIndexOf(",")
                            '26_11_2010   RAGAVA
                            If strLine.EndsWith("END") Or strLine.EndsWith("end") Then
                                ''13_12_2010   RAGAVA
                                If strTextWithoutEnd <> "" Then
                                    strTextWithoutEnd += vbCrLf + strLine.Substring(0, intLastIndexofComma)
                                Else
                                    strTextWithoutEnd += strLine.Substring(0, intLastIndexofComma) ' + vbCrLf
                                End If
                                'Till  Here
                            Else
                                '13_12_2010   RAGAVA
                                If strTextWithoutEnd <> "" Then
                                    strTextWithoutEnd += vbCrLf + strLine '+ vbCrLf
                                Else
                                    strTextWithoutEnd += strLine '+ vbCrLf
                                End If
                                'Till  Here
                            End If
                            'strTextWithoutEnd += strLine.Substring(0, intLastIndexofComma) + vbCrLf
                            'Till   Here
                            File.Delete(strFilePath)
                            End If
                    Next

                    strTextWithoutEnd = strTextWithoutEnd.ToString.Replace(vbCrLf, "XXX")
                    strTextWithoutEnd = strTextWithoutEnd.ToString.Replace(vbLf, "")
                    strTextWithoutEnd = strTextWithoutEnd.ToString.Replace("XXX", vbCrLf)


                    Dim fs As New FileStream(strFilePath, FileMode.Create, FileAccess.ReadWrite)
                    Dim sw As New StreamWriter(fs)
                    sw.BaseStream.Seek(0, SeekOrigin.End)
                    sw.NewLine.Trim()
                    sw.WriteLine(strTextWithoutEnd)
                    sw.Close()
                    GC.Collect()
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub MoveDirectoryToW()
        Try
            '14_07_2011   RAGAVA
            Dim strPath As String = String.Empty
            If IsNew_Revision_Released = "Released" Then
                strPath = "W:\TIEROD\CMS\"
            Else
                'strPath = "C:\MONARCH_TESTING\CMS_TEMP\"
                strPath = "K:\USR\_CYLINDER\CYLOEM\IFL DWG NR\TIEROD\CMS\"

            End If
            If Not Directory.Exists(strPath) Then
                Directory.CreateDirectory(strPath)
            End If
            If Directory.Exists(_strDirectoryName) Then
                If Directory.Exists(strPath + CylinderCodeNumber + "_CMS") Then
                    Directory.Delete(strPath + CylinderCodeNumber + "_CMS", True)
                End If
                My.Computer.FileSystem.MoveDirectory(_strDirectoryName, strPath + CylinderCodeNumber + "_CMS")
            End If
            'Till  Here



            'If Not Directory.Exists("W:\TIEROD\CMS\") Then
            '    Directory.CreateDirectory("W:\TIEROD\CMS\")
            'End If
            'If Directory.Exists(_strDirectoryName) Then
            '    If Directory.Exists("W:\TIEROD\CMS\" + CylinderCodeNumber + "_CMS") Then
            '        Directory.Delete("W:\TIEROD\CMS\" + CylinderCodeNumber + "_CMS", True)
            '    End If
            '    My.Computer.FileSystem.MoveDirectory(_strDirectoryName, "W:\TIEROD\CMS\" + CylinderCodeNumber + "_CMS")
            'End If
        Catch ex As Exception

        End Try
    End Sub

#End Region

#End Region

#Region "Functions"

    Public Function CMSIntegrationfunctionality() As Boolean
        CMSIntegrationfunctionality = False
        Try
            If CreateExcelObjects() Then
                StartLogic()
                KillExcel()
                CMSIntegrationfunctionality = True
            End If
        Catch ex As Exception
            KillExcel()
            CMSIntegrationfunctionality = False
        End Try
    End Function

#End Region



    '15_09_2010    RAGAVA
    Public Sub METHE_TieRodCylinder_Functionality()
        Try
            Dim oExApp As New Excel.Application
            Dim oExlWrkBk As Excel.Workbook
            Dim oExlWrkSht As Excel.Worksheet
            Dim oRng As Excel.Range
            Dim strQuery As String = String.Empty
            Dim objDT As System.Data.DataTable
CHECK_CMS:
            If File.Exists("C:\CMS.xls") = False Then
                MsgBox("Create Excel File C:\CMS.xls" & vbNewLine & "then click ok")
                GoTo CHECK_CMS
            End If
            'TUBE FILE GENERATION
            If _blnIsNewTube OrElse _IsExistingCodeButNotReleased_Tube Then  'anup 17-02-2011 'anup 10-03-2011
                If SeriesForCosting.Equals("TP-High") OrElse SeriesForCosting.Equals("TP-Low") Then
                    strQuery = "Select * from TUBE_OP_10_1 union select * from TUBE_OP_20_1 union select * from TUBE_OP_25_1 union select * from TUBE_OP_30_1"
                Else
                    strQuery = "Select * from TUBE_OP_10_1 union select * from TUBE_OP_20_1"
                End If
                objDT = oDataClass.GetDataTable(strQuery)
                Dim strfile As String = "C:\CMS.xls"
                oExlWrkBk = oExApp.Workbooks.Open(strfile)
                oExlWrkSht = oExlWrkBk.Worksheets("Sheet1")
                Dim icount As Integer = 1
                For Each dr As DataRow In objDT.Rows
                    oRng = oExlWrkSht.Range("A" & icount.ToString)
                    oRng.Value2 = strBoreCodeNumber.ToString    '04_10_2010   RAGAVA ' dr(0).ToString
                    oRng = oExlWrkSht.Range("B" & icount.ToString)
                    oRng.Value2 = "'" & Format(Val(dr(1)), "00").ToString
                    oRng = oExlWrkSht.Range("C" & icount.ToString)
                    oRng.Value2 = "'" & Format(Val(dr(2)), "00").ToString
                    oRng = oExlWrkSht.Range("D" & icount.ToString)
                    oRng.Value2 = dr(3).ToString.Replace("'", "")
                    icount = icount + 1
                Next
                '14_07_2011  RAGAVA
                If IsNew_Revision_Released = "Released" Then
                    oExlWrkBk.SaveAs("W:\TIEROD\CMS\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strBoreCodeNumber.ToString & ".csv", XlFileFormat.xlCSVMSDOS)
                Else
                    'oExlWrkBk.SaveAs("C:\MONARCH_TESTING\CMS_TEMP\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strBoreCodeNumber.ToString & ".csv", XlFileFormat.xlCSVMSDOS)
                    oExlWrkBk.SaveAs("K:\USR\_CYLINDER\CYLOEM\IFL DWG NR\TIEROD\CMS\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strBoreCodeNumber.ToString & ".csv", XlFileFormat.xlCSVMSDOS)
                End If
                'oExlWrkBk.SaveAs("W:\TIEROD\CMS\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strBoreCodeNumber.ToString & ".csv", XlFileFormat.xlCSVMSDOS)
                'Till   Here
                objDT.Clear()
            End If
            'ROD FILE GENERATION
            If _blnIsNewRod OrElse _IsExistingCodeButNotReleased_Rod Then  'anup 17-02-2011'anup 10-03-2011
                If RodMaterialForCosting = "Chrome" Or UCase(RodMaterialForCosting).IndexOf("LION") <> -1 Then
                    strQuery = "Select * from ROD_OP_10_1"
                ElseIf RodMaterialForCosting = "Nitro Steel" Then
                    strQuery = "Select * from ROD_OP_10_2"
                Else
                    '15_12_2010    RAGAVA
                    If strStyleModified.Equals("NON ASAE") Then
                        strQuery = "Select * from ROD_OP_20_7"
                    ElseIf strStyleModified.Equals("ASAE") Then
                        strQuery = "Select * from ROD_OP_20_8"
                    End If
                    'strQuery = "Select * from ROD_OP_20_7"
                    'Till   Here
                End If
                If (SeriesForCosting.Equals("TL (TC)") OrElse SeriesForCosting.Equals("TH (TD)") OrElse SeriesForCosting.Equals("TP-High") OrElse SeriesForCosting.Equals("TP-Low") OrElse SeriesForCosting.Equals("LN")) AndAlso (strStyleModified.Equals("NON ASAE")) Then
                    If (RodMaterialForCosting = "Chrome" OrElse UCase(RodMaterialForCosting).IndexOf("LION") <> -1) Then
                        strQuery = strQuery & " union Select * from ROD_OP_20_1"
                    ElseIf RodMaterialForCosting = "Nitro Steel" Then
                        strQuery = strQuery & " union Select * from ROD_OP_20_2"
                    End If
                ElseIf (SeriesForCosting.Equals("TL (TC)") OrElse SeriesForCosting.Equals("TH (TD)") OrElse SeriesForCosting.Equals("TP-High") OrElse SeriesForCosting.Equals("TP-Low") OrElse SeriesForCosting.Equals("LN")) AndAlso (strStyleModified.Equals("ASAE")) Then
                    If (RodMaterialForCosting = "Chrome" OrElse UCase(RodMaterialForCosting).IndexOf("LION") <> -1) Then
                        strQuery = strQuery & " union Select * from ROD_OP_20_3"
                    ElseIf RodMaterialForCosting = "Nitro Steel" Then
                        strQuery = strQuery & " union Select * from ROD_OP_20_4"
                    End If
                ElseIf SeriesForCosting.Equals("TX (TXC)") AndAlso strStyleModified.Equals("ASAE") Then
                    If (RodMaterialForCosting = "Chrome" OrElse UCase(RodMaterialForCosting).IndexOf("LION") <> -1) Then
                        strQuery = strQuery & " union Select * from ROD_OP_20_5"
                    End If
                ElseIf SeriesForCosting.Equals("TX (TXC)") AndAlso strStyleModified.Equals("NON ASAE") Then
                    If (RodMaterialForCosting = "Chrome" OrElse UCase(RodMaterialForCosting).IndexOf("LION") <> -1) Then
                        strQuery = strQuery & " union Select * from ROD_OP_20_6"
                    End If

                    '15_12_2010    RAGAVA
                ElseIf strStyleModified.Equals("NON ASAE") Then
                    strQuery = strQuery & " union Select * from ROD_OP_20_7"           '20_01_2011  RAGAVA   strQuery & added
                ElseIf strStyleModified.Equals("ASAE") Then
                    strQuery = strQuery & " union Select * from ROD_OP_20_8"           '20_01_2011  RAGAVA   strQuery & added
                    'Till  Here

                End If
                objDT = oDataClass.GetDataTable(strQuery)
                Dim strfile As String = "C:\CMS.xls"
                oExlWrkBk = oExApp.Workbooks.Open(strfile)
                oExlWrkSht = oExlWrkBk.Worksheets("Sheet1")
                Dim icount As Integer = 1
                For Each dr As DataRow In objDT.Rows
                    oRng = oExlWrkSht.Range("A" & icount.ToString)
                    oRng.Value2 = strRodCodeNumber  '04_10_2010  RAGAVA     ' dr(0).ToString
                    oRng = oExlWrkSht.Range("B" & icount.ToString)
                    If objDT.Rows.Count = 30 Then
                        oRng.Value2 = "'" & Format((Val(dr(1)) - 1), "00").ToString
                    Else
                        oRng.Value2 = "'" & Format(Val(dr(1)), "00").ToString
                    End If
                    oRng = oExlWrkSht.Range("C" & icount.ToString)
                    oRng.Value2 = "'" & Format(Val(dr(2)), "00").ToString
                    oRng = oExlWrkSht.Range("D" & icount.ToString)
                    oRng.Value2 = dr(3).ToString.Replace("'", "")
                    icount = icount + 1
                Next
                '14_07_2011  RAGAVA
                If IsNew_Revision_Released = "Released" Then
                    oExlWrkBk.SaveAs("W:\TIEROD\CMS\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strRodCodeNumber.ToString & ".csv", XlFileFormat.xlCSVMSDOS)
                Else
                    'oExlWrkBk.SaveAs("C:\MONARCH_TESTING\CMS_TEMP\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strRodCodeNumber.ToString & ".csv", XlFileFormat.xlCSVMSDOS)
                    oExlWrkBk.SaveAs("K:\USR\_CYLINDER\CYLOEM\IFL DWG NR\TIEROD\CMS\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strRodCodeNumber.ToString & ".csv", XlFileFormat.xlCSVMSDOS)
                End If
                'oExlWrkBk.SaveAs("W:\TIEROD\CMS\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strRodCodeNumber.ToString & ".csv", XlFileFormat.xlCSVMSDOS)
                'Till   Here

                objDT.Clear()
            End If
            'anup 02-03-2011 start
            KillExcel()
            '14_07_2011  RAGAVA
            If IsNew_Revision_Released = "Released" Then
                CVSFileUtil.DoesFileContainDoubleQuotation("W:\TIEROD\CMS\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strBoreCodeNumber.ToString & ".csv")
                CVSFileUtil.DoesFileContainDoubleQuotation("W:\TIEROD\CMS\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strRodCodeNumber.ToString & ".csv")
            Else
                'CVSFileUtil.DoesFileContainDoubleQuotation("C:\MONARCH_TESTING\CMS_TEMP\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strBoreCodeNumber.ToString & ".csv")
                'CVSFileUtil.DoesFileContainDoubleQuotation("C:\MONARCH_TESTING\CMS_TEMP\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strRodCodeNumber.ToString & ".csv")

                CVSFileUtil.DoesFileContainDoubleQuotation("K:\USR\_CYLINDER\CYLOEM\IFL DWG NR\TIEROD\CMS\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strBoreCodeNumber.ToString & ".csv")
                CVSFileUtil.DoesFileContainDoubleQuotation("K:\USR\_CYLINDER\CYLOEM\IFL DWG NR\TIEROD\CMS\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strRodCodeNumber.ToString & ".csv")
            End If
            'CVSFileUtil.DoesFileContainDoubleQuotation("W:\TIEROD\CMS\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strBoreCodeNumber.ToString & ".csv")
            'CVSFileUtil.DoesFileContainDoubleQuotation("W:\TIEROD\CMS\" + CylinderCodeNumber + "_CMS" & "\" & "METHDE" & strRodCodeNumber.ToString & ".csv")
            'Till  Here
        Catch ex As Exception

        End Try
    End Sub



#Region "Excel Functions"

    Private Function CreateExcelObjects() As Boolean
        CreateExcelObjects = False
        Try
            If CheckForExcel() Then
                If CopyTheMasterFile() Then
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

    Private Function CopyTheMasterFile() As Boolean
        CopyTheMasterFile = False
        Dim blnIsProcessSuccessfull As Boolean = False
        Dim sErrorMessage As String = "Report Master file does not exist"
        Try
            KillExcel()
            ' This function checks if the master report format exists
            If IsMasterReportFileExists() Then
                'Check if directory already exists
                If Directory.Exists(_strDirectoryName) Then
                    'Delete the directory first
                    Directory.Delete(_strDirectoryName, True)
                End If
                Directory.CreateDirectory(_strDirectoryName)
                File.Copy(_strCMSInterationMasterFilePath, _strCMSInterationChildFilePath)
                CopyTheMasterFile = True
                blnIsProcessSuccessfull = True
            End If
            If Not blnIsProcessSuccessfull Then
                CopyTheMasterFile = False
                MessageBox.Show(sErrorMessage, "Error in file creation", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Return blnIsProcessSuccessfull
        Catch ex As Exception
            CopyTheMasterFile = False
        End Try
    End Function

    Private Function IsMasterReportFileExists() As Boolean
        IsMasterReportFileExists = File.Exists(_strCMSInterationMasterFilePath)
    End Function

    Private Function CreateExcel() As Boolean
        CreateExcel = True
        Try
            _oExApplication = New Excel.Application

            _oExApplication.Visible = False

            _oExWorkbook = _oExApplication.Workbooks.Open(_strCMSInterationChildFilePath)

            _oExcelSheet_STKMM_TieRodCylinder = _oExApplication.Sheets(1)

            _oExcelSheet_STKMM_Tube = _oExApplication.Sheets(2)

            _oExcelSheet_STKMM_Rod = _oExApplication.Sheets(3)

            _oExcelSheet_STKMP_TieRod = _oExApplication.Sheets(4)

            _oExcelSheet_STKMP_StopTube = _oExApplication.Sheets(5)

            _oExcelSheet_STKA_TieRodCylinder = _oExApplication.Sheets(6)

            _oExcelSheet_STKA_Tube = _oExApplication.Sheets(7)

            _oExcelSheet_STKA_Rod = _oExApplication.Sheets(8)

            _oExcelSheet_STKA_TieRod = _oExApplication.Sheets(9)

            _oExcelSheet_STKA_StopTube = _oExApplication.Sheets(10)

            _oExcelSheet_METHDM_TieRodCylinder = _oExApplication.Sheets(11)

            _oExcelSheet_METHDM_Tube = _oExApplication.Sheets(12)

            _oExcelSheet_METHDM_Rod = _oExApplication.Sheets(13)

            _oExcelSheet_METHDR_TieRodCylinder = _oExApplication.Sheets(14)

            _oExcelSheet_METHDR_Tube = _oExApplication.Sheets(15)

            _oExcelSheet_METHDR_Rod = _oExApplication.Sheets(16)

            _oExcelSheet_MTHL_TieRodCylinder = _oExApplication.Sheets(17)

            _oExcelSheet_MTHL_Tube = _oExApplication.Sheets(18)

            _oExcelSheet_MTHL_Rod = _oExApplication.Sheets(19)

        Catch ex As Exception
            CreateExcel = False
            MessageBox.Show("Unable to open Excel sheet", "Information", MessageBoxButtons.OK, _
            MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Function

    Private Function KillExcel() As Boolean
        Try
            For Each oProcess As Process In Process.GetProcessesByName("Excel")
                oProcess.Kill()
                GC.Collect()
                System.Threading.Thread.Sleep(100)
            Next
            KillExcel = True
        Catch ex As Exception
            KillExcel = False
        End Try
    End Function

#End Region

End Class
