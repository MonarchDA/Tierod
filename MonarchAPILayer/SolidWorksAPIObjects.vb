Imports MonarchSolidworksLayer
Imports System.io
Imports MonarchFunctionalLayer
Imports MonarchDatabaseLayer               '22_09_2009  ragava
Imports System.Data.OleDb               '22_09_2009  ragava

Public Module SolidWorksAPIObjects

#Region "Variables"
    Dim _oSolidWorksBaseClass As MonarchSolidworksLayer.IFLSolidWorksBaseClass
    Public oTieRodCylinder As New TieRodCylinder
    Private _strErrorMessage As String

    Dim strRodDrawingPath As String = String.Empty           '22_09_2009  ragava
    Dim strTieRodDrawingPath As String = String.Empty           '22_09_2009  ragava
    Dim strStopTubeDrawingPath As String = String.Empty           '22_09_2009  ragava
    Dim strBoreDrawingPath As String = String.Empty           '22_09_2009  ragava

#End Region

#Region "Properties"
    Public Property ErrorMessage() As String
        Get
            Return _strErrorMessage
        End Get
        Set(ByVal value As String)
            _strErrorMessage = value
        End Set
    End Property

    Public Sub SolidWorksBaseClassNothing()

        _oSolidWorksBaseClass = Nothing

    End Sub

    Public ReadOnly Property IFLSolidWorksBaseClassObject() As Object
        Get
            If _oSolidWorksBaseClass Is Nothing Then
                '_oSolidWorksBaseClass = New MonarchSolidworksLayer.IFLSolidWorksBaseClass(True)
                _oSolidWorksBaseClass = New MonarchSolidworksLayer.IFLSolidWorksClass
                _oSolidWorksBaseClass.ConnectSolidWorks()
            End If
            Return _oSolidWorksBaseClass
        End Get
    End Property
#End Region
    Public Sub ProcessDirectory(ByVal targetDirectory As String, Optional ByVal blnSearchSubDir As Boolean = False)

        Dim arrAsmFileEntries As String()
        Dim arrPartFileEntries As String()
        Dim sCompName As String()
        Dim intI As Integer
        Dim intJ As Integer
        intI = 0
        intJ = 0
        Dim strSplit
        Dim strFileName As String
        arrAsmFileEntries = Nothing
        arrPartFileEntries = Nothing
        'IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CommandInProgress = True
        Dim blnRet As Boolean = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.SetCurrentWorkingDirectory(DestinationFilePath)
        'IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CommandInProgress = False
        Dim arrFileEntries As String() = Directory.GetFiles(targetDirectory)
        ReDim sCompName(0)
        strFileName = Nothing

        For intCount As Integer = 0 To UBound(arrFileEntries)
            strSplit = Split(arrFileEntries(intCount), ".")
            If UCase(strSplit(UBound(strSplit))) = UCase("SLDASM") Then
                If arrAsmFileEntries Is Nothing Then
                    ReDim arrAsmFileEntries(intI)
                    arrAsmFileEntries(intI) = arrFileEntries(intCount)
                Else
                    intI = intI + 1
                    ReDim Preserve arrAsmFileEntries(intI)
                    arrAsmFileEntries(intI) = arrFileEntries(intCount)
                End If
            ElseIf UCase(strSplit(UBound(strSplit))) = UCase("SLDPRT") Then
                If arrPartFileEntries Is Nothing Then
                    ReDim arrPartFileEntries(intJ)
                    arrPartFileEntries(intJ) = arrFileEntries(intCount)
                Else
                    intJ = intJ + 1
                    ReDim Preserve arrPartFileEntries(intJ)
                    arrPartFileEntries(intJ) = arrFileEntries(intCount)
                End If
            End If
        Next intCount

        If Not arrPartFileEntries Is Nothing Then
            For intCount As Integer = 0 To UBound(arrPartFileEntries)
                If arrPartFileEntries(intCount).ToString.IndexOf("~$") = -1 Then         '02_10_2009    ragava
                    updatePartModels(arrPartFileEntries(intCount))
                End If
            Next
        End If
        If Not arrAsmFileEntries Is Nothing Then
            For intCount As Integer = 0 To UBound(arrAsmFileEntries)
                If arrAsmFileEntries(intCount).ToString.IndexOf("~$") = -1 Then         '02_10_2009    ragava
                    'Dim strpath As String = arrAsmFileEntries(intCount).Replace("~$", "")


                    '09_04_2010   RAGAVA
                    If arrAsmFileEntries(intCount).ToString.IndexOf("TIE_ROD_ASSEMBLY") <> -1 Then
                        Try
                            KillExcel()
                        Catch ex As Exception

                        End Try
                        Try
                            IFLSolidWorksBaseClassObject.KillAllSolidWorksServices()
                        Catch ex As Exception

                        End Try
                        Try
                            IFLSolidWorksBaseClassObject.ConnectSolidWorks()
                            System.Threading.Thread.Sleep(2000)
                            'IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CommandInProgress = False
                        Catch ex As Exception

                        End Try
                    End If

                    blnRet = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.SetCurrentWorkingDirectory(DestinationFilePath)

                    IFLSolidWorksBaseClassObject.openDocument(arrAsmFileEntries(intCount))
                    IFLSolidWorksBaseClassObject.SolidWorksModel.ViewZoomtofit2()            '02_09_2009   ragava

                    Try
                        updateCustomProperties(arrAsmFileEntries(intCount))         '08_09_2009  ragava
                    Catch ex As Exception
                        MsgBox("Error in Updating Custom Properties")
                    End Try
                    IFLSolidWorksBaseClassObject.SolidWorksModel = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.ActivateDoc(arrAsmFileEntries(intCount))
                    IFLSolidWorksBaseClassObject.SolidWorksModel.EditRebuild3()             '02_09_2009   ragava
                    Try
                        IFLSolidWorksBaseClassObject.DeleteDesignTable(arrAsmFileEntries(intCount))        '02_09_2009   ragava
                    Catch ex As Exception

                    End Try
                    IFLSolidWorksBaseClassObject.Common_TraversAndDeletions_And_SuppressionParts()
                    If arrAsmFileEntries(intCount).ToString.IndexOf("TIE_ROD_ASSEMBLY") <> -1 Then
                        ''31_08_2012   RAGAVA
                        'Try
                        '    IFLSolidWorksBaseClassObject.SolidWorksModel.EditRebuild3()
                        '    IFLSolidWorksBaseClassObject.SolidWorksModel.SaveSilent()
                        '    Dim strDrawingFileName As String = ""
                        '    If ClevisCapPortOrientation.IndexOf("90") <> -1 Then
                        '        strDrawingFileName = DestinationFilePath + "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY_90.SLDDRW"
                        '    Else
                        '        strDrawingFileName = DestinationFilePath + "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDDRW"
                        '    End If
                        '    IFLSolidWorksBaseClassObject.openAssemblyDrawingDocument(strDrawingFileName)
                        '    IFLSolidWorksBaseClassObject.SolidWorksModel = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.ActiveDoc
                        '    IFLSolidWorksBaseClassObject.SolidWorksModel.EditRebuild3()
                        '    IFLSolidWorksBaseClassObject.SolidWorksModel.SaveSilent()
                        '    IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.ActivateDoc(arrAsmFileEntries(intCount))
                        'Catch ex As Exception
                        'End Try
                        'Dim strDrwName As String = arrAsmFileEntries(intCount)
                        'Dim iVersion As Integer = 0
                        'IFLSolidWorksBaseClassObject.SolidWorksModel.SaveAs2(strDrwName, iVersion, False, True)
                        IFLSolidWorksBaseClassObject.SaveAndKill()
                    Else
                        IFLSolidWorksBaseClassObject.SaveAndCloseAllDocuments()
                    End If
                End If
                    'Opening Drawing
                    'IFLSolidWorksBaseClassObject.openDocument(arrAsmFileEntries(intCount).Replace(".SLDASM", ".SLDDRW"))           '02_09_2009   ragava
            Next
       
        End If

        If blnSearchSubDir = True Then
            Dim subdirectoryEntries As String() = Directory.GetDirectories(targetDirectory)
            Dim subdirectory As String
            For Each subdirectory In subdirectoryEntries
                ProcessDirectory(subdirectory)
            Next subdirectory
        End If

    End Sub

    'Public Function getRevisionTableData() As ArrayList
    '    getRevisionTableData = Nothing
    '    getRevisionTableData = New ArrayList
    '    Dim strSQL As String
    '    Dim intCount As Integer
    '    Dim objDT As DataTable
    '    Dim oDataClass As New DataClass
    '    strSQL = "select  *  from revisionTable  where contractNumber='" & ContractNumber & "' and RevisionNumber=" & revisionNumber + 1
    '    objDT = oDataClass.GetDataTable(strSQL)
    '    If Not IsNothing(objDT) AndAlso objDT.Rows.Count < 7 Then
    '        intCount = objDT.Rows.Count
    '    Else
    '        intCount = 7
    '    End If
    '    For intj As Integer = 0 To intCount - 1
    '        getRevisionTableData.Add(New Object(0) {objDT.Rows(intj)("RevisionNumber")})
    '        getRevisionTableData.Add(New Object(0) {objDT.Rows(intj)("ECR_Number")})
    '        getRevisionTableData.Add(New Object(0) {objDT.Rows(intj)("DESCRIPTION")})
    '        getRevisionTableData.Add(New Object(0) {objDT.Rows(intj)("DATE")})
    '        getRevisionTableData.Add(New Object(0) {objDT.Rows(intj)("REVISEDBy")})
    '    Next
    '    Return getRevisionTableData
    'End Function
    Public Function getRevisionTableData() As ArrayList

        getRevisionTableData = Nothing
        getRevisionTableData = New ArrayList
        Dim strSQL As String
        Dim intCount As Integer
        Dim objDT As DataTable
        Dim oDataClass As New DataClass
        strSQL = "select  top 7  * from revisionTable  where contractNumber='" & ContractNumber & "'"
        objDT = oDataClass.GetDataTable(strSQL)
        If Not IsNothing(objDT) AndAlso objDT.Rows.Count < 7 Then
            intCount = objDT.Rows.Count
        Else
            intCount = 7
        End If
        For intj As Integer = 0 To intCount - 1
            getRevisionTableData.Add(New Object(4) {objDT.Rows(intj)("RevisionNumber"), objDT.Rows(intj)("ECR_Number"), objDT.Rows(intj)("DESCRIPTION"), objDT.Rows(intj)("DATE"), objDT.Rows(intj)("REVISEDBy")})
        Next
        Return getRevisionTableData

    End Function

    Public Sub UpdateRevisionTable(ByVal Description As String, ByVal Revision As String)

        Try
            Dim strDate As String = UCase(Format(Date.Today, "dMMMyy"))
            Dim strDescription As String
            Dim PropertyName As String
            PropertyName = "DESCRIPTION_"

            'Checking For Empty Description
            For i As Integer = 1 To 7
                strDescription = IFLSolidWorksBaseClassObject.SolidWorksModel.GetCustomInfoValue("", PropertyName & i.ToString)
                If Trim(strDescription) = "" Then
                    checkProperty("DESCRIPTION_" & i.ToString, "ADD CODE # " & Description)
                    checkProperty("NO_" & i.ToString, Revision)
                    checkProperty("DATE_" & i.ToString, strDate)
                    Exit Sub
                End If
            Next
            'Checking For RevisionNumber
            PropertyName = "NO_"
            Dim iRevision, iRevision7 As Integer
            Dim strRevision As String = String.Empty
            For i As Integer = 7 To 1 Step -1        '23_10_2009  ragava
                strRevision = IFLSolidWorksBaseClassObject.SolidWorksModel.GetCustomInfoValue("", PropertyName & i.ToString)
                If Trim(strRevision) = "-" Then
                    checkProperty("DESCRIPTION_" & i.ToString, "ADD CODE # " & Description)
                    checkProperty("NO_" & i.ToString, Revision)
                    checkProperty("DATE_" & i.ToString, strDate)
                    Exit Sub
                End If
            Next
            strRevision = IFLSolidWorksBaseClassObject.SolidWorksModel.GetCustomInfoValue("", PropertyName & "7")
            iRevision7 = Convert.ToInt16(strRevision)
            For i As Integer = 2 To 7                 '23_10_2009   ragava
                strRevision = IFLSolidWorksBaseClassObject.SolidWorksModel.GetCustomInfoValue("", PropertyName & i.ToString)
                If Trim(strRevision) <> "-" Then
                    iRevision = Convert.ToInt16(strRevision)
                    If iRevision <= iRevision7 Then
                        '23_10_2009   ragava
                        If i < 7 Then
                            For k As Integer = i To 6
                                strRevision = IFLSolidWorksBaseClassObject.SolidWorksModel.GetCustomInfoValue("", "NO_" & (k + 1).ToString)
                                strDescription = IFLSolidWorksBaseClassObject.SolidWorksModel.GetCustomInfoValue("", "DESCRIPTION_" & (k + 1).ToString)
                                Dim MyDate As String = IFLSolidWorksBaseClassObject.SolidWorksModel.GetCustomInfoValue("", "DATE_" & (k + 1).ToString)
                                checkProperty("DESCRIPTION_" & k.ToString, "ADD CODE # " & strDescription)
                                checkProperty("NO_" & k.ToString, strRevision)
                                checkProperty("DATE_" & k.ToString, MyDate)
                            Next
                        End If
                        '23_10_2009   ragava    Till   Here
                        checkProperty("DESCRIPTION_" & i.ToString, "ADD CODE # " & Description)
                        checkProperty("NO_" & i.ToString, Revision)
                        checkProperty("DATE_" & i.ToString, strDate)
                        Exit Sub
                    End If
                End If
            Next
        Catch ex As Exception
        End Try

    End Sub

    Public Sub updatePartModels(ByVal fileName As String)

        Dim strpath As String = fileName.Replace("~$", "")
        If Not String.Compare(fileName, strpath) = 0 Then
            Exit Sub
        End If
        Dim blnRet As Boolean = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.SetCurrentWorkingDirectory(DestinationFilePath)
        'IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CommandInProgress = True
        Try
            'IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CommandInProgress = True
            '10_09_2009   ragava
            Try
                Dim strRodDia As String = String.Empty
                If fileName.IndexOf("\ROD\") <> -1 Then
                    strRodDia = dblRodDiameter.ToString
                    If dblRodDiameter = 1.12 Then
                        strRodDia = "1.125"
                    ElseIf dblRodDiameter = 1.5 Then
                        strRodDia = "1.50"
                    ElseIf dblRodDiameter = 1.38 Then
                        strRodDia = "1.375"
                    ElseIf dblRodDiameter = 2 Then
                        strRodDia = "2.00"
                    End If
                    If (fileName.IndexOf("RodDia_" & strRodDia & IIf(Series.IndexOf("TX") <> -1, "TX", "")) = -1) Then
                        Try
                            'File.Delete(fileName)
                        Catch ex As Exception

                        End Try
                        Exit Sub
                    End If
                End If
                '10_09_2009   ragava   Till  Here
            Catch ex As Exception

            End Try
            '13_04_2010   RAGAVA
            Try
                If fileName.IndexOf("\ROD\") <> -1 Then
                    Dim strRodFile As String
                    strRodFile = fileName.Substring(0, fileName.LastIndexOf("\")) & "\RodDia_1.125.SLDPRT"
                    IFLSolidWorksBaseClassObject.openDocument(strRodFile)
                    'IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CommandInProgress = False
                    IFLSolidWorksBaseClassObject.SolidWorksModel.EditRebuild3()
                    IFLSolidWorksBaseClassObject.DeleteDesignTable(strRodFile)
                    IFLSolidWorksBaseClassObject.SaveAndCloseAllDocuments()
                End If
            Catch ex As Exception

            End Try
            IFLSolidWorksBaseClassObject.openDocument(fileName)
            'IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CommandInProgress = False
            If Not IFLSolidWorksBaseClassObject.SolidWorksModel Is Nothing Then
                IFLSolidWorksBaseClassObject.SolidWorksModel.ViewZoomtofit2()
                Try
                    updateCustomProperties(fileName)         '08_09_2009  ragava
                Catch ex As Exception
                    MsgBox("Error in Updating Custom Properties")
                End Try
                IFLSolidWorksBaseClassObject.SolidWorksModel.EditRebuild3()
                Try
                    IFLSolidWorksBaseClassObject.DeleteDesignTable(fileName)
                    'IFLSolidWorksBaseClassObject.Save()
                    IFLSolidWorksBaseClassObject.SaveAndCloseAllDocuments()
                Catch ex As Exception
                End Try
            End If
        Catch ex As Exception
        End Try
    End Sub

    '08_09_2009   Ragava
    Public Sub updateCustomProperties(ByVal strFileName As String)
        'Updating Custom Properties
        Try
            Dim strMaterial, strDescription, strCode, strMar, strDesignation As String
            Dim strAssyNotes, strPaintNotes, strGeneralNotes As String
            strAssyNotes = String.Empty      '11_09_2009  ragava
            strPaintNotes = String.Empty      '11_09_2009  ragava
            strGeneralNotes = String.Empty      '11_09_2009  ragava
            strMaterial = String.Empty
            strDescription = String.Empty
            strCode = String.Empty
            strMar = String.Empty
            strDesignation = String.Empty      '11_09_2009  ragava
            'strDrawingNumber = String.Empty          '11_09_2009  ragava
            If strFileName.IndexOf("\ROD\") <> -1 Then
                If UCase(strFileName).IndexOf(".SLDPRT") <> -1 Then
                    Volume_Rod = Val(IFLSolidWorksBaseClassObject.SolidWorksModel.GetCustomInfoValue(Nothing, "Volume"))
                    Weight_Rod = Math.Round(Volume_Rod * 0.27812, 4)       '04_10_2010   RAGAVA
                End If
                strMaterial = strRodMaterial
                strCode = RodDrawingNumber    '11_09_2009  ragava
                strMar = RodMaterialNumber    '11_09_2009  ragava
                'strDescription = "ROD CYL " & Format(dblRodDiameter, "0.00").ToString & "-" & Format(dblStrokeLength, "00.00").ToString & "-" & Format(dblPistonThreadSize, "0.00").ToString & "-" & Format(dblRodThreadSize, "0.00").ToString        '11_09_2009  ragava
                strDescription = "ROD CYL " & dblRodDiameter & "-" & Format(dblStrokeLength, "00.00").ToString & "-" & Format(dblPistonThreadSize, "0.00").ToString & "-" & Format(dblRodThreadSize, "0.00").ToString        '01_10_2009   ragava
            ElseIf strFileName.IndexOf("\Bore\") <> -1 Then
                If UCase(strFileName).IndexOf(".SLDASM") <> -1 Then
                    Volume_Bore = Val(IFLSolidWorksBaseClassObject.SolidWorksModel.GetCustomInfoValue(Nothing, "Volume"))
                    Weight_Bore = Math.Round(Volume_Bore * 0.27812, 4)       '04_10_2010   RAGAVA
                End If
                strMaterial = ""
                strCode = BoreDrawingNumber    '11_09_2009  ragava
            ElseIf strFileName.IndexOf("\Tie Rod\") <> -1 Then
                If UCase(strFileName).IndexOf(".SLDPRT") <> -1 Then
                    Volume_TieRod = Val(IFLSolidWorksBaseClassObject.SolidWorksModel.GetCustomInfoValue(Nothing, "Volume"))
                    Weight_TieRod = Math.Round(Volume_TieRod * 0.27812, 4)       '04_10_2010   RAGAVA
                End If
                strMaterial = ""
                strCode = TieRodDrawingNumber    '11_09_2009  ragava
            ElseIf strFileName.IndexOf("\Stoptube\") <> -1 Then
                If UCase(strFileName).IndexOf(".SLDPRT") <> -1 Then
                    Volume_StopTube = Val(IFLSolidWorksBaseClassObject.SolidWorksModel.GetCustomInfoValue(Nothing, "Volume"))
                    Weight_StopTube = Math.Round(Volume_StopTube * 0.27812, 4)       '04_10_2010   RAGAVA
                End If
                strMaterial = ""
            ElseIf strFileName.IndexOf("\TIE_ROD_ASSEMBLY\") <> -1 Then
                If UCase(strFileName).IndexOf(".SLDASM") <> -1 Then
                    Volume_Assembly = Val(IFLSolidWorksBaseClassObject.SolidWorksModel.GetCustomInfoValue(Nothing, "Volume"))
                    Weight_Assembly = Math.Round(Volume_Assembly * 0.27812, 4)       '04_10_2010   RAGAVA
                End If
                strMaterial = ""
                strDesignation = strCylinderDescription             '11_09_2009  ragava
                strDescription = "<MOD-DIAM>" & Format(dblBoreDiameter, "0.00").ToString & " CYLINDER"   '31_08_2012  RAGAVA "<MOD-DIAM>" & added        '11_09_2009  ragava
                strCode = PartCode
                strAssyNotes = strAssemblyNotes           '11_09_2009  ragava
                strPaintNotes = strPaintPackagingNotes           '11_09_2009  ragava
                strGeneralNotes = GeneralNotes           '11_09_2009  ragava
            End If
            If strFileName.IndexOf(".SLDASM") <> -1 Then
                If strFileName.IndexOf("\Bore\") <> -1 Then     '11_09_2009  ragava
                    strMaterial = "C1026"          '11_09_2009  ragava
                End If
                strMar = ""               '11_09_2009  ragava
                If strFileName.IndexOf("\TIE_ROD_ASSEMBLY\") <> -1 Then
                    checkProperty("Painting_Optional_Note", strPaintNotes)
                    checkProperty("Assembly_Optional_Note", strAssyNotes)
                    checkProperty("General_Note", strGeneralNotes)        '14_09_2009  ragava

                    '23_02_2010   RAGAVA
                    If dblRodThreadSize = 1 Then
                        checkProperty("Thread_Description", "1-14 UNS-3A THREAD" & vbNewLine & "INSTALL VINYL TUBE ON" & vbNewLine & "EXPOSED THREADS" & vbNewLine & "AFTER WASHING")
                    ElseIf dblRodThreadSize = 1.12 Then
                        checkProperty("Thread_Description", "1 1/8-12UNF-3A THREAD" & vbNewLine & "INSTALL VINYL TUBE ON" & vbNewLine & "EXPOSED THREADS" & vbNewLine & "AFTER WASHING")
                    ElseIf dblRodThreadSize = 1.25 Then
                        checkProperty("Thread_Description", "1 1/4-12UNF-3A THREAD" & vbNewLine & "INSTALL VINYL TUBE ON" & vbNewLine & "EXPOSED THREADS" & vbNewLine & "AFTER WASHING")
                    ElseIf dblRodThreadSize = 1.38 Then
                        checkProperty("Thread_Description", "1 3/8-12UNF-3A THREAD" & vbNewLine & "INSTALL VINYL TUBE ON" & vbNewLine & "EXPOSED THREADS" & vbNewLine & "AFTER WASHING")
                    ElseIf dblRodThreadSize = 1.5 Then
                        checkProperty("Thread_Description", "1 1/2-12UNF-3A THREAD" & vbNewLine & "INSTALL VINYL TUBE ON" & vbNewLine & "EXPOSED THREADS" & vbNewLine & "AFTER WASHING")
                    End If
                    'Till    Here


                    checkProperty("Customer_Name", CustomerName)
                    checkProperty("Designation", strDesignation)          '11_09_2009  ragava
                    '14_09_2009  ragava
                    If strRodMaterial.IndexOf("CHROME PLATED") <> -1 Then
                        checkProperty("Rod_Material", "CHROME")      '01_10_2009   ragava
                    ElseIf strRodMaterial.IndexOf("NITRO") <> -1 Then
                        'checkProperty("Rod_Material", "NITROSTEEL")     '01_10_2009   ragava
                        checkProperty("Rod_Material", "NITRIDE")     '31_08_2012   RAGAVA
                    ElseIf strRodMaterial.IndexOf("LION") <> -1 Then     '08_07_2010   ragava
                        checkProperty("Rod_Material", "LION CHROME")     '08_07_2010   ragava
                    ElseIf strRodMaterial.IndexOf("ROD BLANK ") <> -1 Then
                        checkProperty("Rod_Material", "INDUCTION HARDENED")     '01_10_2009   ragava
                    End If
                    If ClevisCapPort = "9/16 ORB" Then
                        checkProperty("Clevis_Cap_Port", "PORT 9/16-18UNF ORB")
                    ElseIf ClevisCapPort = "3/4 ORB" Then
                        checkProperty("Clevis_Cap_Port", "PORT 3/4-16UNF ORB")
                    Else
                        checkProperty("Clevis_Cap_Port", "PORT " & ClevisCapPort)
                    End If
                    If RodCapPort = "9/16 ORB" Then
                        checkProperty("Rod_Cap_Port", "PORT 9/16-18UNF ORB")
                    ElseIf RodCapPort = "3/4 ORB" Then
                        checkProperty("Rod_Cap_Port", "PORT 3/4-16UNF ORB")
                    Else
                        checkProperty("Rod_Cap_Port", "PORT " & RodCapPort)
                    End If

                    '02_11_2009  Ragava
                    If pinHoleType = "Bushing" Then
                        checkProperty("CLEVIS_PIN_HOLE_NOTE", "<MOD-DIAM>" & Format(PinHoleSize, "0.00").ToString & " PIN WITH" + vbNewLine + "HARDENED BUSHINGS")
                    Else
                        '11_11_2009   Ragava
                        'checkProperty("CLEVIS_PIN_HOLE_NOTE", Format(PinHoleSize + 0.015, "0.000").ToString & " PIN HOLE")
                        If ClevisPins = False Then
                            checkProperty("CLEVIS_PIN_HOLE_NOTE", "<MOD-DIAM>" & Format(PinHoleSize + 0.015, "0.000").ToString & " PIN HOLE")
                        Else
                            '07_04_2010    RAGAVA
                            If blnPinsPlasticBag = True Then
                                checkProperty("CLEVIS_PIN_HOLE_NOTE", "<MOD-DIAM>" & Format(PinHoleSize + 0.015, "0.000").ToString & " PIN HOLE")
                            Else
                                checkProperty("CLEVIS_PIN_HOLE_NOTE", "<MOD-DIAM>" & Format(PinHoleSize, "0.00").ToString & " PIN C/W " & UCase(ClevisPinClips))
                            End If
                            'Till   Here
                        End If
                        '11_11_2009   Ragava   Till  Here
                    End If
                    If RodClevisPinHoleType = "Bushing" Then
                        checkProperty("ROD_PIN_HOLE_NOTE", "<MOD-DIAM>" & Format(PinHoleSize, "0.00").ToString & " PIN WITH" + vbNewLine + "HARDENED BUSHINGS")
                    Else
                        '11_11_2009   Ragava
                        'checkProperty("ROD_PIN_HOLE_NOTE", Format(PinHoleSize + 0.015, "0.000").ToString & " PIN HOLE")
                        '07_04_2010    RAGAVA
                        If RodClevisPins = False Then
                            checkProperty("ROD_PIN_HOLE_NOTE", "<MOD-DIAM>" & Format(PinHoleSize + 0.015, "0.000").ToString & " PIN HOLE")
                        Else
                            If blnPinsPlasticBag = True Then
                                checkProperty("ROD_PIN_HOLE_NOTE", "<MOD-DIAM>" & Format(PinHoleSize + 0.015, "0.000").ToString & " PIN HOLE")
                            Else
                                checkProperty("ROD_PIN_HOLE_NOTE", "<MOD-DIAM>" & Format(PinHoleSize, "0.00").ToString & " PIN C/W " & UCase(RodPinClips))
                            End If
                        End If
                        'Till  Here
                    End If
                    '02_11_2009  Ragava      Till   Here

                    '14_09_2009  ragava  Till  Here
                End If
            End If
            checkProperty("Drawn", "IDOLA FORI")
            'checkProperty("Drawn", "IFL")
            'checkProperty("Designed", Environment.UserName.ToString)     '01_10_2009   ragava      Will Be Used In FINAL VERSION
            checkProperty("Designed", Environment.UserName.ToString)     '03_11_2009  Ragava
            checkProperty("Approved", "")
            checkProperty("Customer_Name", CustomerName)         '17_08_2012   RAGAVA
            'checkProperty("Date", System.DateTime.Today)
            checkProperty("Date", Format(Date.Today, "dMMMyy"))     '10_02_2010    RAGAVA

            checkProperty("Material", strMaterial)
            'checkProperty("Mar#", "")
            checkProperty("Mar#", strMar)               '11_09_2009   ragava
            checkProperty("Description", strDescription)
            checkProperty("Code", strCode)          '11_09_2009  ragava
            '01_10_2009  ragava
            If strFileName.IndexOf("\ROD\") <> -1 Then
                Dim strRodDia As String = String.Empty
                Dim strSeries As String = String.Empty
                strRodDia = dblRodDiameter.ToString
                If dblRodDiameter = 1.12 Then
                    strRodDia = "1.125"
                ElseIf dblRodDiameter = 1.5 Then
                    strRodDia = "1.50"
                ElseIf dblRodDiameter = 1.38 Then
                    strRodDia = "1.375"
                ElseIf dblRodDiameter = 2 Then
                    strRodDia = "2.00"
                End If
                If Series.IndexOf("TX") <> -1 Then
                    strSeries = "TX"
                End If
                If strRodDia & strSeries = "1.25" Then
                    If dblPistonThreadSize = 0.75 Then
                        checkProperty("Piston_Thread_Size", "3/4-16UNF-2A")
                    ElseIf dblPistonThreadSize = 1 Then
                        checkProperty("Piston_Thread_Size", "1-14UNS-2A")
                    End If
                    If dblRodThreadSize = 1.12 Then
                        checkProperty("Rod_Thread_Size", "1-1/8-12UNF-3A")
                    ElseIf dblRodThreadSize = 1.25 Then
                        checkProperty("Rod_Thread_Size", "1-1/4-12UNF-3A")
                    End If
                ElseIf strRodDia & strSeries = "1.25TX" Then
                    If dblPistonThreadSize = 1.13 Then
                        checkProperty("Piston_Thread_Size", "1-1/8-12UNF-3A  TYP. BOTH ENDS")
                    ElseIf dblPistonThreadSize = 1 Then
                        checkProperty("Piston_Thread_Size", "1-14UNS-3A")
                    End If
                ElseIf strRodDia & strSeries = "1.375" Then
                    If dblPistonThreadSize = 1 Then
                        checkProperty("Piston_Thread_Size", "1-14UNS-2A")
                    ElseIf dblPistonThreadSize = 1.12 Then
                        checkProperty("Piston_Thread_Size", "1-1/8-12UNF-2A")
                    End If
                    If dblRodThreadSize = 1.12 Then
                        checkProperty("Rod_Thread_Size", "1-1/8-12UNF-3A")
                    ElseIf dblRodThreadSize = 1.25 Then
                        checkProperty("Rod_Thread_Size", "1-1/4-12UNF-3A")
                    ElseIf dblRodThreadSize = 1.38 Then
                        checkProperty("Rod_Thread_Size", "1-3/8-12UNF-3A")
                    End If
                ElseIf strRodDia & strSeries = "1.50" Then
                    If dblPistonThreadSize = 1 Then
                        checkProperty("Piston_Thread_Size", "1-14UNS-2A")
                    ElseIf dblPistonThreadSize = 1.12 Then
                        checkProperty("Piston_Thread_Size", "1-1/8-12UNF-2A")
                    ElseIf dblPistonThreadSize = 1.13 Then
                        checkProperty("Piston_Thread_Size", "1-1/8-12UNF-2A")
                    ElseIf dblPistonThreadSize = 1.25 Then
                        checkProperty("Piston_Thread_Size", "1-1/4-12UNF-2A")
                    End If
                    If dblRodThreadSize = 1.12 Then
                        checkProperty("Rod_Thread_Size", "1-1/8-12UNF-3A")
                    ElseIf dblRodThreadSize = 1.25 Then
                        checkProperty("Rod_Thread_Size", "1-1/4-12UNF-3A")
                    ElseIf dblRodThreadSize = 1.5 Then
                        checkProperty("Rod_Thread_Size", "1-1/2-12UNF-3A")
                    End If
                ElseIf strRodDia & strSeries = "1.75" Then
                    If dblPistonThreadSize = 1 Then
                        checkProperty("Piston_Thread_Size", "1-14UNS-2A")
                    ElseIf dblPistonThreadSize = 1.12 Then
                        checkProperty("Piston_Thread_Size", "1-1/8-12UNF-2A")
                    ElseIf dblPistonThreadSize = 1.5 Then
                        checkProperty("Piston_Thread_Size", "1-1/2-12UNF-2A")
                    End If
                    If dblRodThreadSize = 1.25 Then
                        checkProperty("Rod_Thread_Size", "1-1/4-12UNF-3A")
                    ElseIf dblRodThreadSize = 1.5 Then
                        checkProperty("Rod_Thread_Size", "1-1/2-12UNF-3A")
                    End If
                ElseIf strRodDia & strSeries = "2.00" Then
                    If dblPistonThreadSize = 1 Then
                        checkProperty("Piston_Thread_Size", "1-14UNS-2A")
                    ElseIf dblPistonThreadSize = 1.12 Then
                        checkProperty("Piston_Thread_Size", "1-1/8-12UNF-2A")
                    ElseIf dblPistonThreadSize = 1.25 Then
                        checkProperty("Piston_Thread_Size", "1-1/4-12UNF-2A")
                    ElseIf dblPistonThreadSize = 1.5 Then
                        checkProperty("Piston_Thread_Size", "1-1/2-12UNF-2A")
                    End If
                    If dblRodThreadSize = 1.25 Then
                        checkProperty("Rod_Thread_Size", "1-1/4-12UNF-3A")
                    ElseIf dblRodThreadSize = 1.5 Then
                        checkProperty("Rod_Thread_Size", "1-1/2-12UNF-3A")
                    End If
                End If
            End If
            '01_10_2009  ragava   Till  Here

        Catch ex As Exception
            MsgBox("Error in Updating Notes to CustomProperties")
        End Try

    End Sub
    ''08_09_2009   Ragava
    'Public Sub updateCustomProperties(ByVal strFileName As String)
    '    'Updating Custom Properties
    '    Try
    '        Dim strMaterial, strDescription, strCode As String
    '        strMaterial = String.Empty
    '        strDescription = String.Empty
    '        strCode = String.Empty
    '        If strFileName.IndexOf("\ROD\") <> -1 Then
    '            strMaterial = strRodMaterial
    '        ElseIf strFileName.IndexOf("\Bore\") <> -1 Then
    '            strMaterial = "BLANK"
    '        ElseIf strFileName.IndexOf("\Tie Rod\") <> -1 Then
    '            strMaterial = "BLANK"
    '        ElseIf strFileName.IndexOf("\Stoptube\") <> -1 Then
    '            strMaterial = "BLANK"
    '        ElseIf strFileName.IndexOf("\TIE_ROD_ASSEMBLY\") <> -1 Then
    '            strMaterial = "BLANK"
    '            strDescription = strCylinderDescription
    '            strCode = PartCode
    '        End If
    '        If strFileName.IndexOf(".SLDASM") <> -1 Then
    '            checkProperty("Painting_Optional_Note", strPaintPackagingNotes)
    '            checkProperty("Assembly_Optional_Note", strAssemblyNotes)
    '            checkProperty("Customer_Name", CustomerName)
    '        End If
    '        checkProperty("Drawn", "IDOLAFORI")
    '        checkProperty("Designed", CustomerName)
    '        checkProperty("Approved", "BLANK")
    '        checkProperty("Date", System.DateTime.Today)
    '        checkProperty("Material", strMaterial)
    '        checkProperty("Mar#", "BLANK")
    '        checkProperty("Description", strDescription)
    '        checkProperty("Code", strCode)


    '    Catch ex As Exception
    '        MsgBox("Error in Updating Notes to CustomProperties")
    '    End Try
    'End Sub
    '08_09_2009 Ragava
    Public Sub checkProperty(ByVal propertyName As String, ByVal value As Object)

        Try
            IFLSolidWorksBaseClassObject.SolidWorksModel.DeleteCustomInfo(propertyName)
            IFLSolidWorksBaseClassObject.SolidWorksModel.AddCustomInfo(propertyName, "Text", value)
        Catch ex As Exception
        End Try

    End Sub

    '15_09_2009   ragava
    Public Sub DrawingUpdation()

        Try
            'IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CommandInProgress = True
            IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CloseAllDocuments(True)               '06_10_2009    ragava
            'IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CommandInProgress = False
            System.Threading.Thread.Sleep(2000)     '09_04_2010   RAGAVA
        Catch ex As Exception
        End Try
        'Rod Drawing
        Dim blnRet As Boolean = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.SetCurrentWorkingDirectory(DestinationFilePath)
        Dim arrFileEntries_Rod As String() = Directory.GetFiles(DestinationFilePath + "\ROD")
        Dim arrFileEntries_Bore As String() = Directory.GetFiles(DestinationFilePath + "\Bore")       '16_09_2009   ragava
        Dim arrFileEntries_TieRod As String() = Directory.GetFiles(DestinationFilePath + "\Tie Rod")

        Dim strFileType As String()
        Dim strRodDrawingFile As String = String.Empty
        Dim strRodDia As String = String.Empty
        Dim strSeries As String = String.Empty
        strRodDia = dblRodDiameter.ToString
        If dblRodDiameter = 1.12 Then
            strRodDia = "1.125"
        ElseIf dblRodDiameter = 1.5 Then
            strRodDia = "1.50"
        ElseIf dblRodDiameter = 1.38 Then
            strRodDia = "1.375"
        ElseIf dblRodDiameter = 2 Then
            strRodDia = "2.00"
        End If
        If Series.IndexOf("TX") <> -1 Then
            strSeries = "TX"
        End If


        'For Rod Drawing Updation
        Try
            For intCount As Integer = 0 To UBound(arrFileEntries_Rod)
                strFileType = Split(arrFileEntries_Rod(intCount), ".")
                If UCase(strFileType(UBound(strFileType))) = UCase("SLDDRW") Then
                    If arrFileEntries_Rod(intCount).IndexOf("RodDia_" & strRodDia & strSeries & ".") <> -1 Then
                        Dim strDrawingFileName As String = arrFileEntries_Rod(intCount)
                        IFLSolidWorksBaseClassObject.openDocument(strDrawingFileName)
                        'IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CommandInProgress = False
                        System.Threading.Thread.Sleep(2000)     '09_04_2010   RAGAVA
                        If Not IFLSolidWorksBaseClassObject.SolidWorksModel Is Nothing Then
                            IFLSolidWorksBaseClassObject.SolidWorksModel.EditRebuild3()
                            System.Threading.Thread.Sleep(1000)         '06_10_2009   ragava
                            strRodDrawingPath = strDrawingFileName          '22_09_2009   ragava
                            '16_09_2009  ragava
                            Try
                                If RodLength >= 27 Then
                                    IFLSolidWorksBaseClassObject.BreakView("Drawing View2", RodLength, 25)
                                End If
                            Catch ex As Exception
                                MsgBox("Error in Breaking the View")
                            End Try


                            'IFLSolidWorksBaseClassObject.SolidWorksDrawingDocument = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.ActiveDoc      'for Testing
                            Dim strRodMat As String = String.Empty
                            If strRodDia & strSeries = "1.125" Then
                                If strRodMaterial.IndexOf("CHROME PLATED") <> -1 Then
                                    strRodMat = "(HARD CHROME PLATE)"
                                ElseIf strRodMaterial.IndexOf("NITRO") <> -1 Then
                                    strRodMat = "(NITROSTEEL TREATMENT)"
                                End If
                                IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD2@Drawing View2", strRodMat, "Drawing View2")
                                If RodDrawingNumber = "492179" Then
                                    IFLSolidWorksBaseClassObject.DeleteDetailItem("DetailItem150@Drawing View2", "Drawing View2")
                                Else
                                    IFLSolidWorksBaseClassObject.DeleteDetailItem("DetailItem151@Drawing View2", "Drawing View2")
                                End If
                            ElseIf strRodDia & strSeries = "1.25" Then
                                If strRodMaterial.IndexOf("CHROME PLATED") <> -1 Then
                                    strRodMat = "(HARD CHROME PLATE)"
                                ElseIf strRodMaterial.IndexOf("NITRO") <> -1 Then
                                    strRodMat = "(NITROSTEEL TREATMENT)"
                                ElseIf strRodMaterial.IndexOf("-08-I") <> -1 Then
                                    strRodMat = "(INDUCTION HARDENED)"
                                End If
                                IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD5@Drawing View2", strRodMat, "Drawing View2")
                                If dblPistonThreadSize = 0.75 Then
                                    checkProperty("Piston_Thread_Size", "3/4-16UNF-2A")
                                    IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD7@Drawing View2", "(Ø.751 INSIDE THIS AREA)", "Drawing View2")
                                ElseIf dblPistonThreadSize = 1 Then
                                    checkProperty("Piston_Thread_Size", "1-14UNS-2A")
                                    IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD7@Drawing View2", "(Ø1.001 INSIDE THIS AREA)", "Drawing View2")
                                End If
                                If dblRodThreadSize = 1.12 Then
                                    checkProperty("Rod_Thread_Size", "1-1/8-12UNF-3A")
                                ElseIf dblRodThreadSize = 1.25 Then
                                    checkProperty("Rod_Thread_Size", "1-1/4-12UNF-3A")
                                End If
                                If strRodMaterial.IndexOf("-08-I") = -1 Then
                                    IFLSolidWorksBaseClassObject.DeleteDimension("RD23@Drawing View2", "Drawing View2")
                                    IFLSolidWorksBaseClassObject.DeleteDimension("RD24@Drawing View2", "Drawing View2")
                                    IFLSolidWorksBaseClassObject.DeleteDimension("RD25@Drawing View2", "Drawing View2")
                                    IFLSolidWorksBaseClassObject.DeleteDimension("RD26@Drawing View2", "Drawing View2")
                                    IFLSolidWorksBaseClassObject.DeleteDimension("RD27@Drawing View2", "Drawing View2")
                                End If
                            ElseIf strRodDia & strSeries = "1.25TX" Then
                                If strRodMaterial.IndexOf("CHROME PLATED") <> -1 Then
                                    strRodMat = "(HARD CHROME PLATE)"
                                ElseIf strRodMaterial.IndexOf("NITRO") <> -1 Then
                                    strRodMat = "(NITROSTEEL TREATMENT)"
                                End If
                                IFLSolidWorksBaseClassObject.OverwriteDimensionNote("D2@Sketch9@RodDia_1.25TX.SLDDRW", strRodMat, "Drawing View2")
                                If dblPistonThreadSize = 1.13 Then
                                    checkProperty("Piston_Thread_Size", "1-1/8-12UNF-3A  TYP. BOTH ENDS")
                                    IFLSolidWorksBaseClassObject.DeleteView("Drawing View15")
                                    IFLSolidWorksBaseClassObject.DeleteView("Drawing View16")
                                    IFLSolidWorksBaseClassObject.DeleteView("Drawing View17")
                                ElseIf dblPistonThreadSize = 1 Then
                                    checkProperty("Piston_Thread_Size", "1-14UNS-3A")
                                    IFLSolidWorksBaseClassObject.DeleteView("Drawing View12")
                                    IFLSolidWorksBaseClassObject.DeleteView("Drawing View13")
                                    IFLSolidWorksBaseClassObject.DeleteView("Drawing View14")
                                End If
                            ElseIf strRodDia & strSeries = "1.375" Then
                                If strRodMaterial.IndexOf("CHROME PLATED") <> -1 Then
                                    strRodMat = "(HARD CHROME PLATE)"
                                ElseIf strRodMaterial.IndexOf("NITRO") <> -1 Then
                                    strRodMat = "(NITRO TEC TREATED .0010 THK 64-71 RC FINISH 16 RMS)"
                                End If
                                IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD10@Drawing View2", strRodMat, "Drawing View2")
                                IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD17@Drawing View2", strRodMat, "Drawing View2")
                                IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD24@Drawing View2", strRodMat, "Drawing View2")
                                If dblPistonThreadSize = 1 Then
                                    checkProperty("Piston_Thread_Size", "1-14UNS-2A")
                                    IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD13@Drawing View2", "(Ø1.001 INSIDE THIS AREA)", "Drawing View2")
                                ElseIf dblPistonThreadSize = 1.12 Then
                                    checkProperty("Piston_Thread_Size", "1-1/8-12UNF-2A")
                                    IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD13@Drawing View2", "(Ø1.126 INSIDE THIS AREA)", "Drawing View2")
                                End If
                                If dblRodThreadSize = 1.12 Then
                                    checkProperty("Rod_Thread_Size", "1-1/8-12UNF-3A")
                                ElseIf dblRodThreadSize = 1.25 Then
                                    checkProperty("Rod_Thread_Size", "1-1/4-12UNF-3A")
                                ElseIf dblRodThreadSize = 1.38 Then
                                    checkProperty("Rod_Thread_Size", "1-3/8-12UNF-3A")
                                End If
                                If RodDrawingNumber = "493883" Or RodDrawingNumber = "494844" Or RodDrawingNumber = "494843" Then
                                    'IFLSolidWorksBaseClassObject.DeleteDetailItem("DetailItem151@Drawing View2", "Drawing View2")
                                    IFLSolidWorksBaseClassObject.DeleteDetailItem("DetailItem146@Drawing View2", "Drawing View2")    '01_10_2009   ragava
                                Else
                                    IFLSolidWorksBaseClassObject.DeleteDetailItem("DetailItem142@Drawing View2", "Drawing View2")
                                End If
                                '01_10_2009   ragava
                                If Not (RodDrawingNumber = "492800" Or RodDrawingNumber = "494843") Then
                                    IFLSolidWorksBaseClassObject.DeleteDetailItem("DetailItem144@Drawing View2", "Drawing View2")
                                End If
                                '01_10_2009   ragava  Till  Here
                            ElseIf strRodDia & strSeries = "1.50" Then
                                If strRodMaterial.IndexOf("CHROME PLATED") <> -1 Then
                                    strRodMat = "(HARD CHROME PLATE)"
                                ElseIf strRodMaterial.IndexOf("NITRO") <> -1 Then
                                    strRodMat = "(NITRO TEC TREATED .0010 THK 64-71 RC FINISH 16 RMS)"
                                End If
                                IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD4@Drawing View2", strRodMat, "Drawing View2")
                                IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD16@Drawing View2", strRodMat, "Drawing View2")
                                If dblPistonThreadSize = 1 Then
                                    checkProperty("Piston_Thread_Size", "1-14UNS-2A")
                                    IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD14@Drawing View2", "(Ø1.001 INSIDE THIS AREA)", "Drawing View2")
                                ElseIf dblPistonThreadSize = 1.12 Then
                                    checkProperty("Piston_Thread_Size", "1-1/8-12UNF-2A")
                                    IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD14@Drawing View2", "(Ø1.126 INSIDE THIS AREA)", "Drawing View2")
                                ElseIf dblPistonThreadSize = 1.13 Then
                                    checkProperty("Piston_Thread_Size", "1-1/8-12UNF-2A")
                                    IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD14@Drawing View2", "(Ø1.126 INSIDE THIS AREA)", "Drawing View2")
                                ElseIf dblPistonThreadSize = 1.25 Then
                                    checkProperty("Piston_Thread_Size", "1-1/4-12UNF-2A")
                                    IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD14@Drawing View2", "(Ø1.251 INSIDE THIS AREA)", "Drawing View2")
                                End If
                                If dblRodThreadSize = 1.12 Then
                                    checkProperty("Rod_Thread_Size", "1-1/8-12UNF-3A")
                                ElseIf dblRodThreadSize = 1.25 Then
                                    checkProperty("Rod_Thread_Size", "1-1/4-12UNF-3A")
                                ElseIf dblRodThreadSize = 1.5 Then
                                    checkProperty("Rod_Thread_Size", "1-1/2-12UNF-3A")
                                End If
                                If Not (RodDrawingNumber = "492059" Or RodDrawingNumber = "492162" Or RodDrawingNumber = "493796" Or RodDrawingNumber = "492068" Or RodDrawingNumber = "493847" Or RodDrawingNumber = "493915") Then
                                    IFLSolidWorksBaseClassObject.DeleteNote("DetailItem150@Drawing View2", "Drawing View2")
                                End If
                                If Not (RodDrawingNumber = "492059" Or RodDrawingNumber = "492162" Or RodDrawingNumber = "493796") Then
                                    IFLSolidWorksBaseClassObject.DeleteNote("DetailItem152@Drawing View2", "Drawing View2")
                                End If
                                If Not (RodDrawingNumber = "493056") Then
                                    'DeleteNote("DetailItem152@Drawing View2", "Drawing View2")
                                End If
                                If strRodMaterial.IndexOf("-08-I") = -1 Then
                                    IFLSolidWorksBaseClassObject.DeleteDimension("RD27@Drawing View2", "Drawing View2")
                                    IFLSolidWorksBaseClassObject.DeleteDimension("RD28@Drawing View2", "Drawing View2")
                                    IFLSolidWorksBaseClassObject.DeleteDimension("RD29@Drawing View2", "Drawing View2")
                                    IFLSolidWorksBaseClassObject.DeleteDimension("RD30@Drawing View2", "Drawing View2")
                                    IFLSolidWorksBaseClassObject.DeleteDimension("RD31@Drawing View2", "Drawing View2")
                                End If
                                If Not (RodDrawingNumber = "494845") Then
                                    IFLSolidWorksBaseClassObject.DeleteNote("DetailItem157@Drawing View2", "Drawing View2")
                                End If
                                If Not (RodDrawingNumber = "492059" Or RodDrawingNumber = "492219" Or RodDrawingNumber = "493796") Then
                                    IFLSolidWorksBaseClassObject.DeleteNote("DetailItem155@Drawing View2", "Drawing View2")
                                End If
                                If RodDrawingNumber = "493274" Or RodDrawingNumber = "493742" Or RodDrawingNumber = "492059" Or RodDrawingNumber = "492162" Or RodDrawingNumber = "492219" Or RodDrawingNumber = "493056" Or RodDrawingNumber = "493796" Or RodDrawingNumber = "493847" Then
                                    IFLSolidWorksBaseClassObject.DeleteDetailItem("DetailItem146@Drawing View2", "Drawing View2")
                                Else
                                    IFLSolidWorksBaseClassObject.DeleteDetailItem("DetailItem158@Drawing View2", "Drawing View2")
                                End If
                            ElseIf strRodDia & strSeries = "1.75" Then
                                If strRodMaterial.IndexOf("CHROME PLATED") <> -1 Then
                                    strRodMat = "(HARD CHROME PLATE)"
                                ElseIf strRodMaterial.IndexOf("NITRO") <> -1 Then
                                    strRodMat = "(NITRO TEC TREATED .0010 THK 64-71 RC FINISH 16 RMS)"
                                End If
                                IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD9@Drawing View2", strRodMat, "Drawing View2")
                                If dblPistonThreadSize = 1 Then
                                    checkProperty("Piston_Thread_Size", "1-14UNS-2A")
                                    IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD13@Drawing View2", "(Ø1.001 INSIDE THIS AREA)", "Drawing View2")
                                ElseIf dblPistonThreadSize = 1.12 Then
                                    checkProperty("Piston_Thread_Size", "1-1/8-12UNF-2A")
                                    IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD13@Drawing View2", "(Ø1.126 INSIDE THIS AREA)", "Drawing View2")
                                ElseIf dblPistonThreadSize = 1.5 Then
                                    checkProperty("Piston_Thread_Size", "1-1/2-12UNF-2A")
                                    IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD13@Drawing View2", "(Ø1.501 INSIDE THIS AREA)", "Drawing View2")
                                End If
                                If dblRodThreadSize = 1.25 Then
                                    checkProperty("Rod_Thread_Size", "1-1/4-12UNF-3A")
                                ElseIf dblRodThreadSize = 1.5 Then
                                    checkProperty("Rod_Thread_Size", "1-1/2-12UNF-3A")
                                End If
                                If Not (RodDrawingNumber = "492074" Or RodDrawingNumber = "492163" Or RodDrawingNumber = "493832") Then
                                    IFLSolidWorksBaseClassObject.DeleteNote("DetailItem147@Drawing View2", "Drawing View2")
                                End If
                                If RodDrawingNumber = "492074" Or RodDrawingNumber = "492163" Or RodDrawingNumber = "493832" Then
                                    IFLSolidWorksBaseClassObject.DeleteDetailItem("DetailItem142@Drawing View2", "Drawing View2")
                                Else
                                    IFLSolidWorksBaseClassObject.DeleteDetailItem("DetailItem148@Drawing View2", "Drawing View2")
                                End If

                            ElseIf strRodDia & strSeries = "2.00" Then
                                If strRodMaterial.IndexOf("CHROME PLATED") <> -1 Then
                                    strRodMat = "(HARD CHROME PLATE)"
                                ElseIf strRodMaterial.IndexOf("NITRO") <> -1 Then
                                    strRodMat = "(NITRO TEC TREATED .0010 THK 64-71 RC FINISH 16 RMS)"
                                End If
                                IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD4@Drawing View2", strRodMat, "Drawing View2")
                                If dblPistonThreadSize = 1 Then
                                    checkProperty("Piston_Thread_Size", "1-14UNS-2A")
                                    IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD12@Drawing View2", "(Ø1.001 INSIDE THIS AREA)", "Drawing View2")
                                ElseIf dblPistonThreadSize = 1.12 Then
                                    checkProperty("Piston_Thread_Size", "1-1/8-12UNF-2A")
                                    IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD12@Drawing View2", "(Ø1.126 INSIDE THIS AREA)", "Drawing View2")
                                ElseIf dblPistonThreadSize = 1.25 Then
                                    checkProperty("Piston_Thread_Size", "1-1/4-12UNF-2A")
                                    IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD12@Drawing View2", "(Ø1.251 INSIDE THIS AREA)", "Drawing View2")
                                ElseIf dblPistonThreadSize = 1.5 Then
                                    checkProperty("Piston_Thread_Size", "1-1/2-12UNF-2A")
                                    IFLSolidWorksBaseClassObject.OverwriteDimensionNote("RD12@Drawing View2", "(Ø1.501 INSIDE THIS AREA)", "Drawing View2")
                                End If
                                If dblRodThreadSize = 1.25 Then
                                    checkProperty("Rod_Thread_Size", "1-1/4-12UNF-3A")
                                ElseIf dblRodThreadSize = 1.5 Then
                                    checkProperty("Rod_Thread_Size", "1-1/2-12UNF-3A")
                                End If
                                If RodDrawingNumber = "492077" Then
                                    IFLSolidWorksBaseClassObject.DeleteDetailItem("DetailItem141@Drawing View2", "Drawing View2")
                                Else
                                    'IFLSolidWorksBaseClassObject.DeleteDetailItem("DetailItem146@Drawing View1", "Drawing View1")
                                    IFLSolidWorksBaseClassObject.DeleteDetailItem("DetailItem145@Drawing View2", "Drawing View1")     '01_10_2009   ragava
                                End If
                            End If
                            IFLSolidWorksBaseClassObject.SolidWorksModel.EditRebuild3()
                            IFLSolidWorksBaseClassObject.Saveas_detached(strDrawingFileName)  '06_09_2012  RAGAVA
                            IFLSolidWorksBaseClassObject.DeleteDanglingDimension(strDrawingFileName)  '07_09_2012   RAGAVA
                            IFLSolidWorksBaseClassObject.SolidWorksModel.SaveSilent()
                            IFLSolidWorksBaseClassObject.SaveAndCloseAllDocuments()
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("Error in Updating Rod Drawing")
        End Try

        'For Bore Drawing Updation
        Try
            For intCount As Integer = 0 To UBound(arrFileEntries_Bore)
                strFileType = Split(arrFileEntries_Bore(intCount), ".")
                If UCase(strFileType(UBound(strFileType))) = UCase("SLDDRW") Then
                    If arrFileEntries_Bore(intCount).IndexOf("Bore-Assy.") <> -1 Then
                        Dim strDrawingFileName As String = arrFileEntries_Bore(intCount)
                        IFLSolidWorksBaseClassObject.openDocument(strDrawingFileName)
                        'IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CommandInProgress = False
                        If Not IFLSolidWorksBaseClassObject.SolidWorksModel Is Nothing Then
                            IFLSolidWorksBaseClassObject.SolidWorksModel.EditRebuild3()
                            System.Threading.Thread.Sleep(1000)         '06_10_2009   ragava
                            strBoreDrawingPath = strDrawingFileName          '22_09_2009   ragava
                            '16_09_2009  ragava
                            Try
                                If TubeLength >= 15 Then
                                    IFLSolidWorksBaseClassObject.BreakView("Section View K-K", TubeLength, 14)
                                End If
                            Catch ex As Exception
                                MsgBox("Error in Breaking the Bore Drawing View")
                            End Try
                            Dim strRodMat As String = String.Empty
                            If Series.IndexOf("TX") <> -1 Or Series.IndexOf("TL") <> -1 Or Series.IndexOf("TH") <> -1 Then
                                'NOTES
                                '1
                                IFLSolidWorksBaseClassObject.DeleteNote("DetailItem192@Sheet1", "")
                                IFLSolidWorksBaseClassObject.DeleteNote("DetailItem193@Sheet1", "")
                                IFLSolidWorksBaseClassObject.DeleteNote("DetailItem195@Sheet1", "")
                                '3
                                IFLSolidWorksBaseClassObject.DeleteView("Drawing View10")
                                IFLSolidWorksBaseClassObject.DeleteDetailedCircle("Detail Circle1")
                                'DETAIL ITEMS
                                IFLSolidWorksBaseClassObject.DeleteNote("DetailItem265@Drawing View10", "Drawing View10")
                                IFLSolidWorksBaseClassObject.DeleteNote("DetailItem266@Drawing View10", "Drawing View10")
                                IFLSolidWorksBaseClassObject.DeleteNote("DetailItem267@Drawing View10", "Drawing View10")
                                '4
                                IFLSolidWorksBaseClassObject.DeleteView("Drawing View9")
                                IFLSolidWorksBaseClassObject.DeleteDetailedCircle("Detail Circle2")
                            ElseIf Series.IndexOf("TP-High") <> -1 Then
                                If dblBoreDiameter >= 3.25 AndAlso dblBoreDiameter <= 5 Then
                                    If strRephasing = "At Extension" Then
                                        IFLSolidWorksBaseClassObject.DeleteNote("DetailItem265@Drawing View10", "Drawing View10")
                                        IFLSolidWorksBaseClassObject.DeleteNote("DetailItem266@Drawing View10", "Drawing View10")
                                        IFLSolidWorksBaseClassObject.DeleteNote("DetailItem267@Drawing View10", "Drawing View10")

                                        IFLSolidWorksBaseClassObject.DeleteView("Drawing View9")
                                        IFLSolidWorksBaseClassObject.DeleteDetailedCircle("Detail Circle2")
                                    ElseIf strRephasing = "At Retraction" Then
                                        'NOTES
                                        '1
                                        IFLSolidWorksBaseClassObject.DeleteNote("DetailItem192@Sheet1", "")
                                        IFLSolidWorksBaseClassObject.DeleteNote("DetailItem193@Sheet1", "")
                                        IFLSolidWorksBaseClassObject.DeleteNote("DetailItem195@Sheet1", "")
                                        '3
                                        IFLSolidWorksBaseClassObject.DeleteView("Drawing View10")
                                        IFLSolidWorksBaseClassObject.DeleteDetailedCircle("Detail Circle1")
                                    End If
                                End If
                            ElseIf Series.IndexOf("TP-Low") <> -1 Then
                                If strRephasing = "Both" Then
                                    'DETAIL ITEMS
                                    IFLSolidWorksBaseClassObject.DeleteNote("DetailItem265@Drawing View10", "Drawing View10")
                                    IFLSolidWorksBaseClassObject.DeleteNote("DetailItem266@Drawing View10", "Drawing View10")
                                    IFLSolidWorksBaseClassObject.DeleteNote("DetailItem267@Drawing View10", "Drawing View10")
                                ElseIf strRephasing = "At Extension" Then
                                    IFLSolidWorksBaseClassObject.DeleteNote("DetailItem265@Drawing View10", "Drawing View10")
                                    IFLSolidWorksBaseClassObject.DeleteNote("DetailItem266@Drawing View10", "Drawing View10")
                                    IFLSolidWorksBaseClassObject.DeleteNote("DetailItem267@Drawing View10", "Drawing View10")

                                    IFLSolidWorksBaseClassObject.DeleteView("Drawing View9")
                                    IFLSolidWorksBaseClassObject.DeleteDetailedCircle("Detail Circle2")
                                ElseIf strRephasing = "At Retraction" Then
                                    'NOTES
                                    '1
                                    IFLSolidWorksBaseClassObject.DeleteNote("DetailItem192@Sheet1", "")
                                    IFLSolidWorksBaseClassObject.DeleteNote("DetailItem193@Sheet1", "")
                                    IFLSolidWorksBaseClassObject.DeleteNote("DetailItem195@Sheet1", "")
                                    '3
                                    IFLSolidWorksBaseClassObject.DeleteView("Drawing View10")
                                    IFLSolidWorksBaseClassObject.DeleteDetailedCircle("Detail Circle1")
                                End If

                            End If
                            IFLSolidWorksBaseClassObject.SolidWorksModel.EditRebuild3()
                            IFLSolidWorksBaseClassObject.Saveas_detached(strDrawingFileName)  '06_09_2012  RAGAVA
                            IFLSolidWorksBaseClassObject.DeleteDanglingDimension(strDrawingFileName)  '07_09_2012   RAGAVA
                            IFLSolidWorksBaseClassObject.SolidWorksModel.SaveSilent()
                            IFLSolidWorksBaseClassObject.SaveAndCloseAllDocuments()
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("Error in Updating Bore Drawing")
        End Try

        'For Tie Rod Drawing Updation
        Try
            For intCount As Integer = 0 To UBound(arrFileEntries_TieRod)
                strFileType = Split(arrFileEntries_TieRod(intCount), ".")
                If UCase(strFileType(UBound(strFileType))) = UCase("SLDDRW") Then
                    If arrFileEntries_TieRod(intCount).IndexOf("Tie-Rod.") <> -1 Then
                        Dim strDrawingFileName As String = arrFileEntries_TieRod(intCount)
                        IFLSolidWorksBaseClassObject.openDocument(strDrawingFileName)
                        'IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CommandInProgress = False
                        If Not IFLSolidWorksBaseClassObject.SolidWorksModel Is Nothing Then
                            IFLSolidWorksBaseClassObject.SolidWorksModel.EditRebuild3()
                            System.Threading.Thread.Sleep(1000)         '06_10_2009   ragava
                            strTieRodDrawingPath = strDrawingFileName          '22_09_2009   ragava
                            '16_09_2009  ragava
                            Try
                                If TieRodLength >= 13 Then
                                    IFLSolidWorksBaseClassObject.BreakView("Drawing View1", TieRodLength, 12)
                                End If
                            Catch ex As Exception
                                MsgBox("Error in Breaking the Bore Drawing View")
                            End Try
                            IFLSolidWorksBaseClassObject.SolidWorksModel.EditRebuild3()
                            IFLSolidWorksBaseClassObject.Saveas_detached(strDrawingFileName)  '06_09_2012  RAGAVA
                            IFLSolidWorksBaseClassObject.DeleteDanglingDimension(strDrawingFileName)  '07_09_2012   RAGAVA
                            IFLSolidWorksBaseClassObject.SolidWorksModel.SaveSilent()
                            IFLSolidWorksBaseClassObject.SaveAndCloseAllDocuments()
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("Error in Updating Bore Drawing")
        End Try

        'Main Assembly Drawing Updation
        Try
            Dim strDrawingFileName As String = String.Empty
            If ClevisCapPortOrientation.IndexOf("90") <> -1 Then
                strDrawingFileName = DestinationFilePath + "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY_90.SLDDRW"
            Else
                strDrawingFileName = DestinationFilePath + "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDDRW"
            End If
            IFLSolidWorksBaseClassObject.openAssemblyDrawingDocument(strDrawingFileName)
            'IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CommandInProgress = False
            System.Threading.Thread.Sleep(1000)         '06_10_2009   ragava
            If Not IFLSolidWorksBaseClassObject.SolidWorksModel Is Nothing Then
                IFLSolidWorksBaseClassObject.SolidWorksModel.EditRebuild3()
                Try
                    '01_10_2009   ragava
                    If strPistonSealPackage <> "PSPSeal + WearRing2" Then
                        IFLSolidWorksBaseClassObject.DeleteNote("DetailItem556@Drawing View6", "Drawing View6")
                    End If
                    '13_10_2009

                    If ClevisCapPortOrientation = "Inline" Then
                        '19_02_2010     RAGAVA
                        If rdbRodClevis = False Then
                            IFLSolidWorksBaseClassObject.DeleteLineFromAssyDrawing("Line6", "Drawing View1")
                            IFLSolidWorksBaseClassObject.DeleteLineFromAssyDrawing("Line7", "Drawing View6")
                        Else
                            '23_02_2010   RAGAVA
                            IFLSolidWorksBaseClassObject.DeleteNotes("DetailItem1539@Drawing View1", "Sheet1", "NOTE")
                            'Till  Here
                        End If
                        'Till    Here

                        '12_10_2009   ragava
                        If ClevisCapCodeNumber = "292668" Or ClevisCapCodeNumber = "292672" Or ClevisCapCodeNumber = "292674" Or ClevisCapCodeNumber = "292675" Or ClevisCapCodeNumber = "292688" Or ClevisCapCodeNumber = "292721" Or ClevisCapCodeNumber = "492516" Or ClevisCapCodeNumber = "493181" Then
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-2")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-3")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-4")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-7")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-8")
                        ElseIf ClevisCapCodeNumber = "493238" Or ClevisCapCodeNumber = "493156" Or ClevisCapCodeNumber = "293028" Or ClevisCapCodeNumber = "492815" Or ClevisCapCodeNumber = "293005" Or ClevisCapCodeNumber = "293004" _
                        Or ClevisCapCodeNumber = "293029" Or ClevisCapCodeNumber = "492840" Or ClevisCapCodeNumber = "493065" Or ClevisCapCodeNumber = "293000" Or ClevisCapCodeNumber = "292810" Or ClevisCapCodeNumber = "292807" _
                        Or ClevisCapCodeNumber = "493112" Or ClevisCapCodeNumber = "493137" Or ClevisCapCodeNumber = "292805" Or ClevisCapCodeNumber = "492579" Or ClevisCapCodeNumber = "492831" Or ClevisCapCodeNumber = "492830" _
                        Or ClevisCapCodeNumber = "492820" Or ClevisCapCodeNumber = "492818" Or ClevisCapCodeNumber = "492817" Or ClevisCapCodeNumber = "492816" Then
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-1")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-3")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-4")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-7")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-8")
                        ElseIf ClevisCapCodeNumber = "293008" Or ClevisCapCodeNumber = "293036" Or ClevisCapCodeNumber = "491047" Or ClevisCapCodeNumber = "492115" Or ClevisCapCodeNumber = "492588" Or ClevisCapCodeNumber = "493999" _
                                           Or ClevisCapCodeNumber = "493772" Or ClevisCapCodeNumber = "493906" Or ClevisCapCodeNumber = "493833" Or ClevisCapCodeNumber = "493755" Or ClevisCapCodeNumber = "493183" Or ClevisCapCodeNumber = "493834" _
                                           Or ClevisCapCodeNumber = "492676" Or ClevisCapCodeNumber = "492681" Or ClevisCapCodeNumber = "492689" Or ClevisCapCodeNumber = "492706" Or ClevisCapCodeNumber = "493185" Or ClevisCapCodeNumber = "492720" _
                                           Or ClevisCapCodeNumber = "492709" Or ClevisCapCodeNumber = "493155" Or ClevisCapCodeNumber = "493722" Or ClevisCapCodeNumber = "492747" Or ClevisCapCodeNumber = "492707" Or ClevisCapCodeNumber = "492833" _
                                           Or ClevisCapCodeNumber = "492696" Or ClevisCapCodeNumber = "492682" Or ClevisCapCodeNumber = "492678" Or ClevisCapCodeNumber = "492660" Or ClevisCapCodeNumber = "492118" Or ClevisCapCodeNumber = "492702" _
                                           Or ClevisCapCodeNumber = "492109" Or ClevisCapCodeNumber = "293037" Or ClevisCapCodeNumber = "293014" Or ClevisCapCodeNumber = "293015" Or ClevisCapCodeNumber = "492683" Or ClevisCapCodeNumber = "492680" _
                                           Or ClevisCapCodeNumber = "491035" Or ClevisCapCodeNumber = "492113" Or ClevisCapCodeNumber = "492371" Or ClevisCapCodeNumber = "492662" Or ClevisCapCodeNumber = "494122" Or ClevisCapCodeNumber = "494123" Then

                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-1")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-2")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-4")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-7")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-8")

                        ElseIf ClevisCapCodeNumber = "492708" Then
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-1")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-2")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-3")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-7")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-8")

                        ElseIf ClevisCapCodeNumber = "493997" Then
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-1")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-2")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-3")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-4")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-8")
                        ElseIf ClevisCapCodeNumber = "493998" Or ClevisCapCodeNumber = "494124" Then
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-1")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-2")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-3")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-4")
                            IFLSolidWorksBaseClassObject.DeleteBlocksFromAssyDrawing("Block3-7")
                        End If
                        '12_10_2009   ragava      Till  Here

                        '06_04_2010   RAGAVA
                        If ClevisPins = True And blnPinsPlasticBag = False Then          '07_04_2010   RAGAVA
                            IFLSolidWorksBaseClassObject.DeleteNotes("DetailItem1541@Drawing View1", "Sheet1", "NOTE")
                        End If
                        If (RodClevisPins = True And blnPinsPlasticBag = False) Or rdbRodClevis = False Then          '07_04_2010   RAGAVA
                            IFLSolidWorksBaseClassObject.DeleteNotes("DetailItem1542@Drawing View1", "Sheet1", "NOTE")
                        End If
                        'Till   Here
                    Else
                        '19_02_2010     RAGAVA
                        If rdbRodClevis = False Then
                            IFLSolidWorksBaseClassObject.DeleteLineFromAssyDrawing("Line9", "Drawing View1")
                            IFLSolidWorksBaseClassObject.DeleteLineFromAssyDrawing("Line10", "Drawing View6")
                            '06_04_2010   RAGAVA
                        Else
                            IFLSolidWorksBaseClassObject.DeleteNotes("DetailItem1112@Drawing View1", "Sheet1", "NOTE")
                            'Till  Here
                        End If
                        'Till    Here
                        '06_04_2010   RAGAVA
                        If ClevisPins = True And blnPinsPlasticBag = False Then          '07_04_2010   RAGAVA
                            IFLSolidWorksBaseClassObject.DeleteNotes("DetailItem750@Drawing View6", "Sheet1", "NOTE")
                        End If
                        If (RodClevisPins = True And blnPinsPlasticBag = False) Or rdbRodClevis = False Then          '14_04_2010   RAGAVA
                            IFLSolidWorksBaseClassObject.DeleteNotes("DetailItem780@Drawing View6", "Sheet1", "NOTE")
                        End If
                        'Till   Here

                    End If
                    


                    '01_10_2009   ragava   Till  Here
                    'If dblStrokeLength >= 11 Then
                    '    IFLSolidWorksBaseClassObject.BreakView("Drawing View1", dblStrokeLength, 10)
                    '    IFLSolidWorksBaseClassObject.BreakView("Section View A-A", dblStrokeLength, 10)
                    'End If
                    '29_10_2009   ragava  Testing For Ramakrishna
                    If dblStrokeLength >= 16 Then
                        IFLSolidWorksBaseClassObject.BreakView("Drawing View1", dblStrokeLength, 15)
                        IFLSolidWorksBaseClassObject.BreakView("Section View A-A", dblStrokeLength, 15)
                    End If
                    '29_10_2009   ragava    Till  Here
                Catch ex As Exception
                    MsgBox("Error in Breaking the MAIN ASSEMBLY Drawing View")
                End Try
                '20_10_2009  ragava
                Try
                    IFLSolidWorksBaseClassObject.EditRetractedDimension(ExtendedLength)
                Catch ex As Exception
                End Try
                '20_10_2009  ragava       Till  Here

                Try
                    If strRodDia = "1.125" Then
                        If ClevisCapPortOrientation = "Inline" Then
                            IFLSolidWorksBaseClassObject.EditDimension("RD2@Drawing View6")        '19_04_2010   RAGAVA
                        Else
                            IFLSolidWorksBaseClassObject.EditDimension("RD1@Drawing View6")        '19_04_2010   RAGAVA
                        End If
                    End If
                Catch ex As Exception

                End Try
                IFLSolidWorksBaseClassObject.SolidWorksModel.EditRebuild3()
                Try
                    Dim strSQL As String = ""
                    Dim oDataClass As New DataClass
                    strSQL = "Select  max(RevisionNumber) as RevisionNumber From revisionTable where ContractNumber='" _
                                                & ContractNumber & "'"
                    Dim objDT As DataTable = oDataClass.GetDataTable(strSQL)
                    intContractRevisionNumber = objDT.Rows(0)("RevisionNumber")
                Catch ex As Exception

                End Try
                Try
                    IFLSolidWorksBaseClassObject.InsertTableRowDrawing(iAssyNotesCount, iPaintingNotesCount, _
                            intContractRevisionNumber, getRevisionTableData)              '12_10_2009   ragava
                    'IFLSolidWorksBaseClassObject.InsertTableRowDrawing(iPaintingNotesCount)
                Catch ex As Exception
                End Try
                '13_10_2009
                Try
                    If iPaintingNotesCount = 0 Then
                        If ClevisCapPortOrientation = "Inline" Then
                            IFLSolidWorksBaseClassObject.DeleteNotes("DetailItem566@Sheet1", "Sheet1", "ANNOTATIONTABLES")
                            IFLSolidWorksBaseClassObject.DeleteNotes("DetailItem525@Sheet1", "Sheet1", "NOTE")

                        Else
                            IFLSolidWorksBaseClassObject.DeleteNotes("DetailItem751@Sheet1", "Sheet1", "ANNOTATIONTABLES")
                            IFLSolidWorksBaseClassObject.DeleteNotes("DetailItem525@Sheet1", "Sheet1", "NOTE")
                        End If
                    End If
                Catch ex As Exception

                End Try
                Try
                    Dim strDrawingName As String = String.Empty
                    If ClevisCapPortOrientation.IndexOf("90") <> -1 Then
                        strDrawingName = "MAIN_ASSEMBLY_90.SLDDRW"
                    Else
                        strDrawingName = "MAIN_ASSEMBLY.SLDDRW"
                    End If
                    Dim strPinSize As Double = dblPinSize
                    If (RodClevisPins = True OrElse ClevisPins = True) AndAlso blnInstallPinsandClips_Checked = True Then          '19_10_2011   RAGAVA
                        IFLSolidWorksBaseClassObject.InsertViewFromexternalPart _
                                (strDrawingFileName, "X:\TieRodModels\TIE_ROD_STD_MODELS\UPDATED Pin kit subassembly.SLDASM", _
                                strDrawingName, strPinSize, RodPinClips, strPinKitId)       '10_09_2011   RAGAVA
                    End If
                Catch ex As Exception

                End Try
                IFLSolidWorksBaseClassObject.Saveas_detached(strDrawingFileName)  '06_09_2012  RAGAVA
                IFLSolidWorksBaseClassObject.DeleteDanglingDimension(strDrawingFileName)  '07_09_2012   RAGAVA
                IFLSolidWorksBaseClassObject.SolidWorksModel.SaveSilent()
                IFLSolidWorksBaseClassObject.SaveAndCloseAllDocuments()
            End If
        Catch ex As Exception
            'MsgBox("Error in Updating MAIN ASSEMBLY Drawing")
        End Try

    End Sub

    '22_09_2009  ragava
    Public Sub RenamePartFile(ByVal strOldName As String, ByVal strNewName As String, ByVal strReferencingDoc As String)

        Try
            Dim bret As Boolean = False
            Rename(strOldName, strNewName)
            bret = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.ReplaceReferencedDocument _
                                        (strReferencingDoc, strOldName, strNewName)
        Catch ex As Exception

        End Try

    End Sub

    '25_09_2009   ragava
    Public Sub FolderStructure(ByVal strReferenceDoc As String, ByVal strNewPartPath As String, ByVal strNetworkPath As String)

        Try
            Dim bRet As Boolean = False
            Dim strPart() As String = strNewPartPath.Split("\")
            Dim strPartName As String = strPart(UBound(strPart))
            'If File.Exists(strNetworkPath) = False Then
            FileCopy(strNewPartPath, strNetworkPath & strPartName)
            bRet = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.ReplaceReferencedDocument _
                                        (strReferenceDoc, strNewPartPath, strNetworkPath & strPartName)
        Catch ex As Exception

        End Try

    End Sub

    '25_09_2009   ragava
    Public Sub MoveDrawingFile(ByVal strOldDrawingPath As String, ByVal strNewPartPath As String, _
                        ByVal strNetworkDrawingPath As String, ByVal strNetworkPartPath As String)

        Try
            Dim bRet As Boolean = False

            'If File.Exists(strNetworkPath) = False Then
            FileCopy(strOldDrawingPath, strNetworkDrawingPath)
            bRet = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.ReplaceReferencedDocument _
                                    (strNetworkDrawingPath, strNewPartPath, strNetworkPartPath)
            Dim strCodeNumber As String = strNetworkDrawingPath.Substring(strNetworkDrawingPath.LastIndexOf("\") + 1, 6)
            If strCodeNumber_BeforeApplicationStart >= Val(strCodeNumber) Then
                IFLSolidWorksBaseClassObject.SaveAs_detached(strNetworkDrawingPath)
            End If
        Catch ex As Exception

        End Try

    End Sub

    '22_09_2009  ragava
    Public Sub RenameDrawingFile(ByVal strReferencingDoc As String, ByVal strOldName As String, ByVal strNewName As String)

        Try
            Dim bret As Boolean = False
            bret = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.ReplaceReferencedDocument _
                                                        (strReferencingDoc, strOldName, strNewName)
            strOldName = strReferencingDoc     '25_09_2009  ragava
            If strNewName.IndexOf(".SLDPRT") <> -1 Then
                Rename(strOldName, strNewName.Replace(".SLDPRT", ".SLDDRW"))
                'MoveDrawingFile(strNewName.Replace(".SLDPRT", ".SLDDRW"),strNewName,
            Else
                Rename(strOldName, strNewName.Replace(".SLDASM", ".SLDDRW"))
            End If
        Catch ex As Exception

        End Try

    End Sub

    '22_09_2009  ragava
    Public Sub insert265BOM(ByVal strLinkExcel As String, ByVal xpos As Double, ByVal ypos As Double, ByVal zpos As Double)

        Try
            Dim blnRet As Boolean = False
            KillExcel()
            oExcelClass.objApp = Nothing        '24_09_2009  ragava
            System.Threading.Thread.Sleep(1000)            '06_10_2009    ragava
            Try
                'IFLSolidWorksBaseClassObject.solidworksdrawingdocument = IFLSolidWorksBaseClassObject.solidworksapplicationobject.activedoc
                IFLSolidWorksBaseClassObject.SolidWorksModel = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.ActiveDoc
                IFLSolidWorksBaseClassObject.solidworksmodel.EditTemplate()
                IFLSolidWorksBaseClassObject.selectOLE()
                IFLSolidWorksBaseClassObject.solidworksmodel.editdelete()
            Catch ex As Exception
            End Try
            IFLSolidWorksBaseClassObject.swModelExt = IFLSolidWorksBaseClassObject.SolidWorksModel.Extension
            IFLSolidWorksBaseClassObject.swOleObj = IFLSolidWorksBaseClassObject.swModelExt.InsertObjectFromFile _
                        (strLinkExcel, False, 1, xpos, ypos, zpos)  ' XPos, YPos, ZPos)       '22_03_2009  ragava   True To False
            blnRet = IFLSolidWorksBaseClassObject.SolidWorksModel.EditRebuild3()
            IFLSolidWorksBaseClassObject.solidworksmodel.EditSheet()          '01_10_2009   RAGAVA
            System.Threading.Thread.Sleep(1000)            '06_10_2009    ragava
            IFLSolidWorksBaseClassObject.SolidWorksModel.SaveSilent()
        Catch ex As Exception
            MsgBox("Error in Inserting ExcelSheet into Drawing")
        End Try

    End Sub
    '22_09_2009  ragava
    Public Sub CreateExcel(ByVal strFileName As String, ByVal ObjDt As DataTable)

        Try
            'Dim strFileName As String = "C:\DESIGN_TABLES\TableDrawing.xls"
            Dim icount As Integer = 2
            'If File.Exists(strFileName) = True Then
            '    File.Delete(strFileName)
            'End If
            oExcelClass.checkExcelInstance()
            System.Threading.Thread.Sleep(1000)            '06_10_2009    ragava
            oExcelClass.objBook = oExcelClass.objApp.Workbooks.Open(strFileName)
            oExcelClass.objSheet = oExcelClass.objBook.Worksheets("Sheet1")
            oExcelClass.objSheet.Cells(1, 1).value = "CODE#"
            oExcelClass.objSheet.Cells(1, 2).value = "Dim-A"
            oExcelClass.objSheet.Cells(1, 3).value = "Stroke"
            oExcelClass.objSheet.Cells(1, 4).value = "Revision"
            oExcelClass.objBook.Save()
            For Each dr As DataRow In ObjDt.Rows
                oExcelClass.objSheet.Range("A" & icount.ToString).Value = dr(1).ToString
                oExcelClass.objSheet.Range("B" & icount.ToString).Value = dr(2).ToString
                oExcelClass.objSheet.Range("C" & icount.ToString).Value = dr(3).ToString
                oExcelClass.objSheet.Range("D" & icount.ToString).Value = dr(4).ToString
                icount += 1
            Next
            oExcelClass.objBook.Save()
            oExcelClass.objApp.Quit()
        Catch ex As Exception
        End Try

    End Sub
    '22_09_2009  ragava
    Public Sub OpenDrawingAndActivateSheet(ByVal strDrawingFile As String, Optional ByVal strSheet As String = "Sheet1")

        Try
            IFLSolidWorksBaseClassObject.openDocument(strDrawingFile)
            IFLSolidWorksBaseClassObject.SolidWorksModel = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.ActiveDoc
            IFLSolidWorksBaseClassObject.SolidWorksDrawingDocument = IFLSolidWorksBaseClassObject.SolidWorksModel
            Dim bRet As Boolean = IFLSolidWorksBaseClassObject.SolidWorksDrawingDocument.ActivateSheet(strSheet)
            'IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CommandInProgress = True
        Catch ex As Exception
        End Try

    End Sub

    '22_09_2009  ragava
    Public Sub UpdateTableDrawing()

        Dim strNetworkPath As String = "X:\"
        'Dim strNetworkPath As String = "C:\Monarch_CDA\"       '02_10_2009 ragava
        Dim strModelNetworkPath As String = "X:\TieRodModels\"
        'Dim strModelNetworkPath As String = "C:\Monarch_CDA\TieRodModels\"       '02_10_2009 ragava
        Dim strDrawingNetworkPath As String = String.Empty        '25_09_2009  ragava


        'Tie Rod Table Drawing
        Try
            Dim strCodeNumber As String = String.Empty
            Dim oDataClass As New DataClass
            Dim strDefaultPath As String = "C:\TableDrawingFile\"
            strStopTubeDrawingPath = DestinationFilePath & "\Stoptube\Stop_tube.SLDDRW"
            'strRodDrawingPath    strTieRodDrawingPath      strStopTubeDrawingPath       strBoreDrawingPath
            Dim strQuery As String = "Select * from TieRodSizes where DrawingNumber = '" & _
                                TieRodDrawingNumber.ToString & "' and TableDrawing = 'Yes'"
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            TieRodLength = Math.Round(TieRodLength, 2)
            'For Rod Drawing
            If objDT.Rows.Count > 0 Then

                '25_09_2009  ragava
                If TieRodDrawingNumber.ToString.StartsWith("4") = True Or TieRodDrawingNumber.ToString.StartsWith("7") = True Then
                    strDrawingNetworkPath = strNetworkPath & "400\" '& TieRodDrawingNumber.ToString & ".SLDDRW"
                ElseIf TieRodDrawingNumber.ToString.StartsWith("6") = True Or TieRodDrawingNumber.ToString.StartsWith("8") = True Then
                    strDrawingNetworkPath = strNetworkPath & "600\" '& TieRodDrawingNumber.ToString & ".SLDDRW"
                ElseIf TieRodDrawingNumber.ToString.StartsWith("2") = True Then
                    strDrawingNetworkPath = strNetworkPath & "200\" '& TieRodDrawingNumber.ToString & ".SLDDRW"
                End If
                '25_09_2009  ragava   Till  Here


                If File.Exists(strDrawingNetworkPath & TieRodDrawingNumber.ToString & ".SLDDRW") = False Then
                    MsgBox("Drawing File " & strDrawingNetworkPath & TieRodDrawingNumber.ToString _
                            & ".SLDDRW doesn't exist, Please Copy the File to specified Location and then Click Ok")
                End If
                'OpenDrawingAndActivateSheet(strDrawingNetworkPath & TieRodDrawingNumber.ToString & ".SLDDRW")       '20_10_2009  ragava
                strQuery = ""

                '03_12_2009   Ragava
                'strQuery = "Select CodeNumber,Dim_A,Revision from TieRodTableDrawing where DrawingNumber = '" & TieRodDrawingNumber & "'"
                Dim dblTieRodLength As Double = dblStrokeLength + dblTieRodStrokeDifference + dblStopTubeLength
                strQuery = "Select CodeNumber,Dim_A,Revision from TieRodTableDrawing where DrawingNumber = '" & _
                                TieRodDrawingNumber & "' and Dim_A = " & Math.Round(dblTieRodLength, 2).ToString
                '03_12_2009   Ragava    Till  Here

                Dim objDT2 As DataTable = oDataClass.GetDataTable(strQuery)
                Dim blnInsert As Boolean = True
                If objDT2.Rows.Count > 0 Then
                    For Each dr As DataRow In objDT2.Rows
                        'If dr(1).ToString = Format(TieRodLength, "00.00").ToString Then
                        If dr(1).ToString = Format(TieRodLength, "0.00").ToString Then        '12_10_2009   ragava
                            blnInsert = False
                            'ANUP 26-10-2010 START
                            IsRowInserted_Tierod = False
                            'ANUP 26-10-2010 TILL HERE
                            strCodeNumber = dr(0).ToString
                            ht_CodeNumbers.Clear()
                            ht_CodeNumbers.Add("TIEROD", strCodeNumber)        '04_10_2010   RAGAVA   TESTING  need to do somethinglike this
                            Exit For
                        End If
                    Next
                    If blnInsert = True Then
                        'ANUP 26-10-2010 START
                        IsRowInserted_Tierod = True
                        'ANUP 26-10-2010 TILL HERE
                        OpenDrawingAndActivateSheet(strDrawingNetworkPath & TieRodDrawingNumber.ToString & ".SLDDRW")       '20_10_2009  ragava
                        'strQuery = "Select CodeNumber,Type from CodeNumberDetails where Type = 'TieROD'"
                        strQuery = "Select CodeNumber,Type,MaxCodeNumber from CodeNumberDetails where Type = 'TieROD'"         '12_10_2009  ragava
                        Dim objDT6 As DataTable = oDataClass.GetDataTable(strQuery)
                        strCodeNumber = objDT6.Rows(0).Item(0).ToString()
                        '12_10_2009  ragava
                        If Val(strCodeNumber) >= objDT6.Rows(0).Item(2) Then
                            MsgBox("Generated CodeNumber is Invalid asper MonarchIndustries... " & strCodeNumber)
                            Exit Try
                            'strCodeNumber = "749999"
                        End If
                        '12_10_2009  ragava    Till   Here
                        '07_10_2009   ragava
                        Dim strRevision As String = "1"
                        Try
                            strQuery = "select max(Revision) from TieRodTableDrawing where DrawingNumber = '" & _
                                                TieRodDrawingNumber.ToString & "'"
                            Dim objDT7 As DataTable = oDataClass.GetDataTable(strQuery)
                            strRevision = (Val(objDT7.Rows(0).Item(0)) + 1).ToString
                            strQuery = ""
                        Catch ex As Exception
                        End Try
                        '07_10_2009   ragava      Till    Here

                        'strQuery = "Insert into TieRodTableDrawing Values('" & TieRodDrawingNumber.ToString & "','" & strCodeNumber.ToString & "','" & TieRodLength.ToString & "','" & (TieRodLength - dblTieRodStrokeDifference).ToString & "','1')"
                        strQuery = "Insert into TieRodTableDrawing Values('" & TieRodDrawingNumber.ToString _
                                & "','" & strCodeNumber.ToString & "','" & TieRodLength.ToString & "','" & _
                                (TieRodLength - dblTieRodStrokeDifference).ToString & "','" & strRevision & "')"               '07_10_2009   ragava

                        Dim objDT5 As DataTable = oDataClass.GetDataTable(strQuery)
                        strQuery = ""
                        'strQuery = "Update CodeNumberDetails set CodeNumber ='" & (Val(strCodeNumber) + 1).ToString & "' Where Type = 'TieRod'"
                        strQuery = "Update CodeNumberDetails set CodeNumber ='" & (Val(strCodeNumber) + 1).ToString & "'"       '30_09_2009  ragava
                        Dim objDT_Temp As DataTable = oDataClass.GetDataTable(strQuery)
                        'End If
                        Try
                            ht_CodeNumbers.Clear()
                            ht_CodeNumbers.Add("TIEROD", strCodeNumber)          '20_01_2011   RAGAVA
                        Catch ex As Exception
                        End Try
                        strQuery = ""
                        strQuery = "Select * from TieRodTableDrawing where DrawingNumber = '" & TieRodDrawingNumber & "' order by Dim_A"
                        Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery)
                        If objDT1.Rows.Count > 0 Then
                            If File.Exists("C:\DESIGN_TABLES\TableDrawing\TieRod.xls") = False Then
                                MsgBox("Please Create an Excel file at following Path : C:\DESIGN_TABLES\TableDrawing\TieRod.xls  and then Click Ok")
                            End If
                            CreateExcel("C:\DESIGN_TABLES\TableDrawing\TieRod.xls", objDT1)
                            System.Threading.Thread.Sleep(1000)            '06_10_2009   ragava
                            insert265BOM("C:\DESIGN_TABLES\TableDrawing\TieRod.xls", 0.015, 0.25, 0)    '05_10_2009   ragava
                            System.Threading.Thread.Sleep(1000)            '06_10_2009   ragava
                            UpdateRevisionTable(strCodeNumber, strRevision)        '21_10_2009   ragava
                        End If
                        'insert265BOM("C:\DESIGN_TABLES\TableDrawing\TieRod.xls", 0.015, 0.25, 0)    '05_10_2009   ragava
                        'IFLSolidWorksBaseClassObject.SolidWorksModel.SaveSilent()          '04_02_2010     RAGAVA

                        '29_01_2010   ragava
                        Try
                            If Directory.Exists("w:\") = True Then
                                'IFLSolidWorksBaseClassObject.SolidWorksModel.SaveAs("w:\" & TieRodDrawingNumber.ToString & ".SLDDRW")
                                IFLSolidWorksBaseClassObject.SolidWorksModel.SaveAs2("w:\" & _
                                        TieRodDrawingNumber.ToString & ".SLDDRW", 3, True, True)
                            End If
                        Catch ex As Exception
                            MsgBox("ERROR IN SAVING DRAWING FILE : " & "w:\" & TieRodDrawingNumber.ToString & ".SLDDRW")
                        End Try
                        'Till Here

                        'IFLSolidWorksBaseClassObject.SaveAndCloseAllDocuments()         '04_02_2010     RAGAVA
                        IFLSolidWorksBaseClassObject.CloseAllDocuments()        '04_02_2010     RAGAVA
                    End If       '20_10_2009  ragava
                    Dim strPart() As String = strTieRodDrawingPath.Split(".")
                    Dim strPartPath As String
                    If strPart.Length > 2 Then
                        strPartPath = strPart(LBound(strPart)) & "." & strPart(LBound(strPart) + 1)
                    Else
                        strPartPath = strPart(LBound(strPart))
                    End If
                    Dim strNewPartPath As String = strPartPath.Substring(0, strPartPath.LastIndexOf("\")) _
                                                    & "\" & strCodeNumber.ToString & ".SLDPRT"
                    'Renaming Tie Rod Part
                    RenamePartFile(strPartPath & ".SLDPRT", strNewPartPath, DestinationFilePath & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM")

                    FolderStructure(DestinationFilePath & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM", _
                                        strNewPartPath, strModelNetworkPath)     '25_09_2009  ragava
                Else
                    GoTo TieRodCodeNumber                    '03_12_2009   Ragava
                End If
            Else
TieRodCodeNumber:  '03_12_2009   Ragava


                '10_12_2010   RAGAVA
                Try
                    TieRodLength = Math.Round(TieRodLength, 2)
                    Dim strQuery1 As String = "Select TieRodPartNumber from TieRodSizes where DrawingNumber = '" & _
                            TieRodDrawingNumber.ToString & "' and TableDrawing = 'No' and [Dimension-A] ='" & _
                            Format(TieRodLength, "0.00").ToString & "'"
                    Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery1)
                    If objDT1.Rows.Count > 0 Then
                        strCodeNumber = objDT1.Rows(0)("TieRodPartNumber")
                        ht_CodeNumbers.Clear()
                        ht_CodeNumbers.Add("TIEROD", strCodeNumber)
                    End If
                Catch ex As Exception
                End Try
                'Till   Here

                OpenDrawingAndActivateSheet(strTieRodDrawingPath)
                IFLSolidWorksBaseClassObject.SolidWorksModel.SaveSilent()
                IFLSolidWorksBaseClassObject.SaveAndCloseAllDocuments()
                If strCodeNumber = "" Then
                    strQuery = ""
                    'strQuery = "Select CodeNumber,Type from CodeNumberDetails where Type = 'TieROD'"
                    strQuery = "Select CodeNumber,Type,MaxCodeNumber from CodeNumberDetails where Type = 'TieROD'"      '12_10_2009  ragava
                    Dim objDT3 As DataTable = oDataClass.GetDataTable(strQuery)
                    strCodeNumber = objDT3.Rows(0).Item(0).ToString()
                    '12_10_2009  ragava
                    If Val(strCodeNumber) >= objDT3.Rows(0).Item(2) Then
                        MsgBox("Generated CodeNumber is Invalid asper MonarchIndustries... " & strCodeNumber)
                        Exit Try
                        'strCodeNumber = "749999"
                    End If
                    Try
                        ht_CodeNumbers.Clear()
                        ht_CodeNumbers.Add("TIEROD", strCodeNumber)         '20_01_2011   RAGAVA
                    Catch ex As Exception
                    End Try
                    '12_10_2009  ragava    Till   Here
                    'strQuery = "Update CodeNumberDetails set CodeNumber ='" & (Val(strCodeNumber) + 1).ToString & "' Where Type = 'TieRod'" 
                    strQuery = "Update CodeNumberDetails set CodeNumber ='" & (Val(strCodeNumber) + 1).ToString & "'"            '30_09_2009  ragava
                    objDT3.Clear()
                    objDT3 = oDataClass.GetDataTable(strQuery)
                End If
                Dim strPart() As String = strTieRodDrawingPath.Split(".")
                'Dim strPartPath As String = strPart(LBound(strPart))
                Dim strPartPath As String
                If strPart.Length > 2 Then
                    strPartPath = strPart(LBound(strPart)) & "." & strPart(LBound(strPart) + 1)
                Else
                    strPartPath = strPart(LBound(strPart))
                End If
                Dim strNewPartPath As String = strPartPath.Substring(0, strPartPath.LastIndexOf("\")) & "\" _
                    & strCodeNumber.ToString & ".SLDPRT"
                'Renaming Tie Rod Part
                RenamePartFile(strPartPath & ".SLDPRT", strNewPartPath, DestinationFilePath & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM")
                FolderStructure(DestinationFilePath & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM", strNewPartPath, strModelNetworkPath)     '25_09_2009  ragava
                'Renaming Drawing
                RenameDrawingFile(strTieRodDrawingPath, strPartPath & ".SLDPRT", strNewPartPath)
                MoveDrawingFile(strNewPartPath.Replace(".SLDPRT", ".SLDDRW"), strNewPartPath, strNetworkPath _
                    & "400\" & strCodeNumber & ".SLDDRW", strModelNetworkPath & strCodeNumber & ".SLDPRT")    '25_09_2009  ragava
            End If
        Catch ex As Exception
        End Try

        '***************************************************************************************************************

        'Tube Table Drawing
        Try
            Dim strCodeNumber As String = String.Empty
            Dim oDataClass As New DataClass
            Dim strDefaultPath As String = "C:\TableDrawingFile\"
            strStopTubeDrawingPath = DestinationFilePath & "\Stoptube\Stop_tube.SLDDRW"
            'strRodDrawingPath    strTieRodDrawingPath      strStopTubeDrawingPath       strBoreDrawingPath
            Dim strQuery As String = "Select * from BoreDiameterDetails where DrawingPartNumber = '" & _
                                        BoreDrawingNumber.ToString & "' and TableDrawing = 'Yes'"
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            TubeLength = Math.Round(TubeLength, 2)

            'For Tube Drawing
            If objDT.Rows.Count > 0 Then
                '25_09_2009  ragava
                If BoreDrawingNumber.ToString.StartsWith("4") = True Or BoreDrawingNumber.ToString.StartsWith("7") = True Then
                    strDrawingNetworkPath = strNetworkPath & "400\" '& BoreDrawingNumber.ToString & ".SLDDRW"
                ElseIf BoreDrawingNumber.ToString.StartsWith("6") = True Or BoreDrawingNumber.ToString.StartsWith("8") = True Then
                    strDrawingNetworkPath = strNetworkPath & "600\" '& BoreDrawingNumber.ToString & ".SLDDRW"
                ElseIf BoreDrawingNumber.ToString.StartsWith("2") = True Then
                    strDrawingNetworkPath = strNetworkPath & "200\" '& BoreDrawingNumber.ToString & ".SLDDRW"
                End If
                '25_09_2009  ragava   Till  Here
                If File.Exists(strDrawingNetworkPath & BoreDrawingNumber.ToString & ".SLDDRW") = False Then
                    MsgBox("Drawing File " & strDrawingNetworkPath & BoreDrawingNumber.ToString _
                        & ".SLDDRW doesn't exist, Please Copy the File to specified Location and then Click Ok")
                End If
                'OpenDrawingAndActivateSheet(strDrawingNetworkPath & BoreDrawingNumber.ToString & ".SLDDRW")       '20_10_2009  ragava
                strQuery = ""
                strQuery = "Select CodeNumber,Dim_A,Revision from TubeTableDrawing where DrawingNumber = '" & BoreDrawingNumber & "'"
                Dim objDT2 As DataTable = oDataClass.GetDataTable(strQuery)
                Dim blnInsert As Boolean = True
                If objDT2.Rows.Count > 0 Then
                    For Each dr As DataRow In objDT2.Rows
                        'If dr(1).ToString = Format(TubeLength, "00.00").ToString Then 
                        If dr(1).ToString = Format(TubeLength, "0.00").ToString Then                  '12_10_2009   ragava
                            blnInsert = False
                            'ANUP 26-10-2010 START
                            IsRowInserted_Tube = False
                            'ANUP 26-10-2010 TILL HERE
                            strCodeNumber = dr(0).ToString
                            ht_CodeNumbers.Add("TUBE", strCodeNumber)        '04_10_2010   RAGAVA   TESTING
                            Exit For
                        End If
                    Next
                    If blnInsert = True Then
                        'ANUP 26-10-2010 START
                        IsRowInserted_Tube = True
                        'ANUP 26-10-2010 TILL HERE
                        'strQuery = "Select CodeNumber,Type from CodeNumberDetails where Type = 'Tube'"
                        OpenDrawingAndActivateSheet(strDrawingNetworkPath & BoreDrawingNumber.ToString & ".SLDDRW")        '20_10_2009  ragava
                        strQuery = "Select CodeNumber,Type,MaxCodeNumber from CodeNumberDetails where Type = 'Tube'"         '12_10_2009   ragava
                        Dim objDT6 As DataTable = oDataClass.GetDataTable(strQuery)
                        strCodeNumber = objDT6.Rows(0).Item(0).ToString()
                        '12_10_2009  ragava
                        If Val(strCodeNumber) >= objDT6.Rows(0).Item(2) Then
                            MsgBox("Generated CodeNumber is Invalid asper MonarchIndustries... " & strCodeNumber)
                            Exit Try
                            'strCodeNumber = "749999"
                        End If
                        '12_10_2009  ragava    Till   Here
                        '07_10_2009   ragava
                        Dim strRevision As String = "1"
                        Try
                            strQuery = "select max(Revision) from TubeTableDrawing where DrawingNumber = '" & _
                                        BoreDrawingNumber.ToString & "'"
                            Dim objDT7 As DataTable = oDataClass.GetDataTable(strQuery)
                            strRevision = (Val(objDT7.Rows(0).Item(0)) + 1).ToString
                            strQuery = ""
                        Catch ex As Exception
                        End Try
                        '07_10_2009   ragava      Till    Here

                        'strQuery = "Insert into TubeTableDrawing Values('" & BoreDrawingNumber.ToString & "','" & strCodeNumber.ToString & "','" & TubeLength.ToString & "','" & (TubeLength - dblTubeStrokeDifference).ToString & "','1')"
                        strQuery = "Insert into TubeTableDrawing Values('" & BoreDrawingNumber.ToString _
                                & "','" & strCodeNumber.ToString & "','" & TubeLength.ToString & "','" & _
                                (TubeLength - dblTubeStrokeDifference).ToString & "','" & strRevision & "')"            '07_10_2009   ragava

                        Dim objDT5 As DataTable = oDataClass.GetDataTable(strQuery)
                        strQuery = ""
                        strQuery = "Update CodeNumberDetails set CodeNumber ='" & (Val(strCodeNumber) + 1).ToString & "'"          '30_09_2009  ragava
                        Dim objDT_Temp As DataTable = oDataClass.GetDataTable(strQuery)
                        'End If        '20_10_2009  ragava
                        Try
                            ht_CodeNumbers.Add("TUBE", strCodeNumber)        '20_01_2011   RAGAVA
                        Catch ex As Exception
                        End Try
                        strQuery = ""
                        strQuery = "Select * from TubeTableDrawing where DrawingNumber = '" & BoreDrawingNumber & "' order by Dim_A"
                        Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery)
                        If objDT1.Rows.Count > 0 Then
                            If File.Exists("C:\DESIGN_TABLES\TableDrawing\Tube.xls") = False Then
                                MsgBox("Please Create an Excel file at following Path : C:\DESIGN_TABLES\TableDrawing\Tube.xls  and then Click Ok")
                            End If
                            CreateExcel("C:\DESIGN_TABLES\TableDrawing\Tube.xls", objDT1)
                            System.Threading.Thread.Sleep(1000)            '06_10_2009   ragava
                            insert265BOM("C:\DESIGN_TABLES\TableDrawing\Tube.xls", 0.015, 0.25, 0)    '05_10_2009   ragava
                            System.Threading.Thread.Sleep(1000)            '06_10_2009   ragava
                            UpdateRevisionTable(strCodeNumber, strRevision)        '21_10_2009   ragava
                        End If
                        'insert265BOM("C:\DESIGN_TABLES\TableDrawing\Tube.xls", 0.015, 0.25, 0)    '05_10_2009   ragava
                        'IFLSolidWorksBaseClassObject.SolidWorksModel.SaveSilent()          '04_02_2010     RAGAVA

                        '29_01_2010   ragava
                        Try
                            If Directory.Exists("w:\") = True Then
                                IFLSolidWorksBaseClassObject.SolidWorksModel.SaveAs("w:\" & BoreDrawingNumber.ToString & ".SLDDRW")
                            End If
                        Catch ex As Exception
                            MsgBox("ERROR IN SAVING DRAWING FILE : " & "w:\" & BoreDrawingNumber.ToString & ".SLDDRW")
                        End Try
                        'Till Here

                        'IFLSolidWorksBaseClassObject.SaveAndCloseAllDocuments()         '04_02_2010     RAGAVA
                        IFLSolidWorksBaseClassObject.CloseAllDocuments()        '04_02_2010     RAGAVA
                    End If        '20_10_2009  ragava
                    Dim strPart() As String = strBoreDrawingPath.Split(".")
                    Dim strPartPath As String
                    If strPart.Length > 2 Then
                        strPartPath = strPart(LBound(strPart)) & "." & strPart(LBound(strPart) + 1)
                    Else
                        strPartPath = strPart(LBound(strPart))
                    End If
                    Dim strNewPartPath As String = strPartPath.Substring(0, strPartPath.LastIndexOf("\")) _
                                                    & "\" & strCodeNumber.ToString & ".SLDASM"
                    'Renaming Tube Assy & Part

                    '29_09_2009  ragava
                    RenamePartFile(DestinationFilePath & "\Bore\Bore.SLDPRT", strNewPartPath.Replace _
                                (".SLDASM", ".SLDPRT"), DestinationFilePath & "\Bore\Bore-Assy.SLDASM")
                    FolderStructure(DestinationFilePath & "\Bore\Bore-Assy.SLDASM", strNewPartPath.Replace _
                                (".SLDASM", ".SLDPRT"), strModelNetworkPath)
                    '29_09_2009  ragava      Till  Here
                    RenamePartFile(strPartPath & ".SLDASM", strNewPartPath, DestinationFilePath _
                                            & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM")
                    FolderStructure(DestinationFilePath & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM", _
                                                    strNewPartPath, strModelNetworkPath)     '25_09_2009  ragava
                End If
            Else


                '10_12_2010   RAGAVA
                Try
                    TubeLength = Math.Round(TubeLength, 2)
                    Dim strQuery1 As String = "Select PartNumber from BoreDiameterDetails where DrawingPartNumber = '" _
                            & BoreDrawingNumber.ToString & "' and TableDrawing = 'No' and TubeLength ='" & _
                            Format(TubeLength, "0.00").ToString & "'"
                    Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery1)
                    If objDT1.Rows.Count > 0 Then
                        strCodeNumber = objDT1.Rows(0)("PartNumber")
                        ht_CodeNumbers.Add("TUBE", strCodeNumber)
                    End If
                Catch ex As Exception
                End Try
                'Till   Here

                OpenDrawingAndActivateSheet(strBoreDrawingPath)
                IFLSolidWorksBaseClassObject.SolidWorksModel.SaveSilent()
                IFLSolidWorksBaseClassObject.SaveAndCloseAllDocuments()
                If strCodeNumber = "" Then
                    strQuery = ""
                    'strQuery = "Select CodeNumber,Type from CodeNumberDetails where Type = 'Tube'"
                    strQuery = "Select CodeNumber,Type,MaxCodeNumber from CodeNumberDetails where Type = 'Tube'"          '12_10_2009   ragava
                    Dim objDT3 As DataTable = oDataClass.GetDataTable(strQuery)
                    strCodeNumber = objDT3.Rows(0).Item(0).ToString()
                    '12_10_2009  ragava
                    If Val(strCodeNumber) >= objDT3.Rows(0).Item(2) Then
                        MsgBox("Generated CodeNumber is Invalid asper MonarchIndustries... " & strCodeNumber)
                        Exit Try
                        'strCodeNumber = "749999"
                    End If
                    Try
                        ht_CodeNumbers.Add("TUBE", strCodeNumber)         '20_01_2011   RAGAVA
                    Catch ex As Exception
                    End Try
                    '12_10_2009  ragava    Till   Here
                    'strQuery = "Update CodeNumberDetails set CodeNumber ='" & (Val(strCodeNumber) + 1).ToString & "' Where Type = 'Tube'" 
                    strQuery = "Update CodeNumberDetails set CodeNumber ='" & (Val(strCodeNumber) + 1).ToString & "'"          '30_09_2009  ragava
                    objDT3.Clear()
                    objDT3 = oDataClass.GetDataTable(strQuery)
                End If
                Dim strPart() As String = strBoreDrawingPath.Split(".")
                'Dim strPartPath As String = strPart(LBound(strPart))
                Dim strPartPath As String
                If strPart.Length > 2 Then
                    strPartPath = strPart(LBound(strPart)) & "." & strPart(LBound(strPart) + 1)
                Else
                    strPartPath = strPart(LBound(strPart))
                End If
                Dim strNewPartPath As String = strPartPath.Substring(0, strPartPath.LastIndexOf("\")) & "\" _
                                            & strCodeNumber.ToString & ".SLDASM"
                'Renaming Tube Assy & Part

                '29_09_2009  ragava
                RenamePartFile(DestinationFilePath & "\Bore\Bore.SLDPRT", strNewPartPath.Replace(".SLDASM", ".SLDPRT"), _
                                        DestinationFilePath & "\Bore\Bore-Assy.SLDASM")
                FolderStructure(DestinationFilePath & "\Bore\Bore-Assy.SLDASM", strNewPartPath.Replace(".SLDASM", ".SLDPRT"), _
                                                strModelNetworkPath)
                '29_09_2009  ragava      Till  Here


                RenamePartFile(strPartPath & ".SLDASM", strNewPartPath, DestinationFilePath & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM")
                FolderStructure(DestinationFilePath & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM", _
                                    strNewPartPath, strModelNetworkPath)     '25_09_2009  ragava
                'Renaming Drawing
                RenameDrawingFile(strBoreDrawingPath, strPartPath & ".SLDASM", strNewPartPath)
                MoveDrawingFile(strNewPartPath.Replace(".SLDASM", ".SLDDRW"), strNewPartPath, _
                        strNetworkPath & "400\" & strCodeNumber & ".SLDDRW", strModelNetworkPath & strCodeNumber & ".SLDASM")    '25_09_2009  ragava
            End If
        Catch ex As Exception
        End Try

        '*************************************************************************************************

        'Rod Table Drawing 


        Try
            Dim strCodeNumber As String = String.Empty
            Dim oDataClass As New DataClass
            'Dim strDefaultPath As String = "C:\TableDrawingFile\"
            strStopTubeDrawingPath = DestinationFilePath & "\Stoptube\Stop_tube.SLDDRW"
            Dim strQuery As String = "Select * from RodDiameterDetails where DrawingPartNumber = '" _
                            & RodDrawingNumber.ToString & "' and TableDrawing = 'Yes'"
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            RodLength = Math.Round(RodLength, 2)

            'For Rod Drawing
            If objDT.Rows.Count > 0 AndAlso strRodMaterial.IndexOf("-08-I") = -1 Then


                '25_09_2009  ragava
                If RodDrawingNumber.ToString.StartsWith("4") = True Or RodDrawingNumber.ToString.StartsWith("7") = True Then
                    strDrawingNetworkPath = strNetworkPath & "400\" '& RodDrawingNumber.ToString & ".SLDDRW"
                ElseIf RodDrawingNumber.ToString.StartsWith("6") = True Or RodDrawingNumber.ToString.StartsWith("8") = True Then
                    strDrawingNetworkPath = strNetworkPath & "600\" '& RodDrawingNumber.ToString & ".SLDDRW"
                ElseIf RodDrawingNumber.ToString.StartsWith("2") = True Then
                    strDrawingNetworkPath = strNetworkPath & "200\" '& RodDrawingNumber.ToString & ".SLDDRW"
                End If
                '25_09_2009  ragava   Till  Here


                'If File.Exists(strDefaultPath & RodDrawingNumber.ToString & ".SLDDRW") = False Then
                If File.Exists(strDrawingNetworkPath & RodDrawingNumber.ToString & ".SLDDRW") = False Then
                    MsgBox("Drawing File " & strDrawingNetworkPath & RodDrawingNumber.ToString _
                        & ".SLDDRW doesn't exist, Please Copy the File to specified Location and then Click Ok")
                End If
                'OpenDrawingAndActivateSheet(strDrawingNetworkPath & RodDrawingNumber.ToString & ".SLDDRW")     '20_10_2009  ragava
                strQuery = ""
                strQuery = "Select CodeNumber,Dim_A,Revision from RodTableDrawing where DrawingNumber = '" & RodDrawingNumber & "'"
                Dim objDT2 As DataTable = oDataClass.GetDataTable(strQuery)
                Dim blnInsert As Boolean = True
                If objDT2.Rows.Count > 0 Then
                    For Each dr As DataRow In objDT2.Rows
                        'If dr(1).ToString = Format(RodLength, "00.00").ToString Then 
                        If dr(1).ToString = Format(RodLength, "0.00").ToString Then               '12_10_2009   ragava
                            blnInsert = False
                            'ANUP 26-10-2010 START
                            IsRowInserted_Rod = False
                            'ANUP 26-10-2010 TILL HERE
                            strCodeNumber = dr(0).ToString
                            ht_CodeNumbers.Add("ROD", strCodeNumber)        '04_10_2010   RAGAVA   TESTING
                            Exit For
                        End If
                    Next
                    If blnInsert = True Then
                        'ANUP 26-10-2010 START
                        IsRowInserted_Rod = True
                        'ANUP 26-10-2010 TILL HERE
                        'strQuery = "Select CodeNumber,Type from CodeNumberDetails where Type = 'ROD'"
                        strQuery = "Select CodeNumber,Type,MaxCodeNumber from CodeNumberDetails where Type = 'ROD'"      '12_10_2009   ragava
                        Dim objDT6 As DataTable = oDataClass.GetDataTable(strQuery)
                        strCodeNumber = objDT6.Rows(0).Item(0).ToString()
                        '12_10_2009  ragava
                        If Val(strCodeNumber) >= objDT6.Rows(0).Item(2) Then
                            MsgBox("Generated CodeNumber is Invalid asper MonarchIndustries... " & strCodeNumber)
                            Exit Try
                            'strCodeNumber = "749999"
                        End If
                        '12_10_2009  ragava    Till   Here
                        '07_10_2009   ragava
                        Dim strRevision As String = "1"
                        Try
                            strQuery = "select max(Revision) from RodTableDrawing where DrawingNumber = '" _
                                                        & RodDrawingNumber.ToString & "'"
                            Dim objDT7 As DataTable = oDataClass.GetDataTable(strQuery)
                            strRevision = (Val(objDT7.Rows(0).Item(0)) + 1).ToString
                            strQuery = ""
                        Catch ex As Exception
                        End Try
                        '07_10_2009   ragava      Till    Here
                        'strQuery = "Insert into RodTableDrawing Values('" & RodDrawingNumber.ToString & "','" & strCodeNumber.ToString & "','" & RodLength.ToString & "','" & (RodLength - dblRodStrokeDifference).ToString & "','1')"
                        strQuery = "Insert into RodTableDrawing Values('" & RodDrawingNumber.ToString & "','" _
                                & strCodeNumber.ToString & "','" & RodLength.ToString & "','" & _
                                (RodLength - dblRodStrokeDifference).ToString & "','" & strRevision & "')"

                        Dim objDT5 As DataTable = oDataClass.GetDataTable(strQuery)
                        strQuery = ""
                        strQuery = "Update CodeNumberDetails set CodeNumber ='" & (Val(strCodeNumber) + 1).ToString & "'"          '30_09_2009  ragava
                        Dim objDT_Temp As DataTable = oDataClass.GetDataTable(strQuery)
                        'End If      '20_10_2009   ragava
                        Try
                            ht_CodeNumbers.Add("ROD", strCodeNumber)         '20_01_2011   RAGAVA
                        Catch ex As Exception
                        End Try
                        strQuery = ""
                        strQuery = "Select * from RodTableDrawing where DrawingNumber = '" & RodDrawingNumber & "' order by Dim_A"
                        Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery)
                        If objDT1.Rows.Count > 0 Then
                            OpenDrawingAndActivateSheet(strDrawingNetworkPath & RodDrawingNumber.ToString & ".SLDDRW")        '20_10_2009  ragava
                            If File.Exists("C:\DESIGN_TABLES\TableDrawing\Rod.xls") = False Then
                                MsgBox("Please Create an Excel file at following Path : C:\DESIGN_TABLES\TableDrawing\Rod.xls  and then Click Ok")
                            End If
                            CreateExcel("C:\DESIGN_TABLES\TableDrawing\Rod.xls", objDT1)
                            System.Threading.Thread.Sleep(1000)            '06_10_2009   ragava
                            insert265BOM("C:\DESIGN_TABLES\TableDrawing\Rod.xls", 0.015, 0.25, 0)      '05_10_2009   ragava
                            System.Threading.Thread.Sleep(1000)            '06_10_2009   ragava
                            UpdateRevisionTable(strCodeNumber, strRevision)        '21_10_2009   ragava
                        End If
                        'insert265BOM("C:\DESIGN_TABLES\TableDrawing\Rod.xls", 0.015, 0.25, 0)    '05_10_2009   ragava
                        System.Threading.Thread.Sleep(1000)            '06_10_2009   ragava
                        'IFLSolidWorksBaseClassObject.SolidWorksModel.SaveSilent()          '04_02_2010     RAGAVA

                        '29_01_2010   ragava
                        Try
                            If Directory.Exists("w:\") = True Then
                                IFLSolidWorksBaseClassObject.SolidWorksModel.SaveAs("w:\" & RodDrawingNumber.ToString & ".SLDDRW")
                            End If
                        Catch ex As Exception
                            MsgBox("ERROR IN SAVING DRAWING FILE : " & "w:\" & RodDrawingNumber.ToString & ".SLDDRW")
                        End Try
                        'Till Here

                        'IFLSolidWorksBaseClassObject.SaveAndCloseAllDocuments()         '04_02_2010     RAGAVA
                        IFLSolidWorksBaseClassObject.CloseAllDocuments()        '04_02_2010     RAGAVA
                    End If          '20_10_2009  ragava
                    Dim strPart() As String = strRodDrawingPath.Split(".")
                    Dim strPartPath As String
                    If strPart.Length > 2 Then
                        strPartPath = strPart(LBound(strPart)) & "." & strPart(LBound(strPart) + 1)
                    Else
                        strPartPath = strPart(LBound(strPart))
                    End If
                    Dim strNewPartPath As String = strPartPath.Substring(0, strPartPath.LastIndexOf("\")) & "\" _
                            & strCodeNumber.ToString & ".SLDPRT"
                    'Renaming Rod Part
                    RenamePartFile(strPartPath & ".SLDPRT", strNewPartPath, DestinationFilePath _
                                & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM")
                    FolderStructure(DestinationFilePath & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM", _
                                strNewPartPath, strModelNetworkPath)     '25_09_2009  ragava

                End If
            Else

                '10_12_2010   RAGAVA
                Try
                    strCodeNumber = ""
                    RodLength = Math.Round(RodLength, 2)
                    Dim strQuery1 As String = "Select PartNumber from RodDiameterDetails where DrawingPartNumber = '" _
                        & RodDrawingNumber.ToString & "' and TableDrawing = 'No' and OverAllRodLength ='" & _
                        Format(RodLength, "0.00").ToString & "'"
                    Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery1)
                    If objDT1.Rows.Count > 0 Then
                        strCodeNumber = objDT1.Rows(0)("PartNumber")
                        ht_CodeNumbers.Add("ROD", strCodeNumber)
                    End If
                Catch ex As Exception
                End Try
                'Till   Here

                OpenDrawingAndActivateSheet(strRodDrawingPath)
                IFLSolidWorksBaseClassObject.SolidWorksModel.SaveSilent()
                IFLSolidWorksBaseClassObject.SaveAndCloseAllDocuments()
                '01_10_2009   ragava
                'If strStyle.IndexOf("ASAE") <> -1 Then
                If strStyle.IndexOf("ASAE") <> -1 AndAlso strCodeNumber = "" Then       '10_12_2010   RAGAVA
                    strCodeNumber = RodCodeNumber
                End If
                '01_10_2009   ragava         Till  Here
                If strCodeNumber = "" Then
                    strQuery = ""
                    'strQuery = "Select CodeNumber,Type from CodeNumberDetails where Type = 'ROD'"
                    strQuery = "Select CodeNumber,Type,MaxCodeNumber from CodeNumberDetails where Type = 'ROD'"       '12_10_2009   ragava
                    Dim objDT3 As DataTable = oDataClass.GetDataTable(strQuery)
                    strCodeNumber = objDT3.Rows(0).Item(0).ToString()
                    '12_10_2009  ragava
                    If Val(strCodeNumber) >= objDT3.Rows(0).Item(2) Then
                        MsgBox("Generated CodeNumber is Invalid asper MonarchIndustries... " & strCodeNumber)
                        Exit Try
                        'strCodeNumber = "749999"
                    End If
                    Try
                        ht_CodeNumbers.Add("ROD", strCodeNumber)         '20_01_2011   RAGAVA
                    Catch ex As Exception
                    End Try
                    '12_10_2009  ragava    Till   Here
                    strQuery = "Update CodeNumberDetails set CodeNumber ='" & (Val(strCodeNumber) + 1).ToString & "'"          '30_09_2009  ragava
                    objDT3.Clear()
                    objDT3 = oDataClass.GetDataTable(strQuery)
                End If
                Dim strPart() As String = strRodDrawingPath.Split(".")
                'Dim strPartPath As String = strPart(LBound(strPart))
                Dim strPartPath As String
                If strPart.Length > 2 Then
                    strPartPath = strPart(LBound(strPart)) & "." & strPart(LBound(strPart) + 1)
                Else
                    strPartPath = strPart(LBound(strPart))
                End If
                Dim strNewPartPath As String = strPartPath.Substring(0, strPartPath.LastIndexOf("\")) & "\" _
                                & strCodeNumber.ToString & ".SLDPRT"
                'Renaming Rod Part
                RenamePartFile(strPartPath & ".SLDPRT", strNewPartPath, DestinationFilePath _
                        & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM")
                FolderStructure(DestinationFilePath & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM", _
                            strNewPartPath, strModelNetworkPath)     '25_09_2009  ragava
                'Renaming Drawing  strDrawingNetworkPath
                RenameDrawingFile(strRodDrawingPath, strPartPath & ".SLDPRT", strNewPartPath)
                MoveDrawingFile(strNewPartPath.Replace(".SLDPRT", ".SLDDRW"), strNewPartPath, _
                    strNetworkPath & "400\" & strCodeNumber & ".SLDDRW", strModelNetworkPath & strCodeNumber & ".SLDPRT")
                'DrawingFolderStructure(strNetworkPath & "400\" & strCodeNumber & ".SLDDRW",
            End If
        Catch ex As Exception
        End Try

        '******************************************************************************************************

       
        'Stop Tube Table Drawing
        Try
            If dblStopTubeLength > 0 Then
                Dim strCodeNumber As String = String.Empty
                Dim strQuery As String = String.Empty
                Dim oDataClass As New DataClass
                Dim strStopTubePartPath As String = DestinationFilePath & "\Stoptube\Stop_tube.SLDPRT"
                Dim strDefaultPath As String = "C:\TableDrawingFile\"
                StopTubeDrawingNumber = "495598"

                '25_09_2009  ragava
                If StopTubeDrawingNumber.ToString.StartsWith("4") = True Or StopTubeDrawingNumber.ToString.StartsWith("7") = True Then
                    strDrawingNetworkPath = strNetworkPath & "400\" '& StopTubeDrawingNumber.ToString & ".SLDDRW"
                ElseIf TieRodDrawingNumber.ToString.StartsWith("6") = True Or StopTubeDrawingNumber.ToString.StartsWith("8") = True Then
                    strDrawingNetworkPath = strNetworkPath & "600\" '& StopTubeDrawingNumber.ToString & ".SLDDRW"
                ElseIf TieRodDrawingNumber.ToString.StartsWith("2") = True Then
                    strDrawingNetworkPath = strNetworkPath & "200\" '& TieRodDrawingNumber.ToString & ".SLDDRW"
                End If
                '25_09_2009  ragava   Till  Here


                If File.Exists(strDrawingNetworkPath & StopTubeDrawingNumber.ToString & ".SLDDRW") = False Then
                    MsgBox("Drawing File " & strDrawingNetworkPath & StopTubeDrawingNumber.ToString & ".SLDDRW doesn't exist, Please Copy the File to specified Location and then Click Ok")
                End If
                'OpenDrawingAndActivateSheet(strDrawingNetworkPath & StopTubeDrawingNumber.ToString & ".SLDDRW")      '20_10_2009  ragava
                strQuery = ""
                strQuery = "Select Dim_A,Dim_B,Dim_C,CodeNumber from StopTubeTableDrawing where DrawingNumber = '" _
                            & StopTubeDrawingNumber & "'"
                Dim objDT2 As DataTable = oDataClass.GetDataTable(strQuery)
                Dim blnInsert As Boolean = True
                Dim dblDim_B As Double = 0
                Dim dblDim_C As Double = Math.Round((dblRodDiameter + 0.015), 2)
                If dblRodDiameter <= 1.12 Then
                    dblDim_B = Math.Round((dblRodDiameter + 0.015) + (2 * 0.19), 2)
                Else
                    dblDim_B = Math.Round((dblRodDiameter + 0.015) + (2 * 0.25), 2)
                End If
                If objDT2.Rows.Count > 0 Then
                    For Each dr As DataRow In objDT2.Rows
                        'If (dr(0).ToString = Format(dblStopTubeLength, "00.00").ToString) AndAlso (dr(1).ToString = Format(dblDim_B, "00.00").ToString) AndAlso (dr(2).ToString = Format(dblDim_C, "00.00").ToString) Then
                        If (dr(0).ToString = Format(dblStopTubeLength, "0.00").ToString) AndAlso (dr(1).ToString = _
                            Format(dblDim_B, "0.00").ToString) AndAlso (dr(2).ToString = Format(dblDim_C, "0.00").ToString) Then         '12_10_2009   ragava
                            blnInsert = False
                            'ANUP 26-10-2010 START
                            IsRowInserted_StopTube = False
                            'ANUP 26-10-2010 TILL HERE
                            strCodeNumber = dr(3).ToString
                            ht_CodeNumbers.Add("STOPTUBE", strCodeNumber)        '04_10_2010   RAGAVA   TESTING
                            Exit For
                        End If
                    Next
                    If blnInsert = True Then
                        'ANUP 26-10-2010 START
                        IsRowInserted_StopTube = True
                        'ANUP 26-10-2010 TILL HERE
                        OpenDrawingAndActivateSheet(strDrawingNetworkPath & StopTubeDrawingNumber.ToString & ".SLDDRW")      '20_10_2009  ragava
                        'strQuery = "Select CodeNumber,Type from CodeNumberDetails where Type = 'StopTube'"
                        strQuery = "Select CodeNumber,Type,MaxCodeNumber from CodeNumberDetails where Type = 'StopTube'"          '12_10_2009  ragava
                        Dim objDT6 As DataTable = oDataClass.GetDataTable(strQuery)
                        strCodeNumber = objDT6.Rows(0).Item(0).ToString()
                        '12_10_2009  ragava
                        If Val(strCodeNumber) >= objDT6.Rows(0).Item(2) Then
                            MsgBox("Generated CodeNumber is Invalid asper Monarch Industries... " & strCodeNumber)
                            Exit Try
                            'strCodeNumber = "749999"
                        End If
                        '12_10_2009  ragava    Till   Here

                        Try
                            ht_CodeNumbers.Add("STOPTUBE", strCodeNumber)         '21_07_2011    RAGAVA
                        Catch ex As Exception
                        End Try

                        '07_10_2009   ragava
                        Dim strRevision As String = "1"
                        Try
                            strQuery = "select max(Revision) from StopTubeTableDrawing where DrawingNumber = '" _
                                            & StopTubeDrawingNumber.ToString & "'"
                            Dim objDT7 As DataTable = oDataClass.GetDataTable(strQuery)
                            strRevision = (Val(objDT7.Rows(0).Item(0)) + 1).ToString
                            strQuery = ""
                        Catch ex As Exception
                        End Try
                        '07_10_2009   ragava      Till    Here

                        'strQuery = "Insert into StopTubeTableDrawing Values('" & StopTubeDrawingNumber.ToString & "','" & strCodeNumber.ToString & "','" & (dblStopTubeLength.ToString) & "','" & (dblDim_B).ToString & "','" & (dblDim_C).ToString & "','1')"
                        strQuery = "Insert into StopTubeTableDrawing Values('" & StopTubeDrawingNumber.ToString _
                            & "','" & strCodeNumber.ToString & "','" & (dblStopTubeLength.ToString) & "','" & _
                            (dblDim_B).ToString & "','" & (dblDim_C).ToString & "','" & strRevision & "')"                '07_10_2009  ragava

                        Dim objDT5 As DataTable = oDataClass.GetDataTable(strQuery)
                        strQuery = ""
                        strQuery = "Update CodeNumberDetails set CodeNumber ='" & (Val(strCodeNumber) + 1).ToString & "'"          '30_09_2009  ragava
                        Dim objDT_Temp As DataTable = oDataClass.GetDataTable(strQuery)
                        'End If      '20_10_2009  ragava
                        strQuery = ""
                        'strQuery = "Select * from StopTubeTableDrawing where DrawingNumber = '" & StopTubeDrawingNumber & "' order by Dim_A,Dim_B,Dim_C"
                        strQuery = "Select * from StopTubeTableDrawing where DrawingNumber = '" & _
                            StopTubeDrawingNumber & "' order by Dim_C,Dim_A,Dim_B"

                        Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery)
                        If objDT1.Rows.Count > 0 Then
                            If File.Exists("C:\DESIGN_TABLES\TableDrawing\StopTube.xls") = False Then
                                MsgBox("Please Create an Excel file at following Path : C:\DESIGN_TABLES\TableDrawing\StopTube.xls  and then Click Ok")
                            End If
                            CreateExcel("C:\DESIGN_TABLES\TableDrawing\StopTube.xls", objDT1)
                            System.Threading.Thread.Sleep(1000)            '06_10_2009   ragava
                            insert265BOM("C:\DESIGN_TABLES\TableDrawing\StopTube.xls", 0.015, 0.25, 0)    '05_10_2009   ragava
                            System.Threading.Thread.Sleep(1000)            '06_10_2009   ragava
                            UpdateRevisionTable(strCodeNumber, strRevision)        '21_10_2009   ragava
                        End If
                        'insert265BOM("C:\DESIGN_TABLES\TableDrawing\StopTube.xls", 0.015, 0.25, 0)    '05_10_2009   ragava
                        'IFLSolidWorksBaseClassObject.SolidWorksModel.SaveSilent()          '04_02_2010     RAGAVA

                        '29_01_2010   ragava
                        Try
                            If Directory.Exists("w:\") = True Then
                                IFLSolidWorksBaseClassObject.SolidWorksModel.SaveAs("w:\" & StopTubeDrawingNumber.ToString & ".SLDDRW")
                            End If
                        Catch ex As Exception
                            MsgBox("ERROR IN SAVING DRAWING FILE : " & "w:\" & StopTubeDrawingNumber.ToString & ".SLDDRW")
                        End Try
                        'Till Here

                        'IFLSolidWorksBaseClassObject.SaveAndCloseAllDocuments()         '04_02_2010     RAGAVA
                        IFLSolidWorksBaseClassObject.CloseAllDocuments()        '04_02_2010     RAGAVA
                    End If      '20_10_2009  ragava
                    Dim strPart() As String = strStopTubePartPath.Split(".")
                    Dim strPartPath As String
                    strPartPath = strPart(LBound(strPart))
                    Dim strNewPartPath As String = strPartPath.Substring(0, strPartPath.LastIndexOf("\")) & "\" _
                        & strCodeNumber.ToString & ".SLDPRT"
                    'Renaming StopTube Part
                    RenamePartFile(strStopTubePartPath, strNewPartPath, DestinationFilePath & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM")
                    System.Threading.Thread.Sleep(1000)            '06_10_2009   ragava
                    FolderStructure(DestinationFilePath & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM", strNewPartPath, strModelNetworkPath)     '25_09_2009  ragava
                End If
            End If
        Catch ex As Exception
        End Try


        '25_09_2009  ragava
        'Renaming Main Assy
        Try
            '01_10_2009  ragava
            Dim strMainAssyDrawing As String = String.Empty
            If ClevisCapPortOrientation.IndexOf("Inline") <> -1 Then
                strMainAssyDrawing = "MAIN_ASSEMBLY.SLDDRW"
            ElseIf ClevisCapPortOrientation.IndexOf("90 Degrees") <> -1 Then
                strMainAssyDrawing = "MAIN_ASSEMBLY_90.SLDDRW"
            End If
            '01_10_2009  ragava   Till  Here

            Dim bret As Boolean = False
            Rename(DestinationFilePath & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM", DestinationFilePath _
                    & "\TIE_ROD_ASSEMBLY\" & PartCode.ToString & ".SLDASM")
            'bret = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.ReplaceReferencedDocument(DestinationFilePath & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDDRW", DestinationFilePath & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM", DestinationFilePath & "\TIE_ROD_ASSEMBLY\" & PartCode.ToString & ".SLDASM")
            bret = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.ReplaceReferencedDocument _
                (DestinationFilePath & "\TIE_ROD_ASSEMBLY\" & strMainAssyDrawing, DestinationFilePath _
                & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDASM", DestinationFilePath & "\TIE_ROD_ASSEMBLY\" & PartCode.ToString & ".SLDASM")
            'Rename(DestinationFilePath & "\TIE_ROD_ASSEMBLY\MAIN_ASSEMBLY.SLDDRW", DestinationFilePath & "\TIE_ROD_ASSEMBLY\" & PartCode.ToString & ".SLDDRW")
            Rename(DestinationFilePath & "\TIE_ROD_ASSEMBLY\" & strMainAssyDrawing, DestinationFilePath _
                        & "\TIE_ROD_ASSEMBLY\" & PartCode.ToString & ".SLDDRW")
            File.Copy(DestinationFilePath & "\TIE_ROD_ASSEMBLY\" & PartCode.ToString & ".SLDASM", _
                strModelNetworkPath & PartCode.ToString & ".SLDASM", True)
            bret = False      '30_09_2009  ragava
            bret = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.ReplaceReferencedDocument _
                (DestinationFilePath & "\TIE_ROD_ASSEMBLY\" & PartCode.ToString & ".SLDDRW", DestinationFilePath _
                & "\TIE_ROD_ASSEMBLY\" & PartCode.ToString & ".SLDASM", strModelNetworkPath & PartCode.ToString & ".SLDASM")     '30_09_2009  ragava

            '10_02_2010   ragava
            Try
                If Directory.Exists("w:\") = True Then
                    File.Copy(DestinationFilePath & "\TIE_ROD_ASSEMBLY\" & PartCode.ToString & ".SLDDRW", "W:\" _
                                & PartCode.ToString & ".SLDDRW", True)
                End If
            Catch ex As Exception
                MsgBox("ERROR IN SAVING DRAWING FILE : " & "w:\" & PartCode.ToString & ".SLDDRW")
            End Try
            'File.Copy(DestinationFilePath & "\TIE_ROD_ASSEMBLY\" & PartCode.ToString & ".SLDDRW", "X:\600\" & PartCode.ToString & ".SLDDRW", True)
            'Till Here

            'File.Copy(DestinationFilePath & "\TIE_ROD_ASSEMBLY\" & PartCode.ToString & ".SLDDRW", "C:\Monarch_CDA\600\" & PartCode.ToString & ".SLDDRW", True)        '02_10_2009  ragava
            System.Threading.Thread.Sleep(2000)            '06_10_2009    ragava
            '30_09_2009  ragava

            IFLSolidWorksBaseClassObject.openDocument("X:\TieRodModels\" & PartCode.ToString & ".SLDASM")    '06_10_2009   ragava
            System.Threading.Thread.Sleep(3000)            '06_10_2009   ragava
            'IFLSolidWorksBaseClassObject.openDocument("X:\600\" & PartCode.ToString & ".SLDDRW")

            'IFLSolidWorksBaseClassObject.openAssemblyDrawingDocument("X:\600\" & PartCode.ToString & ".SLDDRW")
            ' Sugandhi        'IFLSolidWorksBaseClassObject.openAssemblyDrawingDocument("W:\" & PartCode.ToString & ".SLDDRW")   '10_02_2010   RAGAVA

            'IFLSolidWorksBaseClassObject.openDocument("C:\Monarch_CDA\600\" & PartCode.ToString & ".SLDDRW")    '02_10_2009  ragava
            IFLSolidWorksBaseClassObject.SolidWorksModel = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.ActiveDoc
            IFLSolidWorksBaseClassObject.SolidWorksModel.EditRebuild3()
            IFLSolidWorksBaseClassObject.SolidWorksModel.SaveSilent()
            IFLSolidWorksBaseClassObject.SolidWorksModel.SaveAs2("W:\" & PartCode.ToString & ".SLDDRW", 3, True, True)           '03_02_2011   RAGAVA
            IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CloseAllDocuments(True)
            '30_09_2009  ragava    Till  Here
        Catch ex As Exception
        End Try
        '25_09_2009  ragava   Till  Here
        '01_10_2009  ragava
        Try
            KillExcel()
        Catch ex As Exception
        End Try
        Try
            IFLSolidWorksBaseClassObject.KillAllSolidWorksServices()
        Catch ex As Exception
        End Try
        'Try
        '    Directory.Delete("C:\MONARCH_TESTING", True)
        'Catch ex As Exception
        '    'MsgBox(ex.Message)
        'End Try
        'Try
        '    IFLSolidWorksBaseClassObject.ConnectSolidWorks()
        '    IFLSolidWorksBaseClassObject.openDocument("X:\600\" & PartCode.ToString & ".SLDDRW")
        '    'IFLSolidWorksBaseClassObject.openDocument("C:\Monarch_CDA\600\" & PartCode.ToString & ".SLDDRW")     '02_10_2009  ragava

        'Catch ex As Exception
        'End Try
        '01_10_2009  ragava   Till  Here
    End Sub

End Module
