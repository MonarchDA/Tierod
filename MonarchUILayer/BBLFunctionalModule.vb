Imports System.io
Imports System.Drawing.Imaging
Imports System.Threading
Imports System
Imports Microsoft.Office.Interop.Excel


Public Module BBLFunctionalModule
    'Public ofrmAdminModule As New frmAdminModule ' Coded by satish 16-06-09
    'Public ofrmCoreDetails As New frmCoreDetails
    ' Public ofa As New frmCoreDetails
    'Public ofrmDeckPlatePermawoodDetailsvb As New frmDeckPlatePermawoodDetails
    'Public ofrmDuctandDrawingOutput As New frmDuctandDrawingOutput
    'Public ofrmInputs As New frmInputs
    'Public ofrmWorkOrder As New frmWorkOrder
    'Public ofrmDBLogin As New frmDBLogin
    'Dim oMagneticCircuitClass As New MagneticCircuitClass
    'Private _oBottomFrameBaseClass As New BottomFrameBaseClass '4-06
    'Public oFunctionalClass As New FunctionalClass
    'Dim oCoreFrameHVLVBaseClass As New CoreFrameHVLVBaseClass
    Dim filePath, libPath, strFileName As String

    Dim fso As New Scripting.FileSystemObject
    'Private _oTopFrameHVLVSingleBand As New TopFrameLvHvSideSingleband
    'Private _oTopFrameHVLVTwoBand As New TopFrameLvHvTwobandClass
    ''  Private _oCoreFrameHVLVBaseClass As New CoreFrameHVLVBaseClass
    'Dim oDeckAssembly As New DeckPlateClass
    'Dim oCrossFlat As New CrossFlatClass
    'Dim oBridgeClass As New BridgeClass
    'Dim oYokeClass As New YokeClass
    'Dim oPermaWoodClass As New PermawoodSegmentClass
    'Private oStandardPartsClass As New StandardPartsClass
    Dim _oThreadProgressBarStepping As System.Threading.Thread
    'Dim _oBBLDataClass As New BBLDataClass
    Private _alUserInputs As New ArrayList
    Private _oExcelApplication As Microsoft.Office.Interop.Excel.Application
    Private _oExcelWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Private _oExcelWorkSheet As Microsoft.Office.Interop.Excel.Worksheet



#Region "Properties"
    Private _oStopWatch As New Diagnostics.Stopwatch
    Private _pbBBL As ProgressBar

    'Public ReadOnly Property BBLDataClassObject() As BBLDataClass
    '    Get
    '        If _oBBLDataClass Is Nothing Then
    '            _oBBLDataClass = New BBLDataClass
    '        End If
    '        Return _oBBLDataClass
    '    End Get
    'End Property
    Public Property BBLStopWatch() As Diagnostics.Stopwatch
        Get
            Return _oStopWatch
        End Get
        Set(ByVal value As Diagnostics.Stopwatch)
            _oStopWatch = value
        End Set
    End Property
    Public WriteOnly Property StopWatchAndProgressBar() As String
        Set(ByVal value As String)
            If value = "Start" Then
                BBLStopWatch.Reset()
                BBLStopWatch.Start()
                Control.CheckForIllegalCrossThreadCalls = False
                ' StartStepingProgressBar()
                '_oThreadProgressBarStepping = New System.Threading.Thread(New Threading.ThreadStart(AddressOf StartStepingProgressBar))
                '_oThreadProgressBarStepping.IsBackground = True
                '_oThreadProgressBarStepping.Start()
            ElseIf value = "Stop" Then
                BBLStopWatch.Stop()
                If _oThreadProgressBarStepping.IsAlive Then
                    SetPB.Value = SetPB.Maximum
                    SetPB.Value = 0
                    _oThreadProgressBarStepping.Resume()
                    _oThreadProgressBarStepping.Abort()
                End If
            ElseIf value = "Suspend" Then
                BBLStopWatch.Stop()
                _oThreadProgressBarStepping.Suspend()
            ElseIf value = "Resume" Then
                BBLStopWatch.Start()
                _oThreadProgressBarStepping.Resume()
            End If
        End Set
    End Property
    Public Property SetPB() As ProgressBar
        Get
            Return _pbBBL
        End Get
        Set(ByVal value As ProgressBar)
            _pbBBL = value
        End Set
    End Property


    Public ReadOnly Property MasterUserInputsFilePath() As String
        Get
            'MasterUserInputsFilePath = Execution_Path + "\Input Screen_final.xls"
            Return MasterUserInputsFilePath
        End Get
    End Property

    Public ReadOnly Property WorkOrderUserInputsFilepath() As String
        Get
            'WorkOrderUserInputsFilepath = DestinationFilePath + "\Exceloutput_" + WorkOrder + "\Input Screen_final.xls"
            Return WorkOrderUserInputsFilepath
        End Get
    End Property

    'Public ReadOnly Property UserInputs() As ArrayList
    '    Get
    '        If _alUserInputs.Count = 0 Then

    '            _alUserInputs.Add(New Object(1) {"W_O_NO", WorkOrder})
    '            _alUserInputs.Add(New Object(1) {"CUSTOMER", CustomerName})
    '            _alUserInputs.Add(New Object(1) {"PREPARED_BY", PreparedBy})
    '            _alUserInputs.Add(New Object(1) {"CHECKED_BY", CheckedBy})
    '            _alUserInputs.Add(New Object(1) {"DATE", System.DateTime.Today})

    '            _alUserInputs.Add(New Object(1) {"Core_Dia_in_mm", FunctionalClassObject.getValueGridValues("Dfe", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"Window_width_in_mm", FunctionalClassObject.getValueGridValues("B", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"Leg_Centre_in_mm", FunctionalClassObject.getValueGridValues("L.C.", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"Window_Height_in_mm", FunctionalClassObject.getValueGridValues("A", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"Total_Stack_Height_including_ducts_in_mm", FunctionalClassObject.getValueGridValues("H", arListInputs)})

    '            _alUserInputs.Add(New Object(1) {"Innermost_winding_I_D_in_mm", FunctionalClassObject.getValueGridValues("D1", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"Innermost_winding_O_D_in_mm", FunctionalClassObject.getValueGridValues("D2", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"Innermost_winding_axial_height_in_mm", FunctionalClassObject.getValueGridValues("WL1", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"Innermost_winding_position_above_bottom_yoke_top_in_mm", FunctionalClassObject.getValueGridValues("M2", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"Outermost_winding_I_D_in_mm", FunctionalClassObject.getValueGridValues("D3", arListInputs)})

    '            _alUserInputs.Add(New Object(1) {"Outermost_winding_O_D_in_mm", FunctionalClassObject.getValueGridValues("D4", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"Outermost_winding_axial_height_in_mm", FunctionalClassObject.getValueGridValues("WL2", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"Outermost_winding_position_above_bottom_yoke_top_in_mm", FunctionalClassObject.getValueGridValues("M4", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"Core_&_Winding_Weight_in_kg", FunctionalClassObject.getValueGridValues("C&CWT", arListInputs)})

    '            _alUserInputs.Add(New Object(1) {"Coil_clamping_force_in_kg", FunctionalClassObject.getValue("frmDeckPlatePermawoodDetails", "CoilClampingForce")})
    '            _alUserInputs.Add(New Object(1) {"Deck_Plate_Insulation_thk_in_mm", FunctionalClassObject.getValue("frmDeckPlatePermawoodDetails", "DPInsulationThk")})
    '            _alUserInputs.Add(New Object(1) {"Deck_Plate_Flat_Width", FunctionalClassObject.getValue("frmDeckPlatePermawoodDetails", "DPFlatWidth")})
    '            _alUserInputs.Add(New Object(1) {"Deck_Plate_Flat_thk_", FunctionalClassObject.getValue("frmDeckPlatePermawoodDetails", "DPFlatThickness")})
    '            _alUserInputs.Add(New Object(1) {"No_of_Flats_/_Deck_Plate", FunctionalClassObject.getValue("frmDeckPlatePermawoodDetails", "NoofFlats")})
    '            _alUserInputs.Add(New Object(1) {"Permawood_Bottom_Segment_/_Ring", FunctionalClassObject.getValue("frmDeckPlatePermawoodDetails", "BottomSegment")})

    '            _alUserInputs.Add(New Object(1) {"Permawood_Top_Segment_/_Ring", FunctionalClassObject.getValue("frmDeckPlatePermawoodDetails", "TopSegmentRing")})
    '            _alUserInputs.Add(New Object(1) {"Top_Segment_/_Ring_Thk_in_mm", FunctionalClassObject.getValue("frmDeckPlatePermawoodDetails", "TopSegmentRingThk")})
    '            _alUserInputs.Add(New Object(1) {"Coil_Clamping_Screw_/_Pressure_Block", FunctionalClassObject.getValue("frmDeckPlatePermawoodDetails", "CoilClampingScrew")})
    '            _alUserInputs.Add(New Object(1) {"Max__Width_C1", FunctionalClassObject.getValueGridValues("C1", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C2", FunctionalClassObject.getValueGridValues("C2", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C3", FunctionalClassObject.getValueGridValues("C3", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C4", FunctionalClassObject.getValueGridValues("C4", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C5", FunctionalClassObject.getValueGridValues("C5", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C6", FunctionalClassObject.getValueGridValues("C6", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C7", FunctionalClassObject.getValueGridValues("C7", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C8", FunctionalClassObject.getValueGridValues("C8", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C9", FunctionalClassObject.getValueGridValues("C9", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C10", FunctionalClassObject.getValueGridValues("C10", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C11", FunctionalClassObject.getValueGridValues("C11", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C12", FunctionalClassObject.getValueGridValues("C12", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C13", FunctionalClassObject.getValueGridValues("C13", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C14", FunctionalClassObject.getValueGridValues("C14", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C15", FunctionalClassObject.getValueGridValues("C15", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C16", FunctionalClassObject.getValueGridValues("C16", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C17", FunctionalClassObject.getValueGridValues("C17", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C18", FunctionalClassObject.getValueGridValues("C18", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C19", FunctionalClassObject.getValueGridValues("C19", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"C20", FunctionalClassObject.getValueGridValues("C20", arListInputs)})

    '            _alUserInputs.Add(New Object(1) {"H1", FunctionalClassObject.getValueGridValues("H1", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H2", FunctionalClassObject.getValueGridValues("H2", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H3", FunctionalClassObject.getValueGridValues("H3", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H4", FunctionalClassObject.getValueGridValues("H4", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H5", FunctionalClassObject.getValueGridValues("H5", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H6", FunctionalClassObject.getValueGridValues("H6", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H7", FunctionalClassObject.getValueGridValues("H7", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H8", FunctionalClassObject.getValueGridValues("H8", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H9", FunctionalClassObject.getValueGridValues("H9", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H10", FunctionalClassObject.getValueGridValues("H10", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H11", FunctionalClassObject.getValueGridValues("H11", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H12", FunctionalClassObject.getValueGridValues("H12", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H13", FunctionalClassObject.getValueGridValues("H13", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H14", FunctionalClassObject.getValueGridValues("H14", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H15", FunctionalClassObject.getValueGridValues("H15", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H16", FunctionalClassObject.getValueGridValues("H16", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H17", FunctionalClassObject.getValueGridValues("H17", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H18", FunctionalClassObject.getValueGridValues("H18", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H19", FunctionalClassObject.getValueGridValues("H19", arListInputs)})
    '            _alUserInputs.Add(New Object(1) {"H20", FunctionalClassObject.getValueGridValues("H20", arListInputs)})

    '            _alUserInputs.Add(New Object(1) {"Duct_in_centre_step", FunctionalClassObject.getValue("frmDuctandDrawingOutput", "DuctArrangement")})
    '            _alUserInputs.Add(New Object(1) {"Psition_of_ducts_in_side_steps", FunctionalClassObject.getValue("frmDuctandDrawingOutput", "DuctPosition")})
    '            _alUserInputs.Add(New Object(1) {"Duct_thk__in_mm", FunctionalClassObject.getValue("frmDuctandDrawingOutput", "DuctThickness")})
    '            '_alUserInputs.Add(New Object(1) {"Drawing_Nos_", FunctionalClassObject.getValue("frmDuctandDrawingOutput", "DuctArrangement")})
    '            _alUserInputs.Add(New Object(1) {"Frame_Details", FunctionalClassObject.getValue("frmDuctandDrawingOutput", "FrameDetails")})

    '            _alUserInputs.Add(New Object(1) {"Frame_Core_Assy_", FunctionalClassObject.getValue("frmDuctandDrawingOutput", "FrameCoreAssembly")})
    '            _alUserInputs.Add(New Object(1) {"Bottom_Segment_/_Ring", FunctionalClassObject.getValue("frmDuctandDrawingOutput", "BottomSegmentRing")})
    '            _alUserInputs.Add(New Object(1) {"Top_Segment_/_Ring", FunctionalClassObject.getValue("frmDuctandDrawingOutput", "TopSegmentRing")})
    '            _alUserInputs.Add(New Object(1) {"Pressure_Blocks", FunctionalClassObject.getValue("frmDuctandDrawingOutput", "PressureBlocks")})
    '            _alUserInputs.Add(New Object(1) {"Yoke_Liner_Top", FunctionalClassObject.getValue("frmDuctandDrawingOutput", "YokeLinerTop")})
    '            _alUserInputs.Add(New Object(1) {"Yoke_Liner_Bottom", FunctionalClassObject.getValue("frmDuctandDrawingOutput", "YokeLinerBottom")})

    '        End If
    '        Return _alUserInputs
    '    End Get
    'End Property

#End Region
  

    Private Function GetNewFilenameWithoutInvalidCharacters(ByVal strNewFilename As String)
        GetNewFilenameWithoutInvalidCharacters = strNewFilename

        If strNewFilename.Contains("""") OrElse strNewFilename.Contains("?") OrElse strNewFilename.Contains("/") OrElse strNewFilename.Contains("\") OrElse strNewFilename.Contains(":") _
                   OrElse strNewFilename.Contains("*") OrElse strNewFilename.Contains("<") OrElse strNewFilename.Contains(">") OrElse strNewFilename.Contains("|") Then

            GetNewFilenameWithoutInvalidCharacters = strNewFilename.Replace(""""c, "_"c)
            GetNewFilenameWithoutInvalidCharacters = strNewFilename.Replace("?"c, "_"c)
            GetNewFilenameWithoutInvalidCharacters = strNewFilename.Replace("/"c, "_"c)
            GetNewFilenameWithoutInvalidCharacters = strNewFilename.Replace("\"c, "_"c)
            GetNewFilenameWithoutInvalidCharacters = strNewFilename.Replace(":"c, "_"c)
            GetNewFilenameWithoutInvalidCharacters = strNewFilename.Replace("*"c, "_"c)
            GetNewFilenameWithoutInvalidCharacters = strNewFilename.Replace("<"c, "_"c)
            GetNewFilenameWithoutInvalidCharacters = strNewFilename.Replace(">"c, "_"c)
            GetNewFilenameWithoutInvalidCharacters = strNewFilename.Replace("|"c, "_"c)

        End If

    End Function

    Public Sub CreateExcelFileinWorkOrderFolder()

        'If Not Directory.Exists(DestinationFilePath + "\Exceloutput_" + WorkOrder) Then
        '    Directory.CreateDirectory(DestinationFilePath + "\Exceloutput_" + WorkOrder)
        'End If

    End Sub

    Public Sub SavetoExcelFile()
        _oExcelApplication = New Microsoft.Office.Interop.Excel.ApplicationClass
        If WorkOrderInputsFolderExists() Then
            CreateExcelFile()
        End If

        _oExcelWorkBook = _oExcelApplication.Workbooks.Open(WorkOrderUserInputsFilepath)
        _oExcelWorkSheet = _oExcelWorkBook.Worksheets("PARAMETERS")
        SaveUserInputsToExcel()
        _oExcelWorkBook.Save()
        _oExcelWorkBook.Close()
        _oExcelApplication.Quit()
    End Sub

    Private Sub SaveUserInputsToExcel()

        'For Each obj As Object In UserInputs
        '    For iRowIndex As Integer = 1 To 80
        '        If _oExcelWorkSheet.Cells(iRowIndex, 1).Value = obj(Parameters.Name) Then
        '            _oExcelWorkSheet.Cells(iRowIndex, 2).Value = obj(Parameters.Value)
        '            Exit For
        '        End If
        '    Next
        'Next
    End Sub

    Private Sub CreateExcelFile()
        If File.Exists(MasterUserInputsFilePath) Then
            File.Copy(MasterUserInputsFilePath, WorkOrderUserInputsFilepath)
        End If
    End Sub

    Private Function WorkOrderInputsFolderExists() As Boolean
        'If Not Directory.Exists(DestinationFilePath + "\Exceloutput_" + WorkOrder) Then
        '    Directory.CreateDirectory(DestinationFilePath + "\Exceloutput_" + WorkOrder)
        'End If

        Return True
    End Function


    Public Function SaveFolder() As Boolean
        Dim dirBrowser As New FolderBrowserDialog
        SaveFolder = False
        dirBrowser.ShowDialog()
        filePath = dirBrowser.SelectedPath
        'If Not filePath.Equals("") Then
        '    If fso.FolderExists(filePath & "\" & WorkOrder) = True Then
        '        fso.DeleteFolder(filePath & "\" & WorkOrder, True)
        '        fso.CreateFolder(filePath & "\" & WorkOrder)
        '        'fso.DeleteFolder(filePath & "\" & WorkOrder & "\" & "InputImages", True)
        '        fso.CreateFolder(filePath & "\" & WorkOrder & "\" & "InputImages")
        '        'fso.DeleteFolder(filePath & "\" & WorkOrder & "\" & "DRAWINGS", True)
        '        fso.CreateFolder(filePath & "\" & WorkOrder & "\" & "DRAWINGS")
        '    Else
        '        fso.CreateFolder(filePath & "\" & WorkOrder)
        '        fso.CreateFolder(filePath & "\" & WorkOrder & "\" & "InputImages")
        '        fso.CreateFolder(filePath & "\" & WorkOrder & "\" & "DRAWINGS")
        '    End If
        '    DestinationFilePath = filePath & "\" & WorkOrder '& "\Models"
        '    libPath = Execution_Path + "\Models"
        '    fso.CopyFolder(libPath, DestinationFilePath, True)
        '    SaveFolder = True
        'End If
    End Function
    Public Function SaveImagesFolder() As Boolean

        ' DestinationFilePath = DestinationFilePath + "\" + WorkOrder + "\InputImages"
        'libPath = Execution_Path + "\InputImages"
        'fso.CopyFolder(libPath, DestinationFilePath + "\InputImages", True)

    End Function

    'New logics for folder deletion on 8-07-09.....................
    Public Function DeleteFoldersBasedonD4() As Boolean

        'If FunctionalClassObject.getValueGridValues("D4", arListInputs) < 1300 Then  'single band
        '    '\MODELS\MORETHAN_1300WDG
        '    '\MODELS\CORE_FRAME_ASSEMBLY\MORETHAN_1300WDG
        '    If Directory.Exists(DestinationFilePath & "\MORETHAN_1300WDG") Then
        '        Directory.Delete(DestinationFilePath & "\MORETHAN_1300WDG")
        '    End If

        '    If Directory.Exists(DestinationFilePath & "\CORE_FRAME_ASSEMBLY\MORETHAN_1300WDG") Then
        '        Directory.Delete(DestinationFilePath & "\CORE_FRAME_ASSEMBLY\MORETHAN_1300WDG")
        '    End If


        'Else
        '    'Two Band
        '    '\MODELS\UPTO_1300WDG
        '    '\MODELS\CORE_FRAME_ASSEMBLY\UPTO_1300WDG

        '    If Directory.Exists(DestinationFilePath & "\UPTO_1300WDG") Then
        '        Directory.Delete(DestinationFilePath & "\UPTO_1300WDG")
        '    End If

        '    If Directory.Exists(DestinationFilePath & "\CORE_FRAME_ASSEMBLY\UPTO_1300WDG") Then
        '        Directory.Delete(DestinationFilePath & "\CORE_FRAME_ASSEMBLY\UPTO_1300WDG")
        '    End If


        'End If

    End Function

    'newly implemented on 8-07-09 for new folder structure
    'Private Sub RearrangeFolderStructureAfterModelGeneration()
    '    'create Models folder
    '    If Directory.Exists(DestinationFilePath & "\Models") Then
    '        Directory.Delete(DestinationFilePath & "\Models")
    '    End If



    'End Sub
    '.........for Testing .........Start......................
    'Public Sub RearrangeFolderStructureAfterModelGeneration()
    '    'create Models folder
    '    If Not Directory.Exists(DestinationFilePath & "\Models") Then
    '        Directory.CreateDirectory(DestinationFilePath & "\Models")
    '    End If

    '    'copy 6 folders to models folder
    '    If Directory.Exists(DestinationFilePath & "\UPTO_1300WDG") Then
    '        Directory.Move(DestinationFilePath & "\UPTO_1300WDG", DestinationFilePath & "\Models\UPTO_1300WDG")
    '        ' Directory.Delete(DestinationFilePath & "\UPTO_1300WDG")
    '    End If

    '    If Directory.Exists(DestinationFilePath & "\MORETHAN_1300WDG") Then
    '        Directory.Move(DestinationFilePath & "\MORETHAN_1300WDG", DestinationFilePath & "\Models\MORETHAN_1300WDG")
    '    End If

    '    If Directory.Exists(DestinationFilePath & "\CORE_FRAME_ASSEMBLY") Then
    '        Directory.Move(DestinationFilePath & "\CORE_FRAME_ASSEMBLY", DestinationFilePath & "\Models\CORE_FRAME_ASSEMBLY")
    '    End If

    '    If Directory.Exists(DestinationFilePath & "\Exceloutput_4878") Then
    '        Directory.Move(DestinationFilePath & "\Exceloutput_4878", DestinationFilePath & "\Models\Exceloutput_4878")
    '        ' Directory.Delete(DestinationFilePath & "\UPTO_1300WDG")
    '    End If

    '    If Directory.Exists(DestinationFilePath & "\InputImages") Then
    '        Directory.Move(DestinationFilePath & "\InputImages", DestinationFilePath & "\Models\InputImages")
    '    End If

    '    If Directory.Exists(DestinationFilePath & "\DRAWINGS") Then
    '        Directory.Move(DestinationFilePath & "\DRAWINGS", DestinationFilePath & "\Models\DRAWINGS")
    '    End If

    'End Sub
    '.........for Testing .........End......................


    Public Sub captureImages(ByVal formName As Form)
        oclsimgcapture.CaptureScreen()
        Dim pic As PictureBox
        Try
            pic = New PictureBox
            pic.Image = oclsimgcapture.Background
            pic.Name = formName.Name
            Dim alist As New ArrayList
            If alist.Contains(formName.Name) Then
                alist.Remove(formName.Name)
            End If
            alist.Add(New Object(1) {formName.Name, pic.Name})
            'If fso.FolderExists(Application.StartupPath + "\InputImages\") = False Then
            '    fso.CreateFolder(Application.StartupPath + "\InputImages\")
            'End If
            'pic.Image.Save(Application.StartupPath + "\InputImages\" + formName.Name + ".jpg", ImageFormat.Jpeg)
        Catch oException As Exception
            MessageBox.Show(oException.Message)
        End Try
    End Sub
    'Public Sub DataSaveClickEvent(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim aData As Byte() = GetDataToSave(CType(sender, Button))
    '    Dim strGUID As String = getObjectBBLDataClass.GetProjectGUIDValue()
    '    getObjectBBLDataClass.SetAndSaveDataToDataClass(aData, strGUID)
    'End Sub

    Private ReadOnly Property GetWorkOrderInfoForm() As Object
        Get
            '    Return New Object(0) {frmContractDetails}
        End Get
    End Property

    Private ReadOnly Property GetInputsInfoForm() As Object
        Get
            'Return New Object(1) {frmWorkOrder, frmInputs}
        End Get
    End Property

    Private ReadOnly Property GetPermaWoodInfoForm() As Object
        Get
            'Return New Object(2) {frmWorkOrder, frmInputs, frmDeckPlatePermawoodDetails}
        End Get
    End Property

    Private ReadOnly Property GetCoreDetailsInfoForm() As Object
        Get
            'Return New Object(3) {frmWorkOrder, frmInputs, frmDeckPlatePermawoodDetails, frmCoreDetails}
        End Get
    End Property

    Private ReadOnly Property GetDuctandDrawingInfoForm() As Object
        Get
            'Return New Object(4) {frmWorkOrder, frmInputs, frmDeckPlatePermawoodDetails, frmCoreDetails, frmDuctandDrawingOutput}
        End Get
    End Property

    Private ReadOnly Property GetSaveButtonForms() As ArrayList
        Get
            Dim aReturnData As New ArrayList
            aReturnData.Add(New Object(1) {"btnNext", GetDuctandDrawingInfoForm()})
            'aReturnData.Add(New Object(1) {"btnPriliminaryDataSave", GetInputsInfoForm()})
            'aReturnData.Add(New Object(1) {"btnHeadSectionSave", GetPermaWoodInfoForm()})
            'aReturnData.Add(New Object(1) {"btnTipSectionSave", GetCoreDetailsInfoForm()})
            'aReturnData.Add(New Object(1) {"btnIntermediateSection1Save", GetDuctandDrawingInfoForm()})
            Return aReturnData
        End Get
    End Property
#Region "Functions"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="oSaveButton"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataToSave(ByVal oSaveButton As Object) As Byte()
        'Dim oObject As Object = GetSaveFormList(CType(oSaveButton, Button), GetSaveButtonForms)
        'Dim oGetSetUIClass As New IFLGetSetUI.IFLGetSetUIClass
        'Dim oDataSet As New DataSet("BBLSaveData")
        'For Each oForm As Form In oObject
        '    Dim oTable As DataTable = oGetSetUIClass.StoreFormData(oForm)
        '    oDataSet.Tables.Add(oTable)
        'Next
        'oDataSet.WriteXml(Execution_Path + "\BBL.xml")
        'GetDataToSave = GetByteArray(oDataSet)
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="oSaveButton"></param>
    ''' <param name="aSaveFormList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSaveFormList(ByVal oSaveButton As Button, ByVal aSaveFormList As ArrayList) As Object
        Dim strButtonName As String = oSaveButton.Name
        Dim oData As Object
        Dim oReturnData As Object = Nothing

        For Each oData In aSaveFormList
            If strButtonName.ToUpper.Equals(oData(0).ToUpper) Then
                oReturnData = oData(1)
            End If
        Next
        Return oReturnData
    End Function

    ''' <summary>
    ''' Gets the byte array data.
    ''' </summary>
    ''' <param name="oDataSet"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetByteArray(ByVal oDataSet As DataSet) As Byte()
        'oDataSet.WriteXml(Execution_Path + "\BBL.xml")
        'Dim fsBLOBFile As New System.IO.FileStream(Execution_Path + "\BBL.xml", IO.FileMode.Open)
        'Dim bytBLOBData(fsBLOBFile.Length() - 1) As Byte
        'fsBLOBFile.Read(bytBLOBData, 0, bytBLOBData.Length)
        'fsBLOBFile.Close()
        'Return bytBLOBData
    End Function


#End Region
#Region "Enums"
    Private Enum Parameters
        Name
        Value
    End Enum
#End Region

End Module

