Imports IFLBaseDataLayer
Imports IFLCommonLayer
Imports Microsoft.Win32.Registry
Imports Microsoft.Win32.RegistryKey
Imports System.Diagnostics.Process
Imports Microsoft.Office.Interop
Imports Microsoft.Win32
Imports System.IO
Imports MonarchFunctionalLayer
Public Class clsCosting

#Region "Variables"

    Private _oCodeNumber_BeforeAscendimg As DataTable

    Private _oCodeNumber_AfterAscendimg As DataTable

    Private _oCostDetails As DataTable

    Private _oTable As DataTable

    'Private _oCostingConnectionObject As IFLConnectionClass

    Private _strCurrentWorkingDirectory As String = System.Environment.CurrentDirectory

    Private _strMotherExcelPath As String = _strCurrentWorkingDirectory + "\Costing_Master.xls"

    Private _strchildExcelPath As String '= _strCurrentWorkingDirectory + "\Reports\Costing.xls"

    Private _oExApplication As Excel.Application

    Private _oExWorkbook As Excel.Workbook

    Private _oExcelSheet_MainAssembly As Excel.Worksheet

    Private _oExcelSheet_NewTube As Excel.Worksheet

    Private _oExcelSheet_NewRod As Excel.Worksheet

    Private _intTotalCostExcelRange As Integer

    'Private _strMainAssemblyName As String

    Private _strNewTubeName As String

    Private _strNewRodName As String

    Private _blnIsNewTube As Boolean

    Private _blnIsNewRod As Boolean

    Private _strZeroCostParts As String = ""

    Private _strWC1Number_MainAssembly As String = ""

    Private _strWC2Number_MainAssembly As String = ""

#End Region

#Region "Property"

    'For storing Code Numbers before Sorting
    Public Property CodeNumber_BeforeAscendimg() As DataTable
        Get
            Return _oCodeNumber_BeforeAscendimg
        End Get
        Set(ByVal value As DataTable)
            _oCodeNumber_BeforeAscendimg = value
        End Set
    End Property

    'For storing Code Numbers after Sorting
    Public Property CodeNumber_AfterAscending() As DataTable
        Get
            Return _oCodeNumber_AfterAscendimg
        End Get
        Set(ByVal value As DataTable)
            _oCodeNumber_AfterAscendimg = value
        End Set
    End Property

    'For storing Cost Details
    Public Property CostDetails() As DataTable
        Get
            Return _oCostDetails
        End Get
        Set(ByVal value As DataTable)
            _oCostDetails = value
        End Set
    End Property

#End Region

#Region "Enum"

    Public Enum CodeNumberItemOrder
        CodeNumber = 0
        PartName = 1
        IsExisting_New = 2
        Quantity = 3
        Units = 4
        Comment = 5
    End Enum

    Public Enum CostDetailsItemOrder
        CodeNumber = 0
        Description = 1
        Cost = 2
        PartName = 3
        IsExisting_New = 4
        Quantity = 5
        Units = 6
        Comment = 7
    End Enum

    Public Enum PaintDetails
        PaintColor = 0
        PaintCode = 1
    End Enum

    Public Enum PackageDetails
        CodeNumber = 0
        Description = 1
    End Enum

    Public Enum LabelDetails
        CodeNumber = 0
        Description = 1
    End Enum

#End Region

#Region "Functions"

    Public Function AddCodeNumberToDataTable(Optional ByVal strCodeNumber As String = "", Optional ByVal strPartName As String = "", _
                                                                        Optional ByVal dblQuanity As Double = 0, Optional ByVal strUnits As String = "EA", Optional ByVal strComment As String = "") As Boolean
        AddCodeNumberToDataTable = False
        If Not IsNothing(strCodeNumber) AndAlso strCodeNumber <> "" Then
            Try
                Dim strISExisting_New As String = ""
                If strCodeNumber.StartsWith(7) Then
                    strISExisting_New = "New"
                Else
                    strISExisting_New = "Existing"
                End If

                If IsNothing(CodeNumber_BeforeAscendimg) Then
                    CodeNumber_BeforeAscendimg = New DataTable
                    CodeNumber_BeforeAscendimg.Columns.Add("CodeNumber")
                    CodeNumber_BeforeAscendimg.Columns.Add("PartName")
                    CodeNumber_BeforeAscendimg.Columns.Add("IsExisting_New")
                    CodeNumber_BeforeAscendimg.Columns.Add("Quantity")
                    CodeNumber_BeforeAscendimg.Columns.Add("Units")
                    CodeNumber_BeforeAscendimg.Columns.Add("Comment")
                    '06_09_2012   RAGAVA
                    Try
                        Dim oDataRow1 As DataRow = CodeNumber_BeforeAscendimg.NewRow
                        oDataRow1.Item("CodeNumber") = "174040"
                        oDataRow1.Item("PartName") = "MASK INSPECT FOR EXTENDED ROD"
                        oDataRow1.Item("IsExisting_New") = "Existing"
                        oDataRow1.Item("Quantity") = "1"
                        oDataRow1.Item("Units") = "EA"
                        oDataRow1.Item("Comment") = ""
                        CodeNumber_BeforeAscendimg.Rows.Add(oDataRow1)
                        If ofrmTieRod1.cmbClevisCapPinHole.Text = "Bushing" OrElse ofrmTieRod1.cmbRodClevisPinHole.Text = "Bushing" Then
                            oDataRow1 = CodeNumber_BeforeAscendimg.NewRow
                            oDataRow1.Item("CodeNumber") = "174043"
                            oDataRow1.Item("PartName") = "MASK CAP SC 355-24"
                            oDataRow1.Item("IsExisting_New") = "Existing"
                            Dim iBushing As Integer = 0
                            If ofrmTieRod1.cmbClevisCapPinHole.Text = "Bushing" Then
                                iBushing = +1
                            End If
                            If ofrmTieRod1.cmbRodClevisPinHole.Text = "Bushing" Then
                                iBushing = +1
                            End If
                            oDataRow1.Item("Quantity") = iBushing
                            oDataRow1.Item("Units") = "EA"
                            oDataRow1.Item("Comment") = ""
                            CodeNumber_BeforeAscendimg.Rows.Add(oDataRow1)
                        End If
                    Catch ex As Exception

                    End Try
               
                End If

                If CodeNumber_BeforeAscendimg.Rows.Count > 0 Then
                    For Each oDataRow_Duplicates As DataRow In CodeNumber_BeforeAscendimg.Rows
                        If oDataRow_Duplicates(CodeNumberItemOrder.CodeNumber) = strCodeNumber Then
                            oDataRow_Duplicates.Item(CodeNumberItemOrder.PartName) = strPartName
                            oDataRow_Duplicates.Item(CodeNumberItemOrder.IsExisting_New) = strISExisting_New
                            Exit Function
                        End If
                    Next
                End If

                Dim oDataRow As DataRow = CodeNumber_BeforeAscendimg.NewRow
                oDataRow.Item(CodeNumberItemOrder.CodeNumber) = strCodeNumber
                oDataRow.Item(CodeNumberItemOrder.PartName) = strPartName
                oDataRow.Item(CodeNumberItemOrder.IsExisting_New) = strISExisting_New

                'Sandeep 18-03-10 12:30pm
                If dblQuanity <> 0 Then
                    oDataRow.Item(CodeNumberItemOrder.Quantity) = dblQuanity
                Else
                    oDataRow.Item(CodeNumberItemOrder.Quantity) = GetQuantityDetails(oDataRow)
                End If

                oDataRow.Item(CodeNumberItemOrder.Units) = strUnits
                '*****************************
                'TODO:Sunny 08-04-10 10am
                oDataRow.Item(CodeNumberItemOrder.Comment) = strComment

                CodeNumber_BeforeAscendimg.Rows.Add(oDataRow)
                AddCodeNumberToDataTable = True
            Catch oException As Exception
                AddCodeNumberToDataTable = False
                MessageBox.Show("Unable to add CodeNumber to DataTable" + oException.ToString, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Function

    ' Added by sandeep on17-03-10 2pm
    Private Function GetQuantityDetails(ByVal oDataRow As DataRow) As Double
        GetQuantityDetails = 1
        Try
            If oDataRow(CodeNumberItemOrder.PartName).Equals("Tie Rod Code Number") Then
                GetQuantityDetails = 4
            ElseIf oDataRow(CodeNumberItemOrder.PartName).Equals("Tie Rod Nut Code") Then
                If SeriesForCosting = "TX (TXC)" Then
                    GetQuantityDetails = 4
                Else
                    GetQuantityDetails = 8
                End If
            Else
                GetQuantityDetails = 1
            End If
        Catch ex As Exception
            GetQuantityDetails = 1
        End Try
    End Function

    ' Added by sandeep on 18-03-10 10am
    Public Function GetPaint_Package_LabelDetails() As Boolean
        GetPaint_Package_LabelDetails = True

        Dim oPaintDataFromDB As DataRow = Nothing
        Try
            Dim strPaint As String = ofrmTieRod2.cmbPaint.Text
            If Not IsNothing(strPaint) Then
                oPaintDataFromDB = IFLConnectionObject.GetDataRow("Select PaintColor,PaintCode from PaintDetails where PaintColor = '" + strPaint + "'")
                If Not IsNothing(oPaintDataFromDB) Then
                    If oPaintDataFromDB.ItemArray.Length > 0 Then
                        AddCodeNumberToDataTable(oPaintDataFromDB(PaintDetails.PaintCode), oPaintDataFromDB(PaintDetails.PaintColor), 0.1, "LT", "Paint")

                    End If
                End If
            End If
        Catch ex As Exception
            oPaintDataFromDB = Nothing
            GetPaint_Package_LabelDetails = False
        End Try

        'commenting Sugandhi_20120601_Start
        Dim oPackingDataFromDB As DataRow = Nothing
        Try
            If ofrmTieRod3.rbYesBagRequired.Checked = True Then
                Dim strBoreDiameter As String = ofrmTieRod1.cmbBore.Text
                Dim strquery_ As String = "select BagPartNumber  from  MIL_WELDED.dbo.BagChart where BoreDia = '" & BoreDiameter & "' and RetractedLength = " & Val(ofrmTieRod1.txtRetractedLength.Text)
                oPackingDataFromDB = IFLConnectionObject.GetDataRow(strquery_)
                If Not IsNothing(oPackingDataFromDB) Then
                    If oPackingDataFromDB.ItemArray.Length > 0 Then
                        AddCodeNumberToDataTable(oPackingDataFromDB("BagPartNumber"), "PLASTIC BAG", , , "")
                    End If
                End If
            End If
            '06_09_2012   RAGAVA
            If ofrmTieRod3.ChkFluidFilmInternal.Checked = True Then
                AddCodeNumberToDataTable("239811", "Fluid Film Internal Caps And Piston", "1", "EA", "")
            End If
            'If UCase(RodMaterialForCosting).IndexOf("LION") <> -1 Then                                 '17_09_2014 Neeraja Start
            '    AddCodeNumberToDataTable("232990", "DECAL LION 1000 CHROME", "1", "EA", "")           'Neeraja commmentaed 
            'End If                                                                                     '17_09_2014 Neeraja End
            'Till  Here
            '    If ofrmTieRod3.chkPackCylinderInPlasticBag.Checked = True Then     '16_06_2011   RAGAVA
            '        Dim strBoreDiameter As String = ofrmTieRod1.cmbBore.Text
            '        oPackingDataFromDB = IFLConnectionObject.GetDataRow("Select CodeNumber, Description from PackagingDetails where " + strBoreDiameter + "  between BoreDia_Min and BoreDia_Max")
            '        If Not IsNothing(oPackingDataFromDB) Then
            '            If oPackingDataFromDB.ItemArray.Length > 0 Then
            '                AddCodeNumberToDataTable(oPackingDataFromDB(PackageDetails.CodeNumber), oPackingDataFromDB(PackageDetails.Description), , , "Package")
            '            End If
            '        End If
            '    End If
        Catch ex As Exception

        End Try
        'commenting Sugandhi_20120601_end

        Dim oLableDataFromDB As DataRow = Nothing
        Try
            Dim strSeriesType As String = ""
            If SeriesForCosting = "TX (TXC)" Then
                strSeriesType = "TX"
            ElseIf SeriesForCosting = "TL (TC)" Then
                strSeriesType = "TL"
            ElseIf SeriesForCosting = "TH (TD)" Then
                strSeriesType = "TH/TP"
            ElseIf SeriesForCosting = "TP-High" Then
                strSeriesType = "TH/TP"
            ElseIf SeriesForCosting = "TP-Low" Then
                strSeriesType = "TH/TP"
            ElseIf SeriesForCosting = "LN" Then          '21_01_2011          RAGAVA
                strSeriesType = "LN"                               '21_01_2011          RAGAVA
            End If
            oLableDataFromDB = IFLConnectionObject.GetDataRow("Select CodeNumber, Description from LableDetails where Bore= '" + strSeriesType + "'")
            If Not IsNothing(oLableDataFromDB) Then
                If oLableDataFromDB.ItemArray.Length > 0 Then
                    AddCodeNumberToDataTable(oLableDataFromDB(LabelDetails.CodeNumber), oLableDataFromDB(LabelDetails.Description))
                End If
            End If
        Catch ex As Exception

        End Try

    End Function

    Public Function Costingfunctionality() As Boolean
        Costingfunctionality = False
        _strchildExcelPath = _strCurrentWorkingDirectory + "\Reports\" + CylinderCodeNumber + "_Costing.xls"
        Try
            If CreateExcelObjects() Then
                If PerformAscendingOrder_CodeNumbers() Then
                    If GetCostDetails_CodeNumbers() Then
                        If DropDataToExcel() Then
                            If _strZeroCostParts <> "" Then
                                'Dim strMessage As String = "Cost details not available for parts " + _strZeroCostParts                                               '07_03_2011  RAGAVA
                                'MessageBox.Show(strMessage, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2)        '07_03_2011  RAGAVA
                            End If
                            Costingfunctionality = True
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Costingfunctionality = False
        End Try
    End Function

    Private Sub ExcelCommonFunctionality()
        Try
            _oExcelSheet_MainAssembly.Range("E2").Value = SRQ
        Catch ex As Exception

        End Try
    End Sub

    Private Function CreateExcelObjects() As Boolean
        CreateExcelObjects = False
        Try
            If CheckForExcel() Then
                If CopyTheMasterFile() Then
                    If CreateExcel() Then
                        ExcelCommonFunctionality()
                        CreateExcelObjects = True
                    End If
                End If
            End If
        Catch ex As Exception
            CreateExcelObjects = False
        End Try
    End Function

    Private Function PerformAscendingOrder_CodeNumbers() As Boolean
        PerformAscendingOrder_CodeNumbers = False
        Try
            Dim oDataView As DataView
            oDataView = CodeNumber_BeforeAscendimg.DefaultView
            oDataView.Sort = "CodeNumber"

            If IsNothing(CodeNumber_AfterAscending) Then
                CodeNumber_AfterAscending = New DataTable
                CodeNumber_AfterAscending.Columns.Add("CodeNumber")
                CodeNumber_AfterAscending.Columns.Add("PartName")
                CodeNumber_AfterAscending.Columns.Add("IsExisting_New")
                CodeNumber_AfterAscending.Columns.Add("Quantity")
                CodeNumber_AfterAscending.Columns.Add("Units")
                CodeNumber_AfterAscending.Columns.Add("Comment")
            End If

            For Each oDataViewItem As DataRowView In oDataView
                Dim oDataRow As DataRow = CodeNumber_AfterAscending.NewRow
                oDataRow.Item(CodeNumberItemOrder.CodeNumber) = oDataViewItem(CodeNumberItemOrder.CodeNumber)
                oDataRow.Item(CodeNumberItemOrder.PartName) = oDataViewItem(CodeNumberItemOrder.PartName)
                oDataRow.Item(CodeNumberItemOrder.IsExisting_New) = oDataViewItem(CodeNumberItemOrder.IsExisting_New)
                oDataRow.Item(CodeNumberItemOrder.Quantity) = oDataViewItem(CodeNumberItemOrder.Quantity)
                oDataRow.Item(CodeNumberItemOrder.Units) = oDataViewItem(CodeNumberItemOrder.Units)
                oDataRow.Item(CodeNumberItemOrder.Comment) = oDataViewItem(CodeNumberItemOrder.Comment)
                CodeNumber_AfterAscending.Rows.Add(oDataRow)
            Next
            PerformAscendingOrder_CodeNumbers = True
        Catch ex As Exception
            PerformAscendingOrder_CodeNumbers = False
        End Try
    End Function

    Private Function GetCostDetails_CodeNumbers() As Boolean
        GetCostDetails_CodeNumbers = False
        Try
            For Each oCodeNumber_AfterAscendingItem As DataRow In CodeNumber_AfterAscending.Rows

                Dim oResultDataRow As DataRow = IFLConnectionObject.GetDataRow("Select * from CostingDetails where PartCode = '" + oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.CodeNumber) + "'")
                Try

                    If IsNothing(CostDetails) Then
                        CostDetails = New DataTable
                        CostDetails.Columns.Add("CodeNumber")
                        CostDetails.Columns.Add("Description")
                        CostDetails.Columns.Add("Cost")
                        CostDetails.Columns.Add("PartName")
                        CostDetails.Columns.Add("IsExisting_New")
                        CostDetails.Columns.Add("Quantity")
                        CostDetails.Columns.Add("Units")
                        CostDetails.Columns.Add("Comment")
                    End If
                    'TubeMaterialCode_Costing

                    Dim oDataRow As DataRow = CostDetails.NewRow
                    If Not IsNothing(oResultDataRow) AndAlso oResultDataRow.ItemArray.Length > 0 Then
                        If Not IsDBNull(oResultDataRow(CostDetailsItemOrder.CodeNumber)) Then
                            oDataRow.Item(CostDetailsItemOrder.CodeNumber) = oResultDataRow(CostDetailsItemOrder.CodeNumber)
                        Else
                            oDataRow.Item(CostDetailsItemOrder.CodeNumber) = ""
                        End If

                        If Not IsDBNull(oResultDataRow(CostDetailsItemOrder.Description)) Then
                            oDataRow.Item(CostDetailsItemOrder.Description) = oResultDataRow(CostDetailsItemOrder.Description)
                        Else
                            oDataRow.Item(CostDetailsItemOrder.Description) = ""
                        End If

                        If Not IsDBNull(oResultDataRow(CostDetailsItemOrder.Cost)) AndAlso Not Val(oResultDataRow(CostDetailsItemOrder.Cost)) = 0 Then
                            oDataRow.Item(CostDetailsItemOrder.Cost) = oResultDataRow(CostDetailsItemOrder.Cost)
                        Else
                            'TODO:Sunny 08-04-10 10am
                            If oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.Comment) = "Paint" Then
                                oDataRow.Item(CostDetailsItemOrder.Cost) = 0.00001
                            ElseIf oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.Comment) = "Package" Then
                                oDataRow.Item(CostDetailsItemOrder.Cost) = 0.00001
                            Else
                                oDataRow.Item(CostDetailsItemOrder.Cost) = 0
                                _strZeroCostParts += oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.PartName) + vbCrLf
                            End If
                            '************************
                        End If

                        'ToDO: Sunny 28-06-10&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& Cost Management
                        Try
                            If Not IsDBNull(oResultDataRow("CostMangementReferenceNumber")) Then
                                Dim dblFactor As Double = IFLConnectionObject.GetValue("Select Cost from CostManagementDetails where IFLID = '" + oResultDataRow("CostMangementReferenceNumber") + "'")
                                If Not oDataRow.Item(CostDetailsItemOrder.Cost) = 0 Then
                                    oDataRow.Item(CostDetailsItemOrder.Cost) = oDataRow.Item(CostDetailsItemOrder.Cost) * dblFactor
                                End If
                            End If
                        Catch ex As Exception

                        End Try
                        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
                    Else

                        oDataRow.Item(CostDetailsItemOrder.CodeNumber) = oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.CodeNumber)

                        If Not IsNothing(oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.PartName)) Then
                            oDataRow.Item(CostDetailsItemOrder.Description) = oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.PartName)
                        Else
                            oDataRow.Item(CostDetailsItemOrder.Description) = "New Part"
                        End If

                        Dim dblNewCost As Double = GetNewPartCost(oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.PartName), oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.CodeNumber), "New Part")
                        If dblNewCost <> 0 Then
                            oDataRow.Item(CostDetailsItemOrder.Cost) = dblNewCost
                        Else
                            'TODO:Sunny 08-04-10 10am
                            If oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.Comment) = "Paint" Then
                                oDataRow.Item(CostDetailsItemOrder.Cost) = 0.00001
                            Else
                                oDataRow.Item(CostDetailsItemOrder.Cost) = 0
                                _strZeroCostParts += oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.PartName) + vbCrLf
                            End If
                            '************************                          
                        End If

                        'ToDO: Sunny 28-06-10&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& Cost Management
                        Dim strNewPartCode As String = ""
                        If oDataRow.Item(CostDetailsItemOrder.Description).ToString.Contains("Tube") Then
                            strNewPartCode = TubeMaterialCode_Costing
                        ElseIf oDataRow.Item(CostDetailsItemOrder.Description).ToString.Contains("Rod") Then
                            strNewPartCode = RodMaterialCode_Costing
                        End If
                        Dim oCostManagementDataRow As DataRow = Nothing
                        Try
                            oCostManagementDataRow = IFLConnectionObject.GetDataRow("Select * from CostingDetails where PartCode = '" + strNewPartCode + "'")
                        Catch ex As Exception
                        End Try

                        Try
                            If Not IsNothing(oCostManagementDataRow) AndAlso oCostManagementDataRow.ItemArray.Length > 0 Then
                                If Not IsDBNull(oCostManagementDataRow("CostMangementReferenceNumber")) Then
                                    Dim dblFactor As Double = IFLConnectionObject.GetValue("Select Cost from CostManagementDetails where IFLID = '" + oCostManagementDataRow("CostMangementReferenceNumber") + "'")
                                    If Not oDataRow.Item(CostDetailsItemOrder.Cost) = 0 Then
                                        oDataRow.Item(CostDetailsItemOrder.Cost) = oDataRow.Item(CostDetailsItemOrder.Cost) * dblFactor
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
                    End If

                    oDataRow.Item(CostDetailsItemOrder.PartName) = oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.PartName)
                    oDataRow.Item(CostDetailsItemOrder.IsExisting_New) = oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.IsExisting_New)
                    oDataRow.Item(CostDetailsItemOrder.Quantity) = oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.Quantity)
                    oDataRow.Item(CostDetailsItemOrder.Units) = oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.Units)
                    oDataRow.Item(CostDetailsItemOrder.Comment) = oCodeNumber_AfterAscendingItem(CodeNumberItemOrder.Comment)
                    CostDetails.Rows.Add(oDataRow)
                Catch ex As Exception

                End Try
            Next
            GetCostDetails_CodeNumbers = True
        Catch ex As Exception
            GetCostDetails_CodeNumbers = False
        End Try
    End Function

    Private Function GetNewPartCost(ByVal strPartName As String, ByVal strPartCode As String, ByVal strDescription As String) As Double
        GetNewPartCost = 0
        Try
            Dim dblTubeLength As Double = Math.Ceiling(TubeLength)
            '

            If strPartName = "Tube Code Number" Then

                'TODO:Sunny 20-04-10
                Dim oResultDataRow As DataRow = IFLConnectionObject.GetDataRow("Select * from CostingDetails where PartCode = '" + TubeMaterialCode_Costing + "'")
                If Not IsNothing(oResultDataRow) Then
                    If Not IsNothing(TubeMaterialCode_Costing) Then
                        _oExcelSheet_NewTube.Range("A6").Value = 1

                        Dim strChangedPartCode As String = GetPurchasedCode(TubeMaterialCode_Costing)
                        If Not IsNothing(strChangedPartCode) Then
                            _oExcelSheet_NewTube.Range("B6").Value = strChangedPartCode
                        Else
                            _oExcelSheet_NewTube.Range("B6").Value = TubeMaterialCode_Costing
                        End If
                    End If

                    If Not IsDBNull(oResultDataRow("Description")) Then
                        _oExcelSheet_NewTube.Range("C6").Value = oResultDataRow("Description")
                        TubeMaterial1 = oResultDataRow("Description")       '08_10_2010   RAGAVA
                    End If

                    Dim dblQuantityPerTubePart_TubeCode As Double = (dblTubeLength + (3 / 8)) / 12
                    If Not dblQuantityPerTubePart_TubeCode = 0 Then
                        _oExcelSheet_NewTube.Range("E6").Value = dblQuantityPerTubePart_TubeCode
                    End If

                    _oExcelSheet_NewTube.Range("F6").Value = "FT"

                    If Not IsDBNull(oResultDataRow("Cost")) Then
                        _oExcelSheet_NewTube.Range("J6").Value = oResultDataRow("Cost")
                    Else
                        _oExcelSheet_NewTube.Range("J6").Value = 0
                        _oExcelSheet_NewTube.Range("J6").Font.Color = RGB(255, 0, 0)
                        MessageBox.Show("Tube material cost per foot is not available")
                    End If
                End If
                '********************

                'Sunny 21-04-10 3pm
                If SeriesForCosting.Contains("TP") Then
                    Dim dblCost As Double = 0
                    Dim dblQuantity As Double = 0
                    _oExcelSheet_NewTube.Range("A7").Value = 2
                    _oExcelSheet_NewTube.Range("B7").Value = 469832
                    _oExcelSheet_NewTube.Range("F7").Value = "EA"

                    If strRephasing.Contains("Both") Then
                        _oExcelSheet_NewTube.Range("E7").Value = 2
                    Else
                        _oExcelSheet_NewTube.Range("E7").Value = 1
                    End If
                    dblQuantity = _oExcelSheet_NewTube.Range("E7").Value

                    Dim oTPDetails As DataRow = IFLConnectionObject.GetDataRow("Select * from CostingDetails where PartCode = 469832")
                    If Not IsNothing(oTPDetails) Then
                        If Not IsDBNull(oTPDetails("Description")) Then
                            _oExcelSheet_NewTube.Range("C7").Value = oTPDetails("Description")
                            TubeMaterial2 = oTPDetails("Description")       '08_10_2010   RAGAVA
                        End If
                        If Not IsDBNull(oTPDetails("Cost")) Then
                            _oExcelSheet_NewTube.Range("J7").Value = oTPDetails("Cost")
                            dblCost = oTPDetails("Cost")
                        End If
                    End If
                    _oExcelSheet_NewTube.Range("O7").Value = dblCost * dblQuantity
                End If

                GetTPDetails()
                _oExWorkbook.Save()


                Dim strBoreDiameterColumn_TubeCode As String = ""
                Dim dblBoreDiamteter_TubeCode As Double = Val(ofrmTieRod1.cmbBore.Text)
                If dblBoreDiamteter_TubeCode = 2 Then
                    strBoreDiameterColumn_TubeCode = "BoreDiameter_2"
                ElseIf dblBoreDiamteter_TubeCode = 2.5 Then
                    strBoreDiameterColumn_TubeCode = "BoreDiameter_2_5"
                ElseIf dblBoreDiamteter_TubeCode = 3 Then
                    strBoreDiameterColumn_TubeCode = "BoreDiameter_3"
                ElseIf dblBoreDiamteter_TubeCode = 3.5 Then
                    strBoreDiameterColumn_TubeCode = "BoreDiameter_3_5"
                ElseIf dblBoreDiamteter_TubeCode = 4 Then
                    strBoreDiameterColumn_TubeCode = "BoreDiameter_4"
                ElseIf dblBoreDiamteter_TubeCode = 4.5 Then
                    strBoreDiameterColumn_TubeCode = "BoreDiameter_4_5"
                ElseIf dblBoreDiamteter_TubeCode = 5 Then
                    strBoreDiameterColumn_TubeCode = "BoreDiameter_5"

                    'ANUP 14-09-2010 START
                ElseIf dblBoreDiamteter_TubeCode = 2.25 Then
                    strBoreDiameterColumn_TubeCode = "BoreDiameter_2_25"
                ElseIf dblBoreDiamteter_TubeCode = 2.75 Then
                    strBoreDiameterColumn_TubeCode = "BoreDiameter_2_75"
                ElseIf dblBoreDiamteter_TubeCode = 3.25 Then
                    strBoreDiameterColumn_TubeCode = "BoreDiameter_3_25"
                ElseIf dblBoreDiamteter_TubeCode = 3.75 Then
                    strBoreDiameterColumn_TubeCode = "BoreDiameter_3_75"
                ElseIf dblBoreDiamteter_TubeCode = 4.25 Then
                    strBoreDiameterColumn_TubeCode = "BoreDiameter_4_25"
                ElseIf dblBoreDiamteter_TubeCode = 4.75 Then
                    strBoreDiameterColumn_TubeCode = "BoreDiameter_4_75"
                End If

                If Not strBoreDiameterColumn_TubeCode = "" Then
                    Try
                        Dim strQuery1 As String = "Select " + strBoreDiameterColumn_TubeCode + " from TubeCutDetails where TubeLength =" + dblTubeLength.ToString
                        Dim dblWC099_RunStandard As Double
                        Try
                            dblWC099_RunStandard = IFLConnectionObject.GetValue(strQuery1)
                        Catch ex As Exception
                            dblWC099_RunStandard = 0
                        End Try

                        Dim strQuery2 As String = "Select " + strBoreDiameterColumn_TubeCode + " from TubeSkiveDetails where TubeLength =" + dblTubeLength.ToString
                        Dim dblWC087_RunStandard As Double
                        Try
                            dblWC087_RunStandard = IFLConnectionObject.GetValue(strQuery2)
                        Catch ex As Exception
                            dblWC087_RunStandard = 0
                        End Try


                        If Not dblWC099_RunStandard = 0 Then
                            _oExcelSheet_NewTube.Range("F14").Value = dblWC099_RunStandard
                        End If

                        If Not dblWC087_RunStandard = 0 Then
                            _oExcelSheet_NewTube.Range("F16").Value = dblWC087_RunStandard
                        End If

                        GetBurdenCost()

                        _oExWorkbook.Save()

                        GetNewPartCost = _oExcelSheet_NewTube.Range("L23").Value
                    Catch ex As Exception

                    End Try
                Else
                    _oExcelSheet_NewTube.Range("F14").Interior.Color = RGB(255, 0, 0)
                    _oExcelSheet_NewTube.Range("F16").Interior.Color = RGB(255, 0, 0)
                    _oExcelSheet_NewTube.Range("L23").Interior.Color = RGB(255, 0, 0)
                    _oExcelSheet_NewTube.Range("L23").Value = 0
                    _oExWorkbook.Save()
                    GetNewPartCost = 0
                    ' MessageBox.Show("Standard Run Cost is not avilable for BoreDiamteter" + dblBoreDiamteter_TubeCode.ToString, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If

            ElseIf strPartName = "Rod Code Number" Then

                Dim dblMaterialCostPerPound As Double = 0
                Dim dblRodWeightPerFoot As Double = GetRodWeight()
                Dim dblQuantityPerRodPart As Double = (RodLength + 0.25) * (dblRodWeightPerFoot / 12)

                Dim strRodDiameterColumn As String = ""
                Dim dblRodDiamteter As Double = RodDiameter

                Dim oResultDataRow As DataRow = IFLConnectionObject.GetDataRow("Select * from CostingDetails where PartCode = '" + RodMaterialCode_Costing + "'")
                If Not IsNothing(oResultDataRow) Then
                    If Not IsNothing(RodMaterialCode_Costing) Then
                        Dim strChangedPartCode As String = GetPurchasedCode(RodMaterialCode_Costing)
                        If Not IsNothing(strChangedPartCode) Then
                            _oExcelSheet_NewRod.Range("B6").Value = strChangedPartCode
                        Else
                            _oExcelSheet_NewRod.Range("B6").Value = RodMaterialCode_Costing
                        End If
                    End If

                    If Not IsDBNull(oResultDataRow("Description")) Then
                        _oExcelSheet_NewRod.Range("C6").Value = oResultDataRow("Description")
                    End If

                    If Not IsDBNull(oResultDataRow("Cost")) Then
                        _oExcelSheet_NewRod.Range("J6").Value = oResultDataRow("Cost")
                    Else
                        _oExcelSheet_NewRod.Range("J6").Value = 0
                        _oExcelSheet_NewRod.Range("J6").Font.Color = RGB(255, 0, 0)
                        MessageBox.Show("Rod material cost per foot is not available")
                    End If
                End If

                If Not dblQuantityPerRodPart = 0 Then
                    _oExcelSheet_NewRod.Range("E6").Value = dblQuantityPerRodPart
                    'ANUP 14-09-2010 START
                    CType(_oExcelSheet_NewRod.Rows(6, Type.Missing), Excel.Range).Font.Color = RGB(0, 0, 0)
                Else
                    _oExcelSheet_NewRod.Range("E6").Value = 0
                    CType(_oExcelSheet_NewRod.Rows(6, Type.Missing), Excel.Range).Font.Color = RGB(255, 0, 0)
                    'ANUP 14-09-2010 TILL HERE
                End If

                _oExWorkbook.Save()

                If dblRodDiamteter = 1.12 Then
                    strRodDiameterColumn = "BoreDiameter_1_12"
                ElseIf dblRodDiamteter = 1.25 Then
                    strRodDiameterColumn = "BoreDiameter_1_25"
                ElseIf dblRodDiamteter = 1.38 Then
                    strRodDiameterColumn = "BoreDiameter_1_38"
                ElseIf dblRodDiamteter = 1.5 Then
                    strRodDiameterColumn = "BoreDiameter_1_5"
                ElseIf dblRodDiamteter = 1.75 Then
                    strRodDiameterColumn = "BoreDiameter_1_75"
                ElseIf dblRodDiamteter = 2 Then
                    strRodDiameterColumn = "BoreDiameter_2"
                End If

                'ANUP 15-12-2010 START
                Dim oCostingAndCMSCommon As New clsCostingAnsCMSCommon
                'ANUP 15-12-2010 TILL HERE

                If Not strRodDiameterColumn = "" Then
                    Try
                        Dim dblWC083_RunStandard As Double

                        Dim strWCNumber As String = ""
                        Dim dblWCNumberValue As Double
                        Dim dblCostRodLength As Double = Math.Ceiling(RodLength)
                        If RodMaterialForCosting = "Chrome" OrElse UCase(RodMaterialForCosting).IndexOf("LION") <> -1 Then 'anup 13-09-2010
                            Dim strQuery1 As String = "Select " + strRodDiameterColumn + " from TRChromeRodCuttingDetails where TubeLength =" + dblCostRodLength.ToString
                            Try
                                dblWC083_RunStandard = IFLConnectionObject.GetValue(strQuery1)
                            Catch ex As Exception
                                dblWC083_RunStandard = 0
                            End Try

                            Dim strQuery2 As String = "Select " + strRodDiameterColumn + " from " + oCostingAndCMSCommon.GetChromeMachiningTableName + " where TubeLength =" + dblCostRodLength.ToString 'ANUP 15-12-2010 START
                            Try
                                dblWCNumberValue = IFLConnectionObject.GetValue(strQuery2)
                            Catch ex As Exception
                                dblWCNumberValue = 0
                            End Try

                            Dim strQuery3 As String = "Select " + strRodDiameterColumn + " from TRChromeRodMachiningWCDetails where TubeLength =" + dblCostRodLength.ToString
                            Try
                                strWCNumber = IFLConnectionObject.GetValue(strQuery3)
                            Catch ex As Exception
                                strWCNumber = ""
                            End Try
                        ElseIf RodMaterialForCosting = "Nitro Steel" Then
                            Dim strQuery1 As String = "Select " + strRodDiameterColumn + " from TRNitroRodCuttingDetails where TubeLength =" + dblCostRodLength.ToString
                            Try
                                dblWC083_RunStandard = IFLConnectionObject.GetValue(strQuery1)
                            Catch ex As Exception
                                dblWC083_RunStandard = 0
                            End Try

                            Dim strQuery2 As String = "Select " + strRodDiameterColumn + " from " + oCostingAndCMSCommon.GetNitroRodMachiningTableName + " where TubeLength =" + dblCostRodLength.ToString  'ANUP 15-12-2010 START
                            Try
                                dblWCNumberValue = IFLConnectionObject.GetValue(strQuery2)
                            Catch ex As Exception
                                dblWCNumberValue = 0
                            End Try

                            Dim strQuery3 As String = "Select " + strRodDiameterColumn + " from TRNitroRodMachiningWCDetails where TubeLength =" + dblCostRodLength.ToString
                            Try
                                strWCNumber = IFLConnectionObject.GetValue(strQuery3)
                            Catch ex As Exception
                                strWCNumber = ""
                            End Try
                        ElseIf RodMaterialForCosting = "Induction Hardened" Then
                            dblWC083_RunStandard = 0

                            'ANUP 15-12-2010 START
                            ' dblWCNumberValue = 0
                            ' strWCNumber = ""
                            Dim strQuery2 As String = "Select " + strRodDiameterColumn + " from " + oCostingAndCMSCommon.GetInductionHBMachiningTableName + " where TubeLength =" + dblCostRodLength.ToString
                            Try
                                dblWCNumberValue = IFLConnectionObject.GetValue(strQuery2)
                            Catch ex As Exception
                                dblWCNumberValue = 0
                            End Try

                            Dim strQuery3 As String = "Select " + strRodDiameterColumn + " from " + oCostingAndCMSCommon.GetInductionHBMachiningTableName("WCDetails") + " where TubeLength =" + dblCostRodLength.ToString
                            Try
                                strWCNumber = IFLConnectionObject.GetValue(strQuery3)
                            Catch ex As Exception
                                strWCNumber = ""
                            End Try
                            'ANUP 15-12-2010 TILL HERE

                        End If

                        If Not dblWC083_RunStandard = 0 Then
                            _oExcelSheet_NewRod.Range("F14").Value = dblWC083_RunStandard
                        End If

                        If Not strWCNumber = "" Then
                            _oExcelSheet_NewRod.Range("B15").Value = strWCNumber
                        End If

                        If Not dblWCNumberValue = 0 Then
                            _oExcelSheet_NewRod.Range("F16").Value = dblWCNumberValue
                        End If


                        GetBurdenCost()

                        _oExWorkbook.Save()

                        GetNewPartCost = _oExcelSheet_NewRod.Range("L19").Value
                    Catch ex As Exception

                    End Try
                Else
                    _oExcelSheet_NewRod.Range("F14").Interior.Color = RGB(255, 0, 0)
                    _oExcelSheet_NewRod.Range("F16").Interior.Color = RGB(255, 0, 0)
                    _oExcelSheet_NewRod.Range("L19").Interior.Color = RGB(255, 0, 0)
                    _oExcelSheet_NewRod.Range("L19").Value = 0
                    _oExWorkbook.Save()
                    GetNewPartCost = 0
                    ' MessageBox.Show("Standard Run Cost is not avilable for RodDiameter" + dblRodDiamteter.ToString, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If
            End If
        Catch ex As Exception
            GetNewPartCost = 0
        End Try
    End Function

    Private Function GetRodWeight() As Double
        Try
            Dim strQuery1 As String = "select WeightPerFoot from RodWeightDetails where RodDiameter = " + RodDiameter.ToString
            GetRodWeight = IFLConnectionObject.GetValue(strQuery1)
        Catch ex As Exception
            GetRodWeight = 0
        End Try
    End Function

    Private Function DropDataToExcel() As Boolean
        DropDataToExcel = False
        Try
            PassDataToCMSIntegration()
            If PutDataToExcel() Then
                If ChangeSheetNamesAndVisibility() Then
                    '  If GetBurdenCost() Then
                    If DeleteEmptyRows() Then
                        If SaveExcel() Then
                            DropDataToExcel = True
                        End If
                    End If
                End If
            End If
            ' End If
        Catch ex As Exception
            DropDataToExcel = False
        End Try
    End Function

    Private Sub PassDataToCMSIntegration()
        'Sunny 22-06-10 For Integration Purpose

        CostDetails_Costing = New DataTable
        CostDetails_Costing = CostDetails.Clone
        For Each oCostDetailsDataRow As DataRow In CostDetails.Rows
            If oCostDetailsDataRow(CostDetailsItemOrder.Description).ToString.Contains("PAINT") Then

                Dim oNewRow1 As DataRow = CostDetails_Costing.NewRow
                oNewRow1.ItemArray = oCostDetailsDataRow.ItemArray
                CostDetails_Costing.Rows.Add(oNewRow1)

                'Dim oPaintDetails As DataRow = IFLConnectionObject.GetDataRow("Select PrimerCode, PrimerCatalystCode, PaintActivatorCode from PaintDetails where PaintCode = '" + oCostDetailsDataRow(CostDetailsItemOrder.CodeNumber) + "'")
                Dim oPaintDetails As DataRow = IFLConnectionObject.GetDataRow("Select PrimerCode, PaintActivatorCode from PaintDetails where PaintCode = '" + oCostDetailsDataRow(CostDetailsItemOrder.CodeNumber) + "'")
                If Not IsNothing(oPaintDetails) AndAlso oPaintDetails.ItemArray.Length > 0 Then
                    For Each oItem As Object In oPaintDetails.ItemArray
                        If Not oItem.ToString = "N/A" Then
                            Dim oNewPaintDataRow As DataRow = CostDetails.NewRow

                            oNewPaintDataRow.Item(CostDetailsItemOrder.CodeNumber) = oItem.ToString

                            Dim strDescription As String = IFLConnectionObject.GetValue("Select Description from CostingDetails where PartCode = '" + oItem.ToString + "'")
                            If Not IsNothing(strDescription) Then
                                oNewPaintDataRow.Item(CostDetailsItemOrder.Description) = strDescription
                            Else
                                oNewPaintDataRow.Item(CostDetailsItemOrder.Description) = ""
                            End If

                            oNewPaintDataRow.Item(CostDetailsItemOrder.Cost) = oCostDetailsDataRow(CostDetailsItemOrder.Cost)
                            oNewPaintDataRow.Item(CostDetailsItemOrder.PartName) = oCostDetailsDataRow(CostDetailsItemOrder.PartName)
                            oNewPaintDataRow.Item(CostDetailsItemOrder.IsExisting_New) = oCostDetailsDataRow(CostDetailsItemOrder.IsExisting_New)
                            oNewPaintDataRow.Item(CostDetailsItemOrder.Quantity) = oCostDetailsDataRow(CostDetailsItemOrder.Quantity)
                            oNewPaintDataRow.Item(CostDetailsItemOrder.Units) = oCostDetailsDataRow(CostDetailsItemOrder.Units)
                            oNewPaintDataRow.Item(CostDetailsItemOrder.Comment) = oCostDetailsDataRow(CostDetailsItemOrder.Comment)

                            Dim oNewRow2 As DataRow = CostDetails_Costing.NewRow
                            oNewRow2.ItemArray = oNewPaintDataRow.ItemArray
                            CostDetails_Costing.Rows.Add(oNewRow2)
                        End If
                    Next
                End If
            Else
                Try
                    Dim oNewRow3 As DataRow = CostDetails_Costing.NewRow
                    oNewRow3.ItemArray = oCostDetailsDataRow.ItemArray
                    CostDetails_Costing.Rows.Add(oNewRow3)
                Catch ex As Exception

                End Try

            End If
        Next
        'CostDetails_Costing = CostDetails
    End Sub

    Private Function GetBurdenCost() As Boolean
        GetBurdenCost = True
        '*************Main Assembly BurdenCost
        Try
            Dim strWC1Number_MainAssembly As String = _strWC1Number_MainAssembly
            Dim oWC1NumberDataRow As DataRow = IFLConnectionObject.GetDataRow("Select * from BurdenCost where WorkCenter = '" + _strWC1Number_MainAssembly + "'")
            If oWC1NumberDataRow.ItemArray.Length > 0 Then
                If Not IsDBNull(oWC1NumberDataRow("LabourRate")) Then
                    _oExcelSheet_MainAssembly.Range("F59").Value = oWC1NumberDataRow("LabourRate")
                End If
                If Not IsDBNull(oWC1NumberDataRow("FixedBurdenCost")) Then
                    _oExcelSheet_MainAssembly.Range("G59").Value = oWC1NumberDataRow("FixedBurdenCost")
                End If
                If Not IsDBNull(oWC1NumberDataRow("VariableBurdenCost")) Then
                    _oExcelSheet_MainAssembly.Range("H59").Value = oWC1NumberDataRow("VariableBurdenCost")
                End If
            End If
        Catch ex As Exception

        End Try

        Try
            Dim strWC2Value_MainAssembly As String = _oExcelSheet_MainAssembly.Range("B60").Value
            Dim oWC2NumberDataRow As DataRow = IFLConnectionObject.GetDataRow("Select * from BurdenCost where WorkCenter = '" + strWC2Value_MainAssembly + "'")
            If oWC2NumberDataRow.ItemArray.Length > 0 Then
                If Not IsDBNull(oWC2NumberDataRow("LabourRate")) Then
                    _oExcelSheet_MainAssembly.Range("F61").Value = oWC2NumberDataRow("LabourRate")
                End If
                If Not IsDBNull(oWC2NumberDataRow("FixedBurdenCost")) Then
                    _oExcelSheet_MainAssembly.Range("G61").Value = oWC2NumberDataRow("FixedBurdenCost")
                End If
                If Not IsDBNull(oWC2NumberDataRow("VariableBurdenCost")) Then
                    _oExcelSheet_MainAssembly.Range("H61").Value = oWC2NumberDataRow("VariableBurdenCost")
                End If
            End If
        Catch ex As Exception

        End Try
        '****************************************

        '*************New Tube BurdenCost
        Try
            Dim strWC1Value_NewTube As String = _oExcelSheet_NewTube.Range("B13").Value
            Dim oWC1Value_NewTubeDataRow As DataRow = IFLConnectionObject.GetDataRow("Select * from BurdenCost where WorkCenter = '" + strWC1Value_NewTube + "'")
            If oWC1Value_NewTubeDataRow.ItemArray.Length > 0 Then
                If Not IsDBNull(oWC1Value_NewTubeDataRow("LabourRate")) Then
                    _oExcelSheet_NewTube.Range("H14").Value = oWC1Value_NewTubeDataRow("LabourRate")
                End If
                If Not IsDBNull(oWC1Value_NewTubeDataRow("FixedBurdenCost")) Then
                    _oExcelSheet_NewTube.Range("I14").Value = oWC1Value_NewTubeDataRow("FixedBurdenCost")
                End If
                If Not IsDBNull(oWC1Value_NewTubeDataRow("VariableBurdenCost")) Then
                    _oExcelSheet_NewTube.Range("J14").Value = oWC1Value_NewTubeDataRow("VariableBurdenCost")
                End If
            End If
        Catch ex As Exception

        End Try

        Try
            Dim strWC2Value_NewTube As String = _oExcelSheet_NewTube.Range("B15").Value
            Dim oWC2Value_NewTubeDataRow As DataRow = IFLConnectionObject.GetDataRow("Select * from BurdenCost where WorkCenter = '" + strWC2Value_NewTube + "'")
            If oWC2Value_NewTubeDataRow.ItemArray.Length > 0 Then
                If Not IsDBNull(oWC2Value_NewTubeDataRow("LabourRate")) Then
                    _oExcelSheet_NewTube.Range("H16").Value = oWC2Value_NewTubeDataRow("LabourRate")
                End If
                If Not IsDBNull(oWC2Value_NewTubeDataRow("FixedBurdenCost")) Then
                    _oExcelSheet_NewTube.Range("I16").Value = oWC2Value_NewTubeDataRow("FixedBurdenCost")
                End If
                If Not IsDBNull(oWC2Value_NewTubeDataRow("VariableBurdenCost")) Then
                    _oExcelSheet_NewTube.Range("J16").Value = oWC2Value_NewTubeDataRow("VariableBurdenCost")
                End If
            End If
        Catch ex As Exception

        End Try

        '*************New Rod BurdenCost
        Try
            Dim strWC1Value_NewRod As String = _oExcelSheet_NewRod.Range("B13").Value
            Dim oWC1Value_NewRodDataRow As DataRow = IFLConnectionObject.GetDataRow("Select * from BurdenCost where WorkCenter = '" + strWC1Value_NewRod + "'")
            If oWC1Value_NewRodDataRow.ItemArray.Length > 0 Then
                If Not IsDBNull(oWC1Value_NewRodDataRow("LabourRate")) Then
                    _oExcelSheet_NewRod.Range("H14").Value = oWC1Value_NewRodDataRow("LabourRate")
                End If
                If Not IsDBNull(oWC1Value_NewRodDataRow("FixedBurdenCost")) Then
                    _oExcelSheet_NewRod.Range("I14").Value = oWC1Value_NewRodDataRow("FixedBurdenCost")
                End If
                If Not IsDBNull(oWC1Value_NewRodDataRow("VariableBurdenCost")) Then
                    _oExcelSheet_NewRod.Range("J14").Value = oWC1Value_NewRodDataRow("VariableBurdenCost")
                End If
            End If
        Catch ex As Exception

        End Try

        Try
            Dim strWC2Value_NewRod As String = _oExcelSheet_NewRod.Range("B15").Value
            Dim oWC2Value_NewRodDataRow As DataRow = IFLConnectionObject.GetDataRow("Select * from BurdenCost where WorkCenter = '" + strWC2Value_NewRod + "'")
            If oWC2Value_NewRodDataRow.ItemArray.Length > 0 Then
                If Not IsDBNull(oWC2Value_NewRodDataRow("LabourRate")) Then
                    _oExcelSheet_NewRod.Range("H16").Value = oWC2Value_NewRodDataRow("LabourRate")
                End If
                If Not IsDBNull(oWC2Value_NewRodDataRow("FixedBurdenCost")) Then
                    _oExcelSheet_NewRod.Range("I16").Value = oWC2Value_NewRodDataRow("FixedBurdenCost")
                End If
                If Not IsDBNull(oWC2Value_NewRodDataRow("VariableBurdenCost")) Then
                    _oExcelSheet_NewRod.Range("J16").Value = oWC2Value_NewRodDataRow("VariableBurdenCost")
                End If
            End If
        Catch ex As Exception

        End Try
        '****************************************

    End Function

    Private Function GetTPDetails() As Boolean
        'Sunny 21-04-10 4pm
        If SeriesForCosting.Contains("TP") Then
            _oExcelSheet_NewTube.Range("A17").Value = 3
            _oExcelSheet_NewTube.Range("B17").Value = "WC136"
            _oExcelSheet_NewTube.Range("E17").Value = "S"
            _oExcelSheet_NewTube.Range("F17").Value = 0.0
            _oExcelSheet_NewTube.Range("G17").Value = "C"
            _oExcelSheet_NewTube.Range("H17").Value = 21.11
            _oExcelSheet_NewTube.Range("I17").Value = 31.5
            _oExcelSheet_NewTube.Range("J17").Value = 31.5
            _oExcelSheet_NewTube.Range("L17").Value = 0.0
            _oExcelSheet_NewTube.Range("M17").Value = 0.0
            _oExcelSheet_NewTube.Range("N17").Value = 0.0

            _oExcelSheet_NewTube.Range("E18").Value = "R"
            _oExcelSheet_NewTube.Range("F18").Value = 125
            _oExcelSheet_NewTube.Range("G18").Value = "A"
            Try
                Dim strWC3Value_NewTube As String = _oExcelSheet_NewTube.Range("B17").Value
                Dim oWC3Value_NewTubeDataRow As DataRow = IFLConnectionObject.GetDataRow("Select * from BurdenCost where WorkCenter = '" + strWC3Value_NewTube + "'")
                If oWC3Value_NewTubeDataRow.ItemArray.Length > 0 Then
                    If Not IsDBNull(oWC3Value_NewTubeDataRow("LabourRate")) Then
                        _oExcelSheet_NewTube.Range("H18").Value = oWC3Value_NewTubeDataRow("LabourRate")
                    End If
                    If Not IsDBNull(oWC3Value_NewTubeDataRow("FixedBurdenCost")) Then
                        _oExcelSheet_NewTube.Range("I18").Value = oWC3Value_NewTubeDataRow("FixedBurdenCost")
                    End If
                    If Not IsDBNull(oWC3Value_NewTubeDataRow("VariableBurdenCost")) Then
                        _oExcelSheet_NewTube.Range("J18").Value = oWC3Value_NewTubeDataRow("VariableBurdenCost")
                    End If
                End If
            Catch ex As Exception
            End Try
            _oExcelSheet_NewTube.Range("L18").Value = 0.16888
            _oExcelSheet_NewTube.Range("M18").Value = 0.252
            _oExcelSheet_NewTube.Range("N18").Value = 0.252



            _oExcelSheet_NewTube.Range("A19").Value = 4
            _oExcelSheet_NewTube.Range("B19").Value = "WC136"
            _oExcelSheet_NewTube.Range("E19").Value = "S"
            _oExcelSheet_NewTube.Range("F19").Value = 0.5
            _oExcelSheet_NewTube.Range("G19").Value = "C"
            _oExcelSheet_NewTube.Range("H19").Value = 20.26
            _oExcelSheet_NewTube.Range("I19").Value = 45.0
            _oExcelSheet_NewTube.Range("J19").Value = 45.0
            _oExcelSheet_NewTube.Range("L19").Value = 0.2026
            _oExcelSheet_NewTube.Range("M19").Value = 0.45
            _oExcelSheet_NewTube.Range("N19").Value = 0.45

            _oExcelSheet_NewTube.Range("E20").Value = "R"
            '06_10_2010   RAGAVA
            If Trim(ofrmTieRod1.cmbRephasingPortPosition.Text).IndexOf("At Both") <> -1 Then
                _oExcelSheet_NewTube.Range("F20").Value = 15
            ElseIf Trim(ofrmTieRod1.cmbRephasingPortPosition.Text).IndexOf("At Extension") <> -1 OrElse Trim(ofrmTieRod1.cmbRephasingPortPosition.Text).IndexOf("At Retraction") <> -1 Then
                _oExcelSheet_NewTube.Range("F20").Value = 30
            End If
            'Till   Here
            _oExcelSheet_NewTube.Range("G20").Value = "A"
            Try
                Dim strWC4Value_NewTube As String = _oExcelSheet_NewTube.Range("B19").Value
                Dim oWC4Value_NewTubeDataRow As DataRow = IFLConnectionObject.GetDataRow("Select * from BurdenCost where WorkCenter = '" + strWC4Value_NewTube + "'")
                If oWC4Value_NewTubeDataRow.ItemArray.Length > 0 Then
                    If Not IsDBNull(oWC4Value_NewTubeDataRow("LabourRate")) Then
                        _oExcelSheet_NewTube.Range("H20").Value = oWC4Value_NewTubeDataRow("LabourRate")
                    End If
                    If Not IsDBNull(oWC4Value_NewTubeDataRow("FixedBurdenCost")) Then
                        _oExcelSheet_NewTube.Range("I20").Value = oWC4Value_NewTubeDataRow("FixedBurdenCost")
                    End If
                    If Not IsDBNull(oWC4Value_NewTubeDataRow("VariableBurdenCost")) Then
                        _oExcelSheet_NewTube.Range("J20").Value = oWC4Value_NewTubeDataRow("VariableBurdenCost")
                    End If
                End If
            Catch ex As Exception
            End Try
            _oExcelSheet_NewTube.Range("L20").Value = 0.5065
            _oExcelSheet_NewTube.Range("M20").Value = 1.125
            _oExcelSheet_NewTube.Range("N20").Value = 1.125
        End If
      
        '****************************************
    End Function

    Private Function PutDataToExcel() As Boolean
        PutDataToExcel = True
        Try
            _intTotalCostExcelRange = 6
            Dim intSNo As Integer = 1
            Dim cSNo As Char = "A"
            Dim cMaterial As Char = "B"
            Dim cDescription As Char = "C"
            Dim cQuantityPerUnit As Char = "D"
            Dim cUnits As Char = "E"
            Dim cStandardUnitCost As Char = "F"
            For Each oFinalCostDetails As DataRow In CostDetails.Rows

                Dim strCodeNumber As String = ""
                Dim strDescription As String = ""
                Dim dblCostOfPart As Double = 0
                Dim dblQuantityPerUnit As Double = 0
                Dim strUnit As String = ""

                If Not IsNothing(oFinalCostDetails(CostDetailsItemOrder.CodeNumber)) Then
                    Dim strChangedPartCode As String = GetPurchasedCode(oFinalCostDetails(CostDetailsItemOrder.CodeNumber))
                    If Not IsNothing(strChangedPartCode) Then
                        strCodeNumber = strChangedPartCode
                    Else
                        strCodeNumber = oFinalCostDetails(CostDetailsItemOrder.CodeNumber)
                    End If
                Else
                    strCodeNumber = ""
                End If

                If Not oFinalCostDetails(CostDetailsItemOrder.Cost) = 0 Then
                    dblCostOfPart = oFinalCostDetails(CostDetailsItemOrder.Cost)
                Else
                    'dblCostOfPart = 0.00001
                    dblCostOfPart = 0
                End If

                'Added by Sandeep on 17-03-10 
                If Not oFinalCostDetails(CostDetailsItemOrder.Quantity).Equals(DBNull.Value) Then
                    dblQuantityPerUnit = oFinalCostDetails(CostDetailsItemOrder.Quantity)
                Else
                    dblQuantityPerUnit = 1
                End If

                'Added by Sandeep on 18-03-10 3pm
                If Not oFinalCostDetails(CostDetailsItemOrder.Units).Equals(DBNull.Value) Then
                    strUnit = oFinalCostDetails(CostDetailsItemOrder.Units)
                Else
                    strUnit = "N/A"
                End If

                If oFinalCostDetails(CostDetailsItemOrder.IsExisting_New) = "New" Then
                    strDescription = oFinalCostDetails(CostDetailsItemOrder.PartName)

                    If oFinalCostDetails(CostDetailsItemOrder.PartName) = "Tube Code Number" Then
                        _strNewTubeName = strCodeNumber
                        _blnIsNewTube = True
                        _oExcelSheet_NewTube.Range("B2").Value = _strNewTubeName
                        _oExcelSheet_NewTube.Range("C2").Value = oFinalCostDetails(CostDetailsItemOrder.PartName)
                    End If

                    If oFinalCostDetails(CostDetailsItemOrder.PartName) = "Rod Code Number" Then
                        _strNewRodName = strCodeNumber
                        _blnIsNewRod = True
                        _oExcelSheet_NewRod.Range("B2").Value = _strNewRodName
                        _oExcelSheet_NewRod.Range("C2").Value = oFinalCostDetails(CostDetailsItemOrder.PartName)
                    End If
                Else
                    If Not IsNothing(oFinalCostDetails(CostDetailsItemOrder.Description)) Then
                        strDescription = oFinalCostDetails(CostDetailsItemOrder.Description)
                    Else
                        strDescription = ""
                    End If

                    If oFinalCostDetails(CostDetailsItemOrder.PartName) = "Tube Code Number" Then
                        _strNewTubeName = strCodeNumber
                        _blnIsNewTube = False
                    End If

                    If oFinalCostDetails(CostDetailsItemOrder.PartName) = "Rod Code Number" Then
                        _strNewRodName = strCodeNumber
                        _blnIsNewRod = False
                    End If
                End If

                _oExcelSheet_MainAssembly.Range(cSNo + _intTotalCostExcelRange.ToString).Value = intSNo
                _oExcelSheet_MainAssembly.Range(cMaterial + _intTotalCostExcelRange.ToString).Value = strCodeNumber
                _oExcelSheet_MainAssembly.Range(cDescription + _intTotalCostExcelRange.ToString).Value = strDescription
                _oExcelSheet_MainAssembly.Range(cQuantityPerUnit + _intTotalCostExcelRange.ToString).Value = dblQuantityPerUnit
                _oExcelSheet_MainAssembly.Range(cUnits + _intTotalCostExcelRange.ToString).Value = strUnit
                _oExcelSheet_MainAssembly.Range(cStandardUnitCost + _intTotalCostExcelRange.ToString).Value = dblCostOfPart

                If dblCostOfPart = 0 Then
                    _oExcelSheet_MainAssembly.Range(cSNo + _intTotalCostExcelRange.ToString).EntireRow.Font.Color = RGB(255, 0, 0)
                End If

                intSNo = intSNo + 1
                _intTotalCostExcelRange = _intTotalCostExcelRange + 1
            Next
            SetWorkCenter_MainAssembly()
            ''_oExcelSheet_MainAssembly.Range("E2").Value = SRQ
            _oExcelSheet_MainAssembly.Range("B2").Value = CylinderCodeNumber
            _oExcelSheet_MainAssembly.Range("C2").Value = SetCodeDesciption
        Catch ex As Exception
            PutDataToExcel = False
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

    Private Function SetWorkCenter_MainAssembly() As Boolean
        'ToDo:Sandeep 19-04-10
        Dim strBoreDiameterColumn As String = ""
        Dim dblBoreDiamteter As Double = Val(ofrmTieRod1.cmbBore.Text)
        If dblBoreDiamteter = 2 Then
            strBoreDiameterColumn = "BoreDiameter_2"
        ElseIf dblBoreDiamteter = 2.5 Then
            strBoreDiameterColumn = "BoreDiameter_2_5"
        ElseIf dblBoreDiamteter = 2.75 Then
            strBoreDiameterColumn = "BoreDiameter_2_75"
        ElseIf dblBoreDiamteter = 3 Then
            strBoreDiameterColumn = "BoreDiameter_3"
        ElseIf dblBoreDiamteter = 3.25 Then
            strBoreDiameterColumn = "BoreDiameter_3_25"
        ElseIf dblBoreDiamteter = 3.5 Then
            strBoreDiameterColumn = "BoreDiameter_3_5"
        ElseIf dblBoreDiamteter = 3.75 Then
            strBoreDiameterColumn = "BoreDiameter_3_75"
        ElseIf dblBoreDiamteter = 4 Then
            strBoreDiameterColumn = "BoreDiameter_4"
        ElseIf dblBoreDiamteter = 4.25 Then
            strBoreDiameterColumn = "BoreDiameter_4_25"
        ElseIf dblBoreDiamteter = 4.5 Then
            strBoreDiameterColumn = "BoreDiameter_4_5"
        ElseIf dblBoreDiamteter = 4.75 Then
            strBoreDiameterColumn = "BoreDiameter_4_75"
        ElseIf dblBoreDiamteter = 5 Then
            strBoreDiameterColumn = "BoreDiameter_5"
        End If

        If Not strBoreDiameterColumn = "" Then
            Try
                Dim strWCNumber As String = ""
                Dim dblWCNumberValue As Double = 0
                Dim dblCostStrokeLength As Double = Math.Ceiling(StrokeLength)

                If SeriesForCosting = "TX (TXC)" Then
                    Dim strQuery3 As String = "Select " + strBoreDiameterColumn + " from TX_TXC_Details where Stroke =" + dblCostStrokeLength.ToString
                    Try
                        dblWCNumberValue = IFLConnectionObject.GetValue(strQuery3)
                    Catch ex As Exception
                        dblWCNumberValue = 0
                    End Try

                    Dim strQuery4 As String = "Select " + strBoreDiameterColumn + " from TX_TXC_WorkCenterDetails where Stroke =" + dblCostStrokeLength.ToString
                    Try
                        strWCNumber = IFLConnectionObject.GetValue(strQuery4)
                    Catch ex As Exception
                        strWCNumber = ""
                    End Try
                    'ElseIf SeriesForCosting = "TL (TC)" OrElse SeriesForCosting = "TH (TD)" Then     '21_01_2011     RAGAVA
                ElseIf SeriesForCosting = "TL (TC)" OrElse SeriesForCosting = "TH (TD)" OrElse SeriesForCosting = "LN" Then      '21_01_2011     RAGAVA
                    Dim strQuery3 As String = "Select " + strBoreDiameterColumn + " from TC_TD_TH_TL_Details where Stroke =" + dblCostStrokeLength.ToString
                    Try
                        dblWCNumberValue = IFLConnectionObject.GetValue(strQuery3)
                    Catch ex As Exception
                        dblWCNumberValue = 0
                    End Try
                    Dim strQuery4 As String = "Select " + strBoreDiameterColumn + " from TC_TD_TH_TL_WorkCenterDetails where Stroke =" + dblCostStrokeLength.ToString
                    Try
                        strWCNumber = IFLConnectionObject.GetValue(strQuery4)
                    Catch ex As Exception
                        strWCNumber = ""
                    End Try
                ElseIf SeriesForCosting = "TP-High" OrElse SeriesForCosting = "TP-Low" Then
                    Dim strQuery3 As String = "Select " + strBoreDiameterColumn + " from TP_Details where Stroke =" + dblCostStrokeLength.ToString
                    Try
                        dblWCNumberValue = IFLConnectionObject.GetValue(strQuery3)
                    Catch ex As Exception
                        dblWCNumberValue = 0
                    End Try
                    Dim strQuery4 As String = "Select " + strBoreDiameterColumn + " from TP_WorkCenterDetails where Stroke =" + dblCostStrokeLength.ToString
                    Try
                        strWCNumber = IFLConnectionObject.GetValue(strQuery4)
                    Catch ex As Exception
                        strWCNumber = ""
                    End Try
                End If

                'ToDo:Sunny 15-04-10 5pm
                If Not strWCNumber = "" Then
                    _oExcelSheet_MainAssembly.Range("B58").Value = strWCNumber
                    _strWC1Number_MainAssembly = strWCNumber

                    'TODO:Sunny 23-06-10
                    METHDRAssemblyResource = strWCNumber
                End If
                If Not dblWCNumberValue = 0 Then
                    _oExcelSheet_MainAssembly.Range("E59").Value = dblWCNumberValue

                    'TODO:Sunny 23-06-10
                    METHDRAssemblyRunStandard = dblWCNumberValue
                End If
                '*************************

                '06_10_2010   RAGAVA
                Dim stdcost As Double
                If UCase(CustomerName).IndexOf("CNH") <> -1 Then
                    If BoreDiameter <= 3 Then
                        stdcost = 12.63
                    ElseIf BoreDiameter > 3 AndAlso BoreDiameter <= 4 Then
                        stdcost = 11.34
                    ElseIf BoreDiameter > 4 Then
                        stdcost = 9.3
                    End If
                Else
                    If BoreDiameter <= 3 Then
                        stdcost = 16
                    ElseIf BoreDiameter > 3 AndAlso BoreDiameter <= 4 Then
                        stdcost = 14
                    ElseIf BoreDiameter > 4 Then
                        stdcost = 11
                    End If
                End If
                Try
                    Dim TempExcelApp As Excel.Application
                    Dim TempWorkBook As Excel.Workbook
                    Dim tempWorkSheet As Excel.Worksheet
                    Dim Rng As Excel.Range
                    TempExcelApp = New Excel.Application
                    TempExcelApp.Visible = False
                    TempWorkBook = TempExcelApp.Workbooks.Open(System.Environment.CurrentDirectory & "\CMS_PAINTING.xls")
                    tempWorkSheet = TempExcelApp.Sheets("Paint Standards")

                    '10_09_2011   RAGAVA
                    Rng = tempWorkSheet.Range("L28")
                    Rng.Value = "1"
                    Rng = tempWorkSheet.Range("L29")
                    Rng.Value = Trim(ofrmTieRod1.txtRetractedLength.Text)
                    Rng = tempWorkSheet.Range("L30")
                    Rng.Value = BoreDiameter
                    Rng = tempWorkSheet.Range("L31")
                    Rng.Value = 1
                    Rng = tempWorkSheet.Range("L32")
                    If RodClevisPins = True OrElse ClevisPins = True Then
                        If blnInstallPinsandClips_Checked = True Then
                            Rng.Value = 0
                        Else
                            Rng.Value = 1
                        End If
                    Else
                        Rng.Value = 0
                    End If
                    Rng = tempWorkSheet.Range("L33")
                    Rng.Value = Weight_Assembly
                    ' sugandhi start
                    Dim s As String = tempWorkSheet.Range("L34").Value.ToString()
                    Dim words As String() = s.Split(New Char() {"&"c})

                    Dim val As Double = Convert.ToInt32(words(0)) / 17

                    ' Rng = tempWorkSheet.Range("L26")
                    'Till  Here
                    METHDRPaintRunStandard = val.ToString()
                Catch ex As Exception
                End Try

                If ofrmTieRod3.chk100OilTest.Checked = True Then
                    _oExcelSheet_MainAssembly.Range("B60").Value = "WC626"
                    _oExcelSheet_MainAssembly.Range("E61").Value = stdcost
                    _oExcelSheet_MainAssembly.Range("B62").Value = "WC631"
                    _oExcelSheet_MainAssembly.Range("E63").Value = METHDRPaintRunStandard
                Else
                    _oExcelSheet_MainAssembly.Range("B60").Value = "WC631"
                    _oExcelSheet_MainAssembly.Range("E61").Value = METHDRPaintRunStandard
                    _oExcelSheet_MainAssembly.Range("A62", "A63").EntireRow.Delete()
                    '_oExcelSheet_MainAssembly.Range("A63").EntireRow.Clear()
                End If
                'Till   Here

                If ofrmTieRod3.rdbOnlyCosting.Checked = True Then      '07_10_2010   RAGAVA
                    'ToDo:Sunny 18-06-10 5pm
                    Dim strQuery5 As String = "Select " + strBoreDiameterColumn + " from TR_Welded_PaintingDetails where Stroke =" + dblCostStrokeLength.ToString
                    Try
                        _strWC2Number_MainAssembly = IFLConnectionObject.GetValue(strQuery5)
                    Catch ex As Exception
                        strWCNumber = ""
                    End Try
                    If Not _strWC2Number_MainAssembly = "" Then
                        _oExcelSheet_MainAssembly.Range("E61").Value = _strWC2Number_MainAssembly

                        'TODO:Sunny 23-06-10
                        METHDRPaintRunStandard = _strWC2Number_MainAssembly
                    End If
                    '***********************
                End If
                _oExWorkbook.Save()
            Catch ex As Exception
            End Try
        Else
            'MessageBox.Show("Standard Run Cost is not avilable for BoreDiamteter" + dblBoreDiamteter.ToString, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End If
        '************************* by Sandeep 19-04-10
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

            'TODO:Sunny 19-04-10 12:20pm
            If Not Directory.Exists("W:\TIEROD\COSTING\") Then
                Directory.CreateDirectory("W:\TIEROD\COSTING\")
            End If
            If File.Exists(_strchildExcelPath) Then
                If File.Exists("W:\TIEROD\COSTING\" + CylinderCodeNumber + "_Costing.xls") Then
                    File.Delete("W:\TIEROD\COSTING\" + CylinderCodeNumber + "_Costing.xls")
                End If
                File.Move(_strchildExcelPath, "W:\TIEROD\COSTING\" + CylinderCodeNumber + "_Costing.xls")
            End If
            SaveExcel = True
        Catch ex As Exception
            SaveExcel = False
        End Try
    End Function

#End Region

#Region "Excel Functions"

    Public Function CheckForExcel() As Boolean
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

    Public Function CopyTheMasterFile() As Boolean
        CopyTheMasterFile = False
        Dim blnIsProcessSuccessfull As Boolean = False
        Dim sErrorMessage As String = "Report Master file does not exist"
        Try
            ' CloseExcel()
            ' This function checks if the master report format exists
            If IsMasterReportFileExists() Then
                Try
                    ' Check if file already exists
                    If File.Exists(_strchildExcelPath) Then
                        If Not IsNothing(_oExApplication) Then
                            _oExApplication = Nothing
                        End If
                        ' CloseExcel()
                        ' Delete the existing file first
                        File.Delete(_strchildExcelPath)
                    End If
                    File.Copy(_strMotherExcelPath, _strchildExcelPath)
                    CopyTheMasterFile = True
                    blnIsProcessSuccessfull = True
                Catch oException As Exception
                    sErrorMessage = "Unable to copy the source file" + vbCrLf + vbCrLf + oException.Message
                End Try
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

    Public Function IsMasterReportFileExists() As Boolean
        IsMasterReportFileExists = File.Exists(_strMotherExcelPath)
    End Function

    Public Function CreateExcel() As Boolean
        CreateExcel = True
        Try
            _oExApplication = New Excel.Application
            _oExApplication.Visible = False
            _oExWorkbook = _oExApplication.Workbooks.Open(_strchildExcelPath)
            _oExcelSheet_MainAssembly = _oExApplication.Sheets(1)
            _oExcelSheet_NewTube = _oExApplication.Sheets(2)
            _oExcelSheet_NewRod = _oExApplication.Sheets(3)
        Catch ex As Exception
            CreateExcel = False
            MessageBox.Show("Unable to open Excel sheet", "Information", MessageBoxButtons.OK, _
            MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Function

    Private Function ChangeSheetNamesAndVisibility() As Boolean
        ChangeSheetNamesAndVisibility = False
        Try
            _oExcelSheet_MainAssembly.Name = CylinderCodeNumber

            If _blnIsNewTube Then
                If Not IsNothing(_strNewTubeName) Then
                    _oExcelSheet_NewTube.Name = _strNewTubeName
                End If
                _oExcelSheet_NewTube.Visible = Excel.XlSheetVisibility.xlSheetVisible
            Else
                If Not IsNothing(_strNewTubeName) Then
                    _oExcelSheet_NewTube.Name = _strNewTubeName
                End If
                _oExcelSheet_NewTube.Visible = Excel.XlSheetVisibility.xlSheetHidden
            End If

            If _blnIsNewRod Then
                If Not IsNothing(_strNewRodName) Then
                    _oExcelSheet_NewRod.Name = _strNewRodName
                End If
                _oExcelSheet_NewRod.Visible = Excel.XlSheetVisibility.xlSheetVisible
            Else
                If Not IsNothing(_strNewRodName) Then
                    _oExcelSheet_NewRod.Name = _strNewRodName
                End If
                _oExcelSheet_NewRod.Visible = Excel.XlSheetVisibility.xlSheetHidden
            End If

            ChangeSheetNamesAndVisibility = True
        Catch ex As Exception
            ChangeSheetNamesAndVisibility = False
        End Try
    End Function

    Private Function DeleteEmptyRows() As Boolean
        DeleteEmptyRows = False
        Try
            Dim strRowCount As String = 6
            Dim strStart_EmptyCellRange As String = ""

            While (strRowCount > 0)
                If IsNothing(_oExcelSheet_MainAssembly.Range("A" + strRowCount).Value) Then
                    strStart_EmptyCellRange = "A" + strRowCount
                    Exit While
                End If
                strRowCount += 1
            End While

            While (IsNothing(_oExcelSheet_MainAssembly.Range("A" + strRowCount).Value))
                _oExcelSheet_MainAssembly.Range("A" + strRowCount, "A" + strRowCount).EntireRow.Delete()
            End While

            _oExWorkbook.Save()

            DeleteEmptyRows = True
        Catch ex As Exception
            DeleteEmptyRows = False
        End Try
    End Function

#End Region

End Class
 