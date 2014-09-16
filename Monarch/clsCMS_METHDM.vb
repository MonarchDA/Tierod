Imports Microsoft.Office.Interop
Public Class clsCMS_METHDM

#Region "Private Variables"

    Private _objExcelSheet As Excel.Worksheet

    Private _strMethodType As String 'Always 1

    Private _strAlternateMethodNumber As String 'Always 0

    Private _strPlantCode As String 'Always C01

    Private _strPartNumber As String

    Private _strSequenceNumber As String 'Always 10

    Private _strLineNumber As String 'Increament by 10

    Private _strMaterialPartNumber As String 'From Costing Datatable

    Private _strMaterialDescription As String 'Always Blank

    Private _strStockType As String 'Always Blank

    Private _strQuantityPer As String 'From Costing Datatable

    Private _strUnitofMeansure_forQuantityPer As String 'From Costing Datatable

    Private _strQuantityMultiplier As String 'Always 1

    Private _strRequired_orByProduct As String 'Always R

    Private _strAllocation As String 'Always Y

    Private _strBackFlush As String 'Always Y

    Private _strBlowThroughPart As String 'Always Blank except Bag Plastic(1)

    Private _strMajorComponent As String 'Always N

    Private _strScrapPercentage As String 'Always 0

    Private _strSparePartsQuantity As String 'Always 0

    Private _strStockLocation As String

    Private _strDrawFromLoc_AsDefined As String 'Always Y

    Private _strItemNumber As String 'To Be verified

    Private _strItemNumberExtension As String 'Always Blank

    Private _strDrawingNumber As String 'Always Blank

    Private _strMajorGroupCode As String 'Always Blank

    Private _strLeadTime As String 'Always Blank

    Private _strFixUsageFlag As String 'Always Blank

    Private _strWholeUnitConsumptionFlag As String 'Always Blank

    Private _strRoundingCutoffDecimalValue As String 'Always Blank

    Private _strLastappliedECNdtl As String 'Always Blank

    Private _strLastappliedECNline As String 'Always Blank

    Private _strLastappliedECNdate As String 'Always Blank 

    Private _strLastappliedECN As String 'Always Blank

    Private _strECNEffectiveDate As String 'Always Blank

    Private _strLastappliedECNtime As String 'Always Blank

#End Region

#Region "Public Properties"

    Public ReadOnly Property MethodType() As String
        Get
            Return _strMethodType
        End Get
    End Property

    Public ReadOnly Property AlternateMethodNumber() As String
        Get
            Return _strAlternateMethodNumber
        End Get
    End Property

    Public ReadOnly Property PlantCode() As String
        Get
            Return _strPlantCode
        End Get
    End Property

    Public Property PartNumber() As String
        Get
            Return _strPartNumber
        End Get
        Set(ByVal value As String)
            _strPartNumber = value
        End Set
    End Property

    '06_09_2010   RAGAVA  ReadOnly to READ-WRITE modified
    'Public ReadOnly Property SequenceNumber() As String
    '    Get
    '        Return _strSequenceNumber
    '    End Get
    'End Property
    Public Property SequenceNumber() As String
        Get
            Return _strSequenceNumber
        End Get
        Set(ByVal value As String)
            _strSequenceNumber = value
        End Set
    End Property

    Public Property LineNumber() As String
        Get
            Return _strLineNumber
        End Get
        Set(ByVal value As String)
            _strLineNumber = value
        End Set
    End Property

    Public Property MaterialPartNumber() As String
        Get
            Return _strMaterialPartNumber
        End Get
        Set(ByVal value As String)
            _strMaterialPartNumber = value
        End Set
    End Property

    Public Property MaterialDescription() As String
        Get
            Return _strMaterialDescription
        End Get
        Set(ByVal value As String)
            _strMaterialDescription = value
        End Set
    End Property

    Public Property StockType() As String
        Get
            Return _strStockType
        End Get
        Set(ByVal value As String)
            _strStockType = value
        End Set
    End Property

    Public Property QuantityPer() As String
        Get
            Return _strQuantityPer
        End Get
        Set(ByVal value As String)
            _strQuantityPer = value
        End Set
    End Property

    Public Property UnitofMeansure_forQuantityPer() As String
        Get
            Return _strUnitofMeansure_forQuantityPer
        End Get
        Set(ByVal value As String)
            _strUnitofMeansure_forQuantityPer = value
        End Set
    End Property

    Public ReadOnly Property QuantityMultiplier() As String
        Get
            Return _strQuantityMultiplier
        End Get
    End Property

    Public ReadOnly Property Required_orByProduct() As String
        Get
            Return _strRequired_orByProduct
        End Get
    End Property

    Public ReadOnly Property Allocation() As String
        Get
            Return _strAllocation
        End Get
    End Property

    Public ReadOnly Property BackFlush() As String
        Get
            Return _strBackFlush
        End Get
    End Property

    Public Property BlowThroughPart() As String
        Get
            Return _strBlowThroughPart
        End Get
        Set(ByVal value As String)
            _strBlowThroughPart = value
        End Set
    End Property

    Public ReadOnly Property MajorComponent() As String
        Get
            Return _strMajorComponent
        End Get
    End Property

    Public ReadOnly Property ScrapPercentage() As String
        Get
            Return _strScrapPercentage
        End Get
    End Property

    Public ReadOnly Property SparePartsQuantity() As String
        Get
            Return _strSparePartsQuantity
        End Get
    End Property

    Public Property StockLocation() As String
        Get
            Return _strStockLocation
        End Get
        Set(ByVal value As String)
            _strStockLocation = value
        End Set
    End Property

    Public ReadOnly Property DrawFromLoc_AsDefined() As String
        Get
            Return _strDrawFromLoc_AsDefined
        End Get
    End Property

    Public Property ItemNumber() As String
        Get
            Return _strItemNumber
        End Get
        Set(ByVal value As String)
            _strItemNumber = value
        End Set
    End Property

    Public ReadOnly Property ItemNumberExtension() As String
        Get
            Return _strItemNumberExtension
        End Get
    End Property

    Public ReadOnly Property DrawingNumber() As String
        Get
            Return _strDrawingNumber
        End Get
    End Property

    Public ReadOnly Property MajorGroupCode() As String
        Get
            Return _strMajorGroupCode
        End Get
    End Property

    Public ReadOnly Property LeadTime() As String
        Get
            Return _strLeadTime
        End Get
    End Property

    Public ReadOnly Property FixUsageFlag() As String
        Get
            Return _strFixUsageFlag
        End Get
    End Property

    Public ReadOnly Property WholeUnitConsumptionFlag() As String
        Get
            Return _strWholeUnitConsumptionFlag
        End Get
    End Property

    Public ReadOnly Property RoundingCutoffDecimalValue() As String
        Get
            Return _strRoundingCutoffDecimalValue
        End Get
    End Property

    Public ReadOnly Property LastappliedECNdtl() As String
        Get
            Return _strLastappliedECNdtl
        End Get
    End Property

    Public ReadOnly Property LastappliedECNline() As String
        Get
            Return _strLastappliedECNline
        End Get
    End Property

    Public ReadOnly Property LastappliedECNdate() As String
        Get
            Return _strLastappliedECNdate
        End Get
    End Property

    Public ReadOnly Property LastappliedECN() As String
        Get
            Return _strLastappliedECN
        End Get
    End Property

    Public ReadOnly Property ECNEffectiveDate() As String
        Get
            Return _strECNEffectiveDate
        End Get
    End Property

    Public ReadOnly Property LastappliedECNtime() As String
        Get
            Return _strLastappliedECNtime
        End Get
    End Property
   
    'Public Property SetAndGet() As String
    '    Get
    '        Return _strSetAndGet
    '    End Get
    '    Set(ByVal value As String)
    '        _strSetAndGet = value
    '    End Set
    'End Property


    'Public ReadOnly Property OnlyGet() As String
    '    Get
    '        Return _strOnlyGet
    '    End Get
    'End Property

#End Region

#Region "Sub Procedures"

    Public Sub SetCommenPropertyValue()

        _strMethodType = 1

        _strAlternateMethodNumber = 0

        _strPlantCode = "C01"

        _strSequenceNumber = "010"

        _strMaterialDescription = ""

        _strQuantityMultiplier = 1

        _strRequired_orByProduct = "R"

        _strAllocation = "Y"

        _strBackFlush = "Y"

        _strMajorComponent = "N"

        _strScrapPercentage = 0

        _strSparePartsQuantity = 0

        _strDrawFromLoc_AsDefined = "Y"

        _strItemNumberExtension = ""

        _strDrawingNumber = ""

        _strMajorGroupCode = ""

        _strLeadTime = ""

        _strFixUsageFlag = ""

        _strWholeUnitConsumptionFlag = ""

        _strRoundingCutoffDecimalValue = ""

        _strLastappliedECNdtl = ""

        _strLastappliedECNline = ""

        _strLastappliedECNdate = ""

        _strLastappliedECN = ""

        _strECNEffectiveDate = ""

        _strLastappliedECNtime = ""

    End Sub

    Public Sub SetDataToExcel(ByVal oExcelSheet As Excel.Worksheet, ByVal strRowcount As String)
        Try
            _objExcelSheet = oExcelSheet
            SetDataToCell("A" + strRowcount, MethodType)
            SetDataToCell("B" + strRowcount, AlternateMethodNumber)
            SetDataToCell("C" + strRowcount, PlantCode)
            SetDataToCell("D" + strRowcount, PartNumber)
            SetDataToCell("E" + strRowcount, SequenceNumber)
            SetDataToCell("F" + strRowcount, LineNumber)
            SetDataToCell("G" + strRowcount, MaterialPartNumber)
            SetDataToCell("H" + strRowcount, MaterialDescription)
            SetDataToCell("I" + strRowcount, StockType)
            SetDataToCell("J" + strRowcount, QuantityPer)
            SetDataToCell("K" + strRowcount, UnitofMeansure_forQuantityPer)
            SetDataToCell("L" + strRowcount, QuantityMultiplier)
            SetDataToCell("M" + strRowcount, Required_orByProduct)
            SetDataToCell("N" + strRowcount, Allocation)
            SetDataToCell("O" + strRowcount, BackFlush)
            SetDataToCell("P" + strRowcount, BlowThroughPart)
            SetDataToCell("Q" + strRowcount, MajorComponent)
            SetDataToCell("R" + strRowcount, ScrapPercentage)
            SetDataToCell("S" + strRowcount, SparePartsQuantity)
            SetDataToCell("T" + strRowcount, StockLocation)
            SetDataToCell("U" + strRowcount, DrawFromLoc_AsDefined)
            SetDataToCell("V" + strRowcount, ItemNumber)
            SetDataToCell("W" + strRowcount, ItemNumberExtension)
            SetDataToCell("X" + strRowcount, DrawingNumber)
            SetDataToCell("Y" + strRowcount, MajorGroupCode)
            SetDataToCell("Z" + strRowcount, LeadTime)
            SetDataToCell("AA" + strRowcount, FixUsageFlag)
            SetDataToCell("AB" + strRowcount, WholeUnitConsumptionFlag)
            SetDataToCell("AC" + strRowcount, RoundingCutoffDecimalValue)
            SetDataToCell("AD" + strRowcount, LastappliedECNdtl)
            SetDataToCell("AE" + strRowcount, LastappliedECNline)
            SetDataToCell("AF" + strRowcount, LastappliedECNdate)
            SetDataToCell("AG" + strRowcount, ECNEffectiveDate)
            SetDataToCell("AH" + strRowcount, RoundingCutoffDecimalValue)
            SetDataToCell("AI" + strRowcount, LastappliedECNtime)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SetDataToCell(ByVal strCellRange As String, ByVal strValue As String)
        Try
            If Not IsNothing(strValue) Then
                _objExcelSheet.Range(strCellRange).Value = strValue
            End If
        Catch ex As Exception
            Dim strError As String = ex.ToString
        End Try
    End Sub

#End Region

End Class
