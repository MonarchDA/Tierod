Imports Microsoft.Office.Interop
Public Class clsCMS_STKA

#Region "Private Variables"

    Private _objExcelSheet As Excel.Worksheet

    Private _strInternalPartNumber As String

    Private _strPlantCode As String 'Always C01

    Private _strUnitofIssue As String 'Always EA

    Private _strReplenishmentType As String

    Private _strSourcePlant As String 'Always blank

    Private _strVendorLeadTime_TransferLeadTime_inDays As String

    Private _strMinimumTransferQuantity As String 'Always 1

    Private _strTransferPolicy As String 'Always 1

    Private _strTransferMultiplier As String 'Always 1

    Private _strAvailabilityInquiryDisplayUnit As String 'Always blank

    Private _strStatus As String 'Always 1

    Private _strInactiveReasonCode As String 'Always DM

    Private _strMinimumQuantity_inUnitofIssue As String 'Always 0

    Private _strMaximumQuantity_inUnitofIssue As String 'Always 0

    Private _strEstimatedAnnualVolume_inUnitofIssue As String 'Always 0

    Private _strLotNumberMandatory As String 'Always N

    Private _strCreateLot_SerialAssociation As String 'Always N

    Private _strMaintainLotBalance As String 'Always N

    Private _strValidateLotNumbers As String 'Always N

    Private _strSerializedMandatory As String 'Always N

    Private _strABCCode As String 'Always blank

    Private _strCycleCountStartDate As String 'Always 0/00/00

    Private _strCycleCountsPerYear As String 'Always blank

    Private _strSellablePart As String 'Always Y

    Private _strRepriceLock As String 'Always blank

    Private _strPricingUnit As String 'Always EA

    Private _strMinimumOrderQuantity_inUnitofIssue As String

    Private _strDefaultContainer As String 'Always blank

    Private _strDefaultPalletContainer As String 'Always blank

    Private _strStandardPackSize As String 'Always blank

    Private _strStandardPackSizeUOM As String 'Always blank

    Private _strMasterPackSize As String 'Always blank

    Private _strMasterPackSizeUOM As String 'Always blank

    Private _strShipFromLocation As String 'Always blank

    Private _strAllocationTimeFence_inDays As String 'Always 30

    Private _strKitCode As String 'Always blank

    Private _strDirectBuyFlag As String 'Always 3

    Private _strDirectBuyActionFlag As String 'Always 1

    Private _strSCDPart As String 'Always 2

    Private _strPOReceiving_StandalonePrintFlag As String 'Always N

    Private _strPOReceiving_StandaloneNumberofCopies As String 'Always blank

    Private _strPOReceiving_StandaloneLabelFormatCode As String 'Always blank

    Private _strProductionReporting_StandalonePrintFlag As String 'Always N

    Private _strProductionReportingStandaloneNumberofCopies As String 'Always blank

    Private _strProductionReportingStandaloneLabelFormatCode As String 'Always blank

    Private _strCompletedProduction_StandalonePrintFlag As String 'Always N

    Private _strCompletedProduction_StandaloneNumberofCopies As String 'Always blank

    Private _strCompletedProduction_StandaloneLabelFormatCode As String 'Always blank

    Private _strShipping_StandaloneNumberofCopies As String 'Always blank

    Private _strShipping_StandaloneLabelFormatCode As String 'Always blank

    Private _strPOReceiving_MasterPrintFlag As String 'Always N

    Private _strPOReceiving_MasterNumberofCopies As String 'Always blank

    Private _strPOReceiving_MasterLabelFormatCode As String 'Always blank

    Private _strCompletedProduction_MasterPrintFlag As String 'Always N

    Private _strCompletedProduction_MasterNumberofCopies As String 'Always blank

    Private _strCompletedProduction_MasterLabelFormatCode As String 'Always blank

    Private _strShipping_MasterPrintFlag As String 'Always N

    Private _strShipping_MasterNumberofCopies As String 'Always blank

    Private _strShipping_MasterLabelFormatCode As String 'Always blank

    Private _strCreateSingle_MasterSerial_forReceivingSerials As String 'Always 2

    Private _strBuyerCode As String

    Private _strPlannerCode As String

    Private _strReceivingContainer As String 'Always blank

    Private _strReceivingPackSize As String 'Always blank

    Private _strReceivingPackSizeUOM As String 'Always blank

    Private _strReceiveToLocation As String

    Private _strDrawFromLocation As String 'Always blank

    Private _strInspectionProcedure As String

    Private _strMaterialPreparationLeadTime_inDays As String

    Private _strRequisitionerCode As String 'Always blank

    Private _strRequisitionsRequiresApproval As String 'Always blank

    Private _strCountryOfOrigin As String

    Private _strProvinceOfOrigin As String 'Always blank

    Private _strPurchaseOrderUnit As String 'Always EA

    Private _strApprovedSupplierRequired As String 'Always 2

    Private _strGenerateMonthendAccrualCOGS_forServiceCharged As String 'Always 2

    Private _strScheduleType As String

    Private _strOptimumRun_PurchaseSize As String

    Private _strMinimumRun_PurchaseSize As String

    Private _strProductionMultiplier As String 'Always 0

    Private _strShrinkageFactor_inPercentage As String 'Always 0

    Private _strManufacturingLeadTime_inDays As String 'Always blank

    Private _strLongestLeadTime_inDays As String 'Always blank

    Private _strOrderReleaseLeadTime_inDays As String 'Always 1

    Private _strOrderPolicy As String 'Always 2

    Private _strSalesForecastTimeFence_inDays As String

    Private _strSecuredTimeFence_inDays As String 'Always blank

    Private _strOrderLookAhead_inDays As String 'Always blank

    Private _strLoadingTimeFence_inDays As String 'Always blank

    Private _strMRPAutoReleaseProduction As String 'Always N

    Private _strRepetitiveControl As String

    Private _strPercentComplete_inPercentage As String 'Always 100

    Private _strBalance_BuildOutDate As String 'Always 0/00/00

    Private _strCommonMaterialPart As String 'Always blank

    Private _strRestrictCurrentStandardCost As String 'Always blank

    Private _strMaterialIssueLocation As String 'Always blank

    Private _strMattecPart As String 'Always blank

    Private _strMattecMaterialType As String 'Always blank

    Private _strStorageClass As String 'Always blank

    Private _strVelocityCode As String 'Always blank

    Private _strNumberofReceivingPack_perPallet As String 'Always blank

    Private _strReceivingPackComment As String 'Always blank

    Private _strNumberofMasterPackperPallet As String 'Always blank

    Private _strMasterPackComment As String 'Always blank

    '06_10_2010    RAGAVA   added for getting commas as per client
    Private _strVariable1 As String = String.Empty
    Private _strVariable2 As String = String.Empty
    Private _strVariable3 As String = String.Empty
    Private _strVariable4 As String = String.Empty
    Private _strVariable5 As String = String.Empty
    Private _strVariable6 As String = String.Empty
    Private _strVariable7 As String = String.Empty
    Private _strVariable8 As String = String.Empty
    Private _strVariable9 As String = String.Empty
    Private _strVariable10 As String = String.Empty
    Private _strVariable11 As String = String.Empty
    Private _strVariable12 As String = String.Empty
    Private _strVariable13 As String = String.Empty
    Private _strVariable14 As String = String.Empty
    Private _strVariable15 As String = String.Empty
    Private _strVariable16 As String = String.Empty
    Private _strVariable17 As String = String.Empty
    Private _strVariable18 As String = String.Empty
    'Till  Here
#End Region

#Region "Public Properties"

    Public Property InternalPartNumber() As String
        Get
            Return _strInternalPartNumber
        End Get
        Set(ByVal value As String)
            _strInternalPartNumber = value
        End Set
    End Property

    Public ReadOnly Property PlantCode() As String
        Get
            Return _strPlantCode
        End Get
    End Property

    Public ReadOnly Property UnitofIssue() As String
        Get
            Return _strUnitofIssue
        End Get
    End Property

    Public Property ReplenishmentType() As String
        Get
            Return _strReplenishmentType
        End Get
        Set(ByVal value As String)
            _strReplenishmentType = value
        End Set
    End Property

    Public ReadOnly Property SourcePlant() As String
        Get
            Return _strSourcePlant
        End Get
    End Property

    Public Property VendorLeadTime_TransferLeadTime_inDays() As String
        Get
            Return _strVendorLeadTime_TransferLeadTime_inDays
        End Get
        Set(ByVal value As String)
            _strVendorLeadTime_TransferLeadTime_inDays = value
        End Set
    End Property

    Public ReadOnly Property MinimumTransferQuantity() As String
        Get
            Return _strMinimumTransferQuantity
        End Get
    End Property

    Public ReadOnly Property TransferPolicy() As String
        Get
            Return _strTransferPolicy
        End Get
    End Property

    Public ReadOnly Property TransferMultiplier() As String
        Get
            Return _strTransferMultiplier
        End Get
    End Property

    Public ReadOnly Property AvailabilityInquiryDisplayUnit() As String
        Get
            Return _strAvailabilityInquiryDisplayUnit
        End Get
    End Property

    Public ReadOnly Property Status() As String
        Get
            Return _strStatus
        End Get
    End Property

    Public ReadOnly Property InactiveReasonCode() As String
        Get
            Return _strInactiveReasonCode
        End Get
    End Property

    Public ReadOnly Property MinimumQuantity_inUnitofIssue() As String
        Get
            Return _strMinimumQuantity_inUnitofIssue
        End Get
    End Property

    Public ReadOnly Property MaximumQuantity_inUnitofIssue() As String
        Get
            Return _strMaximumQuantity_inUnitofIssue
        End Get
    End Property

    Public ReadOnly Property EstimatedAnnualVolume_inUnitofIssue() As String
        Get
            Return _strEstimatedAnnualVolume_inUnitofIssue
        End Get
        'Set(ByVal value As String)
        '    _strEstimatedAnnualVolume_inUnitofIssue = value
        'End Set
    End Property

    Public ReadOnly Property LotNumberMandatory() As String
        Get
            Return _strLotNumberMandatory
        End Get
    End Property

    Public ReadOnly Property CreateLot_SerialAssociation() As String
        Get
            Return _strCreateLot_SerialAssociation
        End Get
    End Property

    Public ReadOnly Property MaintainLotBalance() As String
        Get
            Return _strMaintainLotBalance
        End Get
    End Property

    Public ReadOnly Property ValidateLotNumbers() As String
        Get
            Return _strValidateLotNumbers
        End Get
    End Property

    Public ReadOnly Property SerializedMandatory() As String
        Get
            Return _strSerializedMandatory
        End Get
    End Property

    Public ReadOnly Property ABCCode() As String
        Get
            Return _strABCCode
        End Get
        'Set(ByVal value As String)
        '    _strABCCode = value
        'End Set
    End Property

    Public ReadOnly Property CycleCountStartDate() As String
        Get
            Return _strCycleCountStartDate
        End Get
        'Set(ByVal value As String)
        '    _strCycleCountStartDate = value
        'End Set
    End Property

    Public ReadOnly Property CycleCountsPerYear() As String
        Get
            Return _strCycleCountsPerYear
        End Get
        'Set(ByVal value As String)
        '    _strCycleCountsPerYear = value
        'End Set
    End Property

    Public ReadOnly Property SellablePart() As String
        Get
            Return _strSellablePart
        End Get
    End Property

    Public ReadOnly Property RepriceLock() As String
        Get
            Return _strRepriceLock
        End Get
    End Property

    Public ReadOnly Property PricingUnit() As String
        Get
            Return _strPricingUnit
        End Get
    End Property

    Public Property MinimumOrderQuantity_inUnitofIssue() As String
        Get
            Return _strMinimumOrderQuantity_inUnitofIssue
        End Get
        Set(ByVal value As String)
            _strMinimumOrderQuantity_inUnitofIssue = value
        End Set
    End Property

    Public ReadOnly Property DefaultContainer() As String
        Get
            Return _strDefaultContainer
        End Get
    End Property

    Public ReadOnly Property DefaultPalletContainer() As String
        Get
            Return _strDefaultPalletContainer
        End Get
    End Property

    Public ReadOnly Property StandardPackSize() As String
        Get
            Return _strStandardPackSize
        End Get
    End Property

    Public ReadOnly Property StandardPackSizeUOM() As String
        Get
            Return _strStandardPackSizeUOM
        End Get
    End Property

    Public ReadOnly Property MasterPackSize() As String
        Get
            Return _strMasterPackSize
        End Get
    End Property

    Public ReadOnly Property MasterPackSizeUOM() As String
        Get
            Return _strMasterPackSizeUOM
        End Get
    End Property

    Public ReadOnly Property ShipFromLocation() As String
        Get
            Return _strShipFromLocation
        End Get
    End Property

    Public ReadOnly Property AllocationTimeFence_inDays() As String
        Get
            Return _strAllocationTimeFence_inDays
        End Get
        'Set(ByVal value As String)
        '    _strAllocationTimeFence_inDays = value
        'End Set
    End Property

    Public ReadOnly Property KitCode() As String
        Get
            Return _strKitCode
        End Get
    End Property

    Public ReadOnly Property DirectBuyFlag() As String
        Get
            Return _strDirectBuyFlag
        End Get
    End Property

    Public ReadOnly Property DirectBuyActionFlag() As String
        Get
            Return _strDirectBuyActionFlag
        End Get
    End Property

    Public ReadOnly Property SCDPart() As String
        Get
            Return _strSCDPart
        End Get
    End Property

    Public ReadOnly Property POReceiving_StandalonePrintFlag() As String
        Get
            Return _strPOReceiving_StandalonePrintFlag
        End Get
    End Property

    Public ReadOnly Property POReceiving_StandaloneNumberofCopies() As String
        Get
            Return _strPOReceiving_StandaloneNumberofCopies
        End Get
    End Property

    Public ReadOnly Property POReceiving_StandaloneLabelFormatCode() As String
        Get
            Return _strPOReceiving_StandaloneLabelFormatCode
        End Get
    End Property

    Public ReadOnly Property ProductionReporting_StandalonePrintFlag() As String
        Get
            Return _strProductionReporting_StandalonePrintFlag
        End Get
    End Property

    Public ReadOnly Property ProductionReportingStandaloneNumberofCopies() As String
        Get
            Return _strProductionReportingStandaloneNumberofCopies
        End Get
    End Property

    Public ReadOnly Property ProductionReportingStandaloneLabelFormatCode() As String
        Get
            Return _strProductionReportingStandaloneLabelFormatCode
        End Get
    End Property

    Public ReadOnly Property CompletedProduction_StandalonePrintFlag() As String
        Get
            Return _strCompletedProduction_StandalonePrintFlag
        End Get
    End Property

    Public ReadOnly Property CompletedProduction_StandaloneNumberofCopies() As String
        Get
            Return _strCompletedProduction_StandaloneNumberofCopies
        End Get
    End Property

    Public ReadOnly Property CompletedProduction_StandaloneLabelFormatCode() As String
        Get
            Return _strCompletedProduction_StandaloneLabelFormatCode
        End Get
    End Property

    Public ReadOnly Property Shipping_StandaloneNumberofCopies() As String
        Get
            Return _strShipping_StandaloneNumberofCopies
        End Get
    End Property

    Public ReadOnly Property Shipping_StandaloneLabelFormatCode() As String
        Get
            Return _strShipping_StandaloneLabelFormatCode
        End Get
    End Property

    Public ReadOnly Property POReceiving_MasterPrintFlag() As String
        Get
            Return _strPOReceiving_MasterPrintFlag
        End Get
    End Property

    Public ReadOnly Property POReceiving_MasterNumberofCopies() As String
        Get
            Return _strPOReceiving_MasterNumberofCopies
        End Get
    End Property

    Public ReadOnly Property POReceiving_MasterLabelFormatCode() As String
        Get
            Return _strPOReceiving_MasterLabelFormatCode
        End Get
    End Property

    Public ReadOnly Property CompletedProduction_MasterPrintFlag() As String
        Get
            Return _strCompletedProduction_MasterPrintFlag
        End Get
    End Property

    Public ReadOnly Property CompletedProduction_MasterNumberofCopies() As String
        Get
            Return _strCompletedProduction_MasterNumberofCopies
        End Get
    End Property

    Public ReadOnly Property CompletedProduction_MasterLabelFormatCode() As String
        Get
            Return _strCompletedProduction_MasterLabelFormatCode
        End Get
    End Property

    Public ReadOnly Property Shipping_MasterPrintFlag() As String
        Get
            Return _strShipping_MasterPrintFlag
        End Get
    End Property

    Public ReadOnly Property Shipping_MasterNumberofCopies() As String
        Get
            Return _strShipping_MasterNumberofCopies
        End Get
    End Property

    Public ReadOnly Property Shipping_MasterLabelFormatCode() As String
        Get
            Return _strShipping_MasterLabelFormatCode
        End Get
    End Property

    Public ReadOnly Property CreateSingle_MasterSerial_forReceivingSerials() As String
        Get
            Return _strCreateSingle_MasterSerial_forReceivingSerials
        End Get
    End Property

    Public Property BuyerCode() As String
        Get
            Return _strBuyerCode
        End Get
        Set(ByVal value As String)
            _strBuyerCode = value
        End Set
    End Property

    Public Property PlannerCode() As String
        Get
            Return _strPlannerCode
        End Get
        Set(ByVal value As String)
            _strPlannerCode = value
        End Set
    End Property

    Public ReadOnly Property ReceivingContainer() As String
        Get
            Return _strReceivingContainer
        End Get
    End Property

    Public ReadOnly Property ReceivingPackSize() As String
        Get
            Return _strReceivingPackSize
        End Get
    End Property

    Public ReadOnly Property ReceivingPackSizeUOM() As String
        Get
            Return _strReceivingPackSizeUOM
        End Get
    End Property

    Public Property ReceiveToLocation() As String
        Get
            Return _strReceiveToLocation
        End Get
        Set(ByVal value As String)
            _strReceiveToLocation = value
        End Set
    End Property

    Public ReadOnly Property DrawFromLocation() As String
        Get
            Return _strDrawFromLocation
        End Get
    End Property

    Public Property InspectionProcedure() As String
        Get
            Return _strInspectionProcedure
        End Get
        Set(ByVal value As String)
            _strInspectionProcedure = value
        End Set
    End Property

    Public Property MaterialPreparationLeadTime_inDays() As String
        Get
            Return _strMaterialPreparationLeadTime_inDays
        End Get
        Set(ByVal value As String)
            _strMaterialPreparationLeadTime_inDays = value
        End Set
    End Property

    Public ReadOnly Property RequisitionerCode() As String
        Get
            Return _strRequisitionerCode
        End Get
    End Property

    Public ReadOnly Property RequisitionsRequiresApproval() As String
        Get
            Return _strRequisitionsRequiresApproval
        End Get
    End Property

    Public Property CountryOfOrigin() As String
        Get
            Return _strCountryOfOrigin
        End Get
        Set(ByVal value As String)
            _strCountryOfOrigin = value
        End Set
    End Property

    Public ReadOnly Property ProvinceOfOrigin() As String
        Get
            Return _strProvinceOfOrigin
        End Get
    End Property

    Public ReadOnly Property PurchaseOrderUnit() As String
        Get
            Return _strPurchaseOrderUnit
        End Get
    End Property

    Public ReadOnly Property ApprovedSupplierRequired() As String
        Get
            Return _strApprovedSupplierRequired
        End Get
    End Property

    Public ReadOnly Property GenerateMonthendAccrualCOGS_forServiceCharged() As String
        Get
            Return _strGenerateMonthendAccrualCOGS_forServiceCharged
        End Get
    End Property

    Public Property ScheduleType() As String
        Get
            Return _strScheduleType
        End Get
        Set(ByVal value As String)
            _strScheduleType = value
        End Set
    End Property

    Public Property OptimumRun_PurchaseSize() As String
        Get
            Return _strOptimumRun_PurchaseSize
        End Get
        Set(ByVal value As String)
            _strOptimumRun_PurchaseSize = value
        End Set
    End Property

    Public Property MinimumRun_PurchaseSize() As String
        Get
            Return _strMinimumRun_PurchaseSize
        End Get
        Set(ByVal value As String)
            _strMinimumRun_PurchaseSize = value
        End Set
    End Property

    Public ReadOnly Property ProductionMultiplier() As String
        Get
            Return _strProductionMultiplier
        End Get
    End Property

    Public ReadOnly Property ShrinkageFactor_inPercentage() As String
        Get
            Return _strShrinkageFactor_inPercentage
        End Get
    End Property

    Public ReadOnly Property ManufacturingLeadTime_inDays() As String
        Get
            Return _strManufacturingLeadTime_inDays
        End Get
        'Set(ByVal value As String)
        '    _strManufacturingLeadTime_inDays = value
        'End Set
    End Property

    Public ReadOnly Property LongestLeadTime_inDays() As String
        Get
            Return _strLongestLeadTime_inDays
        End Get
        'Set(ByVal value As String)
        '    _strLongestLeadTime_inDays = value
        'End Set
    End Property

    Public ReadOnly Property OrderReleaseLeadTime_inDays() As String
        Get
            Return _strOrderReleaseLeadTime_inDays
        End Get
        'Set(ByVal value As String)
        '    _strOrderReleaseLeadTime_inDays = value
        'End Set
    End Property

    Public ReadOnly Property OrderPolicy() As String
        Get
            Return _strOrderPolicy
        End Get
        'Set(ByVal value As String)
        '    _strOrderPolicy = value
        'End Set
    End Property

    Public Property SalesForecastTimeFence_inDays() As String
        Get
            Return _strSalesForecastTimeFence_inDays
        End Get
        Set(ByVal value As String)
            _strSalesForecastTimeFence_inDays = value
        End Set
    End Property

    Public ReadOnly Property SecuredTimeFence_inDays() As String
        Get
            Return _strSecuredTimeFence_inDays
        End Get
    End Property

    Public ReadOnly Property OrderLookAhead_inDays() As String
        Get
            Return _strOrderLookAhead_inDays
        End Get
    End Property

    Public ReadOnly Property LoadingTimeFence_inDays() As String
        Get
            Return _strLoadingTimeFence_inDays
        End Get
    End Property

    Public ReadOnly Property MRPAutoReleaseProduction() As String
        Get
            Return _strMRPAutoReleaseProduction
        End Get
    End Property

    Public Property RepetitiveControl() As String
        Get
            Return _strRepetitiveControl
        End Get
        Set(ByVal value As String)
            _strRepetitiveControl = value
        End Set
    End Property

    Public ReadOnly Property PercentComplete_inPercentage() As String
        Get
            Return _strPercentComplete_inPercentage
        End Get
    End Property

    Public ReadOnly Property Balance_BuildOutDate() As String
        Get
            Return _strBalance_BuildOutDate
        End Get
    End Property

    Public ReadOnly Property CommonMaterialPart() As String
        Get
            Return _strCommonMaterialPart
        End Get
    End Property

    Public ReadOnly Property RestrictCurrentStandardCost() As String
        Get
            Return _strRestrictCurrentStandardCost
        End Get
    End Property

    Public ReadOnly Property MaterialIssueLocation() As String
        Get
            Return _strMaterialIssueLocation
        End Get
    End Property

    Public ReadOnly Property MattecPart() As String
        Get
            Return _strMattecPart
        End Get
    End Property

    Public ReadOnly Property MattecMaterialType() As String
        Get
            Return _strMattecMaterialType
        End Get
    End Property

    Public ReadOnly Property StorageClass() As String
        Get
            Return _strStorageClass
        End Get
    End Property

    Public ReadOnly Property VelocityCode() As String
        Get
            Return _strVelocityCode
        End Get
    End Property

    Public ReadOnly Property NumberofReceivingPack_perPallet() As String
        Get
            Return _strNumberofReceivingPack_perPallet
        End Get
    End Property

    Public ReadOnly Property ReceivingPackComment() As String
        Get
            Return _strReceivingPackComment
        End Get
    End Property

    Public ReadOnly Property NumberofMasterPackperPallet() As String
        Get
            Return _strNumberofMasterPackperPallet
        End Get
    End Property

    Public ReadOnly Property MasterPackComment() As String
        Get
            Return _strMasterPackComment
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
        _strPlantCode = "C01"

        _strUnitofIssue = "EA"

        _strSourcePlant = ""

        _strMinimumTransferQuantity = 1

        _strTransferPolicy = 1

        _strTransferMultiplier = 1

        _strAvailabilityInquiryDisplayUnit = ""

        _strStatus = "I"

        _strInactiveReasonCode = "DM"

        _strMinimumQuantity_inUnitofIssue = 0

        _strMaximumQuantity_inUnitofIssue = 0

        _strEstimatedAnnualVolume_inUnitofIssue = 0

        _strLotNumberMandatory = "N"

        _strCreateLot_SerialAssociation = "N"

        _strMaintainLotBalance = "N"

        _strValidateLotNumbers = "N"

        _strSerializedMandatory = "N"

        _strABCCode = ""

        _strCycleCountStartDate = ""         '06_09_2010   RAGAVA     "0/00/00"

        _strCycleCountsPerYear = ""

        _strSellablePart = "Y"

        _strRepriceLock = ""

        _strPricingUnit = "EA"

        _strDefaultContainer = ""

        _strDefaultPalletContainer = ""

        _strStandardPackSize = ""

        _strStandardPackSizeUOM = ""

        _strMasterPackSize = ""

        _strMasterPackSizeUOM = ""

        _strShipFromLocation = ""

        _strAllocationTimeFence_inDays = 30

        _strKitCode = ""

        _strDirectBuyFlag = 3

        _strDirectBuyActionFlag = 1

        _strSCDPart = 2

        _strPOReceiving_StandalonePrintFlag = "N"

        _strPOReceiving_StandaloneNumberofCopies = ""

        _strPOReceiving_StandaloneLabelFormatCode = ""

        _strProductionReporting_StandalonePrintFlag = "N"

        _strProductionReportingStandaloneNumberofCopies = ""

        _strProductionReportingStandaloneLabelFormatCode = ""

        _strCompletedProduction_StandalonePrintFlag = "N"

        _strCompletedProduction_StandaloneNumberofCopies = ""

        _strCompletedProduction_StandaloneLabelFormatCode = ""

        _strShipping_StandaloneNumberofCopies = ""

        _strShipping_StandaloneLabelFormatCode = ""

        _strPOReceiving_MasterPrintFlag = "N"

        _strPOReceiving_MasterNumberofCopies = ""

        _strPOReceiving_MasterLabelFormatCode = "N"     '06_09_2010   RAGAVA ""

        _strCompletedProduction_MasterPrintFlag = "N"

        _strCompletedProduction_MasterNumberofCopies = ""

        _strCompletedProduction_MasterLabelFormatCode = ""

        _strShipping_MasterPrintFlag = "N"

        _strShipping_MasterNumberofCopies = ""

        _strShipping_MasterLabelFormatCode = ""

        _strCreateSingle_MasterSerial_forReceivingSerials = 2

        _strReceivingContainer = ""

        _strReceivingPackSize = ""

        _strReceivingPackSizeUOM = ""

        _strDrawFromLocation = ""

        _strRequisitionerCode = ""

        _strRequisitionsRequiresApproval = ""

        _strProvinceOfOrigin = ""

        _strPurchaseOrderUnit = "EA"

        _strApprovedSupplierRequired = 2

        _strGenerateMonthendAccrualCOGS_forServiceCharged = 2

        _strProductionMultiplier = 0

        _strShrinkageFactor_inPercentage = 0

        _strManufacturingLeadTime_inDays = ""

        _strLongestLeadTime_inDays = ""

        _strOrderReleaseLeadTime_inDays = 1

        _strOrderPolicy = 1

        _strSecuredTimeFence_inDays = ""

        _strOrderLookAhead_inDays = ""

        _strLoadingTimeFence_inDays = ""

        _strMRPAutoReleaseProduction = "N"

        _strPercentComplete_inPercentage = 100

        _strBalance_BuildOutDate = ""     '06_09_2010   RAGAVA "0/00/00"

        _strCommonMaterialPart = ""

        _strRestrictCurrentStandardCost = ""

        _strMaterialIssueLocation = ""

        _strMattecPart = ""

        _strMattecMaterialType = ""

        _strStorageClass = ""

        _strVelocityCode = ""

        _strNumberofReceivingPack_perPallet = ""

        _strReceivingPackComment = ""

        _strNumberofMasterPackperPallet = ""

        _strMasterPackComment = ""

        '06_10_2010    RAGAVA   added for getting commas as per client
        _strVariable1 = ""
        _strVariable2 = ""
        _strVariable3 = ""
        _strVariable4 = ""
        _strVariable5 = ""
        _strVariable6 = ""
        _strVariable7 = ""
        _strVariable8 = ""
        _strVariable9 = ""
        _strVariable10 = ""
        _strVariable11 = ""
        _strVariable12 = ""
        _strVariable13 = ""
        _strVariable14 = ""
        _strVariable15 = ""
        _strVariable16 = ""
        _strVariable17 = ""
        _strVariable18 = ""
        'Till  Here

    End Sub

    Public Sub SetDataToExcel(ByVal oExcelSheet As Excel.Worksheet)
        Try
            _objExcelSheet = oExcelSheet

            SetDataToCell("A2", InternalPartNumber)
            SetDataToCell("B2", PlantCode)
            SetDataToCell("C2", UnitofIssue)
            SetDataToCell("D2", ReplenishmentType)
            SetDataToCell("E2", SourcePlant)
            SetDataToCell("F2", VendorLeadTime_TransferLeadTime_inDays)
            SetDataToCell("G2", MinimumTransferQuantity)
            SetDataToCell("H2", TransferPolicy)
            SetDataToCell("I2", TransferMultiplier)
            SetDataToCell("J2", AvailabilityInquiryDisplayUnit)
            SetDataToCell("K2", Status)
            SetDataToCell("L2", InactiveReasonCode)
            SetDataToCell("M2", MinimumQuantity_inUnitofIssue)
            SetDataToCell("N2", MaximumQuantity_inUnitofIssue)
            SetDataToCell("O2", EstimatedAnnualVolume_inUnitofIssue)
            SetDataToCell("P2", LotNumberMandatory)
            SetDataToCell("Q2", CreateLot_SerialAssociation)
            SetDataToCell("R2", MaintainLotBalance)
            SetDataToCell("S2", ValidateLotNumbers)
            SetDataToCell("T2", SerializedMandatory)
            SetDataToCell("U2", ABCCode)
            SetDataToCell("V2", CycleCountStartDate)
            SetDataToCell("W2", CycleCountsPerYear)
            SetDataToCell("X2", SellablePart)
            SetDataToCell("Y2", RepriceLock)
            SetDataToCell("Z2", PricingUnit)

            SetDataToCell("AA2", MinimumOrderQuantity_inUnitofIssue)
            SetDataToCell("AB2", DefaultContainer)
            SetDataToCell("AC2", DefaultPalletContainer)
            SetDataToCell("AD2", StandardPackSize)
            SetDataToCell("AE2", StandardPackSizeUOM)
            SetDataToCell("AF2", MasterPackSize)
            SetDataToCell("AG2", MasterPackSizeUOM)
            SetDataToCell("AH2", ShipFromLocation)
            SetDataToCell("AI2", AllocationTimeFence_inDays)
            SetDataToCell("AJ2", KitCode)
            SetDataToCell("AK2", DirectBuyFlag)
            SetDataToCell("AL2", DirectBuyActionFlag)
            SetDataToCell("AM2", SCDPart)
            SetDataToCell("AN2", POReceiving_StandalonePrintFlag)
            SetDataToCell("AO2", POReceiving_StandaloneNumberofCopies)
            SetDataToCell("AP2", POReceiving_StandaloneLabelFormatCode)
            SetDataToCell("AQ2", ProductionReporting_StandalonePrintFlag)
            SetDataToCell("AR2", ProductionReportingStandaloneNumberofCopies)
            SetDataToCell("AS2", ProductionReportingStandaloneLabelFormatCode)
            SetDataToCell("AT2", CompletedProduction_StandalonePrintFlag)
            SetDataToCell("AU2", CompletedProduction_StandaloneNumberofCopies)
            SetDataToCell("AV2", CompletedProduction_StandaloneLabelFormatCode)
            SetDataToCell("AW2", Shipping_StandaloneNumberofCopies)
            SetDataToCell("AX2", Shipping_StandaloneLabelFormatCode)
            SetDataToCell("AY2", POReceiving_MasterPrintFlag)
            SetDataToCell("AZ2", POReceiving_MasterNumberofCopies)

            SetDataToCell("BA2", POReceiving_MasterLabelFormatCode)
            SetDataToCell("BB2", CompletedProduction_MasterPrintFlag)
            SetDataToCell("BC2", CompletedProduction_MasterNumberofCopies)
            SetDataToCell("BD2", CompletedProduction_MasterLabelFormatCode)
            SetDataToCell("BE2", Shipping_MasterPrintFlag)
            SetDataToCell("BF2", Shipping_MasterNumberofCopies)
            SetDataToCell("BG2", Shipping_MasterLabelFormatCode)
            SetDataToCell("BH2", CreateSingle_MasterSerial_forReceivingSerials)
            SetDataToCell("BI2", BuyerCode)
            SetDataToCell("BJ2", PlannerCode)
            SetDataToCell("BK2", ReceivingContainer)
            SetDataToCell("BL2", ReceivingPackSize)
            SetDataToCell("BM2", ReceivingPackSizeUOM)
            SetDataToCell("BN2", ReceiveToLocation)
            SetDataToCell("BO2", DrawFromLocation)
            SetDataToCell("BP2", InspectionProcedure)
            SetDataToCell("BQ2", MaterialPreparationLeadTime_inDays)
            SetDataToCell("BR2", RequisitionerCode)
            SetDataToCell("BS2", RequisitionsRequiresApproval)
            SetDataToCell("BT2", CountryOfOrigin)
            SetDataToCell("BU2", ProvinceOfOrigin)
            SetDataToCell("BV2", PurchaseOrderUnit)
            SetDataToCell("BW2", ApprovedSupplierRequired)
            SetDataToCell("BX2", GenerateMonthendAccrualCOGS_forServiceCharged)
            SetDataToCell("BY2", ScheduleType)
            SetDataToCell("BZ2", OptimumRun_PurchaseSize)

            SetDataToCell("CA2", MinimumRun_PurchaseSize)
            SetDataToCell("CB2", ProductionMultiplier)
            SetDataToCell("CC2", ShrinkageFactor_inPercentage)
            SetDataToCell("CD2", ManufacturingLeadTime_inDays)
            SetDataToCell("CE2", LongestLeadTime_inDays)
            SetDataToCell("CF2", OrderReleaseLeadTime_inDays)
            SetDataToCell("CG2", OrderPolicy)
            SetDataToCell("CH2", SalesForecastTimeFence_inDays)
            SetDataToCell("CI2", SecuredTimeFence_inDays)
            SetDataToCell("CJ2", OrderLookAhead_inDays)
            SetDataToCell("CK2", LoadingTimeFence_inDays)
            SetDataToCell("CL2", MRPAutoReleaseProduction)
            SetDataToCell("CM2", RepetitiveControl)
            SetDataToCell("CN2", PercentComplete_inPercentage)
            SetDataToCell("CO2", Balance_BuildOutDate)
            SetDataToCell("CP2", CommonMaterialPart)
            SetDataToCell("CQ2", RestrictCurrentStandardCost)
            SetDataToCell("CR2", MaterialIssueLocation)
            SetDataToCell("CS2", MattecPart)
            SetDataToCell("CT2", MattecMaterialType)
            SetDataToCell("CU2", StorageClass)
            SetDataToCell("CV2", VelocityCode)
            SetDataToCell("CW2", NumberofReceivingPack_perPallet)
            SetDataToCell("CX2", ReceivingPackComment)
            SetDataToCell("CY2", NumberofMasterPackperPallet)
            SetDataToCell("CZ2", MasterPackComment)

            '06_10_2010    RAGAVA   added for getting commas as per client
            SetDataToCell("DA2", _strVariable1)
            SetDataToCell("DB2", _strVariable2)
            SetDataToCell("DC2", _strVariable3)
            SetDataToCell("DD2", _strVariable4)
            SetDataToCell("DE2", _strVariable5)
            SetDataToCell("DF2", _strVariable6)

            SetDataToCell("DG2", _strVariable7)
            SetDataToCell("DH2", _strVariable8)
            SetDataToCell("DI2", _strVariable9)
            SetDataToCell("DJ2", _strVariable10)
            SetDataToCell("DK2", _strVariable11)
            SetDataToCell("DL2", _strVariable12)

            SetDataToCell("DM2", _strVariable13)
            SetDataToCell("DN2", _strVariable14)
            SetDataToCell("DO2", _strVariable15)
            SetDataToCell("DP2", _strVariable16)
            SetDataToCell("DQ2", _strVariable17)
            SetDataToCell("DR2", _strVariable18)
            'Till    Here
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
