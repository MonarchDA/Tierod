Imports Microsoft.Office.Interop
Public Class clsCMS_STKMM_STKMP

#Region "Private Variables"

    Private _strInternalPartNumber As String

    Private _strDefaultPlant As String

    Private _strPartType As String

    Private _strPartDescriptionLine1 As String

    Private _strPartDescriptionLine2 As String

    Private _strPartDescriptionLine3 As String

    Private _strGLSalesCode As String

    Private _strGLExpenseCode As String

    Private _strInventoryUnit As String

    Private _strMajorGroupCode As String

    Private _strMinorGroupCode As String

    Private _strMajorSalesCode As String

    Private _strMinorSalesCode As String

    Private _strCustomerCode As String

    Private _strCustomerPartNumber As String

    Private _strVendorCode As String

    Private _strVendorPartNumber As String

    Private _strCatalogId As String

    Private _strNetWeight As String

    Private _strNetWeightUOM As String

    Private _strShippingVolume As String

    Private _strShippingVolumeUOM As String

    Private _strFreightCodeNMFC As String

    Private _strLump As String

    Private _strConsumedbyLength As String

    Private _strCoil As String

    Private _strHazardous As String

    Private _strMaterialTemplateCode As String

    Private _strUserVerificationTemplateCode As String

    Private _strShelfLifeinDays As String

    Private _strInventoryClass As String

    Private _strHarmonizationCode As String

    Private _strStyleCode As String

    Private _strSizePackageCode As String

    Private _strColorCode As String

    Private _strAssemblyCode As String

    Private _strSubassemblyCode As String

    Private _strUnitDefaults As String

    Private _strApplyGlobalDiscount As String

    Private _strApplyWholeOrderVolDisc As String

    Private _strApplyWholeOrderPrcDisc As String

    Private _strFIFOTracking As String

    Private _strAntiDumpingTracking As String

    Private _strDumpingSubjectIndicator As String

    Private _strServiceChargePart As String

    Private _strOneTimePart As String

    Private _strSeparate_PO_AP_ReceiptperLump As String

    Private _strCommodityCategoryCode As String

    Private _objExcelSheet As Excel.Worksheet

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

    Public ReadOnly Property DefaultPlant() As String
        Get
            Return _strDefaultPlant
        End Get
        'Set(ByVal value As String)
        '    _strDefaultPlant = value
        'End Set
    End Property

    Public Property PartType() As String
        Get
            Return _strPartType
        End Get
        Set(ByVal value As String)
            _strPartType = value
        End Set
    End Property

    Public Property PartDescriptionLine1() As String
        Get
            Return _strPartDescriptionLine1
        End Get
        Set(ByVal value As String)
            _strPartDescriptionLine1 = value
        End Set
    End Property

    Public Property PartDescriptionLine2() As String
        Get
            Return _strPartDescriptionLine2
        End Get
        Set(ByVal value As String)
            _strPartDescriptionLine2 = value
        End Set
    End Property

    Public ReadOnly Property PartDescriptionLine3() As String
        Get
            Return _strPartDescriptionLine3
        End Get
        'Set(ByVal value As String)
        '    _strPartDescriptionLine3 = value
        'End Set
    End Property

    Public ReadOnly Property GLSalesCode() As String
        Get
            Return _strGLSalesCode
        End Get
        'Set(ByVal value As String)
        '    _strGLSalesCode = value
        'End Set
    End Property

    Public Property GLExpenseCode() As String
        Get
            Return _strGLExpenseCode
        End Get
        Set(ByVal value As String)
            _strGLExpenseCode = value
        End Set
    End Property

    Public ReadOnly Property InventoryUnit() As String
        Get
            Return _strInventoryUnit
        End Get
        'Set(ByVal value As String)
        '    _strInventoryUnit = value
        'End Set
    End Property

    Public Property MajorGroupCode() As String
        Get
            Return _strMajorGroupCode
        End Get
        Set(ByVal value As String)
            _strMajorGroupCode = value
        End Set
    End Property

    Public Property MinorGroupCode() As String
        Get
            Return _strMinorGroupCode
        End Get
        Set(ByVal value As String)
            _strMinorGroupCode = value
        End Set
    End Property

    Public Property MajorSalesCode() As String
        Get
            Return _strMajorSalesCode
        End Get
        Set(ByVal value As String)
            _strMajorSalesCode = value
        End Set
    End Property

    Public Property MinorSalesCode() As String
        Get
            Return _strMinorSalesCode
        End Get
        Set(ByVal value As String)
            _strMinorSalesCode = value
        End Set
    End Property

    Public ReadOnly Property CustomerCode() As String
        Get
            Return _strCustomerCode
        End Get
        'Set(ByVal value As String)
        '    _strCustomerCode = value
        'End Set
    End Property

    Public Property CustomerPartNumber() As String
        Get
            Return _strCustomerPartNumber
        End Get
        Set(ByVal value As String)
            _strCustomerPartNumber = value
        End Set
    End Property

    Public ReadOnly Property VendorCode() As String
        Get
            Return _strVendorCode
        End Get
    End Property

    Public ReadOnly Property VendorPartNumber() As String
        Get
            Return _strVendorPartNumber
        End Get
        'Set(ByVal value As String)
        '    _strVendorPartNumber = value
        'End Set
    End Property

    Public Property CatalogId() As String
        Get
            Return _strCatalogId
        End Get
        Set(ByVal value As String)
            _strCatalogId = value
        End Set
    End Property

    Public Property NetWeight() As String
        Get
            Return _strNetWeight
        End Get
        Set(ByVal value As String)
            _strNetWeight = value
        End Set
    End Property

    Public ReadOnly Property NetWeightUOM() As String
        Get
            Return _strNetWeightUOM
        End Get
        'Set(ByVal value As String)
        '    _strNetWeightUOM = value
        'End Set
    End Property

    Public ReadOnly Property ShippingVolume() As String
        Get
            Return _strShippingVolume
        End Get
        'Set(ByVal value As String)
        '    _strShippingVolume = value
        'End Set
    End Property

    Public ReadOnly Property ShippingVolumeUOM() As String
        Get
            Return _strShippingVolumeUOM
        End Get
        'Set(ByVal value As String)
        '    _strShippingVolumeUOM = value
        'End Set
    End Property

    Public ReadOnly Property FreightCodeNMFC() As String
        Get
            Return _strFreightCodeNMFC
        End Get
        'Set(ByVal value As String)
        '    _strFreightCodeNMFC = value
        'End Set
    End Property

    Public ReadOnly Property Lump() As String
        Get
            Return _strLump
        End Get
        'Set(ByVal value As String)
        '    _strLump = value
        'End Set
    End Property

    Public ReadOnly Property ConsumedbyLength() As String
        Get
            Return _strConsumedbyLength
        End Get
        'Set(ByVal value As String)
        '    _strConsumedbyLength = value
        'End Set
    End Property

    Public ReadOnly Property Coil() As String
        Get
            Return _strCoil
        End Get
        'Set(ByVal value As String)
        '    _strCoil = value
        'End Set
    End Property

    Public ReadOnly Property Hazardous() As String
        Get
            Return _strHazardous
        End Get
        'Set(ByVal value As String)
        '    _strHazardous = value
        'End Set
    End Property

    Public ReadOnly Property MaterialTemplateCode() As String
        Get
            Return _strMaterialTemplateCode
        End Get
        'Set(ByVal value As String)
        '    _strMaterialTemplateCode = value
        'End Set
    End Property

    Public Property UserVerificationTemplateCode() As String
        Get
            Return _strUserVerificationTemplateCode
        End Get
        Set(ByVal value As String)
            _strUserVerificationTemplateCode = value
        End Set
    End Property

    Public ReadOnly Property ShelfLifeinDays() As String
        Get
            Return _strShelfLifeinDays
        End Get
        'Set(ByVal value As String)
        '    _strShelfLifeinDays = value
        'End Set
    End Property

    Public ReadOnly Property InventoryClass() As String
        Get
            Return _strInventoryClass
        End Get
        'Set(ByVal value As String)
        '    _strInventoryClass = value
        'End Set
    End Property

    Public Property HarmonizationCode() As String
        Get
            Return _strHarmonizationCode
        End Get
        Set(ByVal value As String)
            _strHarmonizationCode = value
        End Set
    End Property

    Public ReadOnly Property StyleCode() As String
        Get
            Return _strStyleCode
        End Get
        'Set(ByVal value As String)
        '    _strStyleCode = value
        'End Set
    End Property

    Public ReadOnly Property SizePackageCode() As String
        Get
            Return _strSizePackageCode
        End Get
        'Set(ByVal value As String)
        '    _strSizePackageCode = value
        'End Set
    End Property

    Public ReadOnly Property ColorCode() As String
        Get
            Return _strColorCode
        End Get
        'Set(ByVal value As String)
        '    _strColorCode = value
        'End Set
    End Property

    Public ReadOnly Property AssemblyCode() As String
        Get
            Return _strAssemblyCode
        End Get
        'Set(ByVal value As String)
        '    _strAssemblyCode = value
        'End Set
    End Property

    Public ReadOnly Property SubassemblyCode() As String
        Get
            Return _strSubassemblyCode
        End Get
        'Set(ByVal value As String)
        '    _strSubassemblyCode = value
        'End Set
    End Property

    Public ReadOnly Property UnitDefaults() As String
        Get
            Return _strUnitDefaults
        End Get
        'Set(ByVal value As String)
        '    _strUnitDefaults = value
        'End Set
    End Property

    Public ReadOnly Property ApplyGlobalDiscount() As String
        Get
            Return _strApplyGlobalDiscount
        End Get
        'Set(ByVal value As String)
        '    _strApplyGlobalDiscount = value
        'End Set
    End Property

    Public ReadOnly Property ApplyWholeOrderVolDisc() As String
        Get
            Return _strApplyWholeOrderVolDisc
        End Get
        'Set(ByVal value As String)
        '    _strApplyWholeOrderVolDisc = value
        'End Set
    End Property

    Public ReadOnly Property ApplyWholeOrderPrcDisc() As String
        Get
            Return _strApplyWholeOrderPrcDisc
        End Get
        'Set(ByVal value As String)
        '    _strApplyWholeOrderPrcDisc = value
        'End Set
    End Property

    Public ReadOnly Property FIFOTracking() As String
        Get
            Return _strFIFOTracking
        End Get
        'Set(ByVal value As String)
        '    _strFIFOTracking = value
        'End Set
    End Property

    Public Property AntiDumpingTracking() As String
        Get
            Return _strAntiDumpingTracking
        End Get
        Set(ByVal value As String)
            _strAntiDumpingTracking = value
        End Set
    End Property

    Public Property DumpingSubjectIndicator() As String
        Get
            Return _strDumpingSubjectIndicator
        End Get
        Set(ByVal value As String)
            _strDumpingSubjectIndicator = value
        End Set
    End Property

    Public Property ServiceChargePart() As String
        Get
            Return _strServiceChargePart
        End Get
        Set(ByVal value As String)
            _strServiceChargePart = value
        End Set
    End Property

    Public ReadOnly Property OneTimePart() As String
        Get
            Return _strOneTimePart
        End Get
        'Set(ByVal value As String)
        '    _strOneTimePart = value
        'End Set
    End Property

    Public ReadOnly Property Separate_PO_AP_ReceiptperLump() As String
        Get
            Return _strSeparate_PO_AP_ReceiptperLump
        End Get
        'Set(ByVal value As String)
        '    _strSeparate_PO_AP_ReceiptperLump = value
        'End Set
    End Property

    Public ReadOnly Property CommodityCategoryCode() As String
        Get
            Return _strCommodityCategoryCode
        End Get
    End Property

#End Region

#Region "Sub Procedures"

    Public Sub SetCommenPropertyValue()
        _strDefaultPlant = ""   '13_09_2010   RAGAVA    '"C01"
        _strPartDescriptionLine3 = ""
        _strGLSalesCode = "TRC"
        _strInventoryUnit = "EA"
        _strCustomerCode = ""
        _strVendorCode = ""
        _strVendorPartNumber = ""
        _strNetWeightUOM = "LB"
        _strShippingVolume = ""
        _strShippingVolumeUOM = ""
        _strFreightCodeNMFC = ""
        _strLump = 2
        _strConsumedbyLength = 2
        _strCoil = 2
        _strHazardous = 2
        _strMaterialTemplateCode = ""
        _strShelfLifeinDays = ""
        _strInventoryClass = ""
        _strStyleCode = ""
        _strSizePackageCode = ""
        _strColorCode = ""
        _strAssemblyCode = ""
        _strSubassemblyCode = ""
        _strUnitDefaults = 1
        _strApplyGlobalDiscount = 1
        _strApplyWholeOrderVolDisc = 2
        _strApplyWholeOrderPrcDisc = 2
        _strFIFOTracking = 2
        _strOneTimePart = 2
        _strSeparate_PO_AP_ReceiptperLump = 2
        _strCommodityCategoryCode = ""
    End Sub

    Public Sub SetDataToExcel(ByVal oExcelSheet As Excel.Worksheet)
        Try
            _objExcelSheet = oExcelSheet
            SetDataToCell("A2", InternalPartNumber)
            SetDataToCell("B2", DefaultPlant)
            SetDataToCell("C2", PartType)
            SetDataToCell("D2", PartDescriptionLine1)
            SetDataToCell("E2", PartDescriptionLine2)
            SetDataToCell("F2", PartDescriptionLine3)
            SetDataToCell("G2", GLSalesCode)
            SetDataToCell("H2", GLExpenseCode)
            SetDataToCell("I2", InventoryUnit)
            SetDataToCell("J2", MajorGroupCode)
            SetDataToCell("K2", MinorGroupCode)
            SetDataToCell("L2", MajorSalesCode)
            SetDataToCell("M2", MinorSalesCode)
            SetDataToCell("N2", CustomerCode)
            SetDataToCell("O2", CustomerPartNumber)
            SetDataToCell("P2", VendorCode)
            SetDataToCell("Q2", VendorPartNumber)
            SetDataToCell("R2", CatalogId)
            SetDataToCell("S2", NetWeight)
            SetDataToCell("T2", NetWeightUOM)
            SetDataToCell("U2", ShippingVolume)
            SetDataToCell("V2", ShippingVolumeUOM)
            SetDataToCell("W2", FreightCodeNMFC)
            SetDataToCell("X2", Lump)
            SetDataToCell("Y2", ConsumedbyLength)
            SetDataToCell("Z2", Coil)
            SetDataToCell("AA2", Hazardous)
            SetDataToCell("AB2", MaterialTemplateCode)
            SetDataToCell("AC2", UserVerificationTemplateCode)
            SetDataToCell("AD2", ShelfLifeinDays)
            SetDataToCell("AE2", InventoryClass)
            SetDataToCell("AF2", HarmonizationCode)
            SetDataToCell("AG2", StyleCode)
            SetDataToCell("AH2", SizePackageCode)
            SetDataToCell("AI2", ColorCode)
            SetDataToCell("AJ2", AssemblyCode)
            SetDataToCell("AK2", SubassemblyCode)
            SetDataToCell("AL2", UnitDefaults)
            SetDataToCell("AM2", ApplyGlobalDiscount)
            SetDataToCell("AN2", ApplyWholeOrderVolDisc)
            SetDataToCell("AO2", ApplyWholeOrderPrcDisc)
            SetDataToCell("AP2", FIFOTracking)
            SetDataToCell("AQ2", AntiDumpingTracking)
            SetDataToCell("AR2", DumpingSubjectIndicator)
            SetDataToCell("AS2", ServiceChargePart)
            SetDataToCell("AT2", OneTimePart)
            SetDataToCell("AU2", Separate_PO_AP_ReceiptperLump)
            SetDataToCell("AV2", CommodityCategoryCode)
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
