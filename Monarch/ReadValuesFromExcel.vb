Public Class ReadValuesFromExcel         'SUGANDHI

#Region "Variable"

    Private oCustomername As String
    Private oType As String = "Tie Rod Cylinder Assembly"
    Private oCustomerPortCode As String
    Private oRephasingPortPosition As String
    Private oSeries As String
    Private oStyle As String
    Private oBore As Double
    Private oStrokeLength As Double
    Private oRodAdder As Integer
    Private oStopTube As Boolean
    Private oClevisCapPinHole As String
    Private oRodClevisPinHole As String
    Private oStandardRunQty As Integer
    Private oRodMaterials As String
    Private oRodDiameter As Double
    Private oRodDeratedPressureAtmaximumExtension As Double = 0
    Private oPortOrientationForClevisCap As String
    Private oPortOrientationForRodCap As String
    Private oClevisCapPort As String
    Private oRodCapPort As String
    Private oBreathersAtClevisCapPort As Boolean
    Private oBreathersAtRodCapPort As Boolean
    Private oStrokeControl As Boolean
    Private oStrokeControlStages As Integer
    Private oClevisCapPins As Boolean
    Private oRodClevisPins As Boolean
    Private oPinMaterial As String
    Private oClevisCapPinClips As String
    Private oThreadProtected As String
    Private oRodSealPackage As String
    Private oRodClevisCheck As Boolean
    Private oRodEndThreadSize As Double
    Private oRodClevisPinClips As String
    Private oPistonStealPackage As String
    Private oPaint As String
    Private oGenerationType As String
    Private oRetractedLength As Double
    Private oExtendedLength As Double
    Private oPinSizeDetails As Double
    Private oRodClevis As String
    Private oRodWiper As String
    Private oStopTubelength As Double
    Private oNutSize As Double
    Private oNOLABELONCYLINDER As Boolean
    Private oBagRequired As Boolean

#Region "UnUsedCode Variables"
    'Private oTudeSeal As String
    'Private oPackaging As String
    'Private oPinsAreInLineWithPort As Integer
    'Private oRetractedLength1 As Integer
    'Private oExtendedLength1 As Integer
    'Private oRodDiameter1 As Integer
    'Private oPorts As Integer
    'Private oPercentAirTest As Integer
    'Private oPercentOilTest As Integer
    'Private oRephaseOnExtension As Integer
    'Private oRephaseOnRetraction As Integer
    'Private oInstallStrokeControl As Integer
    'Private oStampCustomerPartAndDateCodeOnTube As Integer
    'Private oStampCustomerPartOnTube As Integer
    'Private oStampCountryOfOriginOnTube As Integer
    'Private oRodMaterialsNitroSteel As Integer
    'Private oInstallSteelPlugsInAllPorts As Integer
    'Private oInstallHardenedBushingsAndRodClevisEnd As Integer
    'Private oInstallHardenedBushingsAndClevisCapEnd As Integer
    'Private oAssemblyStopTubeToCylinder As Integer
    'Private oMaskPerBOMAndSOP As Integer
    'Private oMaskBushingsBeforePainting As Integer
    'Private oMaskExposedThreadsAfterWashing As Integer
    'Private oMaskPinHoles As Integer
    'Private oPrime As Integer
    'Private oPaint1 As Integer
    'Private oAffixLabelPerSOP As Integer
    'Private oAffixLabelToBag As Integer
    'Private oIncludePinKitPerBOM As Integer
    'Private oPackCylinderInPlasticBag As Integer
    'Private oPackagePerSOP As Integer
    'Private oCompleteModelOrOnlyCosting As Integer

#End Region

#End Region
#Region "Properties"

    Public Property CustomerName() As String
        Get
            Return oCustomername
        End Get
        Set(ByVal value As String)
            oCustomername = value
        End Set
    End Property

    Public ReadOnly Property Type() As String
        Get
            Return oType
        End Get
    End Property

    Public Property CustomerPortCode() As String
        Get
            Return oCustomerPortCode
        End Get
        Set(ByVal value As String)
            oCustomerPortCode = value
        End Set
    End Property

    Public Property Series() As String
        Get
            Return oSeries
        End Get
        Set(ByVal value As String)
            oSeries = value
        End Set
    End Property

    Public Property RephasingPortPosition() As String
        Get
            Return oRephasingPortPosition
        End Get
        Set(ByVal value As String)
            oRephasingPortPosition = value
        End Set
    End Property

    Public Property Style() As String
        Get
            Return oStyle
        End Get
        Set(ByVal value As String)
            oStyle = value
        End Set
    End Property

    Public Property Bore() As Double
        Get
            Return oBore
        End Get
        Set(ByVal value As Double)
            oBore = value
        End Set
    End Property

    Public Property StrokeLength() As Double
        Get
            Return oStrokeLength
        End Get
        Set(ByVal value As Double)
            oStrokeLength = value
        End Set
    End Property

    Public Property RodAdder() As Integer
        Get
            Return oRodAdder
        End Get
        Set(ByVal value As Integer)
            oRodAdder = value
        End Set
    End Property

    Public Property StopTube() As Boolean
        Get
            Return oStopTube
        End Get
        Set(ByVal value As Boolean)
            oStopTube = value
        End Set
    End Property

    Public Property ClevisCapPinHole() As String
        Get
            Return oClevisCapPinHole
        End Get
        Set(ByVal value As String)
            oClevisCapPinHole = value
        End Set
    End Property

    Public Property RodClevisPinHole() As String
        Get
            Return oRodClevisPinHole
        End Get
        Set(ByVal value As String)
            oRodClevisPinHole = value
        End Set
    End Property

    Public Property StandardRunQty() As Integer
        Get
            Return oStandardRunQty
        End Get
        Set(ByVal value As Integer)
            oStandardRunQty = value
        End Set
    End Property

    Public Property RodMaterials() As String
        Get
            Return oRodMaterials
        End Get
        Set(ByVal value As String)
            oRodMaterials = value
        End Set
    End Property

    Public Property RodDiameter() As Double
        Get
            Return oRodDiameter
        End Get
        Set(ByVal value As Double)
            oRodDiameter = value
        End Set
    End Property

    Public Property RodDeratedPressureAtmaximumExtension() As Double
        Get
            Return oRodDeratedPressureAtmaximumExtension
        End Get
        Set(ByVal value As Double)
            oRodDeratedPressureAtmaximumExtension = value
        End Set
    End Property

    Public Property PortOrientationForClevisCap() As String
        Get
            Return oPortOrientationForClevisCap
        End Get
        Set(ByVal value As String)
            oPortOrientationForClevisCap = value
        End Set
    End Property

    Public Property PortOrientationForRodCap() As String
        Get
            Return oPortOrientationForRodCap
        End Get
        Set(ByVal value As String)
            oPortOrientationForRodCap = value
        End Set
    End Property

    Public Property ClevisCapPort() As String
        Get
            Return oClevisCapPort
        End Get
        Set(ByVal value As String)
            oClevisCapPort = value
        End Set
    End Property

    Public Property RodCapPort() As String
        Get
            Return oRodCapPort
        End Get
        Set(ByVal value As String)
            oRodCapPort = value
        End Set
    End Property

    Public Property StrokeControl() As Boolean
        Get
            Return oStrokeControl
        End Get
        Set(ByVal value As Boolean)
            oStrokeControl = value
        End Set
    End Property

    Public Property StrokeControlStages() As Integer
        Get
            Return oStrokeControlStages
        End Get
        Set(ByVal value As Integer)
            oStrokeControlStages = value
        End Set
    End Property

    Public Property ClevisCapPins() As Boolean
        Get
            Return oClevisCapPins
        End Get
        Set(ByVal value As Boolean)
            oClevisCapPins = value
        End Set
    End Property

    Public Property RodClevisPins() As Boolean
        Get
            Return oRodClevisPins
        End Get
        Set(ByVal value As Boolean)
            oRodClevisPins = value
        End Set
    End Property

    Public Property PinMaterial() As String
        Get
            Return oPinMaterial
        End Get
        Set(ByVal value As String)
            oPinMaterial = value
        End Set
    End Property

    Public Property ClevisCapPinClips() As String
        Get
            Return oClevisCapPinClips
        End Get
        Set(ByVal value As String)
            oClevisCapPinClips = value
        End Set
    End Property

    Public Property ThreadProtected() As String
        Get
            Return oThreadProtected
        End Get
        Set(ByVal value As String)
            oThreadProtected = value
        End Set
    End Property

    Public Property RodSealPackage() As String
        Get
            Return oRodSealPackage
        End Get
        Set(ByVal value As String)
            oRodSealPackage = value
        End Set
    End Property

    Public Property RodClevisCheck() As Boolean
        Get
            Return oRodClevisCheck
        End Get
        Set(ByVal value As Boolean)
            oRodClevisCheck = value
        End Set
    End Property

    Public Property RodEndThreadSize() As Double
        Get
            Return oRodEndThreadSize
        End Get
        Set(ByVal value As Double)
            oRodEndThreadSize = value
        End Set
    End Property

    Public Property RodClevisPinClips() As String
        Get
            Return oRodClevisPinClips
        End Get
        Set(ByVal value As String)
            oRodClevisPinClips = value
        End Set
    End Property

    Public Property PistonStealPackage() As String
        Get
            Return oPistonStealPackage
        End Get
        Set(ByVal value As String)
            oPistonStealPackage = value
        End Set
    End Property

    Public Property Paint() As String
        Get
            Return oPaint
        End Get
        Set(ByVal value As String)
            oPaint = value
        End Set
    End Property

    Public Property GenerationType() As String
        Get
            Return oGenerationType
        End Get
        Set(ByVal value As String)
            oGenerationType = value
        End Set
    End Property

    Public Property PinSizeDetails() As String
        Get
            Return oPinSizeDetails
        End Get
        Set(ByVal value As String)
            oPinSizeDetails = value
        End Set
    End Property

    Public Property RodClevis() As String
        Get
            Return oRodClevis
        End Get
        Set(ByVal value As String)
            oRodClevis = value
        End Set
    End Property

    Public Property RodWiper() As String
        Get
            Return oRodWiper
        End Get
        Set(ByVal value As String)
            oRodWiper = value
        End Set
    End Property

    Public Property RetractedLength() As Double
        Get
            Return oRetractedLength
        End Get
        Set(ByVal value As Double)
            oRetractedLength = value
        End Set
    End Property

    Public Property ExtendedLength() As Double
        Get
            Return oExtendedLength
        End Get
        Set(ByVal value As Double)
            oExtendedLength = value
        End Set
    End Property

    Public Property StopTubeLength() As Double
        Get
            Return oStopTubelength
        End Get
        Set(ByVal value As Double)
            oStopTubelength = value
        End Set
    End Property

    Public Property NutSize() As Double
        Get
            Return oNutSize
        End Get
        Set(ByVal value As Double)
            oNutSize = value
        End Set
    End Property

    Public Property NOLABELONCYLINDER() As Boolean
        Get
            Return oNOLABELONCYLINDER
        End Get
        Set(ByVal value As Boolean)
            oNOLABELONCYLINDER = value
        End Set
    End Property

    Public Property BagRequired() As Boolean
        Get
            Return oBagRequired
        End Get
        Set(ByVal value As Boolean)
            oBagRequired = value
        End Set
    End Property

#Region "UnUsedCode Properties"

    'Public Property PinsAreInLineWithPort() As Integer
    '    Get
    '        Return oPinsAreInLineWithPort
    '    End Get
    '    Set(ByVal value As Integer)
    '        oPinsAreInLineWithPort = value
    '    End Set
    'End Property

    'Public Property RetractedLengthTieRod3() As Integer
    '    Get
    '        Return oRetractedLength1
    '    End Get
    '    Set(ByVal value As Integer)
    '        oRetractedLength1 = value
    '    End Set
    'End Property

    'Public Property ExtendedLengthTieRod3() As Integer
    '    Get
    '        Return oExtendedLength1
    '    End Get
    '    Set(ByVal value As Integer)
    '        oExtendedLength1 = value
    '    End Set
    'End Property

    'Public Property RodDiameterTieRod3() As Integer
    '    Get
    '        Return oRodDiameter1
    '    End Get
    '    Set(ByVal value As Integer)
    '        oRodDiameter1 = value
    '    End Set
    'End Property

    'Public Property Ports() As Integer
    '    Get
    '        Return oPorts
    '    End Get
    '    Set(ByVal value As Integer)
    '        oPorts = value
    '    End Set
    'End Property

    'Public Property PercentAirTest() As Integer
    '    Get
    '        Return oPercentAirTest
    '    End Get
    '    Set(ByVal value As Integer)
    '        oPercentAirTest = value
    '    End Set
    'End Property

    'Public Property PercentOilTest() As Integer
    '    Get
    '        Return oPercentOilTest
    '    End Get
    '    Set(ByVal value As Integer)
    '        oPercentOilTest = value
    '    End Set
    'End Property

    'Public Property RephaseOnExtension() As Integer
    '    Get
    '        Return oRephaseOnExtension
    '    End Get
    '    Set(ByVal value As Integer)
    '        oRephaseOnExtension = value
    '    End Set
    'End Property

    'Public Property RephaseOnRetraction() As Integer
    '    Get
    '        Return oRephaseOnRetraction
    '    End Get
    '    Set(ByVal value As Integer)
    '        oRephaseOnRetraction = value
    '    End Set
    'End Property

    'Public Property InstallStrokeControl() As Integer
    '    Get
    '        Return oInstallStrokeControl
    '    End Get
    '    Set(ByVal value As Integer)
    '        oInstallStrokeControl = value
    '    End Set
    'End Property

    'Public Property StampCustomerPartAndDateCodeOnTube() As Integer
    '    Get
    '        Return oStampCustomerPartAndDateCodeOnTube
    '    End Get
    '    Set(ByVal value As Integer)
    '        oStampCustomerPartAndDateCodeOnTube = value
    '    End Set
    'End Property

    'Public Property StampCustomerPartOnTube() As Integer
    '    Get
    '        Return oStampCustomerPartOnTube
    '    End Get
    '    Set(ByVal value As Integer)
    '        oStampCustomerPartOnTube = value
    '    End Set
    'End Property

    'Public Property StampCountryOfOriginOnTube() As Integer
    '    Get
    '        Return oStampCountryOfOriginOnTube
    '    End Get
    '    Set(ByVal value As Integer)
    '        oStampCountryOfOriginOnTube = value
    '    End Set
    'End Property

    'Public Property RodMaterialsNitroSteel() As Integer
    '    Get
    '        Return oRodMaterialsNitroSteel
    '    End Get
    '    Set(ByVal value As Integer)
    '        oRodMaterialsNitroSteel = value
    '    End Set
    'End Property

    'Public Property InstallSteelPlugsInAllPorts() As Integer
    '    Get
    '        Return oInstallSteelPlugsInAllPorts
    '    End Get
    '    Set(ByVal value As Integer)
    '        oInstallSteelPlugsInAllPorts = value
    '    End Set
    'End Property

    'Public Property InstallHardenedBushingsAndRodClevisEnd() As Integer
    '    Get
    '        Return oInstallHardenedBushingsAndRodClevisEnd
    '    End Get
    '    Set(ByVal value As Integer)
    '        oInstallHardenedBushingsAndRodClevisEnd = value
    '    End Set
    'End Property

    'Public Property InstallHardenedBushingsAndClevisCapEnd() As Integer
    '    Get
    '        Return oInstallHardenedBushingsAndClevisCapEnd
    '    End Get
    '    Set(ByVal value As Integer)
    '        oInstallHardenedBushingsAndClevisCapEnd = value
    '    End Set
    'End Property

    'Public Property AssemblyStopTubeToCylinder() As Integer
    '    Get
    '        Return oAssemblyStopTubeToCylinder
    '    End Get
    '    Set(ByVal value As Integer)
    '        oAssemblyStopTubeToCylinder = value
    '    End Set
    'End Property

    'Public Property MaskPerBOMAndSOP() As Integer
    '    Get
    '        Return oMaskPerBOMAndSOP
    '    End Get
    '    Set(ByVal value As Integer)
    '        oMaskPerBOMAndSOP = value
    '    End Set
    'End Property

    'Public Property MaskBushingsBeforePainting() As Integer
    '    Get
    '        Return oMaskBushingsBeforePainting
    '    End Get
    '    Set(ByVal value As Integer)
    '        oMaskBushingsBeforePainting = value
    '    End Set
    'End Property

    'Public Property MaskExposedThreadsAfterWashing() As Integer
    '    Get
    '        Return oMaskExposedThreadsAfterWashing
    '    End Get
    '    Set(ByVal value As Integer)
    '        oMaskExposedThreadsAfterWashing = value
    '    End Set
    'End Property

    'Public Property MaskPinHoles() As Integer
    '    Get
    '        Return oMaskPinHoles
    '    End Get
    '    Set(ByVal value As Integer)
    '        oMaskPinHoles = value
    '    End Set
    'End Property

    'Public Property Prime() As Integer
    '    Get
    '        Return oPrime
    '    End Get
    '    Set(ByVal value As Integer)
    '        oPrime = value
    '    End Set
    'End Property

    'Public Property PaintTieRod3() As Integer
    '    Get
    '        Return oPaint1
    '    End Get
    '    Set(ByVal value As Integer)
    '        oPaint1 = value
    '    End Set
    'End Property

    'Public Property AffixLabelPerSOP() As Integer
    '    Get
    '        Return oAffixLabelPerSOP
    '    End Get
    '    Set(ByVal value As Integer)
    '        oAffixLabelPerSOP = value
    '    End Set
    'End Property

    'Public Property IncludePinKitPerBOM() As Integer
    '    Get
    '        Return oIncludePinKitPerBOM
    '    End Get
    '    Set(ByVal value As Integer)
    '        oIncludePinKitPerBOM = value
    '    End Set
    'End Property

    'Public Property PackCylinderInPlasticBag() As Integer
    '    Get
    '        Return oPackCylinderInPlasticBag
    '    End Get
    '    Set(ByVal value As Integer)
    '        oPackCylinderInPlasticBag = value
    '    End Set
    'End Property

    'Public Property PackagePerSOP() As Integer
    '    Get
    '        Return oPackagePerSOP
    '    End Get
    '    Set(ByVal value As Integer)
    '        oPackagePerSOP = value
    '    End Set
    'End Property

    'Public Property AffixLabelToBag() As Integer
    '    Get
    '        Return oAffixLabelToBag
    '    End Get
    '    Set(ByVal value As Integer)
    '        oAffixLabelToBag = value
    '    End Set
    'End Property

    'Public Property CompleteModelOrOnlyCosting() As Integer
    '    Get
    '        Return oCompleteModelOrOnlyCosting
    '    End Get
    '    Set(ByVal value As Integer)
    '        oCompleteModelOrOnlyCosting = value
    '    End Set
    'End Property


#End Region
#End Region
End Class
