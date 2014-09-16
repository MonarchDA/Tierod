Imports MonarchDatabaseLayer
Imports System.Diagnostics
Imports System.io
Imports MonarchAPILayer
Imports MonarchFunctionalLayer
Imports System.Threading

Public Module GeneralModule

#Region "Variables"

    Public strRodClevis_Class As String = String.Empty   '19_11_2012    RAGAVA
    Public dblDerateWorkingPressure As Double = 0        '19_11_2012    RAGAVA
    '16_06_2011  RAGAVA
    Public _strPinCodeBE As String = String.Empty
    Public _strPinCodeRE As String = String.Empty
    Public strBaseEndKitCode As String = String.Empty
    Public strRodEndKitCode As String = String.Empty
    Public blnInstallPinsandClips As Boolean = False
    'Public blnInstallPinsandClips_Checked As Boolean = False
    'TILL  HERE

    'Public strCodeNumber_BeforeApplicationStart As String = String.Empty           '21_01_2011         RAGAVA

    Public TubeMaterial1 As String = String.Empty    '08_10_2010   RAGAVA
    Public TubeMaterial2 As String = String.Empty     '08_10_2010   RAGAVA

    Public strNewTableDrawingNumber As String = String.Empty          '02_12_2009   RAGAVA
    Public strNewTieRodTableDrawingNumber As String = String.Empty         '02_12_2009   RAGAVA
    Public strNewTubeTableDrawingNumber As String = String.Empty          '02_12_2009   RAGAVA


    Public dblRodAdderModified As Double      '26_10_2009   ragava
    Public dblStrokeLengthModified As Double      '26_10_2009   ragava
    Public strStyleModified As String = String.Empty      '26_10_2009   ragava

    Declare Sub Sleep Lib "kernel32" (ByVal milliSec As Integer)
    Dim fso As New Scripting.FileSystemObject
    Private _strfilePath As String = String.Empty
    Public htMainParams As New Hashtable
    Public htGUIInputParameters As New Hashtable
    Public _ofrmContractDetails As New frmContractDetails
    Public _ofrmTieRod1 As frmTieRod1
    Public _ofrmTieRod2 As frmTieRod2
    Public _ofrmTieRod3 As frmTieRod3
    Public oclsimgcapture As New clsimgcapture
    Public _ofrmMonarch As New frmMonarch
    Public _ofrmMdiMonarch As New mdiMonarch
    Public oDataClass As New DataClass
    Private _aFormNavigationOrder As ArrayList
    Private _lFormName As Form
    Private _dblRodDiameter As Double
    Private _dblBoreDiameter As Double
    Private _dblStrokeLength As Double
    Private _dblStopTubeLength As Double
    Private _dblRodAdder As Double
    Private _dblWorkingPressure As Double
    Private _dblTempWorkingPressure As Double
    Private _dblColumnLoad As Double
    Private _strCodeDesc As String = String.Empty
    Public htGuiReport As New Hashtable
    Private _dblPinSize As Double
    Private _strSavefiletopath As String = String.Empty
    '31_08_2009
    Private _dblRodStrokeDifference As Double
    Private _dblBoreStrokeDifference As Double
    Private _dblTieRodStrokeDifference As Double
    Private _strCylinderCodeNumber As String = String.Empty
    Private _oCurrentForm As Object
    '04_09_2009  ragava
    Public oCustomerListviewItem As ListViewItem
    Public oContractListviewItem As ListViewItem
    '04_09_2009  ragava   Till  Here
    Private str_SetCodeDesciption As String = String.Empty
    Private _strTieRodNutDescription As String = String.Empty
    Private _strTieRodNutDrawingNumber As String = String.Empty
    Private _strTieRodNutCodeNumber As String = String.Empty

    Private _strBoreDrawingNumber As String = String.Empty
    Private _strBoreCodeNumber As String = String.Empty
    Private _strBoreDescription As String = String.Empty

    Private _strRodDCodeNumber As String = String.Empty
    Private _strRodDDrawingNumber As String = String.Empty
    Private _strRodDDescription As String = String.Empty

    Private _strPistonCodeNumber As String = String.Empty
    Private _strPistonDrawingNumber As String = String.Empty
    Private _strPistonDescription As String = String.Empty

    Private _strPinsCodeNumber As String = String.Empty
    Private _strPinsDrawingNumber As String = String.Empty
    Private _strPinsDescription As String = String.Empty

    Private _strTieRodCodeNumber As String = String.Empty
    Private _strTieRodDrawingNumber As String = String.Empty
    Private _strTieRodDescription As String = String.Empty

    Private _strRodCapCodeNumber As String = String.Empty
    Private _strRodCaprawingNumber As String = String.Empty
    Private _strRodCapDescription As String = String.Empty

    Private _strClevisCapCodeNumber As String = String.Empty
    Private _strClevisCaprawingNumber As String = String.Empty
    Private _strClevisCapDescription As String = String.Empty

    Private _strRodClevisCodeNumber As String = String.Empty
    Private _strRodClevisDrawingNumber As String = String.Empty
    Private _strRodClevisDescription As String = String.Empty

    Private _strstrokeControlCodeNumber As String = String.Empty
    Private _strStrokeDrawingNumber As String = String.Empty
    Private _strStrokeDescription As String = String.Empty
    Private _prbMonarch As ProgressBar
    Private _oThreadProgressBarStepping As Threading.Thread
    Private _dblRecommendedStopTubeLength As Double
    Private _strColumnLoadDeratePressure As String = String.Empty
    Private _strRodClevisPinCodeNumber As String = String.Empty
    Private _strClevisCapPinCodeNumber As String = String.Empty
    Private _dblCylinderPullForce As Double
    Private _blnRevision As Boolean
    Private _blnRodEndThreadSizeNotAvailable As Boolean
    Private _aLgetRevisionTableData As ArrayList
    Private _strStopTubeCodeNumber As String = String.Empty
    Private _strTubeCodeNumber As String = String.Empty
    Private _strTieRodCodeNumber1 As String = String.Empty
    Private _strRodCodeNumber As String = String.Empty
    Private _blnApplicationStop As Boolean = False

    'Added by Sandeep on 04-04-10
    Private _strPistonCode As String = String.Empty

    Private _oClsCosting As New clsCosting

    Private _dblTubeLength As Double

    Private _dblRodLength As Double

    Private _IsNewBtnClicked As Boolean = False     'SUGANDHI


    'Public Property TubeLength() As Double
    '    Get
    '        Return _dblTubeLength
    '    End Get
    '    Set(ByVal value As Double)
    '        _dblTubeLength = value
    '    End Set
    'End Property

    'Public Property RodLength() As Double
    '    Get
    '        Return _dblRodLength
    '    End Get
    '    Set(ByVal value As Double)
    '        _dblRodLength = value
    '    End Set
    'End Property

    Public Property ObjClsCostingDetails() As clsCosting
        Get
            Return _oClsCosting
        End Get
        Set(ByVal value As clsCosting)
            _oClsCosting = value
        End Set
    End Property
    '*********************************

    'Added by Sandeep on 10-03-10
    Private _strRodMaterialForCosting As String = String.Empty

    Private _strSeriesForCosting As String = String.Empty

    Public Property RodMaterialForCosting() As String
        Get
            Return _strRodMaterialForCosting
        End Get
        Set(ByVal value As String)
            _strRodMaterialForCosting = value
        End Set
    End Property

    Public Property SeriesForCosting() As String
        Get
            Return _strSeriesForCosting
        End Get
        Set(ByVal value As String)
            _strSeriesForCosting = value
        End Set
    End Property
    '*********************************

    Public Property ApplicationStop() As Boolean
        Get
            Return _blnApplicationStop
        End Get
        Set(ByVal value As Boolean)
            _blnApplicationStop = value
        End Set
    End Property

    Public Property TieRodCodeNumber1() As String
        Get
            Return _strTieRodCodeNumber1
        End Get
        Set(ByVal value As String)
            _strTieRodCodeNumber1 = value
        End Set
    End Property

    Public Property StopTubeCodeNumber() As String
        Get
            Return _strStopTubeCodeNumber
        End Get
        Set(ByVal value As String)
            _strStopTubeCodeNumber = value
        End Set
    End Property

    Public Property TubeCodeNumber() As String
        Get
            Return _strTubeCodeNumber
        End Get
        Set(ByVal value As String)
            _strTubeCodeNumber = value
        End Set
    End Property

    Public Property RodCodeNumber() As String
        Get
            Return _strRodCodeNumber
        End Get
        Set(ByVal value As String)
            _strRodCodeNumber = value
        End Set
    End Property

    Public Property alGetRevisionTableData() As ArrayList
        Get
            Return _aLgetRevisionTableData
        End Get
        Set(ByVal value As ArrayList)
            _aLgetRevisionTableData = value
        End Set
    End Property

    Public Property blnRodEndThreadSizeNotAvailable() As Boolean
        Get
            Return _blnRodEndThreadSizeNotAvailable
        End Get
        Set(ByVal value As Boolean)
            _blnRodEndThreadSizeNotAvailable = value
        End Set
    End Property

    Public Property blnRevision() As Boolean
        Get
            Return _blnRevision
        End Get
        Set(ByVal value As Boolean)
            _blnRevision = value
        End Set
    End Property

    Public Property IsNewBtnClicked() As Boolean      'SUGANDHI

        Get
            Return _IsNewBtnClicked
        End Get
        Set(ByVal value As Boolean)
            _IsNewBtnClicked = value
        End Set

    End Property

    Public Property dblRecommendedStopTubeLength() As Double
        Get
            Return _dblRecommendedStopTubeLength
        End Get
        Set(ByVal value As Double)
            _dblRecommendedStopTubeLength = value
        End Set
    End Property

    Public Property dblCylinderPullForce() As Double
        Get
            Return _dblCylinderPullForce
        End Get
        Set(ByVal value As Double)
            _dblCylinderPullForce = value
        End Set
    End Property


    'TODO:Sunny 27-04-10 5pm

    Private _strCustomerName As String = String.Empty

    'Public Property CustomerName() As String
    '    Get
    '        Return _strCustomerName
    '    End Get
    '    Set(ByVal value As String)
    '        _strCustomerName = value
    '    End Set
    'End Property


    Private _strPistonThreadSize As String = String.Empty

    Public Property PistonThreadSize() As String
        Get
            Return _strPistonThreadSize
        End Get
        Set(ByVal value As String)
            _strPistonThreadSize = value
        End Set
    End Property


    Private _strTieRodSize As String = String.Empty

    Public Property TieRodSize() As String
        Get
            Return _strTieRodSize
        End Get
        Set(ByVal value As String)
            _strTieRodSize = value
        End Set
    End Property


    Private _dblStopTubeID As Double

    Public Property StopTubeID() As Double
        Get
            Return _dblStopTubeID
        End Get
        Set(ByVal value As Double)
            _dblStopTubeID = value
        End Set
    End Property


    Private _dblStopTubeOD As Double

    Public Property StopTubeOD() As Double
        Get
            Return _dblStopTubeOD
        End Get
        Set(ByVal value As Double)
            _dblStopTubeOD = value
        End Set
    End Property

    Private _blnIsStopTubeSelected As Boolean

    Public Property IsStopTubeSelected() As Boolean
        Get
            Return _blnIsStopTubeSelected
        End Get
        Set(ByVal value As Boolean)
            _blnIsStopTubeSelected = value
        End Set
    End Property

    '----------------------------

    'ANUP 26-10-2010 START 
    Private _blnIsReleaseCylinderChecked As Boolean
    Private _strIsNew_Revision_Released As String = String.Empty

    Public Property IsReleaseCylinderChecked() As Boolean
        Get
            Return _blnIsReleaseCylinderChecked
        End Get
        Set(ByVal value As Boolean)
            _blnIsReleaseCylinderChecked = value
        End Set
    End Property

    Public Property IsNew_Revision_Released() As String
        Get
            Return _strIsNew_Revision_Released
        End Get
        Set(ByVal value As String)
            _strIsNew_Revision_Released = value
        End Set
    End Property

    'ANUP 26-10-2010 TILL HERE

#End Region

#Region "Sunny 20-4-10"

    'TODO: Sunny 20-04-10 10am

    Private _strRodMaterialCode_Costing As String = String.Empty

    Private _strTubeMaterialCode_Costing As String = String.Empty

    Private _strThreadProtected As String = String.Empty

    Private _dblSRQ As Double

    Public Property RodMaterialCode_Costing() As String
        Get
            Return _strRodMaterialCode_Costing
        End Get
        Set(ByVal value As String)
            _strRodMaterialCode_Costing = value
        End Set
    End Property

    Public Property TubeMaterialCode_Costing() As String
        Get
            Return _strTubeMaterialCode_Costing
        End Get
        Set(ByVal value As String)
            _strTubeMaterialCode_Costing = value
        End Set
    End Property

    Public Property ThreadProtected() As String
        Get
            Return _strThreadProtected
        End Get
        Set(ByVal value As String)
            _strThreadProtected = value
        End Set
    End Property

    Public Property SRQ() As Double
        Get
            Return _dblSRQ
        End Get
        Set(ByVal value As Double)
            _dblSRQ = value
        End Set
    End Property

#End Region

#Region "Enums"

    Public Enum EOrderOfFormNavigationArraylist
        CurrentFormName = 0
        CurrentFormObject = 1
        PreviousFormObject = 2
        NextFormObject = 3
    End Enum

    Public Enum Pin_Clip
        Pin = 0
        Clip = 1
    End Enum

#End Region

#Region "Properties"

    Public Property strColumnLoadDeratePressure() As String
        Get
            Return _strColumnLoadDeratePressure
        End Get
        Set(ByVal value As String)
            _strColumnLoadDeratePressure = value
        End Set
    End Property

#Region "Stroke Control"
    Public Property strstrokeControlCodeNumber() As String
        Get
            Return _strstrokeControlCodeNumber
        End Get
        Set(ByVal value As String)
            _strstrokeControlCodeNumber = value
        End Set
    End Property

    Public Property strStrokeControlDrawingNumber() As String
        Get
            Return _strStrokeDrawingNumber
        End Get
        Set(ByVal value As String)
            _strStrokeDrawingNumber = value
        End Set
    End Property
    Public Property strStrokeControlDescription() As String
        Get
            Return _strStrokeDescription
        End Get
        Set(ByVal value As String)
            _strStrokeDescription = value
        End Set
    End Property


#End Region

#Region "Rod Clevis"
    Public Property strRodClevisCodeNumber() As String
        Get
            Return _strRodClevisCodeNumber
        End Get
        Set(ByVal value As String)
            _strRodClevisCodeNumber = value
        End Set
    End Property

    Public Property strRodClevisDrawingNumber() As String
        Get
            Return _strRodClevisDrawingNumber
        End Get
        Set(ByVal value As String)
            _strRodClevisDrawingNumber = value
        End Set
    End Property
    Public Property strRodClevisDescription() As String
        Get
            Return _strRodClevisDescription
        End Get
        Set(ByVal value As String)
            _strRodClevisDescription = value
        End Set
    End Property

#End Region

#Region "Clevis Cap"
    Public Property strClevisCapCodeNumber() As String
        Get
            Return _strClevisCapCodeNumber
        End Get
        Set(ByVal value As String)
            _strClevisCapCodeNumber = value
        End Set
    End Property

    Public Property strClevisCapDrawingNumber() As String
        Get
            Return _strClevisCaprawingNumber
        End Get
        Set(ByVal value As String)
            _strClevisCaprawingNumber = value
        End Set
    End Property
    Public Property strClevisCapDescription() As String
        Get
            Return _strClevisCapDescription
        End Get
        Set(ByVal value As String)
            _strClevisCapDescription = value
        End Set
    End Property
#End Region

#Region "Rod Cap"
    Public Property strRodCapCodeNumber() As String
        Get
            Return _strRodCapCodeNumber
        End Get
        Set(ByVal value As String)
            _strRodCapCodeNumber = value
        End Set
    End Property

    Public Property strRodCapDrawingNumber() As String
        Get
            Return _strRodCaprawingNumber
        End Get
        Set(ByVal value As String)
            _strRodCaprawingNumber = value
        End Set
    End Property
    Public Property strRodCapDescription() As String
        Get
            Return _strRodCapDescription
        End Get
        Set(ByVal value As String)
            _strRodCapDescription = value
        End Set
    End Property
#End Region

#Region "Tie Rod Nut"
    Public Property strTieRodNutDescription() As String
        Get
            Return _strTieRodNutDescription
        End Get
        Set(ByVal value As String)
            _strTieRodNutDescription = value
        End Set
    End Property

    Public Property strTieRodNutDrawingNumber() As String
        Get
            Return _strTieRodNutDrawingNumber
        End Get
        Set(ByVal value As String)
            _strTieRodNutDrawingNumber = value
        End Set
    End Property
    '
    Public Property strTieRodNutCodeNumber() As String
        Get
            Return _strTieRodNutCodeNumber
        End Get
        Set(ByVal value As String)
            _strTieRodNutCodeNumber = value
        End Set
    End Property
#End Region

#Region "Pins"
    Public Property strPinsCodeNumber() As String
        Get
            Return _strPinsCodeNumber
        End Get
        Set(ByVal value As String)
            _strPinsCodeNumber = value
        End Set
    End Property

    Public Property strPinsDrawingNumber() As String
        Get
            Return _strPinsDrawingNumber
        End Get
        Set(ByVal value As String)
            _strPinsDrawingNumber = value
        End Set
    End Property
    Public Property strPinsDescription() As String
        Get
            Return _strPinsDescription
        End Get
        Set(ByVal value As String)
            _strPinsDescription = value
        End Set
    End Property
#End Region

#Region "Piston"
    Public Property strPistonCodeNumber() As String
        Get
            Return _strPistonCodeNumber
        End Get
        Set(ByVal value As String)
            _strPistonCodeNumber = value
        End Set
    End Property

    Public Property strPistonDrawingNumber() As String
        Get
            Return _strPistonDrawingNumber
        End Get
        Set(ByVal value As String)
            _strPistonDrawingNumber = value
        End Set
    End Property
    Public Property strPistonDescription() As String
        Get
            Return _strPistonDescription
        End Get
        Set(ByVal value As String)
            _strPistonDescription = value
        End Set
    End Property
#End Region

#Region "Rod"
    Public Property strRodCodeNumber() As String
        Get
            Return _strRodDCodeNumber
        End Get
        Set(ByVal value As String)
            _strRodDCodeNumber = value
        End Set
    End Property

    Public Property strRodDrawingNumber() As String
        Get
            Return _strRodDDrawingNumber
        End Get
        Set(ByVal value As String)
            _strRodDDrawingNumber = value
        End Set
    End Property

    Public Property strRodDescription() As String
        Get
            Return _strRodDDescription
        End Get
        Set(ByVal value As String)
            _strRodDDescription = value
        End Set
    End Property

#End Region

#Region "Bore or Tube"
    Public Property strBoreDescription() As String
        Get
            Return _strBoreDescription
        End Get
        Set(ByVal value As String)
            _strBoreDescription = value
        End Set
    End Property
    Public Property strBoreCodeNumber() As String
        Get
            Return _strBoreCodeNumber
        End Get
        Set(ByVal value As String)
            _strBoreCodeNumber = value
        End Set
    End Property
    Public Property strBoreDrawingNumber() As String
        Get
            Return _strBoreDrawingNumber
        End Get
        Set(ByVal value As String)
            _strBoreDrawingNumber = value
        End Set
    End Property
#End Region

#Region "Tie Rod"

    Public Property strTieRodDescription() As String
        Get
            Return _strTieRodDescription
        End Get
        Set(ByVal value As String)
            _strTieRodDescription = value
        End Set
    End Property

    Public Property strTieRodDrawingNumber() As String
        Get
            Return _strTieRodDrawingNumber
        End Get
        Set(ByVal value As String)
            _strTieRodDrawingNumber = value
        End Set
    End Property

    Public Property strTieRodCodeNumber() As String
        Get
            Return _strTieRodCodeNumber
        End Get
        Set(ByVal value As String)
            _strTieRodCodeNumber = value
        End Set
    End Property

#End Region

#Region "Rod Clevis Pin"
    Public Property strRodClevisPinCodeNumber() As String
        Get
            Return _strRodClevisPinCodeNumber
        End Get
        Set(ByVal value As String)
            _strRodClevisPinCodeNumber = value
        End Set
    End Property

    Public Property strClevisCapPinCodeNumber() As String
        Get
            Return _strClevisCapPinCodeNumber
        End Get
        Set(ByVal value As String)
            _strClevisCapPinCodeNumber = value
        End Set
    End Property
#End Region

    Public Property SetCodeDesciption() As String
        Get
            Return str_SetCodeDesciption
        End Get
        Set(ByVal value As String)
            str_SetCodeDesciption = value
        End Set
    End Property

    Public Property ObjCurrentForm() As Object
        Get
            Return _oCurrentForm
        End Get
        Set(ByVal value As Object)
            _oCurrentForm = value
        End Set
    End Property

    Public Property ofrmContractDetails() As frmContractDetails
        Get
            Return _ofrmContractDetails
        End Get
        Set(ByVal value As frmContractDetails)
            _ofrmContractDetails = value
        End Set
    End Property

    Public Property ofrmMonarch() As frmMonarch
        Get
            Return _ofrmMonarch
        End Get
        Set(ByVal value As frmMonarch)
            _ofrmMonarch = value
        End Set
    End Property

    Public Property ofrmTieRod1() As frmTieRod1
        Get
            Return _ofrmTieRod1

        End Get
        Set(ByVal value As frmTieRod1)
            _ofrmTieRod1 = value
        End Set
    End Property

    Public Property ofrmTieRod2() As frmTieRod2
        Get
            Return _ofrmTieRod2
        End Get
        Set(ByVal value As frmTieRod2)
            _ofrmTieRod2 = value
        End Set
    End Property

    Public Property ofrmTieRod3() As frmTieRod3
        Get
            Return _ofrmTieRod3
        End Get
        Set(ByVal value As frmTieRod3)
            _ofrmTieRod3 = value
        End Set
    End Property

    Public Property ofrmMdiMonarch() As mdiMonarch
        Get
            Return _ofrmMdiMonarch
        End Get
        Set(ByVal value As mdiMonarch)
            _ofrmMdiMonarch = value
        End Set
    End Property

    Public Property CylinderCodeNumber() As String
        Get
            Return _strCylinderCodeNumber
        End Get
        Set(ByVal value As String)
            _strCylinderCodeNumber = value
        End Set
    End Property

    Public Property SaveWorkFolder() As String
        Get
            Return _strfilePath
        End Get
        Set(ByVal value As String)
            _strfilePath = value
        End Set
    End Property

    Public Property TieRodStrokeDifference() As Double
        Get
            Return _dblTieRodStrokeDifference
        End Get
        Set(ByVal value As Double)
            _dblTieRodStrokeDifference = value
        End Set
    End Property

    Public Property BoreStrokeDifference() As Double
        Get
            Return _dblBoreStrokeDifference
        End Get
        Set(ByVal value As Double)
            _dblBoreStrokeDifference = value
        End Set
    End Property

    Public Property RodStrokeDifference() As Double
        Get
            Return _dblRodStrokeDifference
        End Get
        Set(ByVal value As Double)
            _dblRodStrokeDifference = value
        End Set
    End Property

    Public Property ReportFile() As String
        Get
            Return _strSavefiletopath
        End Get
        Set(ByVal value As String)
            _strSavefiletopath = value
        End Set
    End Property

    Public Property CodeDesc() As String
        Get
            Return _strCodeDesc
        End Get
        Set(ByVal value As String)
            _strCodeDesc = value
        End Set
    End Property

    Public Property RodDiameter() As Double
        Get
            Return _dblRodDiameter
        End Get
        Set(ByVal value As Double)
            _dblRodDiameter = value
        End Set
    End Property

    Public Property BoreDiameter() As Double
        Get
            Return _dblBoreDiameter
        End Get
        Set(ByVal value As Double)
            _dblBoreDiameter = value
        End Set
    End Property

    Public Property StrokeLength() As Double
        Get
            Return _dblStrokeLength
        End Get
        Set(ByVal value As Double)
            _dblStrokeLength = value
        End Set
    End Property

    Public Property StopTubeLength() As Double
        Get
            Return _dblStopTubeLength
        End Get
        Set(ByVal value As Double)
            _dblStopTubeLength = value
        End Set
    End Property

    Public Property RodAdder() As Double
        Get
            Return _dblRodAdder
        End Get
        Set(ByVal value As Double)
            _dblRodAdder = value
        End Set
    End Property

    Public Property TempWorkingPressure() As Double
        Get
            Return _dblTempWorkingPressure
        End Get
        Set(ByVal value As Double)
            _dblTempWorkingPressure = value
        End Set
    End Property

    Public Property WorkingPressure() As Double
        Get
            Return _dblWorkingPressure
        End Get
        Set(ByVal value As Double)
            _dblWorkingPressure = value
        End Set
    End Property

    Public Property ColumnLoad() As Double
        Get
            Return _dblColumnLoad
        End Get
        Set(ByVal value As Double)
            _dblColumnLoad = value
        End Set
    End Property

    Public Property PinSize() As Double
        Get
            Return _dblPinSize
        End Get
        Set(ByVal value As Double)
            _dblPinSize = value
        End Set
    End Property

    Public Property FormName() As Form
        Get
            Return _lFormName
        End Get
        Set(ByVal value As Form)
            _lFormName = value
        End Set
    End Property

    Public WriteOnly Property StopWatchAndProgressBar() As String
        Set(ByVal value As String)
            If value = "Start" Then
                MonarchProgressBar.Visible = True
                Control.CheckForIllegalCrossThreadCalls = False
                _oThreadProgressBarStepping = New Thread(New ThreadStart(AddressOf StartStepingProgressBar))
                _oThreadProgressBarStepping.IsBackground = True
                _oThreadProgressBarStepping.Start()
            ElseIf value = "Stop" Then
                If _oThreadProgressBarStepping.IsAlive Then
                    MonarchProgressBar.Value = MonarchProgressBar.Maximum
                    MonarchProgressBar.Value = 0
                    _oThreadProgressBarStepping.Abort()
                    MonarchProgressBar.Visible = False
                End If
            End If
        End Set
    End Property

    Public Property MonarchProgressBar() As ProgressBar
        Get
            Return _prbMonarch
        End Get
        Set(ByVal value As ProgressBar)
            _prbMonarch = value
        End Set
    End Property

    Private _blnIsCompleteModelGeneration As Boolean

    Public Property IsCompleteModelGeneration() As Boolean
        Get
            Return _blnIsCompleteModelGeneration
        End Get
        Set(ByVal value As Boolean)
            _blnIsCompleteModelGeneration = value
        End Set
    End Property

#End Region

#Region "For CMSIntegration from Costing"

    Private _aCostDetails_Costing As DataTable
    Public Property CostDetails_Costing() As DataTable
        Get
            Return _aCostDetails_Costing
        End Get
        Set(ByVal value As DataTable)
            _aCostDetails_Costing = value
        End Set
    End Property

    Private _strMETHDRAssemblyResource As String
    Public Property METHDRAssemblyResource() As String
        Get
            Return _strMETHDRAssemblyResource
        End Get
        Set(ByVal value As String)
            _strMETHDRAssemblyResource = value
        End Set
    End Property

    Private _dblMETHDRAssemblyRunStandard As Double
    Public Property METHDRAssemblyRunStandard() As String
        Get
            Return _dblMETHDRAssemblyRunStandard
        End Get
        Set(ByVal value As String)
            _dblMETHDRAssemblyRunStandard = value
        End Set
    End Property

    Private _dblMETHDRPaintRunStandard As Double
    Public Property METHDRPaintRunStandard() As String
        Get
            Return _dblMETHDRPaintRunStandard
        End Get
        Set(ByVal value As String)
            _dblMETHDRPaintRunStandard = value
        End Set
    End Property


#End Region

#Region "Sub Procedures"

    Public Sub StartStepingProgressBar()

        While Not MonarchProgressBar.Value = MonarchProgressBar.Maximum + 1
            If MonarchProgressBar.Value = MonarchProgressBar.Maximum Then
                MonarchProgressBar.Value = 0
            End If
            MonarchProgressBar.Value += 1
            Application.DoEvents()
            System.Threading.Thread.Sleep(100)
        End While

    End Sub

    Public Sub clearAllFormData()

        ofrmMonarch = Nothing
        ofrmContractDetails = Nothing
        ofrmTieRod1 = Nothing
        ofrmTieRod2 = Nothing
        ofrmTieRod3 = Nothing

    End Sub

    Public Sub KillAllSolidWorksServices()

        Try
            killSolidWorks("SLDWORKS")
            killSolidWorks("SolidWorksLicTemp.0001")
            killSolidWorks("SolidWorksLicensing")
            killSolidWorks("swvbaserver")

        Catch ex As Exception

        End Try

    End Sub

    Public Sub killSolidWorks(ByVal _strProcessName As String)

        Dim proc As System.Diagnostics.Process
        Try
            For Each proc In System.Diagnostics.Process.GetProcessesByName(_strProcessName)
                If proc.HasExited = False Then
                    proc.Kill()
                End If
            Next
        Catch oException As Exception
            ' MessageBox.Show("Unable to kill the Service" + vbNewLine + "System Generated Error" + oException.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            '_oErrorObject = oException
        End Try

    End Sub

    Public Sub getExcelReport()

        Dim dblRodLength As Double = 0
        updateHashTablevalues(htGuiReport, "Series", SeriesForCosting)
        updateHashTablevalues(htGuiReport, "Rephasing Port Position", IIf(Trim(ofrmTieRod1. _
                        cmbRephasingPortPosition.Text) = "", "None", ofrmTieRod1.cmbRephasingPortPosition.Text))
        updateHashTablevalues(htGuiReport, "Port Orientation at Rod End", ofrmTieRod1.cmbPortOrientationForRodCap.Text)
        updateHashTablevalues(htGuiReport, "Style", ofrmTieRod1.cmbStyle.Text)
        updateHashTablevalues(htGuiReport, "Bore", ofrmTieRod1.cmbBore.Text)
        updateHashTablevalues(htGuiReport, "Stroke Length", StrokeLength)
        updateHashTablevalues(htGuiReport, "Rod Adder", RodAdder)
        updateHashTablevalues(htGuiReport, "Stop Tube", IIf(ofrmTieRod1.rdbStopTubeYes.Checked = True, "Yes", "No"))
        updateHashTablevalues(htGuiReport, "Stop Tube Length", StopTubeLength)
        updateHashTablevalues(htGuiReport, "Rod Material", ofrmTieRod1.cmbRodMaterial.Text)
        updateHashTablevalues(htGuiReport, "Rod Diameter", RodDiameter)
        updateHashTablevalues(htGuiReport, "Nut Size", ofrmTieRod2.txtTieRodNutSize.Text)
        updateHashTablevalues(htGuiReport, "Port Orientation at Clevis Cap", ofrmTieRod1.cmbPortOrientation.Text)
        updateHashTablevalues(htGuiReport, "Pin Hole", ofrmTieRod1.cmbClevisCapPinHole.Text)
        updateHashTablevalues(htGuiReport, "Port", ofrmTieRod1.cmbClevisCapPort.Text)
        updateHashTablevalues(htGuiReport, "Pins", IIf(ofrmTieRod2.optPinsYes.Checked = True, "Yes", "No"))
        updateHashTablevalues(htGuiReport, "Rod Clevis Pins", IIf(ofrmTieRod2.optPinsYes_Rod.Checked = True, "Yes", "No"))        '05_04_2010    RAGAVA

        '14_04_2011  RAGAVA
        Try
            If Trim(ofrmTieRod1.cmbClevisCapPort.Text) <> "" Then
                Dim strClevisPortType As String = Trim(ofrmTieRod1.cmbClevisCapPort.Text).ToString. _
                            Substring(Trim(ofrmTieRod1.cmbClevisCapPort.Text).LastIndexOf(" ") + 1, 3)
                updateHashTablevalues(htGuiReport, "Base End Port Type", strClevisPortType)
                updateHashTablevalues(htGuiReport, "Base End Port Size", "'" & Trim(ofrmTieRod1. _
                    cmbClevisCapPort.Text).ToString.Substring(0, Trim(ofrmTieRod1.cmbClevisCapPort.Text).ToString.IndexOf(" ")))         '19_04_2011   RAGAVA
            End If
            If Trim(ofrmTieRod1.cmbRodCapPort.Text) <> "" Then
                Dim strRodPortType As String = Trim(ofrmTieRod1.cmbRodCapPort.Text).ToString. _
                        Substring(Trim(ofrmTieRod1.cmbRodCapPort.Text).LastIndexOf(" ") + 1, 3)
                updateHashTablevalues(htGuiReport, "Rod End Port Type", strRodPortType)
                updateHashTablevalues(htGuiReport, "Rod End Port Size", "'" & Trim(ofrmTieRod1. _
                            cmbRodCapPort.Text).ToString.Substring(0, Trim(ofrmTieRod1.cmbRodCapPort.Text).ToString.IndexOf(" ")))         '19_04_2011   RAGAVA
            End If
        Catch ex As Exception
        End Try
        'Till  Here

        updateHashTablevalues(htGuiReport, "Pin Material", ofrmTieRod2.cmbPinMaterial.Text)
        updateHashTablevalues(htGuiReport, "Pin Size", PinSize)
        updateHashTablevalues(htGuiReport, "Clips", ofrmTieRod2.cmbClips.Text)
        updateHashTablevalues(htGuiReport, "Piston Seal Package", ofrmTieRod2.cmbPistonSealPackage.Text)
        updateHashTablevalues(htGuiReport, "Rod Seal Package", ofrmTieRod2.cmbRodSealPackage.Text)
        updateHashTablevalues(htGuiReport, "Rod Cap", ofrmTieRod2.txtRodCap.Text)
        updateHashTablevalues(htGuiReport, "Clevis Cap", ofrmTieRod2.txtClevisCap.Text)
        updateHashTablevalues(htGuiReport, "Rod End Thread", ofrmTieRod2.cmbRodEndThread.Text)
        updateHashTablevalues(htGuiReport, "Rod Clevis Check", IIf(ofrmTieRod2.rdbRodClevisYes.Checked = True, "Yes", "No"))
        If ofrmTieRod2.rdbRodClevisYes.Checked = True Then      '18_10_2011   RAGAVA
            updateHashTablevalues(htGuiReport, "Rod Clevis", strRodClevisCodeNumber) 'ofrmTieRod2.cmbRodClevis.Text)
        End If
        updateHashTablevalues(htGuiReport, "Rod Clevis Pin", strRodClevisPinCodeNumber)

        updateHashTablevalues(htGuiReport, "Stroke Control", IIf(ofrmTieRod1.optStrokeControlYes.Checked = True, "Yes", "No"))
        updateHashTablevalues(htGuiReport, "Stroke Length Adder", ofrmTieRod1.cmbStrokeLengthAdder.Text)
        updateHashTablevalues(htGuiReport, "Retracted Length", ofrmTieRod1.txtRetractedLength.Text)
        updateHashTablevalues(htGuiReport, "Extended Length", ofrmTieRod1.txtExtendedLength.Text)
        updateHashTablevalues(htGuiReport, "Tie Rod Size", ofrmTieRod2.txtTieRodSize.Text)
        updateHashTablevalues(htGuiReport, "Tie Rod Nut size", ofrmTieRod2.txtTieRodNutSize.Text)
        updateHashTablevalues(htGuiReport, "Tie Rod Nut Qty", ofrmTieRod2.txtTieRodNutQty.Text)
        updateHashTablevalues(htGuiReport, "Tie Rod Nut", ofrmTieRod2.txtTieRodNutSize.Text)
        updateHashTablevalues(htGuiReport, "Thred Protector", ofrmTieRod2.cmbThreadProtected.Text)
        updateHashTablevalues(htGuiReport, "Paint", ofrmTieRod2.cmbPaint.Text)
        updateHashTablevalues(htGuiReport, "Packaging", ofrmTieRod2.txtPackaging.Text)
        'anup 31-01-2011 start
        ' updateHashTablevalues(htGuiReport, "Rod Wiper", ofrmTieRod2.txtRodWiper.Text)
        updateHashTablevalues(htGuiReport, "Rod Wiper", ofrmTieRod2.cmbRodWiper.Text)
        'anup 31-01-2011 till here
        updateHashTablevalues(htGuiReport, "Tube Seal", ofrmTieRod2.txtTubeSeal1.Text)

        '22_02_2010   RAGAVA
        'If ofrmTieRod3.chkPackPinsAndClipsInPlasticBag.Checked = True Then
        '    updateHashTablevalues(htGuiReport, "Pins in Separate Bag", "Yes")
        'Else
        If blnInstallPinsandClips_Checked = True Then
            updateHashTablevalues(htGuiReport, "Pins in Separate Bag", "No")
        Else
            updateHashTablevalues(htGuiReport, "Pins in Separate Bag", "Yes")
        End If
        'End If
        '22_02_2010   RAGAVA   Till  Here
        'Piston
        Try
            Dim oListViewItem As ListViewItem
            Dim strColumns() As String
            strColumns = (ofrmTieRod2.cmbPistonSealPackage.Text).Split("+")
            Dim StrSql As String
            Dim arrSeries As String()
            arrSeries = SeriesForCosting.ToString.Split(" ")
            Dim strSeries As String = arrSeries(0)
            If SeriesForCosting.ToString.StartsWith("TX") = False Then
                oListViewItem = ofrmTieRod1.LVNutSizeDetails.SelectedItems(0)
                ' oListViewItem = ofrmTieRod1.LVNutSizeDetails.Items(ofrmTieRod1.LVNutSizeDetails.GetCurrentIndex)
                StrSql = "select * from PistonSealDetails where BoreDiameter = " & _
                    Val(ofrmTieRod1.cmbBore.Text) & " and PistonNutSize = " & Val(oListViewItem.SubItems(0).Text) _
                                & " and Series like '%" & strSeries & "%'"
            Else
                StrSql = "select * from PistonSealDetails where BoreDiameter = " & _
                        Val(ofrmTieRod1.cmbBore.Text) & " and Series like '%" & strSeries & "%'" '" and PistonNutSize = " & Val(oListViewItem.SubItems(0).Text)
            End If

            For Each strCol As String In strColumns
                StrSql = StrSql & " and " & Trim(strCol) & " <> ''"
            Next
            Dim objDT As DataTable = oDataClass.GetDataTable(StrSql)
            If objDT.Rows.Count > 0 Then
                updateHashTablevalues(htGuiReport, "Piston", objDT.Rows(0).Item("PartNumber").ToString)
                ObjClsCostingDetails.AddCodeNumberToDataTable(objDT.Rows(0).Item("PartNumber").ToString, "Piston Code") 'Sandeep 04-03-10-4pm
                If Not IsDBNull(objDT.Rows(0).Item("PistonNutCode")) Then
                    ObjClsCostingDetails.AddCodeNumberToDataTable(objDT.Rows(0).Item("PistonNutCode").ToString, "Piston Nut Code") 'Sandeep 30-04-10-10am
                End If
                _strPistonCode = objDT.Rows(0).Item("PartNumber").ToString
            End If
        Catch ex As Exception

        End Try
        'Pin
        Try
            Dim oListViewItem As ListViewItem
            Dim StrSql As String
            oListViewItem = ofrmTieRod2.LVPinSizeDetails.SelectedItems(0)
            ' oListViewItem = ofrmTieRod2.LVPinSizeDetails.Items(ofrmTieRod2.LVPinSizeDetails.GetCurrentIndex)
            StrSql = "select PartNumber from ClevisPinDetails where PinMaterial = '" & _
                Trim(ofrmTieRod2.cmbPinMaterial.Text) & "' and PinHoleSize = " & _
                Val(oListViewItem.SubItems(0).Text) & " and PinType = '" & Trim(ofrmTieRod2.cmbClips.Text) _
                & "' and " & Val(ofrmTieRod1.cmbBore.Text) & ">= BoreDiameterMinimum and " & _
                Val(ofrmTieRod1.cmbBore.Text) & " <= BoreDiameterMaximum"
            Dim objDT As DataTable = oDataClass.GetDataTable(StrSql)
            If objDT.Rows.Count > 0 Then
                updateHashTablevalues(htGuiReport, "Pin", objDT.Rows(0).Item("PartNumber").ToString)
            End If
        Catch ex As Exception

        End Try
        'Rod Drawing Number
        Try
            Dim oListViewItem As ListViewItem
            oListViewItem = ofrmTieRod1.LVRodDiameterDetails.SelectedItems(0)
            'oListViewItem = ofrmTieRod1.LVRodDiameterDetails.Items(ofrmTieRod1.LVRodDiameterDetails.GetCurrentIndex)
            Dim StrSql As String
            Dim dblPistonNutSize As Double = 0
            'ANUP 12-10-2010 START
            If ofrmTieRod1.LVNutSizeDetails.SelectedItems.Count < 1 Then
                dblPistonNutSize = 0
                StrSql = "select rdd.DrawingPartNumber,rdd.OverAllRodLength,rdd.StrokeLength,rdd.RodMaterialNumber from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.PartNumberId = rdd.PartNumber and bdrd.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ")and rdd.series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and MaterialType = '" & Trim(ofrmTieRod1.cmbRodMaterial.Text) & "' and RodDiameter = " & Val(oListViewItem.SubItems(0).Text) & " and RodThreadSize = " & Val(ofrmTieRod2.cmbRodEndThread.Text)
            Else
                Dim oListViewNutSize As ListViewItem = ofrmTieRod1.LVNutSizeDetails.SelectedItems(0)     '05_09_2009  ragava
                ' oListViewNutSize = ofrmTieRod1.LVNutSizeDetails.Items(ofrmTieRod1.LVNutSizeDetails.GetCurrentIndex)
                dblPistonNutSize = Val(oListViewNutSize.SubItems(0).Text)
                StrSql = "select rdd.DrawingPartNumber,rdd.OverAllRodLength,rdd.StrokeLength,rdd.RodMaterialNumber from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.PartNumberId = rdd.PartNumber and bdrd.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ")and rdd.series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and MaterialType = '" & Trim(ofrmTieRod1.cmbRodMaterial.Text) & "' and RodDiameter = " & Val(oListViewItem.SubItems(0).Text) & " and RodThreadSize = " & Val(ofrmTieRod2.cmbRodEndThread.Text) & " And pistonthreadSize= " & dblPistonNutSize
            End If
            'ANUP 12-10-2010 TILL HERE
            If Trim(ofrmTieRod1.cmbStyle.Text) = "ASAE" Then
                StrSql = StrSql + " And StrokeLength =" & Val(ofrmTieRod1.cmbStrokeLength.Text)
            End If
            Dim objDT As DataTable = oDataClass.GetDataTable(StrSql)
            If objDT.Rows.Count > 0 Then
                RodStrokeDifference = Math.Round(Val(objDT.Rows(0).Item("OverAllRodLength").ToString) _
                    - Val(objDT.Rows(0).Item("StrokeLength").ToString), 2)
                updateHashTablevalues(htMainParams, "Rod Drawing Number", objDT.Rows(0).Item("DrawingPartNumber").ToString)
                RodDrawingNumber = objDT.Rows(0).Item("DrawingPartNumber").ToString        '11_09_2009   Ragava
                RodMaterialNumber = objDT.Rows(0).Item("RodMaterialNumber").ToString        '11_09_2009   Ragava
                If Trim(ofrmTieRod1.cmbStyle.Text) = "ASAE" Then
                    Try
                        If Val(objDT.Rows(0).Item("StrokeLength").ToString) = Val(ofrmTieRod1.cmbStrokeLength.Text) Then
                            dblRodLength = Val(objDT.Rows(0).Item("OverAllRodLength").ToString)
                        ElseIf Val(ofrmTieRod1.cmbStrokeLength.Text) > Val(objDT.Rows(0).Item("StrokeLength").ToString) Then
                            dblRodLength = Val(objDT.Rows(0).Item("OverAllRodLength").ToString) + 11.25
                        ElseIf Val(ofrmTieRod1.cmbStrokeLength.Text) < Val(objDT.Rows(0).Item("StrokeLength").ToString) Then
                            dblRodLength = Val(objDT.Rows(0).Item("OverAllRodLength").ToString) - 11.25
                        End If
                    Catch ex As Exception

                    End Try
                End If
            End If
        Catch ex As Exception

        End Try

        'Bore Drawing Number
        Try
            Dim StrSql As String
            Dim arrSeries As String()
            arrSeries = SeriesForCosting.ToString.Split(" ")
            Dim strSeries As String = arrSeries(0)
            If Trim(ofrmTieRod1.cmbRephasingPortPosition.Text) <> "" Then
                If Trim(ofrmTieRod1.cmbRephasingPortPosition.Text) = "At Extension" Then
                    strSeries = strSeries.Insert(2, "E")
                ElseIf Trim(ofrmTieRod1.cmbRephasingPortPosition.Text) = "At Retraction" Then
                    strSeries = strSeries.Insert(2, "B")
                Else
                    strSeries = strSeries.Insert(2, "2")
                End If
            End If
            StrSql = "select * from BoreDiameterDetails where BoreDiameter = " & _
                Val(ofrmTieRod1.cmbBore.Text) & " and Series like '%" & strSeries & "%'"
            Dim objDT As DataTable = oDataClass.GetDataTable(StrSql)
            If objDT.Rows.Count > 0 Then
                BoreStrokeDifference = Math.Round(Val(objDT.Rows(0).Item("TubeLength").ToString) _
                            - Val(objDT.Rows(0).Item("NominalStroke").ToString), 2)
                updateHashTablevalues(htGuiReport, "Bore Drawing Number", objDT.Rows(0).Item("DrawingPartNumber").ToString)
                BoreDrawingNumber = objDT.Rows(0).Item("DrawingPartNumber").ToString        '11_09_2009   Ragava
            End If
        Catch ex As Exception

        End Try
        'TieRod Drawing Number
        Try
            Dim StrSql As String
            'ANUP 12-10-2010 START
            StrSql = "select DrawingNumber,StrokeLength,[Dimension-A] from TieRodSizes trz,BoreDiameter_TieRodSizes bdtr where bdtr.PartNumberID = trz.TieRodPartNumber and bdtr.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ")and trz.series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "'"
            'ANUP 12-10-2010 TILL HERE
            Dim objDT As DataTable = oDataClass.GetDataTable(StrSql)
            If objDT.Rows.Count > 0 Then
                TieRodStrokeDifference = Math.Round(Val(objDT.Rows(0).Item("Dimension-A").ToString) _
                        - Val(objDT.Rows(0).Item("StrokeLength").ToString), 2) 'StrokeLength
                updateHashTablevalues(htGuiReport, "Tie Rod Drawing Number", objDT.Rows(0).Item("DrawingNumber").ToString)
                TieRodDrawingNumber = objDT.Rows(0).Item("DrawingNumber").ToString        '11_09_2009   Ragava
            End If
        Catch ex As Exception

        End Try
        updateHashTablevalues(htGuiReport, "Tube Length", StrokeLength + BoreStrokeDifference + StopTubeLength)
        If Trim(ofrmTieRod1.cmbStyle.Text) = "ASAE" Then
            updateHashTablevalues(htMainParams, "Rod Length", dblRodLength)
        Else
            updateHashTablevalues(htMainParams, "Rod Length", StrokeLength + RodStrokeDifference + StopTubeLength + RodAdder)
        End If
        updateHashTablevalues(htGuiReport, "Tie Rod Length", StrokeLength + TieRodStrokeDifference + StopTubeLength)

    End Sub

    Public Sub updateDesignTables()

        Try

            If Not CopyTheMasterFile(Application.StartupPath & "\GUI_PARAMETERS_report.xls") Then
                Exit Sub
            End If
            Application.DoEvents()
            oExcelClass.checkExcelInstance()
            oExcelClass.objBook = oExcelClass.objApp.Workbooks.Open(ReportFile)
            oExcelClass.objSheet = oExcelClass.objBook.Worksheets("Sheet1")
            getExcelReport()
            oExcelClass.IDEnumerator = htGuiReport.GetEnumerator
            oExcelClass.updateDesign_Parameters("Sheet1")
            Try
                oExcelClass.checkExcelInstance()
                oExcelClass.objBook = oExcelClass.objApp.Workbooks.Open(ReportFile)
                oExcelClass.objSheet = oExcelClass.objBook.Worksheets("Sheet2")
                updateGUIInputParameters()
                oExcelClass.objBook.Close()
                oExcelClass.objApp.Quit()
                oExcelClass.objApp = Nothing
            Catch ex As Exception
            End Try
        Catch ex As Exception
        End Try

    End Sub

    Public Sub updateMainDesignTables()

        Try
            oExcelClass.checkExcelInstance()
            oExcelClass.objBook = oExcelClass.objApp.Workbooks.Open("C:\DESIGN_TABLES\GUI_PARAMETERS.xls")
            oExcelClass.objSheet = oExcelClass.objBook.Worksheets("Sheet1")
            updateMainParams()
            oExcelClass.IDEnumerator = htMainParams.GetEnumerator
            oExcelClass.updateDesign_Parameters("Sheet1")
            Try

                oExcelClass.objApp = Nothing
                Sleep(1000)
                oExcelClass.checkExcelInstance()
                Try
                    oExcelClass.objBook = oExcelClass.objApp.Workbooks.Open("C:\DESIGN_TABLES\MAIN_ASSEMBLY.xls")
                    oExcelClass.objSheet = oExcelClass.objBook.Worksheets("Sheet1")
                    oExcelClass.objBook.Save()
                    oExcelClass.objBook.Close()
                Catch ex As Exception
                End Try

                oExcelClass.objApp.Quit()
                Sleep(1000)
                saveAllExcelFiles()
            Catch ex As Exception
            End Try
        Catch ex As Exception
        End Try

    End Sub

    Public Function CopyTheMasterFile(ByVal fileName As String) As Boolean

        Dim IsProcessSuccessfull As Boolean = False
        Dim sErrorMessage As String = "Report Master file does not exist"
        If IsMasterReportFileExists(fileName) Then
            Try
                If File.Exists(ReportFile) Then
                    File.Delete(ReportFile)
                End If
                File.Copy(fileName, ReportFile)
                IsProcessSuccessfull = True
            Catch oException As Exception
                sErrorMessage = "Unable to copy the source file" + vbCrLf + vbCrLf + oException.Message
            End Try
        End If
        If Not IsProcessSuccessfull Then
            MessageBox.Show(sErrorMessage, "Error in file creation", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        Return IsProcessSuccessfull

    End Function

    Public Function IsMasterReportFileExists(ByVal fileName As String) As Boolean

        IsMasterReportFileExists = File.Exists(fileName)

    End Function

    Public Sub ShowSaveDialog()
        ' Get the report filename from the user

        ReportFile = GetFileNameFromUser()

        ' Check if any valid filename is provided
        If Not ReportFile.Equals("") Then
            ' Go ahead with generating the report
            If GenerateReportToExcelFile() Then
                ' Show message to the user that the report was successfully generated and ask him if he/she likes to 
                'open(the)report to view now
                Dim sMessage As String = "Report successfully generated and saved at: " + vbCrLf
                sMessage += ReportFile + vbCrLf + vbCrLf
                sMessage += "Do you want to view the report file now?"
                If MessageBox.Show(sMessage, "Report generated", MessageBoxButtons.YesNo, MessageBoxIcon.Question, _
                MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                    ' cach any error if excel doesnt starts
                    Try
                        Process.Start(ReportFile)
                    Catch oException As Exception
                        sMessage = "Unable to start MS Excel 2003" + vbCrLf + "Try opening the file manually from here:- " _
                        + vbCrLf + "savefiletopath " + vbCrLf + vbCrLf
                        sMessage += "System generated error:-" + vbCrLf + oException.Message
                        MessageBox.Show(sMessage, "Error opening Excel", MessageBoxButtons.OK, MessageBoxIcon.Hand)
                    End Try
                End If
            End If
        End If

    End Sub

    Public Function GetFileNameFromUser() As String

        GetFileNameFromUser = ""
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.Filter = "Excel files (*.xls)|*.xls"
        SaveFileDialog.Title = "Save Report"
        SaveFileDialog.CheckFileExists = False
        SaveFileDialog.FilterIndex = 2
        SaveFileDialog.RestoreDirectory = True
        SaveFileDialog.ShowDialog()
        If Not SaveFileDialog.FileName.Equals("") Then
            GetFileNameFromUser = (SaveFileDialog.FileName) & ".xls".ToString()
        End If
        SaveFileDialog = Nothing

    End Function

    Public Function GenerateReportToExcelFile() As Boolean

        Try
            updateDesignTables()
            GenerateReportToExcelFile = True
        Catch ex As Exception
            MessageBox.Show("Unable to generate report!" + vbCrLf + _
            "Kindly check Microsoct Excel 2003 is installed in this system.", "Error in report generation", _
            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Function

    Public Sub saveAllExcelFiles()

        oExcelClass.getExcelFiles("C:\DESIGN_TABLES")
        Try
            oExcelClass.objApp.Quit()
            Sleep(3000)
            ' killExcelObjects()
        Catch ex As Exception
        End Try

    End Sub

    Public Sub updateMainParams()
        '05_09_2009  ragava
        Dim dblRodLength As Double = 0

        updateHashTablevalues(htMainParams, "Series", SeriesForCosting)
        updateHashTablevalues(htMainParams, "Rephasing Port Position", IIf(Trim(ofrmTieRod1. _
                cmbRephasingPortPosition.Text) = "", "None", ofrmTieRod1.cmbRephasingPortPosition.Text))
        updateHashTablevalues(htMainParams, "Port Orientation at Rod End", ofrmTieRod1.cmbPortOrientationForRodCap.Text)
        updateHashTablevalues(htMainParams, "Style", ofrmTieRod1.cmbStyle.Text)
        updateHashTablevalues(htMainParams, "Bore", ofrmTieRod1.cmbBore.Text)
        updateHashTablevalues(htMainParams, "Stroke Length", StrokeLength)
        updateHashTablevalues(htMainParams, "Rod Adder", RodAdder)
        updateHashTablevalues(htMainParams, "Stop Tube", IIf(ofrmTieRod1.rdbStopTubeYes.Checked = True, "Yes", "No"))
        updateHashTablevalues(htMainParams, "Stop Tube Length", StopTubeLength)
        updateHashTablevalues(htMainParams, "Rod Material", ofrmTieRod1.cmbRodMaterial.Text)
        updateHashTablevalues(htMainParams, "Rod Diameter", RodDiameter)
        updateHashTablevalues(htMainParams, "Nut Size", ofrmTieRod2.txtTieRodNutSize.Text)
        updateHashTablevalues(htMainParams, "Port Orientation at Clevis Cap", ofrmTieRod1.cmbPortOrientation.Text)
        updateHashTablevalues(htMainParams, "Pin Hole", ofrmTieRod1.cmbClevisCapPinHole.Text)
        updateHashTablevalues(htMainParams, "Port", ofrmTieRod1.cmbClevisCapPort.Text)
        updateHashTablevalues(htMainParams, "Pins", IIf(ofrmTieRod2.optPinsYes.Checked = True, "Yes", "No"))
        updateHashTablevalues(htMainParams, "Rod Clevis Pins", IIf(ofrmTieRod2.optPinsYes_Rod.Checked = True, "Yes", "No"))        '05_04_2010    RAGAVA
        '14_04_2011  RAGAVA
        Try
            If Trim(ofrmTieRod1.cmbClevisCapPort.Text) <> "" Then
                Dim strClevisPortType As String = Trim(ofrmTieRod1.cmbClevisCapPort.Text).ToString. _
                            Substring(Trim(ofrmTieRod1.cmbClevisCapPort.Text).LastIndexOf(" ") + 1, 3)
                updateHashTablevalues(htMainParams, "Base End Port Type", strClevisPortType)
                updateHashTablevalues(htMainParams, "Base End Port Size", "'" & Trim(ofrmTieRod1. _
                    cmbClevisCapPort.Text).ToString.Substring(0, Trim(ofrmTieRod1.cmbClevisCapPort.Text).ToString.IndexOf(" ")))         '19_04_2011   RAGAVA
            End If
            If Trim(ofrmTieRod1.cmbRodCapPort.Text) <> "" Then
                Dim strRodPortType As String = Trim(ofrmTieRod1.cmbRodCapPort.Text).ToString. _
                      Substring(Trim(ofrmTieRod1.cmbRodCapPort.Text).LastIndexOf(" ") + 1, 3)
                updateHashTablevalues(htMainParams, "Rod End Port Type", strRodPortType)
                updateHashTablevalues(htMainParams, "Rod End Port Size", "'" & Trim(ofrmTieRod1. _
                     cmbRodCapPort.Text).ToString.Substring(0, Trim(ofrmTieRod1.cmbRodCapPort.Text).ToString.IndexOf(" ")))         '19_04_2011   RAGAVA
            End If
        Catch ex As Exception
        End Try
        'Till  Here
        updateHashTablevalues(htMainParams, "Pin Material", ofrmTieRod2.cmbPinMaterial.Text)
        updateHashTablevalues(htMainParams, "Pin Size", PinSize)
        updateHashTablevalues(htMainParams, "Clips", ofrmTieRod2.cmbClips.Text)
        updateHashTablevalues(htMainParams, "Piston Seal Package", ofrmTieRod2.cmbPistonSealPackage.Text)
        updateHashTablevalues(htMainParams, "Rod Seal Package", ofrmTieRod2.cmbRodSealPackage.Text)
        updateHashTablevalues(htMainParams, "Rod Cap", ofrmTieRod2.txtRodCap.Text)
        updateHashTablevalues(htMainParams, "Clevis Cap", ofrmTieRod2.txtClevisCap.Text)
        updateHashTablevalues(htMainParams, "Rod End Thread", ofrmTieRod2.cmbRodEndThread.Text)
        updateHashTablevalues(htMainParams, "Rod Clevis Check", IIf(ofrmTieRod2.rdbRodClevisYes.Checked = True, "Yes", "No"))
        If ofrmTieRod2.rdbRodClevisYes.Checked = True Then      '18_10_2011   RAGAVA
            updateHashTablevalues(htMainParams, "Rod Clevis", strRodClevisCodeNumber) ' ofrmTieRod2.cmbRodClevis.Text)
        End If
        updateHashTablevalues(htMainParams, "Stroke Control", IIf(ofrmTieRod1.optStrokeControlYes.Checked = True, "Yes", "No"))
        updateHashTablevalues(htMainParams, "Stroke Length Adder", ofrmTieRod1.cmbStrokeLengthAdder.Text)
        updateHashTablevalues(htMainParams, "Retracted Length", ofrmTieRod1.txtRetractedLength.Text)
        updateHashTablevalues(htMainParams, "Extended Length", ofrmTieRod1.txtExtendedLength.Text)
        updateHashTablevalues(htMainParams, "Tie Rod Size", ofrmTieRod2.txtTieRodSize.Text)
        updateHashTablevalues(htMainParams, "Tie Rod Nut size", ofrmTieRod2.txtTieRodNutSize.Text)
        updateHashTablevalues(htMainParams, "Tie Rod Nut Qty", ofrmTieRod2.txtTieRodNutQty.Text)
        updateHashTablevalues(htMainParams, "Tie Rod Nut", ofrmTieRod2.txtTieRodNutSize.Text)
        updateHashTablevalues(htMainParams, "Thred Protector", ofrmTieRod2.cmbThreadProtected.Text)
        updateHashTablevalues(htMainParams, "Paint", ofrmTieRod2.cmbPaint.Text)
        updateHashTablevalues(htMainParams, "Packaging", ofrmTieRod2.txtPackaging.Text)
        'anup 31-01-2011 start
        'updateHashTablevalues(htMainParams, "Rod Wiper", ofrmTieRod2.txtRodWiper.Text)
        updateHashTablevalues(htMainParams, "Rod Wiper", ofrmTieRod2.cmbRodWiper.Text)
        'anup 31-01-2011 till here
        updateHashTablevalues(htMainParams, "Tube Seal", ofrmTieRod2.txtTubeSeal1.Text)
        updateHashTablevalues(htMainParams, "Rod Clevis Pin", strRodClevisPinCodeNumber) '15_109_2009
        '23_02_2010   RAGAVA
        'If ofrmTieRod3.chkPackPinsAndClipsInPlasticBag.Checked = True Then
        '    updateHashTablevalues(htMainParams, "Pins in Separate Bag", "Yes")
        'Else
        If blnInstallPinsandClips_Checked = True Then
            updateHashTablevalues(htMainParams, "Pins in Separate Bag", "No")
        Else
            updateHashTablevalues(htMainParams, "Pins in Separate Bag", "Yes")
        End If
        'End If
        '23_02_2010   RAGAVA   Till  Here

        '31_08_2009  ragava
        'Piston
        Try
            Dim oListViewItem As ListViewItem
            Dim strColumns() As String
            strColumns = (ofrmTieRod2.cmbPistonSealPackage.Text).Split("+")
            Dim StrSql As String
            'oListViewItem = ofrmTieRod1.LVNutSizeDetails.SelectedItems(0)
            '        oListViewItem = ofrmTieRod1.LVRodDiameterDetails.SelectedItems(0)
            'StrSql = "select * from PistonSealDetails where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & " and PistonNutSize = " & Val(oListViewItem.SubItems(0).Text)
            Dim arrSeries As String()
            arrSeries = SeriesForCosting.ToString.Split(" ")
            Dim strSeries As String = arrSeries(0)
            If SeriesForCosting.ToString.StartsWith("TX") = False Then        '04_09_2009  ragava
                oListViewItem = ofrmTieRod1.LVNutSizeDetails.SelectedItems(0)
                ' oListViewItem = ofrmTieRod1.LVNutSizeDetails.Items(ofrmTieRod1.LVNutSizeDetails.GetCurrentIndex)
                StrSql = "select * from PistonSealDetails where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) _
                    & " and PistonNutSize = " & Val(oListViewItem.SubItems(0).Text) & " and Series like '%" & strSeries & "%'"
            Else
                StrSql = "select * from PistonSealDetails where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) _
                    & " and Series like '%" & strSeries & "%'" '" and PistonNutSize = " & Val(oListViewItem.SubItems(0).Text)
            End If

            For Each strCol As String In strColumns
                StrSql = StrSql & " and " & Trim(strCol) & " <> ''"
            Next
            Dim objDT As DataTable = oDataClass.GetDataTable(StrSql)
            If objDT.Rows.Count > 0 Then
                updateHashTablevalues(htMainParams, "Piston", objDT.Rows(0).Item("PartNumber").ToString)
            End If
        Catch ex As Exception

        End Try
        'Pin
        Try

            Dim oListViewItem As ListViewItem
            Dim StrSql As String
            oListViewItem = ofrmTieRod2.LVPinSizeDetails.SelectedItems(0)
            ' oListViewItem = ofrmTieRod2.LVPinSizeDetails.Items(ofrmTieRod2.LVPinSizeDetails.GetCurrentIndex)
            StrSql = "select PartNumber from ClevisPinDetails where PinMaterial = '" & _
                Trim(ofrmTieRod2.cmbPinMaterial.Text) & "' and PinHoleSize = " & Val(oListViewItem.SubItems(0).Text) _
                    & " and PinType = '" & Trim(ofrmTieRod2.cmbClips.Text) & "' and " & Val(ofrmTieRod1.cmbBore.Text) _
                    & ">= BoreDiameterMinimum and " & Val(ofrmTieRod1.cmbBore.Text) & " <= BoreDiameterMaximum"
            Dim objDT As DataTable = oDataClass.GetDataTable(StrSql)
            If objDT.Rows.Count > 0 Then
                updateHashTablevalues(htMainParams, "Pin", objDT.Rows(0).Item("PartNumber").ToString)
            End If
        Catch ex As Exception

        End Try
        'Rod Drawing Number
        Try
            Dim oListViewItem As ListViewItem
            oListViewItem = ofrmTieRod1.LVRodDiameterDetails.SelectedItems(0)
            'oListViewItem = ofrmTieRod1.LVRodDiameterDetails.Items(ofrmTieRod1.LVRodDiameterDetails.GetCurrentIndex)
            Dim StrSql As String
            Dim dblPistonNutSize As Double = 0
            If ofrmTieRod1.LVNutSizeDetails.SelectedItems.Count < 1 Then
                dblPistonNutSize = 0
                'StrSql = "select rdd.DrawingPartNumber,rdd.OverAllRodLength,rdd.StrokeLength from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.PartNumberId = rdd.PartNumber and bdrd.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ")and rdd.series = '" & IIf(Trim(ofrmTieRod1.cmbSeries.Text.ToString).StartsWith("TX"), "TX", "TL/TH/TP") & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and MaterialType = '" & Trim(ofrmTieRod1.cmbRodMaterial.Text) & "' and RodDiameter = " & Val(oListViewItem.SubItems(0).Text) & " and RodThreadSize = " & Val(ofrmTieRod2.cmbRodEndThread.Text)
                'ANUP 12-10-2010 START
                StrSql = "select rdd.DrawingPartNumber,rdd.OverAllRodLength,rdd.StrokeLength,RDD.RodMaterialNumber,rdd.PartNumber from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.PartNumberId = rdd.PartNumber and bdrd.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ")and rdd.series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and MaterialType = '" & Trim(ofrmTieRod1.cmbRodMaterial.Text) & "' and RodDiameter = " & Val(oListViewItem.SubItems(0).Text) & " and RodThreadSize = " & Val(ofrmTieRod2.cmbRodEndThread.Text)           '01_10_2009  ragava
            Else
                Dim oListViewNutSize As ListViewItem = ofrmTieRod1.LVNutSizeDetails.SelectedItems(0)     '05_09_2009  ragava
                ' oListViewNutSize = ofrmTieRod1.LVNutSizeDetails.Items(ofrmTieRod1.LVNutSizeDetails.GetCurrentIndex)
                dblPistonNutSize = Val(oListViewNutSize.SubItems(0).Text)
                'StrSql = "select rdd.DrawingPartNumber,rdd.OverAllRodLength,rdd.StrokeLength from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.PartNumberId = rdd.PartNumber and bdrd.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ")and rdd.series = '" & IIf(Trim(ofrmTieRod1.cmbSeries.Text.ToString).StartsWith("TX"), "TX", "TL/TH/TP") & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and MaterialType = '" & Trim(ofrmTieRod1.cmbRodMaterial.Text) & "' and RodDiameter = " & Val(oListViewItem.SubItems(0).Text) & " and RodThreadSize = " & Val(ofrmTieRod2.cmbRodEndThread.Text) & " And pistonthreadSize= " & dblPistonNutSize
                StrSql = "select rdd.DrawingPartNumber,rdd.OverAllRodLength,rdd.StrokeLength,RDD.RodMaterialNumber,rdd.PartNumber from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.PartNumberId = rdd.PartNumber and bdrd.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ")and rdd.series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and MaterialType = '" & Trim(ofrmTieRod1.cmbRodMaterial.Text) & "' and RodDiameter = " & Val(oListViewItem.SubItems(0).Text) & " and RodThreadSize = " & Val(ofrmTieRod2.cmbRodEndThread.Text) & " And pistonthreadSize= " & dblPistonNutSize            '01_10_2009  ragava
            End If
            'ANUP 12-10-2010 TILL HERE
            '25_09_2009
            If Trim(ofrmTieRod1.cmbStyle.Text) = "ASAE" Then
                StrSql = StrSql + " And StrokeLength =" & Val(ofrmTieRod1.cmbStrokeLength.Text)
            End If
            'StrSql = "select rdd.DrawingPartNumber,rdd.OverAllRodLength,rdd.StrokeLength from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.PartNumberId = rdd.PartNumber and bdrd.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ")and rdd.series = '" & IIf(Trim(ofrmTieRod1.cmbSeries.Text.ToString).StartsWith("TX"), "TX", "TL/TH/TP") & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and MaterialType = '" & Trim(ofrmTieRod1.cmbRodMaterial.Text) & "' and RodDiameter = " & Val(oListViewItem.SubItems(0).Text) 
            'StrSql = "select rdd.DrawingPartNumber,rdd.OverAllRodLength,rdd.StrokeLength from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.PartNumberId = rdd.PartNumber and bdrd.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ")and rdd.series = '" & IIf(Trim(ofrmTieRod1.cmbSeries.Text.ToString).StartsWith("TX"), "TX", "TL/TH/TP") & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and MaterialType = '" & Trim(ofrmTieRod1.cmbRodMaterial.Text) & "' and RodDiameter = " & Val(oListViewItem.SubItems(0).Text) & " and RodThreadSize = " & Val(ofrmTieRod2.cmbRodEndThread.Text)
            Dim objDT As DataTable = oDataClass.GetDataTable(StrSql)
            If objDT.Rows.Count > 0 Then
                RodStrokeDifference = Math.Round(Val(objDT.Rows(0).Item("OverAllRodLength").ToString) - Val(objDT.Rows(0).Item("StrokeLength").ToString), 2)
                updateHashTablevalues(htMainParams, "Rod Drawing Number", objDT.Rows(0).Item("DrawingPartNumber").ToString)
                RodDrawingNumber = objDT.Rows(0).Item("DrawingPartNumber").ToString          '15_09_2009   ragava
                RodMaterialNumber = objDT.Rows(0).Item("RodMaterialNumber").ToString        '01_10_2009   Ragava
                RodCodeNumber = objDT.Rows(0).Item("PartNumber").ToString          '01_10_2009   ragava
                '05_09_2009  ragava
                If Trim(ofrmTieRod1.cmbStyle.Text) = "ASAE" Then
                    Try
                        'dblRodLength
                        If Val(objDT.Rows(0).Item("StrokeLength").ToString) = Val(ofrmTieRod1.cmbStrokeLength.Text) Then
                            dblRodLength = Val(objDT.Rows(0).Item("OverAllRodLength").ToString)
                        ElseIf Val(ofrmTieRod1.cmbStrokeLength.Text) > Val(objDT.Rows(0).Item("StrokeLength").ToString) Then
                            dblRodLength = Val(objDT.Rows(0).Item("OverAllRodLength").ToString) + 11.25
                        ElseIf Val(ofrmTieRod1.cmbStrokeLength.Text) < Val(objDT.Rows(0).Item("StrokeLength").ToString) Then
                            dblRodLength = Val(objDT.Rows(0).Item("OverAllRodLength").ToString) - 11.25
                        End If
                    Catch ex As Exception

                    End Try
                End If
                '05_09_2009  ragava  Till  Here
            End If
        Catch ex As Exception

        End Try
        'Bore Drawing Number
        Try
            'Dim oListViewItem As ListViewItem
            Dim StrSql As String
            Dim arrSeries As String()
            'If ofrmTieRod1.cmbSeries.Text.ToString.StartsWith("TP") Then
            '    arrSeries = ofrmTieRod1.cmbSeries.Text.ToString.Split("-")
            'Else
            '    arrSeries = ofrmTieRod1.cmbSeries.Text.ToString.Split(" ")
            'End If
            arrSeries = SeriesForCosting.ToString.Split(" ")
            Dim strSeries As String = arrSeries(0)
            '01_09_2009  ragava
            If Trim(ofrmTieRod1.cmbRephasingPortPosition.Text) <> "" Then
                If Trim(ofrmTieRod1.cmbRephasingPortPosition.Text) = "At Extension" Then
                    strSeries = strSeries.Insert(2, "E")
                ElseIf Trim(ofrmTieRod1.cmbRephasingPortPosition.Text) = "At Retraction" Then
                    strSeries = strSeries.Insert(2, "B")
                Else
                    strSeries = strSeries.Insert(2, "2")
                End If
            End If
            StrSql = "select * from BoreDiameterDetails where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) _
                & " and Series like '%" & strSeries & "%'"
            Dim objDT As DataTable = oDataClass.GetDataTable(StrSql)
            If objDT.Rows.Count > 0 Then
                BoreStrokeDifference = Math.Round(Val(objDT.Rows(0).Item("TubeLength").ToString) _
                    - Val(objDT.Rows(0).Item("NominalStroke").ToString), 2)
                updateHashTablevalues(htMainParams, "Bore Drawing Number", objDT.Rows(0).Item("DrawingPartNumber").ToString)
                BoreDrawingNumber = objDT.Rows(0).Item("DrawingPartNumber").ToString        '24_09_2009   Ragava
            End If
        Catch ex As Exception

        End Try
        'TieRod Drawing Number
        Try
            Dim StrSql As String
            'ANUP 12-10-2010 START
            StrSql = "select DrawingNumber,StrokeLength,[Dimension-A] from TieRodSizes trz,BoreDiameter_TieRodSizes bdtr where bdtr.PartNumberID = trz.TieRodPartNumber and bdtr.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ")and trz.series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "'"
            'ANUP 12-10-2010 TILL HERE
            Dim objDT As DataTable = oDataClass.GetDataTable(StrSql)
            If objDT.Rows.Count > 0 Then
                'TieRodStrokeDifference = Math.Round(Val(objDT.Rows(0).Item("StrokeLength").ToString) - Val(objDT.Rows(0).Item("Dimension-A").ToString), 2)
                TieRodStrokeDifference = Math.Round(Val(objDT.Rows(0).Item("Dimension-A").ToString) _
                        - Val(objDT.Rows(0).Item("StrokeLength").ToString), 2) 'StrokeLength
                TieRodDrawingNumber = objDT.Rows(0).Item("DrawingNumber").ToString        '24_09_2009   Ragava
                updateHashTablevalues(htMainParams, "Tie Rod Drawing Number", objDT.Rows(0).Item("DrawingNumber").ToString)
            End If
        Catch ex As Exception

        End Try
        updateHashTablevalues(htMainParams, "Tube Length", StrokeLength + BoreStrokeDifference + StopTubeLength)
        dblTubeStrokeDifference = BoreStrokeDifference         '24_09_2009   ragava
        TubeLength = StrokeLength + BoreStrokeDifference + StopTubeLength          '16_09_2009  ragava
        'updateHashTablevalues(htMainParams, "Rod Length", StrokeLength + RodStrokeDifference + StopTubeLength + RodAdder + IIf(Trim(ofrmTieRod1.cmbStyle.Text) = "ASAE", 2, 0))
        'updateHashTablevalues(htMainParams, "Rod Length", StrokeLength + RodStrokeDifference + StopTubeLength + RodAdder)
        '05_09_2009  ragava
        If Trim(ofrmTieRod1.cmbStyle.Text) = "ASAE" Then
            updateHashTablevalues(htMainParams, "Rod Length", dblRodLength)
            RodLength = dblRodLength        '16_09_2009  ragava
        Else
            updateHashTablevalues(htMainParams, "Rod Length", StrokeLength + RodStrokeDifference + StopTubeLength + RodAdder)
            RodLength = StrokeLength + RodStrokeDifference + StopTubeLength + RodAdder        '16_09_2009  ragava
        End If
        dblRodStrokeDifference = RodStrokeDifference         '22_09_2009  ragava
        updateHashTablevalues(htMainParams, "Tie Rod Length", StrokeLength + TieRodStrokeDifference + StopTubeLength)
        TieRodLength = StrokeLength + TieRodStrokeDifference + StopTubeLength    '16_09_2009   ragava
        dblTieRodStrokeDifference = TieRodStrokeDifference         '24_09_2009   ragava

    End Sub

    Public Sub updateGUIInputParameters()

        oExcelClass.updateGUIParameters("Contract Details", "")
        oExcelClass.updateGUIParameters("CustomerName", Trim(ofrmContractDetails.cmbCustomerName.Text))           '22_02_2010   RAGAVA
        oExcelClass.updateGUIParameters("Customer Part Number", ofrmContractDetails.txtlPartCode.Text)
        oExcelClass.updateGUIParameters("Tie Rod 1 Screen Details", "")
        oExcelClass.updateGUIParameters("Series", SeriesForCosting)
        oExcelClass.updateGUIParameters("Rephasing Port Position", IIf(Trim(ofrmTieRod1.cmbRephasingPortPosition.Text) _
                        = "", "None", ofrmTieRod1.cmbRephasingPortPosition.Text))
        oExcelClass.updateGUIParameters("Port Orientation at Rod End", ofrmTieRod1.cmbPortOrientationForRodCap.Text)
        oExcelClass.updateGUIParameters("Style", ofrmTieRod1.cmbStyle.Text)
        oExcelClass.updateGUIParameters("Bore", ofrmTieRod1.cmbBore.Text)
        oExcelClass.updateGUIParameters("Stroke Length", StrokeLength)
        oExcelClass.updateGUIParameters("Rod Adder", RodAdder)
        oExcelClass.updateGUIParameters("Stop Tube", IIf(ofrmTieRod1.rdbStopTubeYes.Checked = True, "Yes", "No"))
        oExcelClass.updateGUIParameters("Stop Tube Length", StopTubeLength)
        oExcelClass.updateGUIParameters("Recommended Stop Tube Length", Val(ofrmTieRod1.txtRecommendedStoptubeLength.Text))
        oExcelClass.updateGUIParameters("Pin Hole", ofrmTieRod1.cmbClevisCapPinHole.Text)
        oExcelClass.updateGUIParameters("Retracted Length", ofrmTieRod1.txtRetractedLength.Text)
        oExcelClass.updateGUIParameters("Extended Length", ofrmTieRod1.txtExtendedLength.Text)
        oExcelClass.updateGUIParameters("Rod Material", ofrmTieRod1.cmbRodMaterial.Text)
        oExcelClass.updateGUIParameters("Rod Diameter", RodDiameter)
        oExcelClass.updateGUIParameters("Port Orientation for Clevis Cap", ofrmTieRod1.cmbPortOrientation.Text)
        oExcelClass.updateGUIParameters("Port Orientation for Rod Cap", ofrmTieRod1.cmbPortOrientationForRodCap.Text)
        oExcelClass.updateGUIParameters("Clevis Cap Port", ofrmTieRod1.cmbClevisCapPort.Text)
        oExcelClass.updateGUIParameters("Rod Cap Port", ofrmTieRod1.cmbRodCapPort.Text)
        oExcelClass.updateGUIParameters("Nut Size", ofrmTieRod2.txtTieRodNutSize.Text)
        oExcelClass.updateGUIParameters("Stroke Control", IIf(ofrmTieRod1.optStrokeControlYes.Checked = True, "Yes", "No"))
        oExcelClass.updateGUIParameters("Stroke Length Adder", ofrmTieRod1.cmbStrokeLengthAdder.Text)
        oExcelClass.updateGUIParameters("Tie Rod 2 Screen Details", "")
        oExcelClass.updateGUIParameters("Pins", IIf(ofrmTieRod2.optPinsYes.Checked = True, "Yes", "No"))
        oExcelClass.updateGUIParameters("Rod Clevis Pins", IIf(ofrmTieRod2.optPinsYes_Rod.Checked = True, "Yes", "No"))        '05_04_2010    RAGAVA
        '14_04_2011  RAGAVA
        Try
            If Trim(ofrmTieRod1.cmbClevisCapPort.Text) <> "" Then
                Dim strClevisPortType As String = Trim(ofrmTieRod1.cmbClevisCapPort.Text).ToString. _
                            Substring(Trim(ofrmTieRod1.cmbClevisCapPort.Text).LastIndexOf(" ") + 1, 3)
                oExcelClass.updateGUIParameters("Base End Port Type", strClevisPortType)
                oExcelClass.updateGUIParameters("Base End Port Size", "'" & Trim(ofrmTieRod1. _
                        cmbClevisCapPort.Text).ToString.Substring(0, Trim(ofrmTieRod1.cmbClevisCapPort.Text).ToString.IndexOf(" ")))         '19_04_2011   RAGAVA
            End If
            If Trim(ofrmTieRod1.cmbRodCapPort.Text) <> "" Then
                Dim strRodPortType As String = Trim(ofrmTieRod1.cmbRodCapPort.Text).ToString. _
                            Substring(Trim(ofrmTieRod1.cmbRodCapPort.Text).LastIndexOf(" ") + 1, 3)
                oExcelClass.updateGUIParameters("Rod End Port Type", strRodPortType)
                oExcelClass.updateGUIParameters("Rod End Port Size", "'" & Trim(ofrmTieRod1. _
                    cmbRodCapPort.Text).ToString.Substring(0, Trim(ofrmTieRod1.cmbRodCapPort.Text).ToString.IndexOf(" ")))         '19_04_2011   RAGAVA
            End If
        Catch ex As Exception
        End Try
        'Till  Here
        oExcelClass.updateGUIParameters("Pin Size", PinSize)
        oExcelClass.updateGUIParameters("Pin Material", ofrmTieRod2.cmbPinMaterial.Text)
        oExcelClass.updateGUIParameters("Clevis Cap Pin Clips", ofrmTieRod2.cmbClips.Text)
        oExcelClass.updateGUIParameters("Tie Rod Size", ofrmTieRod2.txtTieRodSize.Text)


        oExcelClass.updateGUIParameters("Tie Rod Nut Code", ofrmTieRod2.txtTieRodNutSize.Text)
        ObjClsCostingDetails.AddCodeNumberToDataTable(ofrmTieRod2.txtTieRodNutSize.Text, "Tie Rod Nut Code") 'Sandeep 04-03-10-4pm

        oExcelClass.updateGUIParameters("Tie Rod Nut Qty", ofrmTieRod2.txtTieRodNutQty.Text)
        oExcelClass.updateGUIParameters("Thred Protector", ofrmTieRod2.cmbThreadProtected.Text)
        oExcelClass.updateGUIParameters("Rod Seal Package", ofrmTieRod2.cmbRodSealPackage.Text)


        oExcelClass.updateGUIParameters("Rod Cap", ofrmTieRod2.txtRodCap.Text)
        ObjClsCostingDetails.AddCodeNumberToDataTable(ofrmTieRod2.txtRodCap.Text, "Rod Cap Code") 'Sandeep 04-03-10-4pm
        GetSealWiper(ofrmTieRod2.txtRodCap.Text) 'TODO:Sandeep 20-04-10-2pm


        oExcelClass.updateGUIParameters("Clevis Cap", ofrmTieRod2.txtClevisCap.Text)
        ObjClsCostingDetails.AddCodeNumberToDataTable(ofrmTieRod2.txtClevisCap.Text, "Clevis Cap Code") 'Sandeep 04-03-10-4pm

        oExcelClass.updateGUIParameters("Rod End Thread Size", ofrmTieRod2.cmbRodEndThread.Text)
        oExcelClass.updateGUIParameters("Rod Clevis Check", IIf(ofrmTieRod2.rdbRodClevisYes.Checked = True, "Yes", "No"))

        If ofrmTieRod2.rdbRodClevisYes.Checked = True Then      '18_10_2011   RAGAVA
            oExcelClass.updateGUIParameters("Rod Clevis", strRodClevisCodeNumber)
            ObjClsCostingDetails.AddCodeNumberToDataTable(strRodClevisCodeNumber, "Rod Clevis Code") 'Sandeep 04-03-10-4pm
            GetScrewPartNumber(strRodClevisCodeNumber) 'TODO:Sandeep 20-04-10-2pm
        End If
        If ofrmTieRod2.rdbRodClevisYes.Checked = True Then     '18_10_2011   RAGAVA
            oExcelClass.updateGUIParameters("Rod Clevis Pin Clips", strRodClevisPinCodeNumber)
        End If
        'Sandeep 18-03-10-2pm
        Try

            '16_06_2011  RAGAVA
            If blnInstallPinsandClips_Checked = True AndAlso (strBaseEndKitCode <> "" OrElse strRodEndKitCode <> "") Then
                If strBaseEndKitCode = strRodEndKitCode Then
                    'ObjClsCostingDetails.AddCodeNumberToDataTable(strBaseEndKitCode, "BASE/ROD END KIT", 2)
                    ObjClsCostingDetails.AddCodeNumberToDataTable(strBaseEndKitCode, "BASE/ROD END KIT", 1)      '05_07_2011  RAGAVA
                    GoTo ESC_Pin_And_clips
                End If
                If strBaseEndKitCode <> "" Then
                    ObjClsCostingDetails.AddCodeNumberToDataTable(strBaseEndKitCode, "BASE END KIT", 1)
                End If
                If strRodEndKitCode <> "" Then
                    ObjClsCostingDetails.AddCodeNumberToDataTable(strRodEndKitCode, "ROD END KIT", 1)
                End If
ESC_Pin_And_clips:
                'TILL   HERE
            ElseIf strRodClevisPinCodeNumber.Equals(strClevisCapPinCodeNumber) Then
                Dim strCodeNumbers As String() = strRodClevisPinCodeNumber.Split("-")
                If strCodeNumbers.Length > 1 Then
                    ObjClsCostingDetails.AddCodeNumberToDataTable(strCodeNumbers(Pin_Clip.Pin), "Rod Clevis Pin", 2)
                    ObjClsCostingDetails.AddCodeNumberToDataTable(strCodeNumbers(Pin_Clip.Clip), "Rod Clevis Clip", 4)
                End If
            Else

                Dim strRodClevisCodeNumbers As String() = Nothing
                Dim strClevisCapCodeNumbers As String() = Nothing

                If Not IsNothing(strRodClevisPinCodeNumber) AndAlso Not IsNothing(strClevisCapPinCodeNumber) Then
                    strRodClevisCodeNumbers = strRodClevisPinCodeNumber.Split("-")
                    strClevisCapCodeNumbers = strClevisCapPinCodeNumber.Split("-")
                    If strRodClevisCodeNumbers(Pin_Clip.Pin).Equals(strClevisCapCodeNumbers(Pin_Clip.Pin)) Then
                        ObjClsCostingDetails.AddCodeNumberToDataTable(strRodClevisCodeNumbers(Pin_Clip.Pin), "Rod Clevis Pin", 2)
                        ObjClsCostingDetails.AddCodeNumberToDataTable(strRodClevisCodeNumbers(Pin_Clip.Clip), "Rod Clevis Clip", 2)
                        ObjClsCostingDetails.AddCodeNumberToDataTable(strClevisCapCodeNumbers(Pin_Clip.Clip), "Clevis Cap Clip", 2)
                    ElseIf strRodClevisCodeNumbers(Pin_Clip.Clip).Equals(strClevisCapCodeNumbers(Pin_Clip.Clip)) Then
                        ObjClsCostingDetails.AddCodeNumberToDataTable(strRodClevisCodeNumbers(Pin_Clip.Pin), "Rod Clevis Pin", 1)
                        ObjClsCostingDetails.AddCodeNumberToDataTable(strClevisCapCodeNumbers(Pin_Clip.Pin), "Clevis Cap Pin", 1)
                        ObjClsCostingDetails.AddCodeNumberToDataTable(strRodClevisCodeNumbers(Pin_Clip.Clip), "Rod Clevis Clip", 4)
                    Else
                        ObjClsCostingDetails.AddCodeNumberToDataTable(strRodClevisCodeNumbers(Pin_Clip.Pin), "Rod Clevis Pin", 1)
                        ObjClsCostingDetails.AddCodeNumberToDataTable(strRodClevisCodeNumbers(Pin_Clip.Clip), "Rod Clevis Clip", 2)
                        ObjClsCostingDetails.AddCodeNumberToDataTable(strClevisCapCodeNumbers(Pin_Clip.Pin), "Clevis Cap Pin", 1)
                        ObjClsCostingDetails.AddCodeNumberToDataTable(strClevisCapCodeNumbers(Pin_Clip.Clip), "Clevis Cap Clip", 2)
                    End If
                ElseIf Not IsNothing(strRodClevisPinCodeNumber) AndAlso IsNothing(strClevisCapPinCodeNumber) Then
                    strRodClevisCodeNumbers = strRodClevisPinCodeNumber.Split("-")
                    ObjClsCostingDetails.AddCodeNumberToDataTable(strRodClevisCodeNumbers(Pin_Clip.Pin), "Rod Clevis Pin", 1)
                    ObjClsCostingDetails.AddCodeNumberToDataTable(strRodClevisCodeNumbers(Pin_Clip.Clip), "Rod Clevis Clip", 2)
                ElseIf IsNothing(strRodClevisPinCodeNumber) AndAlso Not IsNothing(strClevisCapPinCodeNumber) Then
                    strClevisCapCodeNumbers = strClevisCapPinCodeNumber.Split("-")
                    ObjClsCostingDetails.AddCodeNumberToDataTable(strClevisCapCodeNumbers(Pin_Clip.Pin), "Clevis Cap Pin", 1)
                    ObjClsCostingDetails.AddCodeNumberToDataTable(strClevisCapCodeNumbers(Pin_Clip.Clip), "Clevis Cap Clip", 2)
                End If

            End If
        Catch ex As Exception

        End Try

        oExcelClass.updateGUIParameters("Piston Seal Package", ofrmTieRod2.cmbPistonSealPackage.Text)
        oExcelClass.updateGUIParameters("Paint", ofrmTieRod2.cmbPaint.Text)
        oExcelClass.updateGUIParameters("Packaging", ofrmTieRod2.txtPackaging.Text)
        'anup 31-01-2011 start
        'oExcelClass.updateGUIParameters("Rod Wiper", ofrmTieRod2.txtRodWiper.Text)
        oExcelClass.updateGUIParameters("Rod Wiper", ofrmTieRod2.cmbRodWiper.Text)
        'anup 31-01-2011 till here
        oExcelClass.updateGUIParameters("Tube Seal", ofrmTieRod2.txtTubeSeal1.Text)
        oExcelClass.updateGUIParameters("Tie Rod 3 Screen Details", "")
        Dim oCtl As Control
        For Each oCtl In ofrmTieRod3.Controls
            If TypeOf (oCtl) Is GroupBox Then
                Dim oCtl1 As Control
                oCtl1 = DirectCast(oCtl, GroupBox)
                For Each octl2 As Control In oCtl1.Controls
                    If TypeOf (octl2) Is CheckBox Then
                        If DirectCast(octl2, CheckBox).Checked = True Then
                            oExcelClass.updateGUIParameters(octl2.Text, "Selected")
                        End If
                    End If
                Next
            End If
        Next

        Try
            SetCodeNumbersToReport()

            'TODO: Sunny 20-04-10 10am
            oExcelClass.updateGUIParameters("Rod Code Number", strRodCodeNumber)
            Dim strQuery_RodMaterial As String
            If strRodCodeNumber.StartsWith(7) Then
                strQuery_RodMaterial = "select distinct(RodMaterialNumber) from RodDiameterDetails where "
                strQuery_RodMaterial += "RodDiameter = " + RodDiameter.ToString + " and MaterialType= '" + RodMaterialForCosting + "'"
                _strRodMaterialCode_Costing = IFLConnectionObject.GetValue(strQuery_RodMaterial)
            End If
            ObjClsCostingDetails.AddCodeNumberToDataTable(strRodCodeNumber, "Rod Code Number") 'Sandeep 04-03-10-4pm
            '**************************

            'TODO: Sunny 20-04-10 10am
            oExcelClass.updateGUIParameters("Tube Code Number", strBoreCodeNumber)
            Dim strQuery_TubeMaterial As String
            If strBoreCodeNumber.StartsWith(7) Then
                If SeriesForCosting.Contains("TX") Then
                    'strQuery_TubeMaterial = "select distinct(TubeMaterialNumber) from BoreDiameterDetails where BoreDiameter = " _
                    '                                        + BoreDiameter.ToString + " and Series = '" + SeriesForCosting + "'"
                    strQuery_TubeMaterial = "select distinct(TubeMaterialNumber) from BoreDiameterDetails where BoreDiameter = " _
                                                            + BoreDiameter.ToString + " and Series like 'TX%'"        '26_11_2010   RAGAVA
                Else
                    strQuery_TubeMaterial = "select distinct(TubeMaterialNumber) from BoreDiameterDetails where BoreDiameter = " _
                                                                         + BoreDiameter.ToString + " and Series <> '" + SeriesForCosting + "'"
                End If
                _strTubeMaterialCode_Costing = IFLConnectionObject.GetValue(strQuery_TubeMaterial)
            End If

            'Sunny 21-04-10
            '08_10_2010   RAGAVA  Commented
            'If SeriesForCosting.Contains("TP") Then
            '    If strRephasing.Contains("Both") Then
            '        ObjClsCostingDetails.AddCodeNumberToDataTable(469832, "Rephase Port Code Number", 2)
            '    Else
            '        ObjClsCostingDetails.AddCodeNumberToDataTable(469832, "Rephase Port Code Number")
            '    End If
            'End If
            ObjClsCostingDetails.AddCodeNumberToDataTable(strBoreCodeNumber, "Tube Code Number") 'Sandeep 04-03-10-4pm
            '**************************

            oExcelClass.updateGUIParameters("Stop Tube Code Number", StopTubeCodeNumber)
            ObjClsCostingDetails.AddCodeNumberToDataTable(StopTubeCodeNumber, "Stop Tube Code Number") 'Sandeep 04-03-10-4pm

            'Sunny 27-04-10
            Dim strQuery As String = "select Dim_C, Dim_B from StopTubeTableDrawing where CodeNumber = '" + StopTubeCodeNumber + "'"
            Dim oDRStopTubeOD_ID As DataRow = IFLConnectionObject.GetDataRow(strQuery)
            If Not IsNothing(oDRStopTubeOD_ID) Then
                StopTubeID = oDRStopTubeOD_ID(0)
                StopTubeOD = oDRStopTubeOD_ID(1)
            End If
            '**************************
            oExcelClass.updateGUIParameters("Tie Rod Code Number", strTieRodCodeNumber)
            ObjClsCostingDetails.AddCodeNumberToDataTable(strTieRodCodeNumber, "Tie Rod Code Number") 'Sandeep 04-03-10-4pm

            GetPistonOringDetails() 'TODO: Sunny 20-04-10 10am
        Catch ex As Exception

        End Try

    End Sub

    Private Sub GetPistonOringDetails()

        Dim strPistonOringCode As String
        If SeriesForCosting.Contains("TX") Then
            If BoreDiameter <= 2.5 Then
                strPistonOringCode = 197601
            Else
                strPistonOringCode = 197600
            End If
            ObjClsCostingDetails.AddCodeNumberToDataTable(strPistonOringCode, "Piston Oring Code") 'Sandeep 04-03-10-4pm
        End If

    End Sub

    'TODO:Sandeep 20-04-10-4pm
    Private Sub GetScrewPartNumber(ByVal strRodClevisCodeNumber As String)

        Try
            Dim strQuery As String = "select * from RodClevisDetails where PartNumber = '" + strRodClevisCodeNumber + "'"
            Dim oDRRodClevisDetails As DataRow = IFLConnectionObject.GetDataRow(strQuery)
            If Not IsNothing(oDRRodClevisDetails) Then
                If Not IsDBNull(oDRRodClevisDetails("SetScrewPartNumber")) AndAlso oDRRodClevisDetails("SetScrewPartNumber") <> "" Then
                    ObjClsCostingDetails.AddCodeNumberToDataTable(oDRRodClevisDetails("SetScrewPartNumber"), "Set Screw Part Number") 'Sandeep 20-04-10-2pm
                Else
                    If Not IsDBNull(oDRRodClevisDetails("BoltNumber")) AndAlso oDRRodClevisDetails("SetScrewPartNumber") <> "" Then
                        ObjClsCostingDetails.AddCodeNumberToDataTable(oDRRodClevisDetails("BoltNumber"), "Bolt Number") 'Sandeep 20-04-10-2pm
                    End If
                    If Not IsDBNull(oDRRodClevisDetails("NutNumber")) AndAlso oDRRodClevisDetails("SetScrewPartNumber") <> "" Then
                        ObjClsCostingDetails.AddCodeNumberToDataTable(oDRRodClevisDetails("NutNumber"), "Nut Number") 'Sandeep 20-04-10-2pm
                    End If
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    'Sandeep 20-04-10-2pm
    Private Sub GetSealWiper(ByVal strRodCapCode As String)

        Try
            'anup 01-02-2011 start
            Dim strQuery As String = String.Empty
            If SeriesForCosting = "LN" Then
                strQuery = "select WiperCodeNumber as SealWiper, Wiper_Description, RodDiameter,Series from dbo.RodWiperDetails where Series='LN' and Wiper_Description='" + ofrmTieRod2.cmbRodWiper.Text + "'"
            Else
                strQuery = "select * from RodCapDetails where PartNumber = '" + strRodCapCode + "'"
            End If

            Dim oDRRodCapDetails As DataRow = IFLConnectionObject.GetDataRow(strQuery)
            If Not IsNothing(oDRRodCapDetails) AndAlso oDRRodCapDetails.ItemArray.Length > 0 Then
                If Not IsDBNull(oDRRodCapDetails("SealWiper")) AndAlso oDRRodCapDetails("SealWiper") <> "" Then
                    ObjClsCostingDetails.AddCodeNumberToDataTable(oDRRodCapDetails("SealWiper"), "Seal Wiper") 'Sandeep 20-04-10-2pm
                End If
            End If

            'anup 01-02-2011 till here
        Catch ex As Exception

        End Try

    End Sub

    Public Sub InitialiseAllChildFormObjects()

        '_ofrmMdiMonarch = New mdiMonarch
        _ofrmMonarch = New frmMonarch
        _ofrmContractDetails = New frmContractDetails
        _ofrmTieRod1 = New frmTieRod1
        _ofrmTieRod2 = New frmTieRod2
        _ofrmTieRod3 = New frmTieRod3

    End Sub

    Public Property FormNavigationOrder() As ArrayList

        Get
            If _aFormNavigationOrder Is Nothing Then
                _aFormNavigationOrder = New ArrayList
                _aFormNavigationOrder.Clear()
                _aFormNavigationOrder.Add(New Object(3) {"frmMonarch", _ofrmMonarch, Nothing, _ofrmContractDetails})
                _aFormNavigationOrder.Add(New Object(3) {"frmContractDetails", _ofrmContractDetails, _ofrmMonarch, _ofrmTieRod1})
                _aFormNavigationOrder.Add(New Object(3) {"frmTieRod1", _ofrmTieRod1, _ofrmContractDetails, _ofrmTieRod2})
                _aFormNavigationOrder.Add(New Object(3) {"frmTieRod2", _ofrmTieRod2, _ofrmTieRod1, _ofrmTieRod3})
                _aFormNavigationOrder.Add(New Object(3) {"frmTieRod3", _ofrmTieRod3, _ofrmTieRod2, Nothing})
            End If
            Return _aFormNavigationOrder
        End Get
        Set(ByVal value As ArrayList)
            _aFormNavigationOrder = value
        End Set

    End Property

    Public Sub GenerateNotes()
        Dim iNotes As Integer = 0       '28_10_2009   Ragava
        '19_10_2009   ragava 
        Dim HT_AssemblyNotes As New Hashtable
        Dim HT_PaintNotes As New Hashtable
        Try
            If ofrmTieRod3.ChkPins.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtPins.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtPins.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.ChkPins.Text))
            End If
            '28_10_2009   Ragava
            If Trim(ofrmTieRod3.txtRetractedLength.Text) <> "" Then
                iNotes = Val(ofrmTieRod3.txtRetractedLength.Text)
            Else
                iNotes += 1
            End If
            '28_10_2009   Ragava   Till  Here
            'HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.ChkRetractedLength.Text) & " = " & ofrmTieRod1.txtRetractedLength.Text)       '21_10_2009   ragava
            HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.ChkRetractedLength.Text) & " ______________________")        '11_11_2009   ragava
            '28_10_2009   Ragava
            If Trim(ofrmTieRod3.txtExtenedLength.Text) <> "" Then
                iNotes = Val(ofrmTieRod3.txtExtenedLength.Text)
            Else
                iNotes += 1
            End If
            '28_10_2009   Ragava   Till  Here
            HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.ChkExtendedLength.Text) & "    ______________________")        '11_11_2009   ragava
            '28_10_2009   Ragava
            If Trim(ofrmTieRod3.txtRodDiameter.Text) <> "" Then
                iNotes = Val(ofrmTieRod3.txtRodDiameter.Text)
            Else
                iNotes += 1
            End If
            '28_10_2009   Ragava   Till  Here
            HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.ChkRodDiameter.Text) & "            ______________________")        '11_11_2009   ragava
            If ofrmTieRod3.ChkPorts.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtPorts.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtPorts.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.ChkPorts.Text))
            End If
            'If ofrmTieRod3.chk100AirTest.Checked = True Then  VAMSI 10-09-2014
            '    '28_10_2009   Ragava
            '    If Trim(ofrmTieRod3.txtAirTest.Text) <> "" Then
            '        iNotes = Val(ofrmTieRod3.txtAirTest.Text)
            '    Else
            '        iNotes += 1
            '    End If
            '    '28_10_2009   Ragava   Till  Here
            '    HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.chk100AirTest.Text))
            'End If
            If ofrmTieRod3.chk100OilTest.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtOilTest.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtOilTest.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.chk100OilTest.Text))
            End If
            If ofrmTieRod3.ChkRephaseExtension.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtRephaseOnExtension.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtRephaseOnExtension.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.ChkRephaseExtension.Text))
            End If
            If ofrmTieRod3.ChkRephaseRetraction.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtRephaseOnRetraction.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtRephaseOnRetraction.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.ChkRephaseRetraction.Text))
            End If
            If ofrmTieRod3.chkInstallStrokeControl.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtInstallStrokeLength.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtInstallStrokeLength.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.chkInstallStrokeControl.Text))
            End If
            If ofrmTieRod3.chkStampCustomerPartandDate.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtStampCustomerPartandDate.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtStampCustomerPartandDate.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.chkStampCustomerPartandDate.Text))
            End If
            If ofrmTieRod3.chkStampCustomerPartOnTube.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtStampCustomerPart.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtStampCustomerPart.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.chkStampCustomerPartOnTube.Text))
            End If
            If ofrmTieRod3.ChkStampCountryOfOrigin.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtStampCountry.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtStampCountry.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.ChkStampCountryOfOrigin.Text))
            End If
            If ofrmTieRod3.ChkRodMaterial.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtRodMaterialNitroSteel.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtRodMaterialNitroSteel.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.ChkRodMaterial.Text))
            End If
            If ofrmTieRod3.ChkInstallSteelPlugs.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtInstallSteelPlugs.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtInstallSteelPlugs.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.ChkInstallSteelPlugs.Text))
            End If
            'If ofrmTieRod3.ChkInstallHardenedBushings.Checked = True Then
            '    '28_10_2009   Ragava
            '    If Trim(ofrmTieRod3.txtInstallHardenedBushings.Text) <> "" Then
            '        iNotes = Val(ofrmTieRod3.txtInstallHardenedBushings.Text)
            '    Else
            '        iNotes += 1
            '    End If
            '    '28_10_2009   Ragava   Till  Here
            '    HT_AssemblyNotes.Add(iNotes, UCase(ofrmTieRod3.ChkInstallHardenedBushings.Text))
            'End If
            If ofrmTieRod3.ChkHardenedBushingsRodClevisEnd.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtInstallHardenedBushingsRodClevis.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtInstallHardenedBushingsRodClevis.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.ChkHardenedBushingsRodClevisEnd.Text))
            End If
            If ofrmTieRod3.ChkHardenedBushingsClevisCapEnd.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtInstallHardenedBushingsClevisCap.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtInstallHardenedBushingsClevisCap.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.ChkHardenedBushingsClevisCapEnd.Text))
            End If
            '02_11_2009   Ragava
            'If ofrmTieRod3.chkAffixLabelToBag.Checked = True Then
            '    '28_10_2009   Ragava
            '    If Trim(ofrmTieRod3.txtAffixLabeltoBag.Text) <> "" Then
            '        iNotes = Val(ofrmTieRod3.txtAffixLabeltoBag.Text)
            '    Else
            '        iNotes += 1
            '    End If
            '    '28_10_2009   Ragava   Till  Here
            '    HT_AssemblyNotes.Add(iNotes, UCase(ofrmTieRod3.chkAffixLabelToBag.Text))
            'End If
            '02_11_2009  Ragava   Till   Here
            If ofrmTieRod3.ChkAssemblyStopTube.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtAssemblyStopTube.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtAssemblyStopTube.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.ChkAssemblyStopTube.Text))
            End If

            '20_10_2009  ragava
            If ofrmTieRod3.chkAssemblyNotes.Checked = True AndAlso Trim(ofrmTieRod3.RichTextBox1.Text) <> "" Then
                Dim strText As String() = ofrmTieRod3.RichTextBox1.Lines
                For Each str As String In strText
                    If Trim(str) <> "" Then
                        Dim strSplit() As String
                        strSplit = str.Split("}")
                        Dim strNumber As String = strSplit(LBound(strSplit))
                        '28_10_2009   Ragava
                        Try
                            If strNumber = "" Then
                                iNotes += 1
                                HT_AssemblyNotes.Add(iNotes.ToString & ".0", UCase(strSplit(UBound(strSplit))))
                            Else
                                HT_AssemblyNotes.Add(strNumber & ".0", UCase(strSplit(UBound(strSplit))))
                            End If
                        Catch ex As Exception
                        End Try
                    End If
                Next
            End If
            Dim AssyNotesCount As Integer = HT_AssemblyNotes.Keys.Count
            For Each Item As Object In HT_AssemblyNotes.Keys
                If AssyNotesCount < Val(Item) Then
                    AssyNotesCount = Val(Item)
                End If
            Next
            '20_10_2009  ragava  Till  Here
            Dim iCount As Integer = 1
            strAssemblyNotes = ""
            For iCount = 1 To AssyNotesCount
                'If Trim(HT_AssemblyNotes(Val(iCount.ToString))) <> "" Then
                If HT_AssemblyNotes.ContainsKey(iCount.ToString & ".0") = True Then             '04_11_2009  Ragava
                    strAssemblyNotes = strAssemblyNotes & iCount.ToString & ". " & _
                                Trim(HT_AssemblyNotes((iCount.ToString & ".0"))) & vbNewLine
                End If
            Next
            iAssyNotesCount = HT_AssemblyNotes.Keys.Count         '20_10_2009   RAGAVA

            '*********  END OF ASSEMBLY NOTES GENERATION **********************************************

            'PAINT NOTES GENERATION
            iCount = 1
            iNotes = 0         '28_10_2009   Ragava
            '29_04_2011  RAGAVA
            If ofrmTieRod3.chkMaskPerBOM.Checked = True Then
                If Trim(ofrmTieRod3.txtMaskPerBOM.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtMaskPerBOM.Text)
                Else
                    iNotes += 1
                End If
                HT_PaintNotes.Add(iNotes.ToString & ".0", ofrmTieRod3.chkMaskPerBOM.Text)
                iCount += 1
            End If
            'Till  Here

            If ofrmTieRod3.ChkMaskBushings.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtMaskBushings.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtMaskBushings.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_PaintNotes.Add(iNotes.ToString & ".0", ofrmTieRod3.ChkMaskBushings.Text)
                iCount += 1
            End If
            If ofrmTieRod3.ChkMaskExposedThreads.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtMaskExposedThreads.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtMaskExposedThreads.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_PaintNotes.Add(iNotes.ToString & ".0", ofrmTieRod3.ChkMaskExposedThreads.Text)
                iCount += 1
            End If
            If ofrmTieRod3.ChkMaskPinHoles.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtMaskPinholes.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtMaskPinholes.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_PaintNotes.Add(iNotes.ToString & ".0", ofrmTieRod3.ChkMaskPinHoles.Text)
                iCount += 1
            End If
            If ofrmTieRod3.ChkPrime.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtPrime.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtPrime.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_PaintNotes.Add(iNotes.ToString & ".0", ofrmTieRod3.ChkPrime.Text)
                iCount += 1
            End If
            If ofrmTieRod3.ChkPaint.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtPaint.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtPaint.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                'HT_PaintNotes.Add(iNotes.ToString & ".0", ofrmTieRod3.ChkPaint.Text & " " & _
                '    UCase(ofrmTieRod2.cmbPaint.Text.ToString.Substring(0, ofrmTieRod2.cmbPaint.Text.ToString.IndexOf("-"))))          '21_10_2009  ragava
                'iCount += 1
                Dim strQuery As String
                Dim PaintDescription As String
                strQuery = "select description from PaintDetails where PaintColor='" & ofrmTieRod2.cmbPaint.Text & "'"
                PaintDescription = IFLConnectionObject.GetValue(strQuery)



                HT_PaintNotes.Add(iNotes.ToString & ".0", ofrmTieRod3.ChkPaint.Text & " " & UCase(PaintDescription))
                iCount += 1
            End If
            If ofrmTieRod3.chkAffixLabel.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtAffixLabel.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtAffixLabel.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_PaintNotes.Add(iNotes.ToString & ".0", ofrmTieRod3.chkAffixLabel.Text)
                iCount += 1
            End If

            ''23_19_2011   Ragava
            'If ofrmTieRod3.chkAffixLabel.Checked = True Then
            '    iNotes += 1
            '    HT_PaintNotes.Add(iNotes.ToString & ".0", "INCLUDE PIN KIT PER BOM")
            '    iCount += 1
            'End If
            'Till   Here
            If ofrmTieRod3.chkInstallPinAndClips.Checked = True Then
                '28_10_2009   Ragava
                If Trim(ofrmTieRod3.txtInstallPinandClips.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtInstallPinandClips.Text)
                Else
                    iNotes += 1
                End If
                '28_10_2009   Ragava   Till  Here
                HT_PaintNotes.Add(iNotes.ToString & ".0", ofrmTieRod3.chkInstallPinAndClips.Text)
                iCount += 1
            End If
            If ofrmTieRod3.chkNoLabelOnCylinder.Checked = True Then   ' Sugandhi_20120601_start

                If Trim(ofrmTieRod3.txtNoLabelOnCylinder.Text) <> "" Then
                    iNotes = Val(ofrmTieRod3.txtNoLabelOnCylinder.Text)
                Else
                    iNotes += 1
                End If

                HT_PaintNotes.Add(iNotes.ToString & ".0", UCase(ofrmTieRod3.chkNoLabelOnCylinder.Text))
                iCount += 1
            End If             ' Sugandhi_20120601_end

            If Trim(ofrmTieRod3.txtPackagePerSOP.Text) <> "" Then

                If iNotes = Val(ofrmTieRod3.txtPackagePerSOP.Text) Then
                    iNotes = iNotes + 1
                Else
                    iNotes = Val(ofrmTieRod3.txtPackagePerSOP.Text)
                End If
            Else
                iNotes += 1
            End If
            HT_PaintNotes.Add(iNotes.ToString & ".0", ofrmTieRod3.chkPackagePerSOP.Text)
            iCount += 1
            'Till  Here

            '20_10_2009  ragava
            If ofrmTieRod3.chkPaintingNote.Checked = True AndAlso Trim(ofrmTieRod3.RichTextBox2.Text) <> "" Then
                Dim strText As String() = ofrmTieRod3.RichTextBox2.Lines
                For Each str As String In strText
                    If Trim(str) <> "" Then
                        Dim strSplit() As String
                        strSplit = str.Split("}")
                        Dim strNumber As String = strSplit(LBound(strSplit))
                        Try
                            If Trim(strNumber) = "" Then
                                iNotes += 1
                                HT_PaintNotes.Add(iNotes.ToString & ".0", UCase(strSplit(UBound(strSplit))))
                            Else
                                HT_PaintNotes.Add(strNumber & ".0", UCase(strSplit(UBound(strSplit))))
                            End If
                        Catch ex As Exception
                        End Try
                    End If
                Next
            End If
            Dim PaintNotesCount As Integer = HT_PaintNotes.Keys.Count
            For Each Item As Object In HT_PaintNotes.Keys
                If PaintNotesCount < Val(Item) Then
                    PaintNotesCount = Val(Item)
                End If
            Next
            iCount = 1
            strPaintPackagingNotes = ""
            For iCount = 1 To PaintNotesCount
                If HT_PaintNotes.ContainsKey(iCount.ToString & ".0") = True Then                '04_11_2009  Ragava
                    strPaintPackagingNotes = strPaintPackagingNotes & iCount.ToString & ". " & UCase(Trim(HT_PaintNotes((iCount.ToString & ".0")))) & vbNewLine
                End If
            Next
            iPaintingNotesCount = HT_PaintNotes.Keys.Count
            '20_10_2009   ragava   Till   Here
            'GeneralNotes = "BORE " & Format(BoreDiameter, "0.00").ToString & " X " & Format(StrokeLength, "0.00").ToString & " STROKE " & vbNewLine        '01_10_2009  ragava
            GeneralNotes = Format(BoreDiameter, "0.00").ToString & " BORE" & " X " & _
                Format(StrokeLength, "0.00").ToString & " STROKE " & vbNewLine        '10_10_2011  ragava
            'dblDerateWorkingPressure
            GeneralNotes = GeneralNotes & "MAX. WORKING PRESSURE " & Format(WorkingPressure, "0").ToString & " PSI" & vbNewLine    '10_02_2010  RAGAVA  RETRACTION removed
            If Trim(strColumnLoadDeratePressure) <> "" AndAlso Trim(strColumnLoadDeratePressure) <> "N/A" Then      '02_10_2009  ragava
                GeneralNotes = GeneralNotes & "MAX. WORKING PRESSURE " & strColumnLoadDeratePressure & " PSI FULL EXTENSION" & vbNewLine
            End If
            '21_11_2012    RAGAVA
            If dblDerateWorkingPressure > 0 Then
                GeneralNotes = GeneralNotes & "DERATED WORKING PRESSURE " & dblDerateWorkingPressure.ToString & " PSI" & vbNewLine
            End If
            If strRodClevis_Class = "Class1" Then
                GeneralNotes = GeneralNotes & "MONARCH CYLINDER CLASS 1-15,000 CYCLES AS PER WI04-E-11" & vbNewLine
            Else
                GeneralNotes = GeneralNotes & "MONARCH CYLINDER CLASS 2-50,000 CYCLES AS PER WI04-E-11" & vbNewLine
            End If
            'DERATED WORKING PRESSURE xxx PSI
            'MONARCH CYLINDER CLASS 1/2-15,000/50,000 CYCLES AS PER WI04-E-11
            'TILL   HERE
            GeneralNotes = GeneralNotes & "CYLINDER CLEANLINESS CONTROLLED AS PER MONARCH WI10-E-50" & vbNewLine         '23_09_2011   RAGAVA
            If Trim(ofrmContractDetails.txtlPartCode.Text) <> "" AndAlso Trim(ofrmContractDetails.txtlPartCode.Text) <> "N/A" Then      '02_10_2009  RAGAVA
                GeneralNotes = GeneralNotes & "CUSTOMER PART # " & Trim(ofrmContractDetails.txtlPartCode.Text)      '20_10_2009  ragava
            Else
                GeneralNotes = GeneralNotes & "CUSTOMER PART # TBA"   '22_12_2012   RAGAVA
            End If
        Catch ex As Exception
            MsgBox("Error in Retrieving Notes")
        End Try

    End Sub

    '21_01_2011    RAGAVA
    Private Function GetCodeNumber_BeforeApplicationStart() As String

        Try
            Dim strQuery As String = "Select CodeNumber from CodeNumberDetails where Type = 'ROD'"
            Dim objDT6 As DataTable = oDataClass.GetDataTable(strQuery)
            GetCodeNumber_BeforeApplicationStart = objDT6.Rows(0).Item(0).ToString()
            Return GetCodeNumber_BeforeApplicationStart
        Catch ex As Exception
        End Try

    End Function

    Public Sub clearLoadInformation()

        Try
            strCodeNumber_BeforeApplicationStart = GetCodeNumber_BeforeApplicationStart()      '21_01_2011          RAGAVA
        Catch ex As Exception
        End Try
        mdiMonarch.lvwGeneralInformation.Columns.Clear()
        mdiMonarch.mdiComponent.Columns.Clear()
        ofrmTieRod1.txtWorkingPressure.Text = 0
        strColumnLoadDeratePressure = 0
        ColumnLoad = 0
        dblCylinderPullForce = 0
        CylinderCodeNumber = ""
        SetCodeDesciption = ""

        strBoreCodeNumber = ""
        strBoreDrawingNumber = ""
        strBoreDescription = ""
        strRodCodeNumber = ""
        strRodDrawingNumber = ""
        strRodDescription = ""

        strPistonCodeNumber = ""
        strPistonDrawingNumber = ""
        strPistonDescription = ""

        strstrokeControlCodeNumber = ""
        strStrokeControlDrawingNumber = ""
        strStrokeControlDescription = ""

        strPinsCodeNumber = ""
        strPinsDrawingNumber = ""
        strPinsDescription = ""

        strTieRodNutCodeNumber = ""
        strTieRodNutDrawingNumber = ""
        strTieRodDescription = ""

        strTieRodCodeNumber = ""
        strTieRodDrawingNumber = ""
        strTieRodDescription = ""

        strRodCapCodeNumber = ""
        strRodCapDrawingNumber = ""
        strRodCapDescription = ""

        strClevisCapCodeNumber = ""
        strClevisCapDrawingNumber = ""
        strClevisCapDescription = ""
        strRodClevisCodeNumber = ""

        strRodClevisDrawingNumber = ""
        strRodClevisDescription = ""
        strRodClevisPinCodeNumber = ""
        'frmTieRod1.Refresh()
        'frmTieRod2.Refresh()
        'frmTieRod3.Refresh()
        'frmContractDetails.Refresh()
        'InitialiseAllChildFormObjects()

    End Sub

    Public Sub LoadInformation()

        mdiMonarch.lvwGeneralInformation.Clear()
        mdiMonarch.lvwGeneralInformation.Items.Clear()
        mdiMonarch.lvwGeneralInformation.Columns.Clear()
        mdiMonarch.lvwGeneralInformation.Columns.Add("Property", 107, HorizontalAlignment.Center)
        mdiMonarch.lvwGeneralInformation.Columns.Add("Value", 200, HorizontalAlignment.Center)
        mdiMonarch.lvwGeneralInformation.Items.Add("Work Pressure")
        'mdiMonarch.lvwGeneralInformation.Items(0).SubItems.Add(IIf(WorkingPressure = 0, TempWorkingPressure, WorkingPressure))
        mdiMonarch.lvwGeneralInformation.Items(0).SubItems.Add(Val(ofrmTieRod1.txtWorkingPressure.Text))
        mdiMonarch.lvwGeneralInformation.Items.Add("Column Load Derate Pressure")
        mdiMonarch.lvwGeneralInformation.Items(1).SubItems.Add(strColumnLoadDeratePressure)
        mdiMonarch.lvwGeneralInformation.Items.Add("Column Load")
        mdiMonarch.lvwGeneralInformation.Items(2).SubItems.Add(ColumnLoad)
        mdiMonarch.lvwGeneralInformation.Items.Add("Cylinder Pull Force")
        mdiMonarch.lvwGeneralInformation.Items(3).SubItems.Add(dblCylinderPullForce)
        mdiMonarch.lvwGeneralInformation.Items.Add("Cylinder Code Number")
        mdiMonarch.lvwGeneralInformation.Items(4).SubItems.Add(CylinderCodeNumber)
        mdiMonarch.lvwGeneralInformation.Items.Add("Code Description")
        mdiMonarch.lvwGeneralInformation.Items(5).SubItems.Add(SetCodeDesciption)
        '--------Component Level---------------
        mdiMonarch.mdiComponent.Clear()
        mdiMonarch.mdiComponent.Items.Clear()
        mdiMonarch.mdiComponent.Columns.Clear()
        mdiMonarch.mdiComponent.Columns.Add("Component", 107, HorizontalAlignment.Center)
        mdiMonarch.mdiComponent.Columns.Add("Code Number", 200, HorizontalAlignment.Center)
        mdiMonarch.mdiComponent.Columns.Add("Drawing Number", 107, HorizontalAlignment.Center)
        mdiMonarch.mdiComponent.Columns.Add("Description", 200, HorizontalAlignment.Center)
        mdiMonarch.mdiComponent.Items.Add("Tube")
        mdiMonarch.mdiComponent.Items.Add("Rod")
        mdiMonarch.mdiComponent.Items.Add("Piston")
        mdiMonarch.mdiComponent.Items.Add("Stroke Control")
        mdiMonarch.mdiComponent.Items.Add("Pins")
        mdiMonarch.mdiComponent.Items.Add("Tie rod Nut")
        mdiMonarch.mdiComponent.Items.Add("Tie Rod")
        mdiMonarch.mdiComponent.Items.Add("Rod Cap")
        mdiMonarch.mdiComponent.Items.Add("Clevis Cap")
        mdiMonarch.mdiComponent.Items.Add("Rod Clevis")
        mdiMonarch.mdiComponent.Items.Add("Rod Clevis Pin")
        mdiMonarch.mdiComponent.Items(0).SubItems.Add(strBoreCodeNumber)
        'mdiMonarch.mdiComponent.Items(0).SubItems.Add(" ")                 '23_11_2009    Ragava   Commented
        mdiMonarch.mdiComponent.Items(0).SubItems.Add(strBoreDrawingNumber)
        mdiMonarch.mdiComponent.Items(0).SubItems.Add(strBoreDescription)
        mdiMonarch.mdiComponent.Items(1).SubItems.Add(strRodCodeNumber)
        'mdiMonarch.mdiComponent.Items(1).SubItems.Add(" ")                        '23_11_2009    Ragava   Commented
        mdiMonarch.mdiComponent.Items(1).SubItems.Add(strRodDrawingNumber)
        mdiMonarch.mdiComponent.Items(1).SubItems.Add(strRodDescription)

        mdiMonarch.mdiComponent.Items(2).SubItems.Add(strPistonCodeNumber)
        mdiMonarch.mdiComponent.Items(2).SubItems.Add(strPistonDrawingNumber)
        mdiMonarch.mdiComponent.Items(2).SubItems.Add(strPistonDescription)

        mdiMonarch.mdiComponent.Items(3).SubItems.Add(strstrokeControlCodeNumber)
        mdiMonarch.mdiComponent.Items(3).SubItems.Add(strStrokeControlDrawingNumber)
        mdiMonarch.mdiComponent.Items(3).SubItems.Add(strStrokeControlDescription)

        mdiMonarch.mdiComponent.Items(4).SubItems.Add(strPinsCodeNumber)
        mdiMonarch.mdiComponent.Items(4).SubItems.Add(strPinsDrawingNumber)
        mdiMonarch.mdiComponent.Items(4).SubItems.Add(strPinsDescription)

        mdiMonarch.mdiComponent.Items(5).SubItems.Add(strTieRodNutCodeNumber)
        mdiMonarch.mdiComponent.Items(5).SubItems.Add("N/A")
        mdiMonarch.mdiComponent.Items(5).SubItems.Add(strTieRodDescription)

        mdiMonarch.mdiComponent.Items(6).SubItems.Add(strTieRodCodeNumber)
        'mdiMonarch.mdiComponent.Items(6).SubItems.Add(" ")                 '23_11_2009    Ragava   Commented
        mdiMonarch.mdiComponent.Items(6).SubItems.Add(strTieRodDrawingNumber)
        mdiMonarch.mdiComponent.Items(6).SubItems.Add(strTieRodDescription)

        mdiMonarch.mdiComponent.Items(7).SubItems.Add(strRodCapCodeNumber)
        mdiMonarch.mdiComponent.Items(7).SubItems.Add(strRodCapDrawingNumber)
        mdiMonarch.mdiComponent.Items(7).SubItems.Add(strRodCapDescription)

        mdiMonarch.mdiComponent.Items(8).SubItems.Add(strClevisCapCodeNumber)
        mdiMonarch.mdiComponent.Items(8).SubItems.Add(strClevisCapDrawingNumber)
        mdiMonarch.mdiComponent.Items(8).SubItems.Add(strClevisCapDescription)
        mdiMonarch.mdiComponent.Items(9).SubItems.Add(strRodClevisCodeNumber)

        mdiMonarch.mdiComponent.Items(9).SubItems.Add(strRodClevisDrawingNumber)
        mdiMonarch.mdiComponent.Items(9).SubItems.Add(strRodClevisDescription)
        mdiMonarch.mdiComponent.Items(10).SubItems.Add(strRodClevisPinCodeNumber)
        mdiMonarch.mdiComponent.Items(10).SubItems.Add("N/A")                   '23_11_2009    Ragava   Commented
        mdiMonarch.mdiComponent.Items(10).SubItems.Add("N/A")                   '23_11_2009    Ragava   Commented

    End Sub

    Public Sub SaveModelFolder()

        Dim dirBrowser As New FolderBrowserDialog
        'SaveModelFolder = False
        'dirBrowser.ShowDialog()

        'SaveWorkFolder = dirBrowser.SelectedPath
        SaveWorkFolder = "C:\MONARCH_TESTING"
        If fso.FolderExists(SaveWorkFolder) = False Then
            fso.CreateFolder(SaveWorkFolder)
        End If

        If Not SaveWorkFolder.Equals("") Then
            If fso.FolderExists(SaveWorkFolder & "\" & CylinderCodeNumber) = True Then
                fso.DeleteFolder(SaveWorkFolder & "\" & CylinderCodeNumber, True)
                fso.CreateFolder(SaveWorkFolder & "\" & CylinderCodeNumber)
                'fso.DeleteFolder(filePath & "\" & WorkOrder & "\" & "InputImages", True)
                'fso.CreateFolder(SaveWorkFolder & "\" & CylinderCodeNumber & "\" & "InputImages")'15_09_2009
                'fso.DeleteFolder(filePath & "\" & WorkOrder & "\" & "DRAWINGS", True)
                'fso.CreateFolder(SaveWorkFolder & "\" & CylinderCodeNumber & "\" & "DRAWINGS")''15_09_2009
            Else
                fso.CreateFolder(SaveWorkFolder & "\" & CylinderCodeNumber)
                'fso.CreateFolder(SaveWorkFolder & "\" & CylinderCodeNumber & "\" & "InputImages")'15_09_2009
                'fso.CreateFolder(SaveWorkFolder & "\" & CylinderCodeNumber & "\" & "DRAWINGS")'15_09_2009
            End If
            DestinationFilePath = SaveWorkFolder & "\" & CylinderCodeNumber '& "\Models"
            fso.CopyFolder(Execution_Path + "\Tie_Rod_lib", DestinationFilePath, True)
            'SaveModelFolder = True

            '"C:\DESIGN_TABLES"
            Try
                If fso.FolderExists("C:\DESIGN_TABLES") = True Then
                    fso.DeleteFolder("C:\DESIGN_TABLES", True)
                End If
                fso.CopyFolder(Execution_Path + "\DESIGN_TABLES_MAIN", "C:\DESIGN_TABLES", True)
            Catch ex As Exception
                MsgBox("Error in Copying Design Tables")
            End Try
        End If

    End Sub

    Public Sub CaluculateLength(ByVal status As Boolean)

        If status = True Then
            Dim dblCastingSum As Double
            Dim intASAEFactor As Integer
            If ofrmTieRod1.cmbStyle.Text = "NON ASAE" Then
                If BoreDiameter <= 4.5 Then
                    dblCastingSum = 10.25
                ElseIf BoreDiameter > 4.5 Then
                    dblCastingSum = 12.25
                End If
            ElseIf ofrmTieRod1.cmbStyle.Text = "ASAE" Then
                If Val(ofrmTieRod1.cmbStrokeLength.Text) = 8 Then
                    dblCastingSum = 12.25
                ElseIf Val(ofrmTieRod1.cmbStrokeLength.Text) = 16 Then
                    dblCastingSum = 15.5
                End If
            End If
            If ofrmTieRod1.cmbStyle.Text = "ASAE" Then
                intASAEFactor = 2
            ElseIf ofrmTieRod1.cmbStyle.Text = "NON ASAE" Then
                intASAEFactor = 0
            End If

            ofrmTieRod1.txtRetractedLength.Text = StrokeLength + StopTubeLength + RodAdder + dblCastingSum
            ofrmTieRod1.txtExtendedLength.Text = Val(ofrmTieRod1.txtRetractedLength.Text) + StrokeLength
        Else
            ofrmTieRod1.txtRetractedLength.Text = Val(ofrmTieRod1.txtRetractedLength.Text) - 2.125
            ofrmTieRod1.txtExtendedLength.Text = Val(ofrmTieRod1.txtRetractedLength.Text) + StrokeLength
            'MessageBox.Show("Needs to be clarrify by Client", "Clarrification")
            'rdbRodClevisYes.Checked = True
        End If

    End Sub

    Public Function calculateRecommendedStopTubeLength() As Double           '26_10_2009  ragava   'From Sub to Function

        Try
            If Trim(ofrmTieRod1.cmbStyle.Text) <> "ASAE" Then   '26_10_2009  ragava 
                calculateRecommendedStopTubeLength = 0 '26_10_2009  ragava 
                Dim dblPushForce As Double = WorkingPressure * ((Math.PI / 4) * (Math.Pow(BoreDiameter, 2)))
                Dim dblDividend As Double = Math.Pow(Math.PI, 2) * Math.Pow(BoreDiameter, 4) * 1470000
                'dblRecommendedStopTubeLength = Val(ofrmTieRod1.txtExtendedLength.Text) - Math.Sqrt(dblDividend / dblPushForce)
                dblRecommendedStopTubeLength = (Val(ofrmTieRod1.txtExtendedLength.Text) - 40) / 10 '23_10_2009
                dblRecommendedStopTubeLength = Math.Round(dblRecommendedStopTubeLength, 2)
                ofrmTieRod1.txtRecommendedStoptubeLength.Text = IIf(dblRecommendedStopTubeLength > 0, _
                                    dblRecommendedStopTubeLength, "N/A")
                If (ofrmTieRod1.txtRecommendedStoptubeLength.Text <> "N/A") Then         '26_10_2009     Ragava
                    If Val(ofrmTieRod1.txtRecommendedStoptubeLength.Text) > 0 Then
                        ofrmTieRod1.txtStopTubeLength.Text = Val(ofrmTieRod1.txtRecommendedStoptubeLength.Text)
                        'Else 
                        '    ofrmTieRod1.txtStopTubeLength.Text = 0.1      '26_10_2009     Ragava
                    End If
                    '26_10_2009   ragava
                    Dim result As Double = Val(ofrmTieRod1.txtStopTubeLength.Text) Mod 0.125
                    If Not result = 0 Then
                        'If blnRevision = False Then
                        '    MessageBox.Show("Entered value is not multiples of 1/8, So Application is Rounding Off", "Information")
                        'End If
                        result = Val(ofrmTieRod1.txtStopTubeLength.Text) / 0.125
                        result = Math.Ceiling(result)
                        result = result * 0.125
                        ofrmTieRod1.txtStopTubeLength.Text = result.ToString
                        ofrmTieRod1.txtRecommendedStoptubeLength.Text = result.ToString
                    End If
                Else
                    ofrmTieRod1.rdbStopTubeNo.Checked = True                   '26_10_2009  ragava 
                End If
                calculateRecommendedStopTubeLength = Val(ofrmTieRod1.txtRecommendedStoptubeLength.Text)                   '26_10_2009  ragava 
                '26_10_2009   ragava   Till  Here
                ofrmTieRod1.txtRecommendedStoptubeLength.ReadOnly = True
            Else
                ofrmTieRod1.txtRecommendedStoptubeLength.Text = "N/A"   '26_10_2009  ragava 
                ofrmTieRod1.rdbStopTubeNo.Checked = True   '26_10_2009  ragava 
            End If
        Catch EX As Exception
        End Try

    End Function

    Private Sub ClearAll(ByVal C As Control)

        Dim Ctrl As Control
        Dim MyCheckBox As CheckBox
        Dim MyComboBox As IFLCustomUILayer.IFLComboBox
        Dim MyListView As IFLCustomUILayer.IFLListView
        Dim myTextBox As IFLCustomUILayer.IFLTextBox
        For Each Ctrl In C.Controls
            Select Case TypeName(Ctrl)
                Case "IFLTextBox"
                    myTextBox = CType(Ctrl, IFLCustomUILayer.IFLTextBox)
                    myTextBox.Text = ""
                Case "RichTextBox"
                    Ctrl.Text = ""
                Case "CheckBox"
                    MyCheckBox = CType(Ctrl, CheckBox)
                    MyCheckBox.Checked = False
                Case "IFLComboBox"
                    'Ctrl.Text = ""
                    MyComboBox = CType(Ctrl, IFLCustomUILayer.IFLComboBox)
                    MyComboBox.SelectedIndex = -1
                Case "IFLListView"
                    MyListView = CType(Ctrl, IFLCustomUILayer.IFLListView)
                    MyListView.Items.Clear()
                Case Else
                    If Ctrl.Controls.Count > 0 Then
                        ClearAll(Ctrl)
                    End If
            End Select
        Next Ctrl

    End Sub

    Public Function GetClevisCapDetails(Optional ByVal strCallingEndName As Boolean = False) As String

        GetClevisCapDetails = Nothing
        If strCallingEndName = False Then          '19_07_2011   RAGAVA
            'FOR CLEVIS CAP DETAILS
            If SeriesForCosting = "TX (TXC)" Then
                GetClevisCapDetails = "Series = 'TX' "
            Else 'If SeriesForCosting = "TL (TC)" Then              '19_07_2011   RAGAVA  commented
                GetClevisCapDetails = "Series <>'TX'"
            End If
        Else
            'ANUP 12-10-2010 START
            'FOR RODCAP DETAILS
            If strCallingEndName Then
                If SeriesForCosting = "LN" Then
                    GetClevisCapDetails = "series = 'LN'"
                    '22_07_2011  RAGAVA
                ElseIf SeriesForCosting = "TX (TXC)" Then
                    GetClevisCapDetails = "Series = 'TX' "
                ElseIf SeriesForCosting = "TL (TC)" Then
                    GetClevisCapDetails = "Series = 'TL' "
                    'Till   Here
                Else
                    GetClevisCapDetails = "series = 'TH/TP'"
                End If
            Else
                GetClevisCapDetails = "series = 'TH/TP/LN'"
            End If

            'ANUP 12-10-2010 TILL HERE
        End If

    End Function
    '02_11_2009   Ragava
    Public Function GetRodClevisDetails() As String

        GetRodClevisDetails = Nothing
        If SeriesForCosting = "TX (TXC)" Then
            GetRodClevisDetails = "Series = 'TX' "
        Else
            GetRodClevisDetails = "series = 'TL'"
        End If

    End Function

    Public Sub EditProjectFunctionality()

        If ofrmMonarch.lvwContractDetails.SelectedItems.Count > 0 Then
            oContractListviewItem = ofrmMonarch.lvwContractDetails.SelectedItems(0)
            ContractNumber = Trim(oContractListviewItem.Text)
            'anup 23-12-2010 start
            If IsNew_Revision_Released = "Released" Then
                intContractRevisionNumber = -1
            Else
                intContractRevisionNumber = Trim(oContractListviewItem.SubItems(2).Text)
            End If
            'anup 23-12-2010 till here
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If DataFetchingForEditAndCopy(ContractNumber) Then
            'Dim intlatestRunNumber As Integer = GETLATESTRUNNUMBER(ApplicationLoginObject.ProjectNumber, IFLConnectionObject)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

    End Sub

    Public Function DataFetchingForEditAndCopy(ByVal ContractNumber) As Boolean

        Try
            Dim strQuery As String
            strQuery = "SELECT PROJECT_XML FROM ContractMaster WHERE ContractNumber='" + ContractNumber + "'"
            Dim strByte() As Byte = IFLConnectionObject.GetValue(strQuery)
            Dim strXMLFilePath As String = Execution_Path1 + "\MIL.xml"
            If IO.File.Exists(strXMLFilePath) Then
                IO.File.Delete(strXMLFilePath)
            End If
            Dim fstream As New System.IO.FileStream(strXMLFilePath, IO.FileMode.OpenOrCreate)
            Dim bw As System.IO.BinaryWriter = New System.IO.BinaryWriter(fstream)
            Dim br As System.IO.BinaryReader = New System.IO.BinaryReader(fstream)
            bw.Write(strByte)
            br.Close()
            Dim oDataSet As New DataSet
            oDataSet.ReadXml(strXMLFilePath)
            If IO.File.Exists(strXMLFilePath) Then
                IO.File.Delete(strXMLFilePath)
            End If
            Dim IFLBaseData As New IFLGetSetUI.IFLGetSetUIClass
            For Each oTable As DataTable In oDataSet.Tables
                Dim oCurrentForm As Object = GetTableRelatedForm(oTable.TableName)
                IFLBaseData.SetDataToForm(oTable, oCurrentForm)
            Next
        Catch ex As Exception
        End Try

    End Function

    Private Function GetTableRelatedForm(ByVal strFormName As String) As Form

        GetTableRelatedForm = Nothing
        For Each oForm As Form In GetTieRod3
            If strFormName.ToUpper.Equals(oForm.Name.ToUpper) Then
                GetTableRelatedForm = oForm
                Exit For
            End If
        Next

    End Function

    Public Sub SetCodeNumbersToReport()

        Try
            Dim strCodeNumber As String = String.Empty
            Dim oDataClass As New DataClass
            Dim strQuery As String = "Select * from RodDiameterDetails where DrawingPartNumber = '" _
                                & RodDrawingNumber.ToString & "' and TableDrawing = 'Yes'"
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            RodLength = Math.Round(RodLength, 2)
            strQuery = ""
            strQuery = "Select CodeNumber,Dim_A,Revision from RodTableDrawing where DrawingNumber = '" & RodDrawingNumber & "'"
            Dim objDT2 As DataTable = oDataClass.GetDataTable(strQuery)
            Dim blnInsert As Boolean = True
            If objDT2.Rows.Count > 0 Then
                For Each dr As DataRow In objDT2.Rows
                    If dr(1).ToString = Format(RodLength, "0.00").ToString Then
                        blnInsert = False
                        strCodeNumber = dr(0).ToString
                        Exit For
                    End If
                Next
                If blnInsert = True Then
                    strQuery = "Select CodeNumber,Type,MaxCodeNumber from CodeNumberDetails where Type = 'ROD'"
                    Dim objDT6 As DataTable = oDataClass.GetDataTable(strQuery)
                    strCodeNumber = objDT6.Rows(0).Item(0).ToString()
                    If Val(strCodeNumber) >= objDT6.Rows(0).Item(2) Then
                        strCodeNumber = strCodeNumber & " - CodeNumber Exceeds the Maximum Limit"
                        ApplicationStop = True
                    End If
                End If
            Else
                If strStyle.IndexOf("ASAE") <> -1 Then
                    strCodeNumber = RodCodeNumber
                End If
                If strCodeNumber = "" Then
                    strQuery = ""
                    strQuery = "Select CodeNumber,Type,MaxCodeNumber from CodeNumberDetails where Type = 'ROD'"
                    Dim objDT3 As DataTable = oDataClass.GetDataTable(strQuery)
                    strCodeNumber = objDT3.Rows(0).Item(0).ToString()
                    If Val(strCodeNumber) >= objDT3.Rows(0).Item(2) Then
                        strCodeNumber = strCodeNumber & " - CodeNumber Exceeds the Maximum Limit"
                        ApplicationStop = True
                    End If
                End If
            End If
            RodCodeNumber = strCodeNumber
        Catch ex As Exception
        End Try

        '******************************************************************************************************

        'Tube Table Drawing
        Try
            Dim strCodeNumber As String = String.Empty
            Dim oDataClass As New DataClass
            Dim strQuery As String = "Select * from BoreDiameterDetails where DrawingPartNumber = '" & _
                                BoreDrawingNumber.ToString & "' and TableDrawing = 'Yes'"
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            TubeLength = Math.Round(TubeLength, 2)
            If objDT.Rows.Count > 0 Then
                strQuery = ""
                strQuery = "Select CodeNumber,Dim_A,Revision from TubeTableDrawing where DrawingNumber = '" & BoreDrawingNumber & "'"
                Dim objDT2 As DataTable = oDataClass.GetDataTable(strQuery)
                Dim blnInsert As Boolean = True
                If objDT2.Rows.Count > 0 Then
                    For Each dr As DataRow In objDT2.Rows
                        If dr(1).ToString = Format(TubeLength, "0.00").ToString Then                  '12_10_2009   ragava
                            blnInsert = False
                            strCodeNumber = dr(0).ToString
                            Exit For
                        End If
                    Next
                    If blnInsert = True Then
                        strQuery = "Select CodeNumber,Type,MaxCodeNumber from CodeNumberDetails where Type = 'Tube'"         '12_10_2009   ragava
                        Dim objDT6 As DataTable = oDataClass.GetDataTable(strQuery)
                        strCodeNumber = objDT6.Rows(0).Item(0).ToString()
                        If Val(strCodeNumber) >= objDT6.Rows(0).Item(2) Then
                            strCodeNumber = strCodeNumber & " - CodeNumber Exceeds the Maximum Limit"
                            ApplicationStop = True
                        End If
                    End If        '20_10_2009  ragava
                End If
            Else
                If strCodeNumber = "" Then
                    strQuery = ""
                    strQuery = "Select CodeNumber,Type,MaxCodeNumber from CodeNumberDetails where Type = 'Tube'"          '12_10_2009   ragava
                    Dim objDT3 As DataTable = oDataClass.GetDataTable(strQuery)
                    strCodeNumber = objDT3.Rows(0).Item(0).ToString()
                    If Val(strCodeNumber) >= objDT3.Rows(0).Item(2) Then
                        strCodeNumber = strCodeNumber & " - CodeNumber Exceeds the Maximum Limit"
                        ApplicationStop = True
                    End If
                End If
            End If
            TubeCodeNumber = strCodeNumber
        Catch ex As Exception
        End Try

        '*************************************************************************************************

        'Tie Rod Table Drawing
        Try
            Dim strCodeNumber As String = String.Empty
            Dim oDataClass As New DataClass
            Dim strQuery As String = "Select * from TieRodSizes where DrawingNumber = '" & _
                        TieRodDrawingNumber.ToString & "' and TableDrawing = 'Yes'"
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            TieRodLength = Math.Round(TieRodLength, 2)
            If objDT.Rows.Count > 0 Then
                strQuery = ""
                strQuery = "Select CodeNumber,Dim_A,Revision from TieRodTableDrawing where DrawingNumber = '" & TieRodDrawingNumber & "'"
                Dim objDT2 As DataTable = oDataClass.GetDataTable(strQuery)
                Dim blnInsert As Boolean = True
                If objDT2.Rows.Count > 0 Then
                    For Each dr As DataRow In objDT2.Rows
                        If dr(1).ToString = Format(TieRodLength, "0.00").ToString Then        '12_10_2009   ragava
                            blnInsert = False
                            strCodeNumber = dr(0).ToString
                            Exit For
                        End If
                    Next
                    If blnInsert = True Then
                        strQuery = "Select CodeNumber,Type,MaxCodeNumber from CodeNumberDetails where Type = 'TieROD'"         '12_10_2009  ragava
                        Dim objDT6 As DataTable = oDataClass.GetDataTable(strQuery)
                        strCodeNumber = objDT6.Rows(0).Item(0).ToString()
                        If Val(strCodeNumber) >= objDT6.Rows(0).Item(2) Then
                            strCodeNumber = strCodeNumber & " - CodeNumber Exceeds the Maximum Limit"
                            ApplicationStop = True
                        End If
                    End If       '20_10_2009  ragava
                End If
            Else
                If strCodeNumber = "" Then
                    strQuery = ""
                    strQuery = "Select CodeNumber,Type,MaxCodeNumber from CodeNumberDetails where Type = 'TieROD'"      '12_10_2009  ragava
                    Dim objDT3 As DataTable = oDataClass.GetDataTable(strQuery)
                    strCodeNumber = objDT3.Rows(0).Item(0).ToString()
                    If Val(strCodeNumber) >= objDT3.Rows(0).Item(2) Then
                        strCodeNumber = strCodeNumber & " - CodeNumber Exceeds the Maximum Limit"
                        ApplicationStop = True
                    End If
                End If
            End If
            TieRodCodeNumber1 = strCodeNumber
        Catch ex As Exception
        End Try

        '***************************************************************************************************************

        'Stop Tube Table Drawing
        Try
            If dblStopTubeLength > 0 Then
                Dim strCodeNumber As String = String.Empty
                Dim strQuery As String = String.Empty
                Dim oDataClass As New DataClass
                StopTubeDrawingNumber = "495598"
                strQuery = ""
                strQuery = "Select Dim_A,Dim_B,Dim_C,CodeNumber from StopTubeTableDrawing where DrawingNumber = '" _
                                                        & StopTubeDrawingNumber & "'"
                Dim objDT2 As DataTable = oDataClass.GetDataTable(strQuery)
                Dim blnInsert As Boolean = True
                Dim dblDim_B As Double = 0
                Dim dblDim_C As Double = Math.Round((dblRodDiameter + 0.015), 2)
                'If dblRodDiameter <= 1.12 Then
                '    dblDim_B = Math.Round((dblRodDiameter + 0.015) + (2 * 0.19), 2)
                'Else
                '    dblDim_B = Math.Round((dblRodDiameter + 0.015) + (2 * 0.25), 2)
                'End If
                If objDT2.Rows.Count > 0 Then
                    For Each dr As DataRow In objDT2.Rows
                        'If (dr(0).ToString = Format(Math.Round(dblStopTubeLength, 2), "0.00").ToString) AndAlso (dr(1).ToString = Format(dblDim_B, "0.00").ToString) AndAlso (dr(2).ToString = Format(dblDim_C, "0.00").ToString) Then         '12_10_2009   ragava
                        '    blnInsert = False
                        '    strCodeNumber = dr(3).ToString
                        '    Exit For
                        'End If

                        If (dr(0).ToString = Format(Math.Ceiling(dblStopTubeLength * 100) / 100, "0.00").ToString) _
                        AndAlso (dr(2).ToString = Format(dblDim_C, "0.00").ToString) Then         '12_10_2009   ragava
                            blnInsert = False
                            strCodeNumber = dr(3).ToString
                            Exit For
                        End If
                    Next
                    If blnInsert = True Then
                        strQuery = "Select CodeNumber,Type,MaxCodeNumber from CodeNumberDetails where Type = 'StopTube'"          '12_10_2009  ragava
                        Dim objDT6 As DataTable = oDataClass.GetDataTable(strQuery)
                        strCodeNumber = objDT6.Rows(0).Item(0).ToString()
                        If Val(strCodeNumber) >= objDT6.Rows(0).Item(2) Then
                            strCodeNumber = strCodeNumber & " - CodeNumber Exceeds the Maximum Limit"
                            ApplicationStop = True
                        End If
                    End If
                End If
                StopTubeCodeNumber = strCodeNumber
            End If
        Catch ex As Exception
        End Try

    End Sub

    'Sandeep 04-03-10-4pm
    Public Sub AddCodeNumbersToCostingExcelRetrivedFromDB()
        Dim strShippingPlug_RodCap As String = ""
        Dim strShippingPlug_ClevisCap As String = ""
        Dim strPerminentPlug_ClevisCap As String = ""

        Dim strORing_PistonSeal As String = ""
        Dim strBackUpRing_PistonSeal As String = ""
        Dim strORing_RodCap As String = ""
        Dim strBackUpRing_RodCap As String = ""
        Dim strORing_ClevisCap As String = ""
        Dim strBackUpRing_ClevisCap As String = ""

        '20_09_2011  RAGAVA
        Dim strDualSeal_RodCap As String = ""
        Dim strDualSeal_ClevisCap As String = ""
        'Till  Here

        Dim strWearRing1_Piston As String = ""
        Dim strWearRing2_Piston As String = ""
        Dim strWearRing1_RodCap As String = ""
        Dim strWearRing2_RodCap As String = ""

        Dim oDataClass As New DataClass
        If Not IsNothing(_strPistonCode) Then
            Dim oPistonDataTable As DataTable = oDataClass.GetDataTable _
                                ("Select * from PistonSealDetails where PartNumber=" + _strPistonCode)
            Dim oPistonDataRow As DataRow = oPistonDataTable.Rows(0)
            strORing_PistonSeal = oPistonDataRow("ORing")
            strBackUpRing_PistonSeal = oPistonDataRow("BackUpRing")
            ObjClsCostingDetails.AddCodeNumberToDataTable(oPistonDataRow("PTFESeal"), "PTFESeal Code") 'Sandeep 04-03-10-4pm
            ObjClsCostingDetails.AddCodeNumberToDataTable(oPistonDataRow("ORingExpander"), "ORingExpander Code") 'Sandeep 04-03-10-4pm
            ObjClsCostingDetails.AddCodeNumberToDataTable(oPistonDataRow("PSPSeal"), "PSPSeal Code") 'Sandeep 04-03-10-4pm
            'ANUP 03-11-2010 START
            If ofrmTieRod2.cmbPistonSealPackage.Text.StartsWith("WynSeal") Then
                ObjClsCostingDetails.AddCodeNumberToDataTable(oPistonDataRow("WynSeal"), "WynSeal Code")
            ElseIf ofrmTieRod2.cmbPistonSealPackage.Text.StartsWith("GlydP") Then
                ObjClsCostingDetails.AddCodeNumberToDataTable(oPistonDataRow("GlydP"), "GlydP Code")
            End If
            'ANUP 03-11-2010 TILL HERE
            strWearRing1_Piston = oPistonDataRow("WearRing1")
            strWearRing2_Piston = oPistonDataRow("WearRing2")
        End If

        If Not IsNothing(ofrmTieRod2.txtRodCap.Text) Then
            Dim oRodCapDataTable As DataTable = oDataClass.GetDataTable _
                                        ("Select * from RodCapDetails where PartNumber=" + ofrmTieRod2.txtRodCap.Text)
            Dim oPistonDataRow As DataRow = oRodCapDataTable.Rows(0)
            '01_05_2012   RAGAVA
            If Not SeriesForCosting = "LN" Then
                strORing_RodCap = oPistonDataRow("ORing")
                strBackUpRing_RodCap = oPistonDataRow("BackUpRing")
            Else
                strDualSeal_RodCap = oPistonDataRow("DualSeal")          '20_09_2011   RAGAVA
            End If
            'strORing_RodCap = oPistonDataRow("ORing")
            'strBackUpRing_RodCap = oPistonDataRow("BackUpRing")
            'strDualSeal_RodCap = oPistonDataRow("DualSeal")          '20_09_2011   RAGAVA
            'Till  Here
            ObjClsCostingDetails.AddCodeNumberToDataTable(oPistonDataRow("Hallite"), "Hallite Code") 'Sandeep 04-03-10-4pm
            ObjClsCostingDetails.AddCodeNumberToDataTable(oPistonDataRow("ZMacro"), "ZMacro Code") 'Sandeep 04-03-10-4pm
            ObjClsCostingDetails.AddCodeNumberToDataTable(oPistonDataRow("RU9"), "Rod Seal Code")       '23_08_2011   RAGAVA
            strWearRing1_RodCap = oPistonDataRow("WearRing1")
            strWearRing2_RodCap = oPistonDataRow("WearRing2")
            strShippingPlug_RodCap = oPistonDataRow("ShippingPlugNumber")
        End If

        If Not IsNothing(ofrmTieRod2.txtClevisCap.Text) Then
            Dim oRodCapDataTable As DataTable = oDataClass.GetDataTable _
                    ("Select * from ClevisCapDetails where PartNumber=" + ofrmTieRod2.txtClevisCap.Text)
            Dim oPistonDataRow As DataRow = oRodCapDataTable.Rows(0)
            '01_05_2012   RAGAVA
            If Not SeriesForCosting = "LN" Then
                strORing_ClevisCap = oPistonDataRow("ORingSeal")
                strBackUpRing_ClevisCap = oPistonDataRow("BackUpSeal")
            Else
                strDualSeal_ClevisCap = oPistonDataRow("DualSeal")
            End If
            'strORing_ClevisCap = oPistonDataRow("ORingSeal")
            'strBackUpRing_ClevisCap = oPistonDataRow("BackUpSeal")
            'strDualSeal_ClevisCap = oPistonDataRow("DualSeal")          '20_09_2011   RAGAVA
            'Till  Here

            ObjClsCostingDetails.AddCodeNumberToDataTable(oPistonDataRow("Sealant"), "Sealant Code", 0.0001) 'Sandeep 04-03-10-4pm
            strShippingPlug_ClevisCap = oPistonDataRow("ShippingPlug")
            strPerminentPlug_ClevisCap = oPistonDataRow("PermanentPlug")
        End If

        '01_05_2012 RAM
        If SeriesForCosting = "LN" Then
            If strDualSeal_ClevisCap.Equals(strDualSeal_RodCap) Then
                ObjClsCostingDetails.AddCodeNumberToDataTable(strDualSeal_ClevisCap, "DualSeal", 2)
            Else
                ObjClsCostingDetails.AddCodeNumberToDataTable(strDualSeal_RodCap, "Dual Seal Rod Cap", 1)
                If strDualSeal_ClevisCap <> "N/A" Then     '18_10_2011   RAGAVA
                    ObjClsCostingDetails.AddCodeNumberToDataTable(strDualSeal_ClevisCap, "Dual Seal Clevis Cap", 1)
                End If
            End If
        End If
        'Till Here

        '20_09_2011   RAGAVA
        If SeriesForCosting.ToString.StartsWith("TX") = True AndAlso Val(ofrmTieRod1.cmbBore.Text) <= 3 Then
            strORing_RodCap = ""
            strBackUpRing_RodCap = ""
            strORing_ClevisCap = ""
            strBackUpRing_ClevisCap = ""
            If strDualSeal_ClevisCap.Equals(strDualSeal_RodCap) Then
                ObjClsCostingDetails.AddCodeNumberToDataTable(strDualSeal_ClevisCap, "DualSeal", 2)
            Else
                ObjClsCostingDetails.AddCodeNumberToDataTable(strDualSeal_RodCap, "Dual Seal Rod Cap", 1)
                If strDualSeal_ClevisCap <> "N/A" Then     '18_10_2011   RAGAVA
                    ObjClsCostingDetails.AddCodeNumberToDataTable(strDualSeal_ClevisCap, "Dual Seal Clevis Cap", 1)
                End If
            End If

        End If
        'Till   Here

        'Sandeep 20-04-10-4pm
        If ThreadProtected = "All Permenant" Then
            ObjClsCostingDetails.AddCodeNumberToDataTable(strPerminentPlug_ClevisCap, "Permanent Plug Code", 3) 'Sandeep 20-04-10-4pm
        Else
            ObjClsCostingDetails.AddCodeNumberToDataTable(strPerminentPlug_ClevisCap, "Permanent Plug Code") 'Sandeep 20-04-10-4pm
            If strShippingPlug_RodCap.Equals(strShippingPlug_ClevisCap) Then
                ObjClsCostingDetails.AddCodeNumberToDataTable(strShippingPlug_RodCap, "Shipping Plug Code", 2) 'Sandeep 20-04-10-4pm
            Else
                ObjClsCostingDetails.AddCodeNumberToDataTable(strShippingPlug_RodCap, "Shipping Plug RodCap Code") 'Sandeep 20-04-10-4pm
                ObjClsCostingDetails.AddCodeNumberToDataTable(strShippingPlug_ClevisCap, "Shipping Plug ClevisCap Code") 'Sandeep 20-04-10-4pm
            End If
        End If
        '*********************

        'Sandeep 20-04-10-4pm
        If strORing_PistonSeal.Equals(strORing_RodCap) AndAlso strORing_PistonSeal.Equals(strORing_ClevisCap) Then
            ObjClsCostingDetails.AddCodeNumberToDataTable(strORing_PistonSeal, "ORing Code", 3) 'Sandeep 20-04-10-4pm
        ElseIf strORing_PistonSeal.Equals(strORing_RodCap) Then
            ObjClsCostingDetails.AddCodeNumberToDataTable(strORing_PistonSeal, "ORing Code", 2) 'Sandeep 20-04-10-4pm
            ObjClsCostingDetails.AddCodeNumberToDataTable(strORing_ClevisCap, "ORing Code Clevis Cap") 'Sandeep 20-04-10-4pm
        ElseIf strORing_PistonSeal.Equals(strORing_ClevisCap) Then
            ObjClsCostingDetails.AddCodeNumberToDataTable(strORing_PistonSeal, "ORing Code", 2) 'Sandeep 20-04-10-4pm
            ObjClsCostingDetails.AddCodeNumberToDataTable(strORing_RodCap, "ORing Code Rod Cap") 'Sandeep 20-04-10-4pm
        ElseIf strORing_RodCap.Equals(strORing_ClevisCap) Then
            ObjClsCostingDetails.AddCodeNumberToDataTable(strORing_RodCap, "ORing Code", 2) 'Sandeep 20-04-10-4pm
            ObjClsCostingDetails.AddCodeNumberToDataTable(strORing_PistonSeal, "ORing Code Piston Cap") 'Sandeep 20-04-10-4pm
        Else
            ObjClsCostingDetails.AddCodeNumberToDataTable(strORing_PistonSeal, "ORing Code Piston Cap") 'Sandeep 20-04-10-4pm
            ObjClsCostingDetails.AddCodeNumberToDataTable(strORing_RodCap, "ORing Code Rod Cap") 'Sandeep 20-04-10-4pm
            ObjClsCostingDetails.AddCodeNumberToDataTable(strORing_ClevisCap, "ORing Code Clevis Cap") 'Sandeep 20-04-10-4pm
        End If

        ''Sandeep 20-04-10-4pm
        'If strBackUpRing_PistonSeal.Equals(strBackUpRing_RodCap) AndAlso strBackUpRing_PistonSeal.Equals(strBackUpRing_ClevisCap) Then
        '    ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_PistonSeal, "BackUpRing Code", 3) 'Sandeep 20-04-10-4pm
        'ElseIf strBackUpRing_PistonSeal.Equals(strBackUpRing_RodCap) Then
        '    ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_PistonSeal, "BackUpRing Code", 2) 'Sandeep 20-04-10-4pm
        '    ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_ClevisCap, "BackUpRing Code Clevis Cap") 'Sandeep 20-04-10-4pm
        'ElseIf strBackUpRing_PistonSeal.Equals(strBackUpRing_ClevisCap) Then
        '    ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_PistonSeal, "BackUpRing Code", 2) 'Sandeep 20-04-10-4pm
        '    ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_RodCap, "BackUpRing Code Rod Cap") 'Sandeep 20-04-10-4pm
        'ElseIf strBackUpRing_RodCap.Equals(strBackUpRing_ClevisCap) Then
        '    ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_RodCap, "BackUpRing Code", 2) 'Sandeep 20-04-10-4pm
        '    ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_PistonSeal, "BackUpRing Code Piston Cap") 'Sandeep 20-04-10-4pm
        'Else
        '    ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_PistonSeal, "BackUpRing Code Piston Cap") 'Sandeep 20-04-10-4pm
        '    ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_RodCap, "BackUpRing Code Rod Cap") 'Sandeep 20-04-10-4pm
        '    ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_ClevisCap, "BackUpRing Code Clevis Cap") 'Sandeep 20-04-10-4pm
        'End If


        '08_08_2011         RAGAVA
        If strBackUpRing_PistonSeal.Equals(strBackUpRing_RodCap) AndAlso strBackUpRing_PistonSeal.Equals(strBackUpRing_ClevisCap) Then
            ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_PistonSeal, "BackUpRing Code", 4)
        ElseIf strBackUpRing_PistonSeal.Equals(strBackUpRing_RodCap) Then
            ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_PistonSeal, "BackUpRing Code", 3)
            ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_ClevisCap, "BackUpRing Code Clevis Cap")
        ElseIf strBackUpRing_PistonSeal.Equals(strBackUpRing_ClevisCap) Then
            ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_PistonSeal, "BackUpRing Code", 3)
            ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_RodCap, "BackUpRing Code Rod Cap")
        ElseIf strBackUpRing_RodCap.Equals(strBackUpRing_ClevisCap) Then
            ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_RodCap, "BackUpRing Code", 2)
            ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_PistonSeal, "BackUpRing Code Piston Cap", 2)
        Else
            ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_PistonSeal, "BackUpRing Code Piston Cap", 2)
            ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_RodCap, "BackUpRing Code Rod Cap")
            ObjClsCostingDetails.AddCodeNumberToDataTable(strBackUpRing_ClevisCap, "BackUpRing Code Clevis Cap")
        End If
        'Till   Here


        'Sandeep 20-04-10-4pm
        If strWearRing1_Piston.Equals(strWearRing1_RodCap) Then
            ObjClsCostingDetails.AddCodeNumberToDataTable(strWearRing1_Piston, "WearRing1 Code", 2) 'Sandeep 20-04-10-4pm
        Else
            ObjClsCostingDetails.AddCodeNumberToDataTable(strWearRing1_Piston, "WearRing1 Code Piston") 'Sandeep 20-04-10-4pm
            ObjClsCostingDetails.AddCodeNumberToDataTable(strWearRing1_RodCap, "WearRing1 Code Rod") 'Sandeep 20-04-10-4pm
        End If

        'Sandeep 20-04-10-4pm
        If strWearRing2_Piston.Equals(strWearRing2_RodCap) Then
            ObjClsCostingDetails.AddCodeNumberToDataTable(strWearRing2_Piston, "WearRing2 Code", 4) 'Sandeep 20-04-10-4pm
        Else
            ObjClsCostingDetails.AddCodeNumberToDataTable(strWearRing2_Piston, "WearRing2 Code Piston", 2) 'Sandeep 20-04-10-4pm
            ObjClsCostingDetails.AddCodeNumberToDataTable(strWearRing2_RodCap, "WearRing2 Code Rod", 2) 'Sandeep 20-04-10-4pm
        End If

    End Sub

    Public Function GetPinKits_Details(ByVal strPinCode As String, ByVal strClip As String, ByVal strBaseOrRod As String) As String
        Try
            GetPinKits_Details = String.Empty
            Dim strQuery As String = ""
            If strBaseOrRod = "BASE" Then
                strQuery = "select PinKitCodeNumber from Pin_Kit_Details where FirstPin = '" & strPinCode _
                            & "' and ClipType = '" & strClip & "'"
            ElseIf strBaseOrRod = "ROD" Then
                strQuery = "select PinKitCodeNumber from Pin_Kit_Details where SecondPin = '" & strPinCode _
                                        & "' and ClipType = '" & strClip & "'"
            End If
            GetPinKits_Details = IFLConnectionObject.GetValue(strQuery)
            'If strPinCode = "268900" AndAlso strClip.IndexOf("Cotter Pin") <> -1 Then
            '    GetPinKits_Details = "235004"
            'ElseIf strPinCode = "134953" AndAlso strClip.IndexOf("R -") <> -1 Then
            '    GetPinKits_Details = "235012"
            'ElseIf strPinCode = "134953" AndAlso strClip.IndexOf("Cotter Pin") <> -1 Then
            '    GetPinKits_Details = "235005"
            'ElseIf strPinCode = "257832" AndAlso strClip.IndexOf("Cotter Pin") <> -1 Then
            '    GetPinKits_Details = "235006"
            'ElseIf strPinCode = "257835" AndAlso strClip.IndexOf("Cotter Pin") <> -1 Then
            '    GetPinKits_Details = "235007"
            '    'ObjClsWeldedCylinderFunctionalClass.CmbClips_BaseEnd.Text = "Hair Pin"
            'End If
        Catch ex As Exception
        End Try

    End Function

    '16_06_2011  RAGAVA
    Public Function Validate_PinandClipsNotes() As Boolean
        Try
            Dim strKitCode_BaseEnd As String = GetPinKits_Details(_strPinCodeBE, Trim(ofrmTieRod2.cmbClips.Text), "BASE")
            strBaseEndKitCode = strKitCode_BaseEnd         '06_06_2011   RAGAVA
            Dim strKitCode_RodEnd As String = GetPinKits_Details(_strPinCodeRE, Trim(ofrmTieRod2.cmbRodClevisPinClips.Text), "ROD")
            strRodEndKitCode = strKitCode_RodEnd         '06_06_2011   RAGAVA
            If strKitCode_BaseEnd = strKitCode_RodEnd AndAlso (strKitCode_BaseEnd <> "" Or strKitCode_RodEnd <> "") Then
                Validate_PinandClipsNotes = True
                blnInstallPinsandClips = False
            Else
                Validate_PinandClipsNotes = False
                blnInstallPinsandClips = True
            End If
        Catch ex As Exception
        End Try
    End Function


#End Region

End Module

