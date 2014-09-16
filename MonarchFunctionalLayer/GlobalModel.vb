Imports IFLBaseDataLayer
Imports MonarchFunctionalLayer

Public Module GeneralModule
#Region "Class Level Variables"

    Public blnGenerateClicked As Boolean = False            '14_07_2011   RAGAVA

    Public strCodeNumber_BeforeApplicationStart As String = String.Empty           '21_01_2011         RAGAVA
    Public ht_CodeNumbers As New Hashtable       '04_10_2010   RAGAVA

    Dim _blnClevisPins As Boolean         '06_04_2010   RAGAVA
    Dim _blnRodClevisPins As Boolean      '06_04_2010   RAGAVA
    Dim _blnPinsPlasticBag As Boolean = False    '07_04_2010   RAGAVA

    Dim _rdbRodClevis As Boolean              '19_02_2010    RAGAVA
    Dim _ClevisCapCodeNumber As String        '12_10_2009     ragava
    Dim _iAssyNotesCount As Integer           '12_10_2009     ragava
    Dim _iPaintingNotesCount As Integer       '12_10_2009     ragava
    Public dblPinSize As Double                  '10_09_2011  RAGAVA
    Public blnInstallPinsandClips_Checked As Boolean = False       '19_10_2011  RAGAVA
    Public strPinKitId As String = ""   '23_09_2011   RAGAVA

    Dim _strRodCodeNumber As String          '01_10_2009   ragava
    Dim _strStyle As String          '01_10_2009   ragava
    Dim _strPistonSealPackage As String      '01_10_2009   ragava

    Dim _dblRodStrokeDifference As Double   '22_09_2009  ragava
    Dim _dblTubeStrokeDifference As Double  '24_09_2009   ragava
    Dim _dblTieRodStrokeDifference As Double   '24_09_2009  ragava
    Dim _dblStopTubeLength As Double        '24_09_2009  ragava

    '16_09_2009  ragava
    Dim dblRodLength As Double
    Dim dblTubeLength As Double
    Dim _strRephasing As String
    Dim dblTieRodLength As Double
    Dim strClevisCapPortOrientation As String
    '16_09_2009  ragava   Till  Here

    '14_09_2009  ragava
    Dim strClevisCapPort As String
    Dim strRodCapPort As String
    Dim dblPinHoleSize As Double
    '14_09_2009  ragava      Till  Here

    '08_09_2009   ragava
    Dim _strRodMaterial As String
    Dim _strCylinderDescription As String
    Dim _strPaintPackagingNotes As String
    Dim _strAssemblyNotes As String
    '08_09_2009   ragava   Till  Here
    '11_09_2009  ragava
    Dim _strGeneralNotes As String
    Dim _dblStrokeLength As Double
    Dim _dblPistonThreadSize As Double
    Dim _dblRodThreadSize As Double
    Dim _dblBoreDiameter As Double

    Dim _strRodMaterialNumber As String
    Dim _strRodDrawingNumber As String
    Dim _strBoreDrawingNumber As String
    Dim _strTieRodDrawingNumber As String
    '11_09_2009  ragava  Till  Here

    Dim _dblRodDiameter As String    '10_09_2009  RAGAVA
    Dim _strSeries As String         '10_09_2009  RAGAVA


    Dim _strCustomerName As String
    Dim _strContractNumber As String
    Dim _strType As String
    Dim _strPartCode As String
    Dim _intRevision As Integer
    Dim _strExecution_Path As String
    Dim _strExecution_Path1 As String
    Dim _lFormName As Form
    Private _oIFLConnection As IFLConnectionClass
    Dim _oFunctionalClass As FunctionalClass
    Dim _arAllFormsCotrolValues As ArrayList
    Public oclsimgcapture As New clsimgcapture
    Dim _arDuctDrawingOutPut As ArrayList
    Dim _arListInputs As ArrayList '29_05_2009
    Dim _strSavefiletopath As String
    Dim _strDestPath As String
    Dim _alDrawingForDXF As ArrayList
    Public oDBLoginObject As New IFLBusinessLayer.IFLDataBaseLoginClass
    Public oExcelClass As New ExcelClass
    Dim _strStopTubeDrawingNumber As String     '22_09_2009  ragava
    Dim _intContractRevisionNumber As Integer
    Dim _strPinHoleType As String
    Dim _strRodClevisPinHoleType As String       '02_11_2009  ragava
    Dim _strPartCode1 As String
    Dim _dblExtendedLength As Double             '20_10_2009   ragava
    Private _blnPins As Boolean        '11_11_2009  Ragava
    Private _strClevisPinClips As String          '11_11_2009  Ragava
    Private _strRodPinClips As String          '11_11_2009  Ragava

    Public Volume_Rod As Double        '25_06_2010   RAGAVA
    Public Volume_TieRod As Double        '25_06_2010   RAGAVA
    Public Volume_StopTube As Double        '25_06_2010   RAGAVA
    Public Volume_Bore As Double        '25_06_2010   RAGAVA
    Public Volume_Assembly As Double        '25_06_2010   RAGAVA

    Public Weight_Rod As Double        '04_10_2010   RAGAVA
    Public Weight_TieRod As Double        '04_10_2010   RAGAVA
    Public Weight_StopTube As Double        '04_10_2010   RAGAVA
    Public Weight_Bore As Double        '04_10_2010   RAGAVA
    Public Weight_Assembly As Double        '04_10_2010   RAGAVA
#End Region
#Region "Public properties"

    '11_11_2009  Ragava
    Public Property ClevisPinClips() As String
        Get
            Return _strClevisPinClips
        End Get
        Set(ByVal value As String)
            _strClevisPinClips = value
        End Set
    End Property

    '11_11_2009  Ragava
    Public Property RodPinClips() As String
        Get
            Return _strRodPinClips
        End Get
        Set(ByVal value As String)
            _strRodPinClips = value
        End Set
    End Property



    '11_11_2009  Ragava
    Public Property blnPins() As Boolean
        Get
            Return _blnPins
        End Get
        Set(ByVal value As Boolean)
            _blnPins = value
        End Set
    End Property



    '20_10_2009   ragava
    Public Property ExtendedLength() As Double       '21_10_2009   ragava      Integer
        Get
            Return _dblExtendedLength
        End Get
        Set(ByVal value As Double)
            _dblExtendedLength = value
        End Set
    End Property

    Public Property intContractRevisionNumber() As Integer
        Get
            Return _intContractRevisionNumber
        End Get
        Set(ByVal value As Integer)
            _intContractRevisionNumber = value
        End Set
    End Property
    '12_10_2009   ragava
    Public Property ClevisCapCodeNumber() As String
        Get
            Return _ClevisCapCodeNumber
        End Get
        Set(ByVal value As String)
            _ClevisCapCodeNumber = value
        End Set
    End Property


    '12_10_2009  ragava
    Public Property iAssyNotesCount() As Integer
        Get
            Return _iAssyNotesCount
        End Get
        Set(ByVal value As Integer)
            _iAssyNotesCount = value
        End Set
    End Property

    '12_10_2009  ragava
    Public Property iPaintingNotesCount() As Integer
        Get
            Return _iPaintingNotesCount
        End Get
        Set(ByVal value As Integer)
            _iPaintingNotesCount = value
        End Set
    End Property
    '06_04_2010     RAGAVA
    Public Property RodClevisPins() As Boolean
        Get
            Return _blnRodClevisPins
        End Get
        Set(ByVal value As Boolean)
            _blnRodClevisPins = value
        End Set
    End Property
    '06_04_2010     RAGAVA
    Public Property ClevisPins() As Boolean
        Get
            Return _blnClevisPins
        End Get
        Set(ByVal value As Boolean)
            _blnClevisPins = value
        End Set
    End Property
    '07_04_2010   RAGAVA
    Public Property blnPinsPlasticBag() As Boolean
        Get
            Return _blnPinsPlasticBag
        End Get
        Set(ByVal value As Boolean)
            _blnPinsPlasticBag = value
        End Set
    End Property
    '19_02_2010     RAGAVA
    Public Property rdbRodClevis() As Boolean
        Get
            Return _rdbRodClevis
        End Get
        Set(ByVal value As Boolean)
            _rdbRodClevis = value
        End Set
    End Property

    '01_10_2009  ragava
    Public Property strPistonSealPackage() As String
        Get
            Return _strPistonSealPackage
        End Get
        Set(ByVal value As String)
            _strPistonSealPackage = value
        End Set
    End Property


    '01_10_2009  ragava
    Public Property strStyle() As String
        Get
            Return _strStyle
        End Get
        Set(ByVal value As String)
            _strStyle = value
        End Set
    End Property

    '01_10_2009  ragava
    Public Property RodCodeNumber() As String
        Get
            Return _strRodCodeNumber
        End Get
        Set(ByVal value As String)
            _strRodCodeNumber = value
        End Set
    End Property

    '24_09_2009  ragava
    Public Property dblStopTubeLength() As Double
        Get
            Return _dblStopTubeLength
        End Get
        Set(ByVal value As Double)
            _dblStopTubeLength = value
        End Set
    End Property

    '24_09_2009  ragava
    Public Property dblTieRodStrokeDifference() As Double
        Get
            Return _dblTieRodStrokeDifference
        End Get
        Set(ByVal value As Double)
            _dblTieRodStrokeDifference = value
        End Set
    End Property

    '24_09_2009  ragava
    Public Property dblTubeStrokeDifference() As Double
        Get
            Return _dblTubeStrokeDifference
        End Get
        Set(ByVal value As Double)
            _dblTubeStrokeDifference = value
        End Set
    End Property

    '22_09_2009  ragava
    Public Property dblRodStrokeDifference() As Double
        Get
            Return _dblRodStrokeDifference
        End Get
        Set(ByVal value As Double)
            _dblRodStrokeDifference = value
        End Set
    End Property

    '22_09_2009  ragava
    Public Property StopTubeDrawingNumber() As String
        Get
            Return _strStopTubeDrawingNumber
        End Get
        Set(ByVal value As String)
            _strStopTubeDrawingNumber = value
        End Set
    End Property

    '16_09_2009  ragava
    Public Property strRephasing() As String
        Get
            Return _strRephasing
        End Get
        Set(ByVal value As String)
            _strRephasing = value
        End Set
    End Property
    '16_09_2009  ragava
    Public Property ClevisCapPortOrientation() As String
        Get
            Return strClevisCapPortOrientation
        End Get
        Set(ByVal value As String)
            strClevisCapPortOrientation = value
        End Set
    End Property
    '16_09_2009  ragava
    Public Property RodLength() As Double
        Get
            Return dblRodLength
        End Get
        Set(ByVal value As Double)
            dblRodLength = value
        End Set
    End Property

    '16_09_2009  ragava
    Public Property TubeLength() As Double
        Get
            Return dblTubeLength
        End Get
        Set(ByVal value As Double)
            dblTubeLength = value
        End Set
    End Property
    '16_09_2009  ragava
    Public Property TieRodLength() As Double
        Get
            Return dblTieRodLength
        End Get
        Set(ByVal value As Double)
            dblTieRodLength = value
        End Set
    End Property

    '14_09_2009  ragava
    Public Property ClevisCapPort() As String
        Get
            Return strClevisCapPort
        End Get
        Set(ByVal value As String)
            strClevisCapPort = value
        End Set
    End Property

    '14_09_2009  ragava
    Public Property RodCapPort() As String
        Get
            Return strRodCapPort
        End Get
        Set(ByVal value As String)
            strRodCapPort = value
        End Set
    End Property

    '14_09_2009  ragava
    Public Property PinHoleSize() As Double
        Get
            Return dblPinHoleSize
        End Get
        Set(ByVal value As Double)
            dblPinHoleSize = value
        End Set
    End Property

    Public Property pinHoleType() As String
        Get
            Return _strPinHoleType
        End Get
        Set(ByVal value As String)
            _strPinHoleType = value
        End Set
    End Property

    '02_11_2009  Ragava
    Public Property RodClevisPinHoleType() As String
        Get
            Return _strRodClevisPinHoleType
        End Get
        Set(ByVal value As String)
            _strRodClevisPinHoleType = value
        End Set
    End Property



    '11_09_2009  ragava
    Public Property dblStrokeLength() As Double
        Get
            Return _dblStrokeLength
        End Get
        Set(ByVal value As Double)
            _dblStrokeLength = value
        End Set
    End Property
    '11_09_2009  ragava
    Public Property dblPistonThreadSize() As Double
        Get
            Return _dblPistonThreadSize
        End Get
        Set(ByVal value As Double)
            _dblPistonThreadSize = value
        End Set
    End Property
    '11_09_2009  ragava
    Public Property dblRodThreadSize() As Double
        Get
            Return _dblRodThreadSize
        End Get
        Set(ByVal value As Double)
            _dblRodThreadSize = value
        End Set
    End Property
    '11_09_2009  ragava
    Public Property dblBoreDiameter() As Double
        Get
            Return _dblBoreDiameter
        End Get
        Set(ByVal value As Double)
            _dblBoreDiameter = value
        End Set
    End Property
    '11_09_2009  ragava
    Public Property RodMaterialNumber() As String
        Get
            Return _strRodMaterialNumber
        End Get
        Set(ByVal value As String)
            _strRodMaterialNumber = value
        End Set
    End Property
    '11_09_2009  ragava
    Public Property BoreDrawingNumber() As String
        Get
            Return _strBoreDrawingNumber
        End Get
        Set(ByVal value As String)
            _strBoreDrawingNumber = value
        End Set
    End Property
    '11_09_2009  ragava
    Public Property RodDrawingNumber() As String
        Get
            Return _strRodDrawingNumber
        End Get
        Set(ByVal value As String)
            _strRodDrawingNumber = value
        End Set
    End Property
    '11_09_2009  ragava
    Public Property TieRodDrawingNumber() As String
        Get
            Return _strTieRodDrawingNumber
        End Get
        Set(ByVal value As String)
            _strTieRodDrawingNumber = value
        End Set
    End Property
    '11_09_2009  ragava
    Public Property GeneralNotes() As String
        Get
            Return _strGeneralNotes
        End Get
        Set(ByVal value As String)
            _strGeneralNotes = value
        End Set
    End Property
    '08_09_2009   ragava
    Public Property strRodMaterial() As String
        Get
            Return _strRodMaterial
        End Get
        Set(ByVal value As String)
            _strRodMaterial = value
        End Set
    End Property
    Public Property strCylinderDescription() As String
        Get
            Return _strCylinderDescription
        End Get
        Set(ByVal value As String)
            _strCylinderDescription = value
        End Set
    End Property
    Public Property strPaintPackagingNotes() As String
        Get
            Return _strPaintPackagingNotes
        End Get
        Set(ByVal value As String)
            _strPaintPackagingNotes = value
        End Set
    End Property
    Public Property strAssemblyNotes() As String
        Get
            Return _strAssemblyNotes
        End Get
        Set(ByVal value As String)
            _strAssemblyNotes = value
        End Set
    End Property
    '10_09_2009  RAGAVA
    Public Property dblRodDiameter() As String
        Get
            Return _dblRodDiameter
        End Get
        Set(ByVal value As String)
            _dblRodDiameter = value
        End Set
    End Property
    '10_09_2009  RAGAVA
    Public Property Series() As String
        Get
            Return _strSeries
        End Get
        Set(ByVal value As String)
            _strSeries = value
        End Set
    End Property

    Public Property IFLConnectionObject() As IFLConnectionClass
        Get
            Return _oIFLConnection
        End Get
        Set(ByVal value As IFLConnectionClass)
            _oIFLConnection = value
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
    Public Property CustomerName() As String
        Get
            Return _strCustomerName
        End Get
        Set(ByVal value As String)
            _strCustomerName = value
        End Set
    End Property
    Public Property ContractNumber() As String
        Get
            Return _strContractNumber
        End Get
        Set(ByVal value As String)
            _strContractNumber = value
        End Set
    End Property
    Public Property AssemblyType() As String
        Get
            Return _strType
        End Get
        Set(ByVal value As String)
            _strType = value
        End Set
    End Property
    Public Property PartCode() As String
        Get
            Return _strPartCode
        End Get
        Set(ByVal value As String)
            _strPartCode = value
        End Set
    End Property
    Public Property PartCode1() As String
        Get
            Return _strPartCode1
        End Get
        Set(ByVal value As String)
            _strPartCode1 = value
        End Set
    End Property
    Public Property revisionNumber() As Integer
        Get
            Return _intRevision
        End Get
        Set(ByVal value As Integer)
            _intRevision = value
        End Set
    End Property

    Public Property Execution_Path() As String
        Get
            Return _strExecution_Path
        End Get
        Set(ByVal value As String)
            _strExecution_Path = value
        End Set
    End Property
    Public Property Execution_Path1() As String
        Get
            Return _strExecution_Path1
        End Get
        Set(ByVal value As String)
            _strExecution_Path1 = value
        End Set
    End Property

    Public ReadOnly Property FunctionalClassObject() As Object
        Get
            If _oFunctionalClass Is Nothing Then
                _oFunctionalClass = New FunctionalClass
            End If

            Return _oFunctionalClass
        End Get
    End Property

    Public Property SaveFilePath() As String
        Get
            Return _strSavefiletopath
        End Get
        Set(ByVal value As String)
            _strSavefiletopath = value
        End Set
    End Property

    Public Property DestinationFilePath() As String
        Get
            Return _strDestPath
        End Get
        Set(ByVal value As String)
            _strDestPath = value
        End Set
    End Property

    Public Property arAllFormsCotrolValues() As ArrayList
        Get
            If _arAllFormsCotrolValues Is Nothing Then
                _arAllFormsCotrolValues = New ArrayList
            End If
            Return _arAllFormsCotrolValues
        End Get
        Set(ByVal value As ArrayList)
            _arAllFormsCotrolValues = value
        End Set
    End Property

    Public Property arListInputs() As ArrayList
        Get
            If _arListInputs Is Nothing Then
                _arListInputs = New ArrayList
            End If
            Return _arListInputs
        End Get
        Set(ByVal value As ArrayList)
            _arListInputs = value
        End Set
    End Property
#End Region
#Region "Procedures"
    '02_09_2009  ragava
    Public Sub KillExcel()
        Dim proc As System.Diagnostics.Process
        Try
            For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
                If proc.HasExited = False Then
                    proc.Kill()
                End If
            Next
        Catch oException As Exception

        End Try
    End Sub
    Public Sub callCalculations()

    End Sub
    Public Function checkConnections() As Boolean
        checkConnections = False
        If GetConnection() Then
            checkConnections = True
        End If
        'If oDBLoginObject.ConnectWithRegistry Then
        '    IFLConnectionObject = oDBLoginObject.CurrentConnection()
        '    checkConnections = True
        'Else
        '    checkConnections = False
        '    Application.Exit()
        'End If
    End Function
    Public Sub updateHashTablevalues(ByVal hashTab As Hashtable, ByVal key As String, ByVal value As Object, Optional ByVal DefaultValue As Object = "")
        If hashTab.Contains(key) = True Then
            hashTab(key) = value
        Else
            hashTab.Add(key, value)
        End If
    End Sub
#End Region

    Private Function GetConnection() As Boolean
        GetConnection = False
        Try
            Dim strXMLFilePath As String = System.Environment.CurrentDirectory + "\MILConnection.xml"
            Dim oDataSet As New DataSet
            oDataSet.ReadXml(strXMLFilePath)
            If Not oDataSet.Tables.Count <= 0 Then
                Dim strServer As String = oDataSet.Tables(0).Rows(0).Item(0).ToString
                Dim strDataBase As String = oDataSet.Tables(0).Rows(0).Item(1).ToString
                IFLConnectionObject = IFLBaseDataLayer.IFLConnectionClass.GetConnectionObject(strServer, strDataBase, "System.Data.SqlClient", "", "", "SSPI")
                GetConnection = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error occured while connecting to server", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Function
#Region "Enums"
    Public Enum FormObjects
        FormName
        ControlName
    End Enum
#End Region

    'ReleaseCylinder
#Region "ANUP 26-10-2010 START"
    Private _blnIsRowInserted_Tube As Boolean
    Private _blnIsRowInserted_Rod As Boolean
    Private _blnIsRowInserted_Tierod As Boolean
    Private _blnIsRowInserted_StopTube As Boolean

    Public Property IsRowInserted_Tube() As Boolean
        Get
            Return _blnIsRowInserted_Tube
        End Get
        Set(ByVal value As Boolean)
            _blnIsRowInserted_Tube = value
        End Set
    End Property
    Public Property IsRowInserted_Rod() As Boolean
        Get
            Return _blnIsRowInserted_Rod
        End Get
        Set(ByVal value As Boolean)
            _blnIsRowInserted_Rod = value
        End Set
    End Property
    Public Property IsRowInserted_Tierod() As Boolean
        Get
            Return _blnIsRowInserted_Tierod
        End Get
        Set(ByVal value As Boolean)
            _blnIsRowInserted_Tierod = value
        End Set
    End Property
    Public Property IsRowInserted_StopTube() As Boolean
        Get
            Return _blnIsRowInserted_StopTube
        End Get
        Set(ByVal value As Boolean)
            _blnIsRowInserted_StopTube = value
        End Set
    End Property
#End Region


End Module

