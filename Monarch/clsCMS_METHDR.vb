Imports Microsoft.Office.Interop
Public Class clsCMS_METHDR

#Region "Private Variables"

    Private _objExcelSheet As Excel.Worksheet

    Private _strMethodType As String 'Always 1

    Private _strAlternateMethodNumber As String 'Always 0

    Private _strPlantCode As String 'Always C01

    Private _strPartNumber As String

    Private _strSeq As String

    Private _strProcess As String 'Always Blank

    Private _strDepartment As String 'TODO

    Private _strResource As String 'TODO

    Private _strOperation As String 'TODO

    Private _strSetupStandard As String 'TODO

    Private _strSetupCrewSize As String 'Always 1

    Private _strOfMen As String 'Always 1

    Private _strOfMachines As String 'Always 1

    Private _strScheduleRunStandard As String 'TODO

    Private _strCostingRunStandard As String 'TODO

    Private _strRunType As String 'Always A

    Private _strLagTime As String 'Always 0

    Private _strCycleTime_SecPart As String 'Always 0

    Private _strTransferBatch As String 'Always 1

    Private _strMultipleParts As String 'Always 0

    Private _strReportingPoint As String 'Always Y

    Private _strOperationEfficiency As String 'Always 100

    Private _strSchedPriorityGroup As String 'Always Blank

    Private _strBurdenDriverRateFactor As String 'Always 0

    Private _strMRP_Create_RepetitiveJobs_By_TransferBatchQty As String 'Always 2

    Private _strStandardCostRollUp_ByTransferBatch As String 'Always 2

    Private _strConcurrentResources As String 'Always Blank

    Private _strLine As String 'Always Blank

    Private _strStdUnits As String 'Always Blank

    Private _strEfficiencyFactorBeforeSeq As String 'Always Blank

    Private _strEfficiencyFactorAfterSeq As String 'Always Blank
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

    Public Property Seq() As String
        Get
            Return _strSeq
        End Get
        Set(ByVal value As String)
            _strSeq = value
            'anup 16-03-2011 
            'anup 10-03-2011 
            'If _strSeq = 10 Then
            '    _strOfMen = 4
            'ElseIf _strSeq = 30 Then 'anup 11-03-2011 20 changed to 30
            '    _strOfMen = 17
            'Else
            '    _strOfMen = 1
            'End If
        End Set
    End Property

    Public ReadOnly Property Process() As String
        Get
            Return _strProcess
        End Get
    End Property

    Public Property Department() As String
        Get
            Return _strDepartment
        End Get
        Set(ByVal value As String)
            _strDepartment = value
        End Set
    End Property

    Public Property Resource() As String
        Get
            Return _strResource
        End Get
        Set(ByVal value As String)
            _strResource = value
        End Set
    End Property

    Public Property Operation() As String
        Get
            Return _strOperation
        End Get
        Set(ByVal value As String)
            _strOperation = value
        End Set
    End Property

    Public Property SetupStandard() As String
        Get
            Return _strSetupStandard
        End Get
        Set(ByVal value As String)
            _strSetupStandard = value
        End Set
    End Property

    Public ReadOnly Property SetupCrewSize() As String
        Get
            Return _strSetupCrewSize
        End Get
    End Property

    Public Property OfMen() As String 'anup 16-03-2011  Men value will be changing so removed readonly
        Get
            Return _strOfMen
        End Get
        Set(ByVal value As String)
            _strOfMen = value
        End Set

    End Property

    Public Property OfMachines() As String  'anup 16-03-2011  Men value will be changing so removed readonly
        Get
            Return _strOfMachines
        End Get
        Set(ByVal value As String)
            _strOfMachines = value
        End Set
    End Property

    Public Property ScheduleRunStandard() As String
        Get
            Return _strScheduleRunStandard
        End Get
        Set(ByVal value As String)
            _strScheduleRunStandard = value
        End Set
    End Property

    Public Property CostingRunStandard() As String
        Get
            Return _strCostingRunStandard
        End Get
        Set(ByVal value As String)
            _strCostingRunStandard = value
        End Set
    End Property

    Public ReadOnly Property RunType() As String
        Get
            Return _strRunType
        End Get
    End Property

    Public ReadOnly Property LagTime() As String
        Get
            Return _strLagTime
        End Get
    End Property

    Public ReadOnly Property CycleTime_SecPart() As String
        Get
            Return _strCycleTime_SecPart
        End Get
    End Property

    Public ReadOnly Property TransferBatch() As String
        Get
            Return _strTransferBatch
        End Get
    End Property

    Public ReadOnly Property MultipleParts() As String
        Get
            Return _strMultipleParts
        End Get
    End Property

    Public ReadOnly Property ReportingPoint() As String
        Get
            Return _strReportingPoint
        End Get
    End Property

    Public ReadOnly Property OperationEfficiency() As String
        Get
            Return _strOperationEfficiency
        End Get
    End Property

    Public ReadOnly Property SchedPriorityGroup() As String
        Get
            Return _strSchedPriorityGroup
        End Get
    End Property

    Public ReadOnly Property BurdenDriverRateFactor() As String
        Get
            Return _strBurdenDriverRateFactor
        End Get
    End Property

    Public ReadOnly Property MRP_Create_RepetitiveJobs_By_TransferBatchQty() As String
        Get
            Return _strMRP_Create_RepetitiveJobs_By_TransferBatchQty
        End Get
    End Property

    Public ReadOnly Property StandardCostRollUp_ByTransferBatch() As String
        Get
            Return _strStandardCostRollUp_ByTransferBatch
        End Get
    End Property

    Public ReadOnly Property ConcurrentResources() As String
        Get
            Return _strConcurrentResources
        End Get
    End Property

    Public ReadOnly Property Line() As String
        Get
            Return _strLine
        End Get
    End Property

    Public ReadOnly Property _StdUnits() As String
        Get
            Return _strStdUnits
        End Get
    End Property

    Public ReadOnly Property EfficiencyFactorBeforeSeq() As String
        Get
            Return _strEfficiencyFactorBeforeSeq
        End Get
    End Property

    Public ReadOnly Property EfficiencyFactorAfterSeq() As String
        Get
            Return _strEfficiencyFactorAfterSeq
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
        _strProcess = ""
        _strSetupCrewSize = 1
        _strOfMen = 1
        _strOfMachines = 1
        _strRunType = "A"
        _strLagTime = 0
        _strCycleTime_SecPart = 0
        _strTransferBatch = 1
        _strMultipleParts = 0
        _strReportingPoint = "Y"
        _strOperationEfficiency = 100
        _strSchedPriorityGroup = ""
        _strBurdenDriverRateFactor = 0
        _strMRP_Create_RepetitiveJobs_By_TransferBatchQty = 2
        _strStandardCostRollUp_ByTransferBatch = 2
        _strConcurrentResources = ""
        _strLine = ""
        _strStdUnits = ""
        _strEfficiencyFactorBeforeSeq = ""
        _strEfficiencyFactorAfterSeq = ""
    End Sub

    Public Sub SetDataToExcel(ByVal oExcelSheet As Excel.Worksheet, ByVal strRowcount As String)
        Try
            _objExcelSheet = oExcelSheet

            SetDataToCell("A" + strRowcount, MethodType)
            SetDataToCell("B" + strRowcount, AlternateMethodNumber)
            SetDataToCell("C" + strRowcount, PlantCode)
            SetDataToCell("D" + strRowcount, PartNumber)
            SetDataToCell("E" + strRowcount, Seq)
            SetDataToCell("F" + strRowcount, Process)
            SetDataToCell("G" + strRowcount, Department)
            SetDataToCell("H" + strRowcount, Resource)
            SetDataToCell("I" + strRowcount, Operation)
            SetDataToCell("J" + strRowcount, SetupStandard)
            SetDataToCell("K" + strRowcount, SetupCrewSize)
            SetDataToCell("L" + strRowcount, OfMen)
            SetDataToCell("M" + strRowcount, OfMachines)
            SetDataToCell("N" + strRowcount, ScheduleRunStandard)
            SetDataToCell("O" + strRowcount, CostingRunStandard)
            SetDataToCell("P" + strRowcount, RunType)
            SetDataToCell("Q" + strRowcount, LagTime)
            SetDataToCell("R" + strRowcount, CycleTime_SecPart)
            SetDataToCell("S" + strRowcount, TransferBatch)
            SetDataToCell("T" + strRowcount, MultipleParts)
            SetDataToCell("U" + strRowcount, ReportingPoint)
            SetDataToCell("V" + strRowcount, OperationEfficiency)
            SetDataToCell("W" + strRowcount, SchedPriorityGroup)
            SetDataToCell("X" + strRowcount, BurdenDriverRateFactor)
            SetDataToCell("Y" + strRowcount, MRP_Create_RepetitiveJobs_By_TransferBatchQty)
            SetDataToCell("Z" + strRowcount, StandardCostRollUp_ByTransferBatch)
            SetDataToCell("AA" + strRowcount, ConcurrentResources)
            SetDataToCell("AB" + strRowcount, Line)
            SetDataToCell("AC" + strRowcount, _StdUnits)
            SetDataToCell("AD" + strRowcount, EfficiencyFactorBeforeSeq)
            SetDataToCell("AE" + strRowcount, EfficiencyFactorAfterSeq)
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
