Imports Microsoft.Office.Interop
Public Class clsCMS_MTHL

#Region "Private Variables"
    Private _objExcelSheet As Excel.Worksheet

    Private _strMethodType As String 'Always 1

    Private _strAlternateMethodNumber As String 'Always 0

    Private _strPlantCode As String 'Always C01

    Private _strPart As String 'ToDo

    Private _strSeq As String 'ToDo

    Private _strLine As String 'ToDo

    Private _strTool As String 'ToDo

    Private _strRequiredSetup As String 'Always Y

    Private _strRequiredRun As String 'Always Y

    Private _strQuantity As String 'Always 1

    Private _strUnits As String 'Always EA
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

    Public Property Part() As String
        Get
            Return _strPart
        End Get
        Set(ByVal value As String)
            _strPart = value
        End Set
    End Property

    Public Property Seq() As String
        Get
            Return _strSeq
        End Get
        Set(ByVal value As String)
            _strSeq = value
        End Set
    End Property

    Public Property Line() As String
        Get
            Return _strLine
        End Get
        Set(ByVal value As String)
            _strLine = value
        End Set
    End Property

    Public Property Tool() As String
        Get
            Return _strTool
        End Get
        Set(ByVal value As String)
            _strTool = value
        End Set
    End Property

    Public ReadOnly Property RequiredSetup() As String
        Get
            Return _strRequiredSetup
        End Get
    End Property

    Public ReadOnly Property RequiredRun() As String
        Get
            Return _strRequiredRun
        End Get
    End Property

    Public ReadOnly Property Quantity() As String
        Get
            Return _strQuantity
        End Get
    End Property

    Public ReadOnly Property Units() As String
        Get
            Return _strUnits
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
        _strRequiredSetup = "Y"
        _strRequiredRun = "Y"
        _strQuantity = 1
        _strUnits = "EA"
    End Sub

    Public Sub SetDataToExcel(ByVal oExcelSheet As Excel.Worksheet, ByVal strRowcount As String)
        Try
            _objExcelSheet = oExcelSheet
            If Tool <> "" Then          '04_10_2010   RAGAVA
                SetDataToCell("A" + strRowcount, MethodType)
                SetDataToCell("B" + strRowcount, AlternateMethodNumber)
                SetDataToCell("C" + strRowcount, PlantCode)
                SetDataToCell("D" + strRowcount, Part)
                SetDataToCell("E" + strRowcount, Seq)
                SetDataToCell("F" + strRowcount, Line)
                SetDataToCell("G" + strRowcount, Tool)
                SetDataToCell("H" + strRowcount, RequiredSetup)
                SetDataToCell("I" + strRowcount, RequiredRun)
                SetDataToCell("J" + strRowcount, Quantity)
                SetDataToCell("K" + strRowcount, Units)
            End If
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
