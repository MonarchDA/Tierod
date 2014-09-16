Imports MonarchFunctionalLayer
Imports System.IO

Public Class clsCNCUtil

    Private _message As String
    Private _dblThreadedLegth As Double
    Private _dblThreadedSize As Double

    Public ReadOnly Property Message() As String
        Get
            Return _message
        End Get
    End Property

    Public Function DoCNCCodeGeneration() As Boolean

        Try
            _message = String.Empty
            Dim oCYLRos As CYL_Rod
            oCYLRos = CreateInstance()
            ' oCYLRos = Test()  '''''USE THIS FOR TESTING


            CNCFilePath()
            '14_07_2011   RAGAVA
            Dim strFilePath As String
            If IsNew_Revision_Released = "Released" Then
                strFilePath = StringConstants.CNCC.FolderPath & oCYLRos.ProgNo
            Else
                If Directory.Exists("C:\MONARCH_TESTING\CNC_TEMP") = False Then
                    Directory.CreateDirectory("C:\MONARCH_TESTING\CNC_TEMP")
                End If
                strFilePath = "C:\MONARCH_TESTING\CNC_TEMP\" & oCYLRos.ProgNo
            End If
            'Dim strFilePath As String = StringConstants.CNCC.FolderPath & oCYLRos.ProgNo
            'Till  Here

            Dim oCYL_RodService As New CYL_RodService
            Dim oDataBase As New CNCDataBaseClass
            If oDataBase.DoesPartCodeExist(oCYLRos) Then
                oDataBase.UpdateCyl_RodData(oCYLRos)
            Else
                oDataBase.InsertCyl_RodData(oCYLRos)
            End If

            If oCYL_RodService.Start(strFilePath, oCYLRos) Then
                Return True
            Else
                _message = oCYL_RodService.Message
                Return False
            End If

        Catch ex As Exception
            _message = "CNC Code Generation Failed." & vbLf & ex.Message
            Return False
        End Try
    End Function

    Private Sub CNCFilePath()
        Try
            If Not Directory.Exists(StringConstants.CNCC.FolderPath) Then
                Directory.CreateDirectory(StringConstants.CNCC.FolderPath)
            End If
        Catch ex As Exception

        End Try
    End Sub


    Private Function Test() As CYL_Rod
        Dim oCYLRos As New CYL_Rod
        oCYLRos.NominalThreadDia = 1
        oCYLRos.LargeDia = 1.25
        oCYLRos.SmallDia = 1.001
        oCYLRos.TH_Per_IN = 14
        oCYLRos.ShoulderType = "Chamfer"
        oCYLRos.RodType = "Chrome"
        oCYLRos.PartNo = "111111"
        oCYLRos.ByName = "MD"
        ' SetRodEndConfigValues(oCYLRos)
        oCYLRos.setDescription("RD CYL 1.00-16.00-1.25-1.13")
        oCYLRos.output2ndop = True
        oCYLRos.Secondthreaddia = 1.25
        oCYLRos.Secondoptype = "Threaded"
        oCYLRos.Secondthreadnum = 14

        oCYLRos.Drawing_Num = Val(oCYLRos.PartNo)
        oCYLRos.Xhome = 20
        oCYLRos.Zhome = 20
        oCYLRos.Operation = 20
        oCYLRos.WorkCenter = 122
        oCYLRos.AutoDoor = True
        oCYLRos.Drawing_Rev = 1
        oCYLRos.Secondopzzero = -0.03
        oCYLRos.Secondthreadcornerrad = 0.06
        oCYLRos.Secondshoulder = -0.03
        oCYLRos.skimdiameter = 1.24
        oCYLRos.chamferdepthofcut = 0.19
        ' SetExcelValues(oCYLRos)
        oCYLRos.Length = 25.26
        oCYLRos.Th_Length = 1.14
        oCYLRos.Secondthreadlength = 1.13

        oCYLRos.Secondchamfer = 0
        oCYLRos.skimlength = 0



        Return oCYLRos
    End Function


    Private Function CreateInstance() As CYL_Rod
        Try

            Dim oCYLRos As New CYL_Rod
            Dim oDataBase As New CNCDataBaseClass

            oCYLRos.NominalThreadDia = PistonThreadSizeValue()
            oCYLRos.LargeDia = RodDiameter
            oCYLRos.SmallDia = oCYLRos.NominalThreadDia + 0.001
            oCYLRos.TH_Per_IN = oDataBase.GetTH_Per_In(oCYLRos.NominalThreadDia)
            oCYLRos.ShoulderType = "Chamfer"
            oCYLRos.RodType = RodMaterialValue()
            oCYLRos.PartNo = strRodCodeNumber
            oCYLRos.ByName = "MD"
            SetRodEndConfigValues(oCYLRos, oCYLRos.NominalThreadDia, oCYLRos.PartNo)
            setSecondThreadValue(oCYLRos)
            oCYLRos.Drawing_Num = Val(oCYLRos.PartNo)
            oCYLRos.Xhome = 20
            oCYLRos.Zhome = 20
            oCYLRos.Operation = 20
            oCYLRos.WorkCenter = 122
            oCYLRos.AutoDoor = AutoDoorValidation(oCYLRos.WorkCenter)
            oCYLRos.Drawing_Rev = 1
            oCYLRos.Secondopzzero = -0.03
            oCYLRos.Secondthreadcornerrad = 0.06
            oCYLRos.skimdiameter = RodDiameter - 0.015
            oCYLRos.chamferdepthofcut = 0.19
            oCYLRos.Length = RodLength
            oCYLRos.Th_Length = Set_Th_LengthValue(oCYLRos.LargeDia)
            oCYLRos.Secondthreadlength = _dblThreadedLegth
            oCYLRos.Secondshoulder = Math.Round(RodLength - oCYLRos.Secondthreadlength, 3)
            Return oCYLRos

        Catch ex As Exception

        End Try
    End Function


    Private Function AutoDoorValidation(ByVal dblWorkCenter As Double) As Boolean
        If dblWorkCenter = 122 Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function Set_Th_LengthValue(ByVal dblRodDiameter As Double) As Double
        Try

            Dim oReadExcel As New ReadExcel
            Dim strExcelPath As String

            If SeriesForCosting = "TX (TXC)" Then
                Set_Th_LengthValue = Th_LengthValue_1()
            Else
                strExcelPath = StringConstants.CNCC.RodWelded_ExcelPath + ExcelFilePath()
                If Not String.IsNullOrEmpty(strExcelPath) Then
                    If oReadExcel.Open(strExcelPath) Then
                        Set_Th_LengthValue = Th_LengthValue_2(oReadExcel)
                    End If
                End If
            End If

        Catch ex As Exception

        End Try

    End Function

    Private Function ExcelFilePath() As String
        Try
            If dblRodDiameter = 1.25 Then
                ExcelFilePath = "RodDia_1.25.xls"
            ElseIf dblRodDiameter = 1.5 Then
                ExcelFilePath = "RodDia_1.50.xls"
            ElseIf dblRodDiameter = 1.75 Then
                ExcelFilePath = "RodDia_1.75.xls"
            ElseIf dblRodDiameter = 1.38 Then
                ExcelFilePath = "RodDia_1.375.xls"
            ElseIf dblRodDiameter = 2 Then
                ExcelFilePath = "RodDia_2.00.xls"
            End If
        Catch ex As Exception

        End Try
    End Function

    Private Function Th_LengthValue_1() As Double
        Try
            If dblRodDiameter = 1.25 Then
                Th_LengthValue_1 = 1.0739
            ElseIf dblRodDiameter = 1.5 Then
                Th_LengthValue_1 = 0.97
            ElseIf dblRodDiameter = 1.13 Then
                Th_LengthValue_1 = 1.0739
            End If
        Catch ex As Exception

        End Try
    End Function

    Private Function Th_LengthValue_2(ByVal oReadExcel As ReadExcel) As Double
        Try
            If dblRodDiameter = 1.25 Then
                Th_LengthValue_2 = Val(oReadExcel.Read("D3"))
            ElseIf dblRodDiameter = 1.5 Then
                Th_LengthValue_2 = Val(oReadExcel.Read("E3"))
            ElseIf dblRodDiameter = 1.75 Then
                Th_LengthValue_2 = Val(oReadExcel.Read("E3"))
            ElseIf dblRodDiameter = 1.13 Then
                Th_LengthValue_2 = 0.94
            ElseIf dblRodDiameter = 1.38 Then
                Th_LengthValue_2 = Val(oReadExcel.Read("E3"))
            ElseIf dblRodDiameter = 2 Then
                Th_LengthValue_2 = Val(oReadExcel.Read("E3"))
            End If
        Catch ex As Exception

        End Try
    End Function

    Private Function PistonThreadSizeValue() As Double
        Dim dblPistonThreadSize As Double = Val(PistonThreadSize)
        If dblPistonThreadSize = 0.5 Then
            PistonThreadSizeValue = 0.5
        ElseIf dblPistonThreadSize = 0.63 Then
            PistonThreadSizeValue = 0.625
        ElseIf dblPistonThreadSize = 0.75 Then
            PistonThreadSizeValue = 0.75
        ElseIf dblPistonThreadSize = 0.87 Then
            PistonThreadSizeValue = 0.875
        ElseIf dblPistonThreadSize = 1.0 Then
            PistonThreadSizeValue = 1.0
        ElseIf dblPistonThreadSize = 1.12 Then
            PistonThreadSizeValue = 1.125
        ElseIf dblPistonThreadSize = 1.26 OrElse dblPistonThreadSize = 1.25 Then
            PistonThreadSizeValue = 1.25
        ElseIf dblPistonThreadSize = 1.5 Then
            PistonThreadSizeValue = 1.5
        End If
    End Function

    'Private Function PistonThreadSizeConversion_FractionToDecimals() As Double
    '    Try
    '        'For Each oNutSizes As DictionaryEntry In NutSizesInFractions
    '        '    If PistonThreadSize = oNutSizes.Key.ToString Then
    '        '        PistonThreadSizeConversion_FractionToDecimals = Val(oNutSizes.Value)
    '        '        Exit For
    '        '    End If
    '        'Next
    '        For Each oNutSizes As Object In NutSizesInFractions
    '            If oNutSizes(0) = PistonThreadSize Then
    '                PistonThreadSize = oNutSizes(1)
    '                Exit For
    '            End If
    '        Next
    '    Catch ex As Exception

    '    End Try
    'End Function


    'Private Sub SetExcelValues(ByVal oCYLRos As CYL_Rod)
    '    Dim oReadExcel As New ReadExcel
    '    If Not oReadExcel.Open(StringConstants.CNCC.RodWelded_ExcelPath) Then
    '        Return
    '    End If
    '    oCYLRos.Length = Val(oReadExcel.Read("C5"))
    ' oCYLRos.Th_Length = Val(oReadExcel.Read("G5"))
    '    oCYLRos.Secondthreadlength = Val(oReadExcel.Read("Z5"))
    '    oReadExcel.Close()
    'End Sub

    Private Function RodMaterialValue() As String
        Try
            If RodMaterialForCosting = "LION 1000" Then
                RodMaterialValue = "Chrome"
            Else
                RodMaterialValue = RodMaterialForCosting
            End If
        Catch ex As Exception

        End Try
    End Function

    Private Sub SetRodEndConfigValues(ByVal oCYLRos As CYL_Rod, ByVal dblNominalThreadDia As Double, ByVal strRodPartCode As String)
        Try
            Dim dblPistonNutS As Double = dblNominalThreadDia
            Dim dblStrokeLength As Double = StrokeLength
            _dblThreadedSize = dblRodThreadSize

            _dblThreadedLegth = RodThreadLength(oCYLRos.LargeDia, oCYLRos.RodType, strStyleModified, _dblThreadedSize)

            oCYLRos.setDescription(dblPistonNutS, dblStrokeLength, _dblThreadedSize, _dblThreadedLegth)
            oCYLRos.output2ndop = True
            oCYLRos.Secondthreaddia = _dblThreadedSize
            oCYLRos.Secondoptype = StringConstants.CNCC.Threaded

            '' ''        oCYLRos.Secondchamfer = dblWeldSize
            '' ''        oCYLRos.skimlength = ObjClsWeldedCylinderFunctionalClass.ObjClsWeldedGlobalVariables.SkimWidth + dblWeldSize
        Catch ex As Exception

        End Try
    End Sub

    Private Sub setSecondThreadValue(ByVal oCYLRos As CYL_Rod)
        Dim oDataBase As New CNCDataBaseClass
        oCYLRos.Secondthreadnum = oDataBase.GetTH_Per_In(_dblThreadedSize)
    End Sub

    Public Function RodThreadLength(ByVal dblRodDia As Double, ByVal strMaterialType As String, ByVal strIsAsae As String, ByVal dblRodThreadSize As Double) As Double
        Try
            Dim _strQuery As String
            Dim _strErrorMessage As String

            _strQuery = "select RodThreadLength from dbo.RodDiameterDetails where RodDiameter = " + dblRodDia.ToString + " and MaterialType ='" + strMaterialType + "'"
            _strQuery += " and IsASAE ='" + strIsAsae + "' and RodThreadSize =" + dblRodThreadSize.ToString
            RodThreadLength = IFLConnectionObject.GetValue(_strQuery)
            If IsNothing(RodThreadLength) Then
                RodThreadLength = 0
                _strErrorMessage = "Data not retrieved from RodDiameterDetails table" + vbCrLf
            End If
        Catch ex As Exception
            RodThreadLength = 0
        End Try
    End Function


End Class
