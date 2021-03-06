Imports MonarchFunctionalLayer
Imports MonarchAPILayer
Public Class frmTieRod1

#Region "Class Level Variables"

    Private oBool As Boolean = False

    Dim _objNutSizeDT As DataTable
    Dim dblCylinderForce As Double
    Dim _dbltempWorkingPressureSeries As Double
    Dim _dbltempWorkingPressureRod As Double
    Dim _dbltempWorkingPressureNut As Double
    Dim _blnRodThreadSize As Boolean = False

#End Region

#Region "Class Level Properties"

    Public Property IsErrorMessageTierod1() As Boolean     
        Get
            Return oBool
        End Get
        Set(ByVal value As Boolean)
            oBool = value
        End Set
    End Property

    Public Property dbltempWorkingPressureRod() As Double
        Get
            Return _dbltempWorkingPressureRod
        End Get
        Set(ByVal value As Double)
            _dbltempWorkingPressureRod = value
        End Set
    End Property

    Public Property dbltempWorkingPressureNut() As Double
        Get
            Return _dbltempWorkingPressureNut
        End Get
        Set(ByVal value As Double)
            _dbltempWorkingPressureNut = value
        End Set
    End Property

    Public Property dbltempWorkingPressureSeries() As Double
        Get
            Return _dbltempWorkingPressureSeries
        End Get
        Set(ByVal value As Double)
            _dbltempWorkingPressureSeries = value
        End Set
    End Property

    Private ReadOnly Property BoreDiameters() As ArrayList
        Get
            BoreDiameters = New ArrayList
            BoreDiameters.Add(New Object() {"TX (TXC)", 2, 2.5, 3, 3.5, 4})
            BoreDiameters.Add(New Object() {"TL (TC)", 2, 2.5, 3, 3.5, 4, 4.5, 5})
            BoreDiameters.Add(New Object() {"TH (TD)", 2, 2.5, 3, 3.5, 4, 4.5, 5})
            BoreDiameters.Add(New Object() {"TP-High", 2, 2.5, 2.75, 3, 3.25, 3.5, 3.75, 4, 4.25, 4.5, 4.75, 5})
            BoreDiameters.Add(New Object() {"TP-Low", 3, 3.25, 3.5, 3.75, 4, 4.25, 4.5, 4.75, 5})
            'ANUP 12-10-2010 START
            BoreDiameters.Add(New Object() {"LN", 2, 2.5, 3, 3.5, 4, 4.5, 5})
            'ANUP 12-10-2010 TILL HERE
            Return BoreDiameters
        End Get
    End Property

    Private ReadOnly Property TensileArea() As ArrayList
        Get
            TensileArea = New ArrayList
            TensileArea.Add(New Object(1) {0.75, 0.373})
            TensileArea.Add(New Object(1) {1, 0.68})
            TensileArea.Add(New Object(1) {1.12, 0.856})
            TensileArea.Add(New Object(1) {1.13, 0.856})
            TensileArea.Add(New Object(1) {1.25, 1.073})
            TensileArea.Add(New Object(1) {1.5, 1.58})
            TensileArea.Add(New Object(1) {1.75, 2.19})
        End Get
    End Property

#End Region

    Private Sub TieRod1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ColorTheForm()
        SetDefaultValues()

        CylinderCodeNumber = GetCylinderCodeNumber()
        LoadInformation()
        CylinderReleasedFunctionality()

    End Sub

    Public Function GetCylinderCodeNumber() As String

        GetCylinderCodeNumber = ""
        Try
            If oContractListviewItem Is Nothing Then
                Try
                    Dim iContractNumber As Long = 800001
                    Dim strQuery As String
                    strQuery = "select top 1 ContractNumber from ContractNumberDetails order by ContractNumber Desc"
                    Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                    If objDT.Rows.Count > 0 Then
                        iContractNumber = Val(objDT.Rows(0).Item("ContractNumber").ToString)
                        iContractNumber = iContractNumber + 1
                    End If
                    Try
                        strQuery = "Insert into ContractNumberDetails Values ('" & iContractNumber.ToString & "')"
                        objDT = oDataClass.GetDataTable(strQuery)
                    Catch ex As Exception
                    End Try
                    GetCylinderCodeNumber = iContractNumber.ToString
                Catch ex As Exception
                End Try
            Else
                GetCylinderCodeNumber = oContractListviewItem.SubItems(0).Text
            End If
            'TODO: Sunny 20-04-10 10am
            'ObjClsCostingDetails.AddCodeNumberToDataTable(GetCylinderCodeNumber, "Cylinder Code Number") 'Sandeep 04-03-10-4pm
        Catch ex As Exception
        End Try
        Return GetCylinderCodeNumber

    End Function

    Private Sub SetDefaultValues()

        If blnRevision = False Then
            txtRodAdder.Text = 0
            rdbStopTubeNo.Checked = True
            optStrokeControlYes.Checked = True
        End If
     
    End Sub

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub cmbStyle_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As _
                    System.EventArgs) Handles cmbStyle.SelectedIndexChanged
        '26_10_2009  ragava  
        If Trim(sender.Text) <> "" Then
            ComboBoxStyle()
        End If
        RetractedLengthCalculation()

    End Sub

    Public Sub ComboBoxStyle()

        If strStyleModified <> Trim(cmbStyle.Text) Then
            strStyleModified = Trim(cmbStyle.Text)
            txtRetractedLength.Clear()
            txtExtendedLength.Clear()
            StrokeLength = 0
            RodAdder = 0
            txtRecommendedStoptubeLength.Clear()
            txtStopTubeLength.Clear()
        End If

        '26_10_2009  ragava  Till  Here

        Try
            loadBoreDiameterValues(SeriesForCosting)
            cmbStrokeLength.Items.Clear()
            cmbClevisCapPinHole.Items.Clear()
            cmbRodMaterial.Items.Clear()
            LVRodDiameterDetails.Clear()
            cmbClevisCapPort.Items.Clear()
            cmbRodCapPort.Items.Clear()
            LVNutSizeDetails.Clear()
            cmbStrokeLengthAdder.Items.Clear()
        Catch ex As Exception

        End Try
        If Trim(cmbStyle.Text) = "ASAE" Then
            txtRodAdder.Text = 0
            txtRodAdder.Enabled = False
            rdbStopTubeYes.Enabled = False
            rdbStopTubeNo.Checked = True
            txtStrokeLength.Visible = False
            txtStrokeLength.Enabled = False
            cmbStrokeLength.Visible = True
            cmbStrokeLength.Enabled = True
            cmbStrokeLength.Items.Clear()
            cmbStrokeLength.Items.Add(" ")
            optStrokeControlNo.Checked = True
            grbStrokeControl.Enabled = True
        Else
            txtRodAdder.Enabled = True
            rdbStopTubeYes.Enabled = True
            txtStrokeLength.Visible = True
            txtStrokeLength.Text = ""
            txtStrokeLength.BringToFront()
            txtStrokeLength.Enabled = True
            cmbStrokeLength.SendToBack()
            cmbStrokeLength.Visible = False
            cmbStrokeLength.Enabled = False
            optStrokeControlNo.Checked = True
            grbStrokeControl.Enabled = False
        End If
        strStyle = Trim(cmbStyle.Text)

    End Sub

    Public Function checkStrokeLength(ByVal boreDiameter As Double) As Boolean
        '14_10_2009
        checkStrokeLength = False
        Try
            Dim strQuery As String = ""
            If Trim(boreDiameter) <> "" AndAlso cmbStyle.Text = "ASAE" Then
                strQuery = "select distinct StrokeLength from RoddiameterDetails  ,BoreDiameter_RodDiameter bdrd" + vbNewLine
                'ANUP 12-10-2010 START
                strQuery = strQuery + " where Series= '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "' and IsASAE = '" & cmbStyle.Text & "'" & " and bdrd.BoreDiameterID=(select BoreDiameterID from BoreDiameterMaster where BoreDiameter =  " & boreDiameter & ")"
                'ANUP 12-10-2010 TILL HERE
                Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                If objDT.Rows.Count >= 1 Then
                    checkStrokeLength = True
                End If
            End If
        Catch ex As Exception
        End Try

    End Function

    Private Sub cmbSeries_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As _
                        System.EventArgs) Handles cmbSeries.SelectedIndexChanged

        If cmbSeries.Text <> "" Then
            ComboBoxSeries()
        End If

    End Sub

    Public Sub ComboBoxSeries()

        If Module1.BtnBrowseClicked Then
            SeriesForCosting = Module1.ReadValuesFromExcel.Series
        Else
            SeriesForCosting = cmbSeries.Text
        End If

        '18_02_2011    RAGAVA
        If SeriesForCosting = "THC" Then
            SeriesForCosting = "LN"
        End If
        'Till   Here

        Try
            CallSeriesLogics()
        Catch ex As Exception

        End Try

        Try
            Dim strQuery As String
            ofrmTieRod2.cmbPaint.Items.Clear()
            ofrmTieRod2.cmbPaint.Items.Add(" ")
            strQuery = ""
            strQuery = "select Distinct PaintColor,IFLID from PaintDetails order by IFLID"           '19_04_2010   RAGAVA
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            For Each dr As DataRow In objDT.Rows
                ofrmTieRod2.cmbPaint.Items.Add(dr(0).ToString)
            Next
            If objDT.Rows.Count = 1 Then
                ofrmTieRod2.cmbPaint.Text = ofrmTieRod2.cmbPaint.Items(1).ToString()
                ofrmTieRod2.cmbPaint.Enabled = False
            Else
                ofrmTieRod2.cmbPaint.Text = ofrmTieRod2.cmbPaint.Items(1).ToString()
                ofrmTieRod2.cmbPaint.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Sub CallSeriesLogics()

        cmbRodMaterial.Items.Clear()
        If LVRodDiameterDetails.Items.Count > 0 Then
            LVRodDiameterDetails.Clear()
        End If
        If LVNutSizeDetails.Items.Count > 0 Then
            LVNutSizeDetails.Clear()
        End If

        If Not cmbSeries.Text Is Nothing Then
            Series = Trim(SeriesForCosting)
            If SeriesForCosting.StartsWith("TP") Then
                cmbRephasingPortPosition.Enabled = True
                loadRephasingPositions()
                'ofrmTieRod3.chk100AirTest.Checked = True  VAMSI 10-09-2014
                ofrmTieRod3.chk100OilTest.Checked = True
            Else
                cmbRephasingPortPosition.Items.Clear()
                cmbRephasingPortPosition.Enabled = False
                'ofrmTieRod3.chk100AirTest.Checked = False VAMSI 10-09-2014
                ofrmTieRod3.chk100OilTest.Checked = False
            End If
            If SeriesForCosting.StartsWith("TX") Then
                LVNutSizeDetails.Enabled = False
            Else
                LVNutSizeDetails.Enabled = True
            End If
            CodeDesc = ""
            If SeriesForCosting = "TX (TXC)" Then
                CodeDesc = "TXC"
                txtWorkingPressure.Text = 2500
            ElseIf SeriesForCosting = "TL (TC)" Then
                CodeDesc = "TC"
                txtWorkingPressure.Text = 2500
            ElseIf SeriesForCosting = "TH (TD)" Then
                CodeDesc = "TD"
                txtWorkingPressure.Text = 3000
            ElseIf SeriesForCosting = "TP-High" Then
                CodeDesc = "TP"
                txtWorkingPressure.Text = 3000
            ElseIf SeriesForCosting = "TP-Low" Then
                txtWorkingPressure.Text = 3000
                CodeDesc = "TP"
                'ANUP 12-10-2010 START
                'ElseIf cmbSeries.Text = "LN" Then
            ElseIf SeriesForCosting = "LN" Then
                txtWorkingPressure.Text = 3000
                CodeDesc = "LN"
                'ANUP 12-10-2010 TILL HERE
            Else
                txtWorkingPressure.Text = ""
            End If
            dbltempWorkingPressureSeries = txtWorkingPressure.Text
            cmbBore.Items.Clear()
            cmbBore.Items.Add(" ")
            For Each oItem1 As Object In BoreDiameters
                If oItem1(0).ToString.Equals(SeriesForCosting) Then
                    For Each oItem2 As Object In oItem1
                        If Not oItem2.ToString.Equals(SeriesForCosting) Then
                            cmbBore.Items.Add(oItem2)
                        End If
                    Next
                End If
            Next
            WorkingPressure = Val(txtWorkingPressure.Text)
            LoadInformation()
        End If

    End Sub

    Private Sub loadRephasingPositions()

        cmbRephasingPortPosition.Items.Clear()
        cmbRephasingPortPosition.Items.Add(" ")
        cmbRephasingPortPosition.Items.Add("At Extension")
        cmbRephasingPortPosition.Items.Add("At Retraction")
        If Trim(SeriesForCosting) <> "TP-High" Then
            cmbRephasingPortPosition.Items.Add("Both")
        End If

    End Sub

    Private Sub cmbRodMaterial_SelectedIndexChanged(ByVal sender As System.Object, ByVal e _
                    As System.EventArgs) Handles cmbRodMaterial.SelectedIndexChanged

        ComboBoxRodMaterial()

    End Sub

    Public Sub ComboBoxRodMaterial()

        Try
            'Sandeep 20-04-10
            If cmbRodMaterial.Text <> "" Then
                RodMaterialForCosting = cmbRodMaterial.Text
                LVNutSizeDetails.Items.Clear()
                PistonThreadSize = ""
                LVRodDiameterDetails.Items.Clear()
                txtWorkingPressure.Text = dbltempWorkingPressureSeries             '03_11_2009   Ragava
                strRodMaterial = Trim(cmbRodMaterial.Text)
                RodDiameterCalcultion()
                LoadTieRodSizes()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RodDiameterCalcultion()

        Try
            Dim strQuery As String = ""
            Dim aColumns As New ArrayList
            strQuery = "select distinct rd.RodDiameter  from RodDiameterDetails rd,BoreDiameter_RodDiameter bdrd ,RodCapDetails rc,PistonSealDetails p " + vbNewLine
            strQuery = strQuery + "where bdrd.BoreDiameterID=(select BoreDiameterID from BoreDiameterMaster where" + vbNewLine
            strQuery = strQuery + "BoreDiameter =  " & Val(cmbBore.Text) & ") and bdrd.PartNumberID = rd.PartNumber and (rc.RodDiameter=rd.RodDiameter)" + vbNewLine
            'ANUP 12-10-2010 START
            strQuery = strQuery + "And rd.Series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "'  and IsASAE = '" & cmbStyle.Text & "'" + vbNewLine
            'ANUP 12-10-2010 TILL HERE
            strQuery = strQuery + "and MaterialType='" & cmbRodMaterial.Text & "' and  (not rc.hallite='' or not rc.Zmacro=''or   not rc.Notes='')"
            strQuery = strQuery + "and (not p.Oring='' or not p.BackUpRing='' or not p.PTFESeal='' or not p.OringExpander='' or not p.PSPSeal='' or not p.WearRing1='' or p.WearRing2='') And  p.BoreDiameter=" & Val(cmbBore.Text)

            If cmbStyle.Text = "ASAE" Then
                strQuery = strQuery + " And StrokeLength=" & Val(cmbStrokeLength.Text)
            End If
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            LVRodDiameterDetails.Columns.Clear()
            aColumns.Add(New Object(2) {"RodDiameter", "Rod Diameter", True})
            LVRodDiameterDetails.DisplayHeaders = aColumns
            Dim oTable As DataTable = objDT

            oTable.Columns.Add("Derate Pressure At Maximum Extension")
            For Each oRow As DataRow In oTable.Rows
                'If checkRodSealDetails(oRow(0), cmbRodMaterial.Text) = True          '20_01_2011        RAGAVA
                If checkRodSealDetails(oRow(0), cmbRodMaterial.Text) = True OrElse SeriesForCosting = "LN" Then
                    Dim dblRatedForce As Double = Math.Pow(3.1416, 2) * 0.049 * Math.Pow(IIf(Val(oRow("RodDiameter")) = 1.12, 1.125, IIf(Val(oRow("RodDiameter")) = 1.38, 1.375, Val(oRow("RodDiameter")))), 4) * 30 * _
                                   (Math.Pow(10, 6) / Math.Pow(Val(txtExtendedLength.Text), 2))
                    Dim dblAllowablePressure As Double = dblRatedForce / (Math.Pow(BoreDiameter, 2) * 0.7854)
                    If dblAllowablePressure <= WorkingPressure Then
                        oRow("Derate Pressure At Maximum Extension") = Math.Round(dblAllowablePressure, 2)
                    Else
                        oRow("Derate Pressure At Maximum Extension") = "N/A"
                    End If
                Else
                    oRow.Delete()
                End If
            Next
            oTable.AcceptChanges()
            LVRodDiameterDetails.FlushListViewData()
            LVRodDiameterDetails.SourceTable = oTable
            LVRodDiameterDetails.Populate()
            LVRodDiameterDetails.Focus()
            '09_10_2009
            If Not Module1.BtnBrowseClicked Then
                If oTable.Rows.Count = 1 Then
                    LVRodDiameterDetails.Items(0).Selected = True
                    LVRodDiameterDetails.Enabled = False
                Else
                    LVRodDiameterDetails.Items(0).Selected = True
                    LVRodDiameterDetails.Enabled = True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub loadBoreDiameterValues(ByVal SeriesType As String)

        cmbBore.Items.Clear()
        cmbBore.Items.Add(" ")
        If SeriesType.StartsWith("TX") Then
            cmbBore.Items.Add("2")
            cmbBore.Items.Add("2.5")
            cmbBore.Items.Add("3")
            cmbBore.Items.Add("3.5")
            cmbBore.Items.Add("4")
            'ANUP 12-10-2010 START
        ElseIf SeriesType.StartsWith("TL") OrElse SeriesType.StartsWith("TH") OrElse SeriesType.StartsWith("LN") Then
            'ANUP 12-10-2010 TILL HERE
            cmbBore.Items.Add("2")
            cmbBore.Items.Add("2.5")
            cmbBore.Items.Add("3") ' 2, 2.25,2.5, 2.75,3, 3.25, 3.5, 3.75, 4, 4.25, 4.5, 4.75, 5
            cmbBore.Items.Add("3.5")
            cmbBore.Items.Add("4")
            cmbBore.Items.Add("4.5")
            cmbBore.Items.Add("5")
        ElseIf SeriesType = "TP-Low" Then
            If Trim(cmbRephasingPortPosition.Text) = "Both" Then
                cmbBore.Items.Add("2.5")
                cmbBore.Items.Add("2.75")
                cmbBore.Items.Add("3")
                cmbBore.Items.Add("3.5")
            ElseIf Trim(cmbRephasingPortPosition.Text) = "At Retraction" Then
                cmbBore.Items.Add("2")
                cmbBore.Items.Add("2.5")
                cmbBore.Items.Add("2.75")
                cmbBore.Items.Add("3")
                cmbBore.Items.Add("3.25")
                cmbBore.Items.Add("4")
            ElseIf Trim(cmbRephasingPortPosition.Text) = "At Extension" Then
                cmbBore.Items.Add("2")
                cmbBore.Items.Add("2.5")
                cmbBore.Items.Add("2.75")
                cmbBore.Items.Add("3")
                cmbBore.Items.Add("3.25")
                cmbBore.Items.Add("3.5")
                cmbBore.Items.Add("3.75")
                cmbBore.Items.Add("4")
                cmbBore.Items.Add("4.25")
                cmbBore.Items.Add("4.5")
                cmbBore.Items.Add("4.75")
                cmbBore.Items.Add("5")
            End If
        ElseIf SeriesType = "TP-High" Then
            If Trim(cmbRephasingPortPosition.Text) = "At Retraction" Then
                cmbBore.Items.Add("3.25")
                cmbBore.Items.Add("3.5")
            ElseIf Trim(cmbRephasingPortPosition.Text) = "At Extension" Then
                cmbBore.Items.Add("3")
                cmbBore.Items.Add("3.25")
                cmbBore.Items.Add("3.5")
                cmbBore.Items.Add("3.75")
                cmbBore.Items.Add("4")
                cmbBore.Items.Add("4.25")
                cmbBore.Items.Add("4.5")
                cmbBore.Items.Add("4.75")
                cmbBore.Items.Add("5")
            End If
        End If
        For intCount As Integer = 1 To cmbBore.Items.Count - 1
            If cmbStyle.Text = "ASAE" Then
                If checkStrokeLength(cmbBore.Items(intCount)) = False Then
                    cmbBore.Items.Remove(cmbBore.Items(intCount))
                End If
            End If
        Next

    End Sub

    Public Sub ComboBoxPinHole()

        Try
            If Not cmbClevisCapPinHole.Text Is Nothing AndAlso cmbClevisCapPinHole.Text <> "" Then
                cmbClevisCapPort.Items.Clear()
                cmbClevisCapPort.Items.Add(" ")
                Dim strQuery As String = ""
                Dim aColumns As New ArrayList
                strQuery = "select DISTINCT Port from ClevisCapDetails where" & vbNewLine
                '13_07_2011  RAGAVA
                'strQuery = strQuery & GetClevisCapDetails() & vbNewLine
                If SeriesForCosting <> "TX (TXC)" Then
                    strQuery = strQuery & "Series <>'TX'"
                Else
                    strQuery = strQuery & GetClevisCapDetails() & vbNewLine
                End If
                'Till   Here
                strQuery = strQuery & " and BoreDiameter=" & cmbBore.Text.ToString & vbNewLine
                strQuery = strQuery & " And PinHoleType='" & cmbClevisCapPinHole.Text & "'"
                Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                For Each dr As DataRow In objDT.Rows
                    cmbClevisCapPort.Items.Add(dr(0).ToString)
                Next
                '09_10_2009
                If objDT.Rows.Count = 1 Then
                    cmbClevisCapPort.SelectedIndex = 1
                    cmbClevisCapPort.Enabled = False
                Else
                    cmbClevisCapPort.Enabled = True
                End If
                pinHoleType = Trim(cmbClevisCapPinHole.Text)
                RodClevisPinHoleType = Trim(cmbRodClevisPinHole.Text)            '02_11_2009   Ragava
                StopTubeLength = Val(txtStopTubeLength.Text)
                CaluculateLength(True)
                Try
                    Call LoadRodMaterial()
                Catch ex As Exception

                End Try
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub cmbBore_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                        Handles cmbBore.SelectedIndexChanged          '03_12_2009  Ragava  EventHandlers Added

        RetractedLengthCalculation()
        ComboBoxBore(sender)

    End Sub

    Public Sub ComboBoxBore(ByVal sender As Object)

        strNewTableDrawingNumber = ""         '02_12_2009   RAGAVA
        BoreDiameter = Val(cmbBore.Text)
        cmbClevisCapPinHole.Items.Clear()
        cmbRodMaterial.Items.Clear()
        LVRodDiameterDetails.Clear()
        cmbClevisCapPort.Items.Clear()
        cmbRodCapPort.Items.Clear()
        LVNutSizeDetails.Clear()
        cmbStrokeLengthAdder.Items.Clear()

        '13_10_2009
        Try
            If sender.Name <> "txtStopTubeLength" AndAlso Not (sender.Name = "txtStrokeLength" Or sender.Name = "cmbStrokeLength") Then    '03_12_2009    Ragava
                If Trim(cmbBore.Text) <> "" AndAlso Trim(cmbStyle.Text) = "ASAE" Then
                    Dim strQuery As String = ""
                    cmbStrokeLength.Items.Clear()
                    cmbStrokeLength.Items.Add(" ")
                    strQuery = "select distinct StrokeLength from RoddiameterDetails  ,BoreDiameter_RodDiameter bdrd" + vbNewLine
                    'ANUP 12-10-2010 START
                    strQuery = strQuery + " where Series= '" & IIf(Trim(SeriesForCosting.ToString). _
                        StartsWith("TX"), "TX", "TL/TH/TP/LN") & "' and IsASAE = '" & cmbStyle.Text & "'" _
                        & " and bdrd.BoreDiameterID=(select BoreDiameterID from BoreDiameterMaster where BoreDiameter =  " _
                        & cmbBore.Text & ")"
                    'ANUP 12-10-2010 TILL HERE
                    Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                    For Each dr As DataRow In objDT.Rows
                        If CheckRodMaterialLoading(dr(0)) = True Then
                            cmbStrokeLength.Items.Add(dr(0).ToString)
                        End If
                    Next
                    '14_10_2009
                    If objDT.Rows.Count = 1 Then
                        cmbStrokeLength.SelectedIndex = 1
                        cmbStrokeLength.Enabled = False
                    Else
                        cmbStrokeLength.Enabled = True
                    End If
                Else
                    txtStrokeLength.Clear()             '26_10_2009   ragava
                End If
            End If
        Catch ex As Exception
        End Try

        Try
            Dim strQuery As String = ""
            If Trim(cmbBore.Text) <> "" Then
                BoreDiameter = Val(cmbBore.Text)
                dblBoreDiameter = BoreDiameter
                cmbClevisCapPinHole.Items.Clear()
                cmbClevisCapPinHole.Items.Add(" ")
                strQuery = "select distinct PinHoleType from clevisCapDetails where BoreDiameter = " & _
                        cmbBore.Text & " and " & GetClevisCapDetails()
                Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                For Each dr As DataRow In objDT.Rows
                    cmbClevisCapPinHole.Items.Add(dr(0).ToString)
                Next
                '09_10_2009
                If objDT.Rows.Count = 1 Then
                    cmbClevisCapPinHole.SelectedIndex = 1
                    cmbClevisCapPinHole.Enabled = False
                Else
                    cmbClevisCapPinHole.Text = "Standard"                   '29_10_2009   Ragava
                    cmbClevisCapPinHole.Enabled = True
                End If

                '02_11_2009   Ragava
                cmbRodClevisPinHole.Items.Clear()
                cmbRodClevisPinHole.Items.Add(" ")
                'strQuery = "select distinct PinHoleType from RodClevisDetails where " & GetRodClevisDetails()
                strQuery = "select distinct PinHoleType from RodClevisDetails"       '14_07_2010    RAGAVA
                Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery)
                For Each dr As DataRow In objDT1.Rows
                    cmbRodClevisPinHole.Items.Add(dr(0).ToString)
                Next
                If objDT1.Rows.Count = 1 Then
                    cmbRodClevisPinHole.SelectedIndex = 1
                    cmbRodClevisPinHole.Enabled = False
                Else
                    cmbRodClevisPinHole.Text = "Standard"
                    cmbRodClevisPinHole.Enabled = True
                End If
                '02_11_2009   Ragava   Till   Here
                Try
                    'cmbBore.SelectedIndexChanged, cmbStrokeLength.SelectedIndexChanged, txtStrokeLength.Leave, txtStopTubeLength.Leave
                    If Trim(cmbBore.Text) <> "" AndAlso (Trim(cmbStrokeLength.Text) <> "" Or _
                        Trim(txtStrokeLength.Text) <> "") Then 'AndAlso Trim(txtStopTubeLength.Text) <> "" Then        '03_12_2009   Ragava
                        Dim StrSql As String
                        'ANUP 12-10-2010 START
                        StrSql = "select * from TieRodSizes trz,BoreDiameter_TieRodSizes bdtr where bdtr.PartNumberID = trz.TieRodPartNumber and bdtr.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(cmbBore.Text) & ")and trz.series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "'"
                        'ANUP 12-10-2010 TILL HERE
                        Dim objDT3 As DataTable = oDataClass.GetDataTable(StrSql)
                        If objDT3.Rows.Count > 0 Then
                            strTieRodCodeNumber = objDT3.Rows(0).Item("TieRodPartNumber").ToString
                            strTieRodDrawingNumber = objDT3.Rows(0).Item("DrawingNumber").ToString
                            strTieRodDescription = objDT3.Rows(0).Item("Description").ToString
                            mdiMonarch.mdiComponent.Items(6).SubItems.Add(" ")
                            mdiMonarch.mdiComponent.Items(6).SubItems.Add(strTieRodDrawingNumber)
                            mdiMonarch.mdiComponent.Items(6).SubItems.Add(strTieRodDescription)
                            '02_12_2009   ragava

                            TieRodStrokeDifference = Math.Round(Val(objDT3.Rows(0).Item("Dimension-A").ToString) _
                                - Val(objDT3.Rows(0).Item("StrokeLength").ToString), 2) 'StrokeLength

                            '06_01_2011    RAGAVA
                            If Trim(cmbStrokeLength.Text) <> "" Then
                                StrokeLength = Val(Trim(cmbStrokeLength.Text))
                            ElseIf Trim(txtStrokeLength.Text) <> "" Then
                                StrokeLength = Val(Trim(txtStrokeLength.Text))
                            End If
                            'Till   Here


                            Dim dblTieRodLength As Double = StrokeLength + TieRodStrokeDifference + StopTubeLength
                            strQuery = "Select CodeNumber,Dim_A,Revision from TieRodTableDrawing where DrawingNumber = '" _
                                & strTieRodDrawingNumber & "' and Dim_A = " & Math.Round(dblTieRodLength, 2).ToString
                            Dim objDT3_Temp As DataTable = oDataClass.GetDataTable(strQuery)
                            If objDT3_Temp.Rows.Count > 0 Then
                                strTieRodCodeNumber = objDT3_Temp.Rows(0).Item("CodeNumber").ToString
                            Else
                                Dim strQuery1 As String = "Select CodeNumber from CodeNumberDetails where Type = 'TieRod'"
                                Dim objDT3_Temp1 As DataTable = oDataClass.GetDataTable(strQuery1)
                                strNewTableDrawingNumber = objDT3_Temp1.Rows(0).Item("CodeNumber").ToString
                                strNewTieRodTableDrawingNumber = strNewTableDrawingNumber
                                strTieRodCodeNumber = strNewTableDrawingNumber
                            End If

                        Else
                            Dim strQuery1 As String = "Select CodeNumber from CodeNumberDetails where Type = 'TieRod'"
                            Dim objDT3_Temp As DataTable = oDataClass.GetDataTable(strQuery1)
                            strNewTableDrawingNumber = objDT3_Temp.Rows(0).Item("CodeNumber").ToString
                            strNewTieRodTableDrawingNumber = strNewTableDrawingNumber
                            '04_10_2010   RAGAVA
                            'strRodCodeNumber = strNewTableDrawingNumber     
                            strTieRodCodeNumber = strNewTableDrawingNumber
                            'Till  Here
                            '02_12_2009   ragava  Till  Here
                        End If
                        RetractedLengthCalculation()             'Temp
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
                LoadInformation()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Try
            callTubeSealLogics()
            LoadRodMaterial()
        Catch ex As Exception
        End Try

    End Sub

    Public Function CheckRodMaterialLoading(ByVal strokeLength As Double) As Boolean

        CheckRodMaterialLoading = False
        If Not SeriesForCosting Is Nothing AndAlso Not cmbBore.Text Is Nothing AndAlso Not cmbStyle.Text Is Nothing Then
            cmbRodMaterial.Enabled = True
            Dim strSQL As String = ""
            strSQL = "SELECT distinct rd.materialtype FROM    " + vbNewLine
            strSQL = strSQL & "RodDiameterDetails rd,BoreDiameter_RodDiameter bdrd " + vbNewLine
            'ANUP 12-10-2010 START
            strSQL = strSQL & "where bdrd.BoreDiameterID=(select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(cmbBore.Text) & ") and  bdrd.PartNumberID = rd.PartNumber AND  IsASAE='" & cmbStyle.Text & "' and Series='" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & " '"
            'ANUP 12-10-2010 TILL HERE
            If cmbStyle.Text = "ASAE" Then
                strSQL = strSQL & "And StrokeLength = " & strokeLength
            End If
            Dim objDT As DataTable = oDataClass.GetDataTable(strSQL)
            If objDT.Rows.Count > 0 Then
                CheckRodMaterialLoading = True
            End If
        End If

    End Function

    Private Sub LVNutSizeDetails_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) _
                                        Handles LVNutSizeDetails.SelectedIndexChanged

        ListViewNutSizeDetails()

    End Sub

    Public Sub ListViewNutSizeDetails()

        dbltempWorkingPressureNut = 0
        If Module1.BtnBrowseClicked Then
            If LVNutSizeDetails.Enabled Then
                For i As Integer = 0 To LVNutSizeDetails.Items.Count - 1
                    If LVNutSizeDetails.Items(i).Text = Module1.ReadValuesFromExcel.NutSize.ToString() Then
                        dblPistonThreadSize = Convert.ToDouble(LVNutSizeDetails.Items(i).Text)
                        PistonThreadSize = dblPistonThreadSize
                        If LVNutSizeDetails.Items(i).SubItems(2).Text <> "N/A" Then
                            If Val(txtWorkingPressure.Text) > Val(LVNutSizeDetails.Items(i).SubItems(2).Text) Then

                                txtWorkingPressure.Text = LVNutSizeDetails.Items(i).SubItems(2).Text
                                dbltempWorkingPressureNut = Val(txtWorkingPressure.Text)
                            End If
                        Else
                            dbltempWorkingPressureNut = dbltempWorkingPressureSeries
                            txtWorkingPressure.Text = dbltempWorkingPressureSeries
                        End If
                        Exit For
                    Else
                        Dim s As String = Convert.ToDouble(Module1.ReadValuesFromExcel.NutSize.ToString())
                        Dim words As String() = s.Split(New Char() {"."c})
                        Dim word As String

                        If words.Length = 2 Then
                            word = ""
                            word = words(1)
                            If Not word.Length = 2 Then
                                word = word + "0"
                                word = words(0) + "." + word
                            End If
                        ElseIf words.Length = 1 Then
                            word = ""
                            word = "00"
                            word = words(0) + "." + word
                        End If
                        If LVNutSizeDetails.Items(i).Text = word Then
                            dblPistonThreadSize = Convert.ToDouble(LVNutSizeDetails.Items(i).Text)
                            PistonThreadSize = dblPistonThreadSize
                            If LVNutSizeDetails.Items(i).SubItems(2).Text <> "N/A" Then
                                If Val(txtWorkingPressure.Text) > Val(LVNutSizeDetails.Items(i).SubItems(2).Text) Then

                                    txtWorkingPressure.Text = LVNutSizeDetails.Items(i).SubItems(2).Text
                                    dbltempWorkingPressureNut = Val(txtWorkingPressure.Text)
                                End If
                            Else
                                dbltempWorkingPressureNut = dbltempWorkingPressureSeries
                                txtWorkingPressure.Text = dbltempWorkingPressureSeries
                            End If
                            Exit For
                        End If
                    End If
                Next

            Else
                If LVNutSizeDetails.Items.Count > 0 Then
                    dblPistonThreadSize = Convert.ToDouble(LVNutSizeDetails.Items(0).Text)
                    PistonThreadSize = dblPistonThreadSize
                    If LVNutSizeDetails.Items(0).SubItems(2).Text <> "N/A" Then
                        If Val(txtWorkingPressure.Text) > Val(LVNutSizeDetails.Items(0).SubItems(2).Text) Then

                            txtWorkingPressure.Text = LVNutSizeDetails.Items(0).SubItems(2).Text
                            dbltempWorkingPressureNut = Val(txtWorkingPressure.Text)
                        End If
                    Else
                        dbltempWorkingPressureNut = dbltempWorkingPressureSeries
                        txtWorkingPressure.Text = dbltempWorkingPressureSeries
                    End If
                End If
            End If
          
        Else
            If LVNutSizeDetails.SelectedItems.Count > 0 Then

                Dim oSelectedListviewItem As ListViewItem = LVNutSizeDetails.SelectedItems(0)
                dblPistonThreadSize = Val(oSelectedListviewItem.SubItems(0).Text)
                PistonThreadSize = dblPistonThreadSize
                If oSelectedListviewItem.SubItems(2).Text <> "N/A" Then
                    If Val(txtWorkingPressure.Text) > Val(oSelectedListviewItem.SubItems(2).Text) Then

                        txtWorkingPressure.Text = oSelectedListviewItem.SubItems(2).Text
                        dbltempWorkingPressureNut = Val(txtWorkingPressure.Text)
                    End If
                Else
                    dbltempWorkingPressureNut = dbltempWorkingPressureSeries
                    txtWorkingPressure.Text = dbltempWorkingPressureSeries
                End If
            End If
        End If

        Try
            LoadTieRodSizes()
            LoadPinSizeDetails()
        Catch ex As Exception

        End Try
        For Each listviewItem As ListViewItem In LVNutSizeDetails.Items
            Dim index As Integer = LVNutSizeDetails.Items.IndexOf(listviewItem)
            LVNutSizeDetails.Items(index).BackColor = Color.Ivory
            LVNutSizeDetails.Items(index).ForeColor = Color.Black
        Next
        For Each listviewItem As ListViewItem In LVNutSizeDetails.SelectedItems
            Dim index As Integer = LVNutSizeDetails.Items.IndexOf(listviewItem)
            LVNutSizeDetails.Items(index).BackColor = Color.CornflowerBlue
            LVNutSizeDetails.Items(index).ForeColor = Color.White
        Next

    End Sub

    Private Sub txtWorkingPressure_TextChanged(ByVal sender As System.Object, ByVal e As _
                        System.EventArgs) Handles txtWorkingPressure.TextChanged

        If txtWorkingPressure.Text <> "" Then
            WorkingPressure = Val(txtWorkingPressure.Text)
        End If

    End Sub

    Private Sub cmbStrokeLength_SelectedIndexChanged(ByVal sender As System.Object, ByVal e _
                            As System.EventArgs) Handles cmbStrokeLength.SelectedIndexChanged

        ComboBoxStrokeLength()

    End Sub

    Public Sub ComboBoxStrokeLength()

        cmbRodMaterial.Items.Clear()
        LVRodDiameterDetails.Clear()
        cmbRodCapPort.Items.Clear()
        LVNutSizeDetails.Clear()
        cmbStrokeLengthAdder.Items.Clear()
        If cmbStrokeLength.Text <> "" Then
            StrokeLength = Val(cmbStrokeLength.Text)
            dblStrokeLength = StrokeLength
            Try
                Call LoadRodMaterial()
            Catch ex As Exception

            End Try
        End If
        '26_10_2009  ragava
        If cmbClevisCapPinHole.Items.Count > 1 Then
            cmbClevisCapPinHole.SelectedIndex = 0
            cmbClevisCapPinHole.Text = "Standard"                 '29_10_2009   Ragava
        Else
            cmbClevisCapPinHole.SelectedIndex = -1
        End If
        RetractedLengthCalculation()

    End Sub

    Private Sub txtStopTubeLength_TextChanged(ByVal sender As System.Object, ByVal e _
                                As System.EventArgs) Handles txtStopTubeLength.TextChanged

        TextBoxStopTubeLength_TextChanged()

    End Sub

    Public Sub TextBoxStopTubeLength_TextChanged()

        dblStopTubeLength = 0
        If txtStopTubeLength.Text <> "" Then
            'If Val(txtStopTubeLength.Text) = 0 Then
            '    MessageBox.Show("Minimum Value Required is 0.1", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            '    txtStopTubeLength.
            'End If
            StopTubeLength = Val(txtStopTubeLength.Text)
            dblStopTubeLength = StopTubeLength
        End If
        ComboBoxBore(txtStopTubeLength)

    End Sub

    Private Sub txtRodAdder_TextChanged(ByVal sender As System.Object, ByVal e As _
                                        System.EventArgs) Handles txtRodAdder.TextChanged

        If txtRodAdder.Text <> "" Then
            RodAdder = Val(txtRodAdder.Text)
        End If

    End Sub

    Private Sub rdbStopTubeYes_CheckedChanged(ByVal sender As System.Object, ByVal e As _
                                    System.EventArgs) Handles rdbStopTubeYes.CheckedChanged

        RadioBtnStopTubeYesCheckedChanged()

    End Sub

    Public Sub RadioBtnStopTubeYesCheckedChanged()

        If rdbStopTubeYes.Checked = True Then
            IsStopTubeSelected = True
            txtStopTubeLength.Enabled = True
            txtStopTubeLength.Text = Val(txtRecommendedStoptubeLength.Text)       '26_10_2009  ragava
            StopTubeLength = Val(txtStopTubeLength.Text)    '26_10_2009  ragava
            CaluculateLength(True)         '26_10_2009  ragava
            'ofrmTieRod1.txtStopTubeLength.Text = 0.1  
            'calculateRecommendedStopTubeLength()        '26_10_2009  ragava
        End If

    End Sub

    Private Sub txtStopTubeLengthLeave()

        If StrokeLength <= Val(txtStopTubeLength.Text) Then
            If blnRevision = False Then
                MessageBox.Show("Stop Tube Length should not be more than the Stroke Length", "Warning", _
                                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtStopTubeLength.Text = ""
                txtStopTubeLength.Focus()
            End If
        End If
        Dim result As Double = Val(txtStopTubeLength.Text) Mod 0.125
        If Not result = 0 Then
            If blnRevision = False Then
                MessageBox.Show("Entered value is not multiples of 1/8, So Application is Rounding Off", "Information")
            End If
            result = Val(txtStopTubeLength.Text) / 0.125
            result = Math.Ceiling(result)        '26_10_2009  ragava  'Math.Round
            result = result * 0.125
            txtStopTubeLength.Text = result
        End If
        RetractedLengthCalculation()

    End Sub

    Private Sub txtStopTubeLength_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
                                        Handles txtStopTubeLength.Leave

        txtStopTubeLengthLeave()

    End Sub

    Private Sub txtExtendedLength_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If txtExtendedLength.Text <> "" Then
            RodDiameterCalcultion()
        End If

    End Sub

    Private Sub rdbStopTubeNo_CheckedChanged(ByVal sender As System.Object, ByVal e As _
                                System.EventArgs) Handles rdbStopTubeNo.CheckedChanged

        RadioBtnStopTubeNoCheckedChanged()

    End Sub

    Public Sub RadioBtnStopTubeNoCheckedChanged()

        If rdbStopTubeNo.Checked Then
            txtStopTubeLength.Text = "0"       '26_10_2009  ragava
            txtStopTubeLength.Enabled = False
            StopTubeLength = Val(txtStopTubeLength.Text)    '26_10_2009  ragava
            CaluculateLength(True)         '26_10_2009  ragava
            '26_10_2009  ragava
            'txtRecommendedStoptubeLength.Text = 0
            'calculateRecommendedStopTubeLength()   
            'txtStopTubeLength.Text = 0
            '26_10_2009  ragava  Till  Here
        Else
            IsStopTubeSelected = False
        End If

    End Sub

    Private Sub optStrokeControlYes_CheckedChanged(ByVal sender As System.Object, ByVal e As _
                                System.EventArgs) Handles optStrokeControlYes.CheckedChanged

        RadioBtnStrokeControlYes()

    End Sub

    Public Sub RadioBtnStrokeControlYes()

        If optStrokeControlYes.Checked Then
            grbStrokeControl.Enabled = True
            cmbStrokeLengthAdder.Enabled = True
            cmbStrokeLengthAdder.Items.Clear()
            cmbStrokeLengthAdder.Items.Add(" ")
            If RodDiameter = 1.12 Then
                cmbStrokeLengthAdder.Items.Add("2 Stage")
                cmbStrokeLengthAdder.Items.Add("3 Stage")
                cmbStrokeLengthAdder.Enabled = True
            Else
                cmbStrokeLengthAdder.Items.Add("2 Stage")
                cmbStrokeLengthAdder.SelectedIndex = 1
                cmbStrokeLengthAdder.Enabled = False
            End If
        Else
            cmbStrokeLengthAdder.Items.Clear()
            cmbStrokeLengthAdder.Enabled = False
        End If

    End Sub

    Private Sub cmbStyle_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbStyle.Leave

        ComboBoxcmbStyleLeave(sender)

    End Sub

    Public Sub ComboBoxcmbStyleLeave(ByVal sender As Object)

        If txtStrokeLength.Text = "" Then
            rdbStopTubeNo.Checked = True
            'Else
            '    rdbStopTubeYes.Checked = True          '26_10_2009   ragava
        End If
        If sender.Text = "" Then
            sender.IFLDataTag = ""
            Exit Sub
        End If
        If cmbStyle.Text <> "" AndAlso cmbBore.Text <> "" AndAlso (cmbStrokeLength.Text <> "" OrElse _
                        txtStrokeLength.Text <> "") AndAlso txtRodAdder.Text <> "" Then       '26_10_2009   ragava 
            If sender.Text <> sender.IFLDataTag Then
                sender.IFLDataTag = sender.Text
                RetractedLengthCalculation()

                '26_10_2009   ragava
                If (sender.Name.ToString.IndexOf("StrokeLength") <> -1 Or sender.Name _
                                        = "txtRodAdder") AndAlso Val(txtRetractedLength.Text) > 0 Then
                    'calculateRecommendedStopTubeLength()
                    If calculateRecommendedStopTubeLength() > 0 Then
                        rdbStopTubeYes.Checked = True
                        rdbStopTubeNo.Checked = True
                    End If
                End If
                '26_10_2009   ragava   Till  Here
            End If
        End If

        '26_10_2009  ragava
        If sender.Name = "txtStrokeLength" Or sender.Name = "txtRodAdder" Then
            If cmbClevisCapPinHole.Items.Count > 1 Then
                cmbClevisCapPinHole.SelectedIndex = 0
                cmbClevisCapPinHole.Text = "Standard"
            Else
                cmbClevisCapPinHole.SelectedIndex = -1
            End If
        End If

    End Sub

    Private Sub RetractedLengthCalculation() 'Handles cmbStyle.SelectedIndexChanged, cmbBore.SelectedIndexChanged, txtStrokeLength.Leave, txtRodAdder.Leave, txtStopTubeLength.Leave, cmbStrokeLength.SelectedIndexChanged

        Try
            strNewTableDrawingNumber = strNewTieRodTableDrawingNumber       '02_12_2009   Ragava
            Dim strsql As String
            Dim arrseries As String()
            arrseries = SeriesForCosting.ToString.Split(" ")
            Dim strseries As String = arrseries(0)
            If Trim(cmbRephasingPortPosition.Text) <> "" Then
                'ANUP 13-10-2010 START
                If Trim(cmbRephasingPortPosition.Text) = "At Extension" Then
                    strseries = strseries.Insert(2, "e")
                ElseIf Trim(cmbRephasingPortPosition.Text) = "At Retraction" Then
                    'ANUP 13-10-2010 TILL HERE
                    strseries = strseries.Insert(2, "b")
                Else
                    strseries = strseries.Insert(2, "2")
                End If
            End If
            strsql = "select * from borediameterdetails where borediameter = " & Val(cmbBore.Text) _
                                & " and series like '%" & strseries & "%'"

            '23_11_2009    ragava
            If Trim(ofrmTieRod1.cmbStyle.Text) = "ASAE" Then
                strsql = strsql + " and nominalstroke =" & Val(ofrmTieRod1.cmbStrokeLength.Text)
            Else
                '02_12_2009   Ragava
                'strsql = strsql + " and nominalstroke =" & Val(ofrmTieRod1.txtStrokeLength.Text)
                Try
                    Dim objDT_temp As DataTable = oDataClass.GetDataTable(strsql)
                    If objDT_temp.Rows.Count > 0 Then
                        BoreStrokeDifference = Math.Round(Val(objDT_temp.Rows(0).Item("TubeLength").ToString) _
                                                - Val(objDT_temp.Rows(0).Item("NominalStroke").ToString), 2)
                    End If
                Catch ex As Exception
                End Try
                Dim dblTubeLength As Double = StrokeLength + BoreStrokeDifference + StopTubeLength
                strsql = strsql + " and TubeLength = " & Math.Round(dblTubeLength, 2).ToString
                '02_12_2009   Ragava    Till  Here
            End If
            '23_11_2009    ragava    till   here

            Dim objdt2 As DataTable = oDataClass.GetDataTable(strsql)
            If objdt2.Rows.Count > 0 Then
                strBoreDrawingNumber = objdt2.Rows(0).Item("drawingpartnumber").ToString
                strBoreDescription = objdt2.Rows(0).Item("description").ToString
                strBoreCodeNumber = objdt2.Rows(0).Item("partnumber").ToString
                mdiMonarch.mdiComponent.Items(0).SubItems.Add(" ")
                mdiMonarch.mdiComponent.Items(0).SubItems.Add(strBoreDrawingNumber)
                mdiMonarch.mdiComponent.Items(0).SubItems.Add(strBoreDescription)
                '02_12_2009   ragava
            Else
                Dim strQuery As String = "Select CodeNumber from CodeNumberDetails where Type = 'Tube'"
                Dim objDT3 As DataTable = oDataClass.GetDataTable(strQuery)
                If strNewTableDrawingNumber <> "" Then
                    strNewTableDrawingNumber = (Val(strNewTableDrawingNumber) + 1).ToString
                Else
                    strNewTableDrawingNumber = objDT3.Rows(0).Item("CodeNumber").ToString
                End If
                strNewTubeTableDrawingNumber = strNewTableDrawingNumber
                strBoreCodeNumber = strNewTableDrawingNumber
                '02_12_2009   ragava  Till  Here
            End If

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
        If cmbStyle.Text <> "" AndAlso cmbBore.Text <> "" AndAlso (cmbStrokeLength.Text <> "" OrElse _
                            txtStrokeLength.Text <> "") AndAlso txtRodAdder.Text <> "" Then
            If txtStrokeLength.Text <> "" Or Trim(cmbStrokeLength.Text) <> "" Then     '26_10_2009   ragava   ' If txtStopTubeLength.Text <> "" Then
                CaluculateLength(True)
            End If
        End If

    End Sub

    Private Sub txtStrokeLength_TextChanged(ByVal sender As System.Object, ByVal e As _
                                        System.EventArgs) Handles txtStrokeLength.TextChanged

        If txtStrokeLength.Text <> "" Then
            StrokeLength = Val(txtStrokeLength.Text)
            dblStrokeLength = StrokeLength
        End If

    End Sub

    Private Sub LoadRodMaterial()

        If Not SeriesForCosting Is Nothing AndAlso Not cmbBore.Text Is Nothing AndAlso Not cmbStyle.Text _
                        Is Nothing OrElse Not cmbStrokeLength.Text Is Nothing Then
            cmbRodMaterial.Enabled = True
            Dim strSQL As String = ""
            strSQL = "SELECT distinct rd.materialtype FROM    " + vbNewLine
            strSQL = strSQL & "RodDiameterDetails rd,BoreDiameter_RodDiameter bdrd " + vbNewLine
            'ANUP 12-10-2010 START
            strSQL = strSQL & "where bdrd.BoreDiameterID=(select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " _
                & Val(cmbBore.Text) & ") and  bdrd.PartNumberID = rd.PartNumber AND  IsASAE='" & cmbStyle.Text _
                & "' and Series='" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & " '"
            'ANUP 12-10-2010 TILL HERE
            If cmbStyle.Text = "ASAE" Then
                strSQL = strSQL & "And StrokeLength = " & Val(cmbStrokeLength.Text)
            End If
            Dim objDT As DataTable = oDataClass.GetDataTable(strSQL)
            If objDT.Rows.Count > 0 Then
                cmbRodMaterial.Items.Clear()
                cmbRodMaterial.Items.Add(" ")
                For Each dr As DataRow In objDT.Rows
                    If checkRodDiameter(dr(0).ToString) = True Then
                        cmbRodMaterial.Items.Add(dr(0).ToString)
                    End If
                Next
            End If

            '09_10_2009
            If objDT.Rows.Count = 1 Then
                cmbRodMaterial.SelectedIndex = 1
                cmbRodMaterial.Enabled = False
            Else
                cmbRodMaterial.Enabled = True
            End If
        End If

    End Sub

    Public Function checkRodDiameter(ByVal material As String) As Boolean

        checkRodDiameter = False
        Try
            Dim strQuery As String = ""
            Dim aColumns As New ArrayList
            strQuery = "select distinct rd.RodDiameter  from RodDiameterDetails rd,BoreDiameter_RodDiameter bdrd ,RodCapDetails rc,PistonSealDetails p " + vbNewLine
            strQuery = strQuery + "where bdrd.BoreDiameterID=(select BoreDiameterID from BoreDiameterMaster where" + vbNewLine
            strQuery = strQuery + "BoreDiameter =  " & Val(cmbBore.Text) & ") and bdrd.PartNumberID = rd.PartNumber and (rc.RodDiameter=rd.RodDiameter)" + vbNewLine
            'ANUP 12-10-2010 START
            strQuery = strQuery + "and rd.Series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "'  and IsASAE = '" & cmbStyle.Text & "'" + vbNewLine
            'ANUP 12-10-2010 TILL HERE
            strQuery = strQuery + "and MaterialType='" & material & "' and  (not rc.hallite='' or not rc.Zmacro=''or   not rc.Notes='')"
            strQuery = strQuery + "and (not p.Oring='' or not p.BackUpRing='' or not p.PTFESeal='' or not p.OringExpander='' or not p.PSPSeal='' or not p.WearRing1='' or p.WearRing2='') And  p.BoreDiameter=" & Val(cmbBore.Text)

            If cmbStyle.Text = "ASAE" Then
                strQuery = strQuery + " And StrokeLength=" & Val(cmbStrokeLength.Text)
            End If
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            If objDT.Rows.Count >= 1 Then
                For Each dr As DataRow In objDT.Rows
                    If checkRodSealDetails(dr(0), material) = True Then
                        checkRodDiameter = True
                    End If
                Next

            End If
        Catch ex As Exception
        End Try

    End Function

    Public Function checkRodSealDetails(ByVal rodDiameter As Double, ByVal material As String) As Boolean

        checkRodSealDetails = False
        Dim strQuery1 As String
        strQuery1 = "Select Hallite,ZMacro,Notes from RodCapDetails where BoreDiameter = " & Val(cmbBore.Text) & " and RodDiameter =" & rodDiameter & " and " & GetClevisCapDetails(True) 'ANUP 21-10-2010 START
        Dim objDT As DataTable = oDataClass.GetDataTable(strQuery1)
        If objDT.Rows.Count >= 1 Then
            If checkNutSizeDetails(rodDiameter, material) = True Then
                checkRodSealDetails = True
            End If
        End If

    End Function

    'Public Function checkPistonSealPackage(ByVal Size As Double, ByVal rodDiameter As Double, ByVal material As String) As Boolean
    '    checkPistonSealPackage = False
    '    Dim strQuery As String = ""
    '    Dim arrSeries As String()
    '    arrSeries = ofrmTieRod1.cmbSeries.Text.ToString.Split(" ")
    '    Dim strSeries As String = arrSeries(0)
    '    If Size > 0 Then
    '        strQuery = "Select Oring,BackUpRing,PTFESeal,OringExpander,PSPSeal,WearRing1,WearRing2 from PistonSealDetails where BoreDiameter = " & _
    '        Val(cmbBore.Text) & " and PistonNutSize= " & Size & " and Series like '%" & strSeries & "%'"
    '    Else
    '        strQuery = "Select Oring,BackUpRing,PTFESeal,OringExpander,PSPSeal,WearRing1,WearRing2 from PistonSealDetails where BoreDiameter = " & _
    '        Val(cmbBore.Text) & " and Series like '%" & strSeries & "%'"
    '    End If
    '    Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
    '    If objDT.Rows.Count >= 1 Then
    '        If checkPortPins(rodDiameter, Size, material) = True Then '27_10_2009
    '            checkPistonSealPackage = True
    '        End If
    '    End If
    'End Function

    'ANUP 02-11-2010 START
    Public Function checkPistonSealPackage(ByVal Size As Double, ByVal rodDiameter As Double, ByVal material As String) As Boolean

        checkPistonSealPackage = False
        Dim strQuery As String = ""
        Dim arrSeries As String()
        arrSeries = SeriesForCosting.ToString.Split(" ")
        Dim strSeries As String = arrSeries(0)
        If Size > 0 Then
            strQuery = "Select Oring,BackUpRing,PTFESeal,OringExpander,PSPSeal,WearRing1,WearRing2,WynSeal,GlydP from PistonSealDetails where BoreDiameter = " & _
            Val(cmbBore.Text) & " and PistonNutSize= " & Size & " and Series like '%" & strSeries & "%'"
        Else
            strQuery = "Select Oring,BackUpRing,PTFESeal,OringExpander,PSPSeal,WearRing1,WearRing2,WynSeal,GlydP from PistonSealDetails where BoreDiameter = " & _
            Val(cmbBore.Text) & " and Series like '%" & strSeries & "%'"
        End If
        Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
        If objDT.Rows.Count >= 1 Then
            If checkPortPins(rodDiameter, Size, material) = True Then '27_10_2009
                checkPistonSealPackage = True
            End If
        End If

    End Function
    'ANUP 02-11-2010 TILL HERE

    Public Function checkNutSizeDetails(ByVal rodDiameter As Double, ByVal material As String) As Boolean

        Try
            _objNutSizeDT = Nothing
            If Not SeriesForCosting.StartsWith("TX") Then
                Dim strQuery As String = ""
                strQuery = "SELECT distinct rd.PistonThreadSize  FROM " + vbNewLine
                strQuery = strQuery & "RodDiameterDetails rd,BoreDiameter_RodDiameter bdrd ,PistonSealDetails p " + vbNewLine
                strQuery = strQuery & "where bdrd.BoreDiameterID=(select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(cmbBore.Text) & ")" + vbNewLine
                strQuery = strQuery & " and bdrd.PartNumberID = rd.PartNumber and RodDiameter = " & rodDiameter & " And IsASAE='" & cmbStyle.Text & "' And rd.Series <> 'TX'  and (rd.PistonThreadSize=p.PistonNutSize)"
                strQuery = strQuery & "and (not p.Oring='' or not p.BackUpRing=''or   not p.PTFESeal='' or not p.OringExpander='' or not p.PSPSeal='' or not p.WearRing1='' or p.WearRing2='' or not p.WynSeal='' or not p.GlydP='')"   'ANUP 02-11-2010
                If cmbStyle.Text = "ASAE" Then
                    strQuery = strQuery + " And StrokeLength=" & Val(cmbStrokeLength.Text)
                End If
                _objNutSizeDT = oDataClass.GetDataTable(strQuery)
                If _objNutSizeDT.Rows.Count >= 1 Then
                    For Each dr As DataRow In _objNutSizeDT.Rows
                        If checkPistonSealPackage(dr(0), rodDiameter, material) = True Then
                            checkNutSizeDetails = True
                        End If
                    Next
                End If
            Else
                checkNutSizeDetails = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function
    '27_10_2009
    Public Function checkPortPins(ByVal rodDiameter As Double, ByVal NutSize As Double, ByVal _
                                    rodMaterial As String) As Boolean

        checkPortPins = False
        For Each item As String In cmbClevisCapPort.Items
            If Trim(SeriesForCosting) <> "" And Trim(cmbBore.Text) <> "" And Trim(cmbClevisCapPinHole.Text) _
                                    <> "" And Trim(item) <> "" Then
                ofrmTieRod2.LVPinSizeDetails.Clear()
                Dim objDT As New DataTable
                Dim strQuery As String = ""
                strQuery = "Select PinHoleSize from ClevisCapDetails where " & GetClevisCapDetails() _
                        & " and BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & " and Port = '" & item _
                                & "' and PinHoleType = '" & Trim(ofrmTieRod1.cmbClevisCapPinHole.Text) & "'"
                objDT = oDataClass.GetDataTable(strQuery)
                For Each oRow As DataRow In objDT.Rows
                    Dim strQuery1 As String = ""
                    'ANUP 12-10-2010 START
                    If NutSize > 0 Then
                        strQuery = "select distinct rdd.RodThreadSize from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.BoreDIameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ") and  bdrd.PartNumberID = rdd.PartNumber and RodDiameter = " & rodDiameter & " and MaterialType = '" & rodMaterial & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and Series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "' and  PistonThreadSize = " & NutSize
                    Else
                        strQuery = "select distinct rdd.RodThreadSize from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.BoreDIameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ") and  bdrd.PartNumberID = rdd.PartNumber and RodDiameter = " & rodDiameter & " and MaterialType = '" & rodMaterial & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and Series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "'"
                    End If
                    'ANUP 12-10-2010 TILL HERE
                    If cmbStyle.Text = "ASAE" Then
                        strQuery = strQuery + " And StrokeLength=" & Val(ofrmTieRod1.cmbStrokeLength.Text)
                    End If
                    strQuery1 = "select ThreadSize from RodClevisDetails where pinHoleSize = " & oRow(0)
                    strQuery = strQuery + " intersect " + strQuery1
                    Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery)
                    If objDT1.Rows.Count > 0 Then
                        checkPortPins = True
                    Else
                        checkPortPins = False
                    End If
                Next
            End If
        Next

    End Function

    Public Sub LoadPinSizeDetails()

        Try
            If Trim(SeriesForCosting) <> "" And Trim(cmbBore.Text) <> "" And Trim(cmbClevisCapPinHole.Text) _
                                <> "" And Trim(cmbClevisCapPort.Text) <> "" Then
                ofrmTieRod2.LVPinSizeDetails.Clear()
                Dim objDT As New DataTable
                Dim strQuery As String = ""
                strQuery = "Select PinHoleSize from ClevisCapDetails where " & GetClevisCapDetails() _
                    & " and BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & " and Port = '" & _
                    Trim(ofrmTieRod1.cmbClevisCapPort.Text) & "' and PinHoleType = '" & _
                    Trim(ofrmTieRod1.cmbClevisCapPinHole.Text) & "'"
                objDT = oDataClass.GetDataTable(strQuery)
                Dim oTable As DataTable = objDT
                oTable.Columns.Add("Safety Factor")
                oTable.Columns.Add("Derate Pressure")
                For Each oRow As DataRow In oTable.Rows
                    If checkPinSizeDetails(oRow(0)) = False Then
                        oRow.Delete()
                    Else
                        Dim dblSafetyFactor As Double
                        Dim dblYieldStrength As Double = 75000
                        Dim dblDeratePressure As Double
                        txtWorkingPressure.Text = IIf(dbltempWorkingPressureNut = 0, dbltempWorkingPressureSeries, _
                                                                    dbltempWorkingPressureNut)
                        'dblCylinderForce = IIf(Val(txtWorkingPressure.Text) = dbltempWorkingPressureSeries, Val(txtWorkingPressure.Text), dbltempWorkingPressureSeries) * (3.1416 / 4) * (Math.Pow(BoreDiameter, 2) - Math.Pow(RodDiameter, 2))
                        'dblCylinderForce = Val(txtWorkingPressure.Text) * (3.1416 / 4) * (Math.Pow(BoreDiameter, 2) - Math.Pow(RodDiameter, 2))
                        dblCylinderForce = Val(txtWorkingPressure.Text) * (3.1416 / 4) * (Math.Pow(BoreDiameter, 2))                 '29_10_2009   Ragava   CylinderPushForce

                        'dblSafetyFactor = dblYieldStrength / dblCylinderForce / 2 * 3.1416 * (Math.Pow(Val(oRow("PinHoleSize")), 2) / 4)   28/10/09
                        dblSafetyFactor = dblYieldStrength / (dblCylinderForce / (2 * 3.1416 * _
                                                    (Math.Pow(Val(oRow("PinHoleSize")), 2) / 4)))
                        If dblSafetyFactor <= 2 Then
                            oRow("Safety Factor") = 2
                            'dblCylinderForce = dblYieldStrength / dblSafetyFactor / (2 * 3.1416 * (Math.Pow(Val(oRow("PinHoleSize")), 2) / 4))
                            dblCylinderForce = dblYieldStrength / (dblSafetyFactor / (2 * 3.1416 * _
                                            (Math.Pow(Val(oRow("PinHoleSize")), 2) / 4)))          '29_10_2009   Ragava
                            'dblDeratePressure = (2 * (Math.Pow(Val(oRow("PinHoleSize")), 2)) * 75000) / (((Math.Pow(BoreDiameter, 2) - Math.Pow(RodDiameter, 2))) * dblSafetyFactor)
                            '14_10_2009
                            dblDeratePressure = (75000 / 2) * 0.5 * (Math.Pow(Val(oRow("PinHoleSize")), 2)) / _
                                                    (Math.Pow(BoreDiameter, 2) - Math.Pow(RodDiameter, 2))
                            If SeriesForCosting.ToString.StartsWith("TX") Then
                                oRow("Derate Pressure") = IIf(dblDeratePressure >= Val(txtWorkingPressure.Text), _
                                                    "N/A", Math.Round(dblDeratePressure, 2))
                            Else
                                oRow("Derate Pressure") = IIf(dblDeratePressure >= dbltempWorkingPressureNut, _
                                                    "N/A", Math.Round(dblDeratePressure, 2))
                            End If
                        Else
                            oRow("Safety Factor") = Math.Round(dblSafetyFactor, 2)
                            oRow("Derate Pressure") = "N/A"
                        End If
                    End If
                Next
                oTable.AcceptChanges()
                ofrmTieRod2.LVPinSizeDetails.FlushListViewData()
                ofrmTieRod2.LVPinSizeDetails.SourceTable = oTable
                Module1.PinSizeDetailsDataTable.Rows.Clear()
                Module1.PinSizeDetailsDataTable = oTable
                ofrmTieRod2.LVPinSizeDetails.Populate()
                '09_10_2009
                'If oTable.Rows.Count = 1 Then
                '    ofrmTieRod2.LVPinSizeDetails.Items(0).Selected = True
                '    ofrmTieRod2.LVPinSizeDetails.Enabled = False
                'Else
                '    ofrmTieRod2.LVPinSizeDetails.Items(0).Selected = True
                '    ofrmTieRod2.LVPinSizeDetails.Enabled = True
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Function checkPinSizeDetails(ByVal pinSize As Double) As Boolean

        checkPinSizeDetails = False
        Dim strQuery As String = ""
        Dim strQuery1 As String = ""

        'If mdiMonarch.IsBtnmygClicked Or ofrmContractDetails.IsBrowseBtnClicked Then
        '    If LVRodDiameterDetails.Items.Count > 0 Then
        '        RodDiameter = Module1.ReadValuesFromExcel.RodDiameter
        '    End If
        'Else
        If LVRodDiameterDetails.SelectedItems.Count > 0 Then
            Dim oSelectedListviewItem As ListViewItem = LVRodDiameterDetails.SelectedItems(0)
            RodDiameter = Val(oSelectedListviewItem.SubItems(0).Text)
        End If

        'End If

        'If mdiMonarch.IsBtnmygClicked Or ofrmContractDetails.IsBrowseBtnClicked Then
        '    If LVNutSizeDetails.SelectedItems.Count > 0 Then
        '        Dim oListViewItemNutSize As ListViewItem = ofrmTieRod1.LVNutSizeDetails.SelectedItems(0)
        '        strQuery = "select distinct rdd.RodThreadSize from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.BoreDIameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ") and  bdrd.PartNumberID = rdd.PartNumber and RodDiameter = " & RodDiameter & " and MaterialType = '" & Trim(ofrmTieRod1.cmbRodMaterial.Text) & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and Series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "' and  PistonThreadSize = " & Val(oListViewItemNutSize.SubItems(0).Text)
        '    End If
        'Else
        If LVNutSizeDetails.SelectedItems.Count > 0 Then
            Dim oListViewItemNutSize As ListViewItem = ofrmTieRod1.LVNutSizeDetails.SelectedItems(0)
            strQuery = "select distinct rdd.RodThreadSize from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.BoreDIameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ") and  bdrd.PartNumberID = rdd.PartNumber and RodDiameter = " & RodDiameter & " and MaterialType = '" & Trim(ofrmTieRod1.cmbRodMaterial.Text) & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and Series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "' and  PistonThreadSize = " & Val(oListViewItemNutSize.SubItems(0).Text)
        Else
            strQuery = "select distinct rdd.RodThreadSize from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.BoreDIameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ") and  bdrd.PartNumberID = rdd.PartNumber and RodDiameter = " & RodDiameter & " and MaterialType = '" & Trim(ofrmTieRod1.cmbRodMaterial.Text) & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and Series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "'"
        End If
        'End If

        'ANUP 12-10-2010 TILL HERE
        If cmbStyle.Text = "ASAE" Then
            strQuery = strQuery + " And StrokeLength=" & Val(ofrmTieRod1.cmbStrokeLength.Text)
        End If
        strQuery1 = "select ThreadSize from RodClevisDetails where pinHoleSize = " & pinSize
        strQuery = strQuery + " intersect " + strQuery1
        Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)

        If objDT.Rows.Count > 0 Then
            checkPinSizeDetails = True
        End If

    End Function

    Public Sub CallRodCapPortLogics()
        Try
            If Trim(cmbRodCapPort.Text) <> "" Then
                If LVRodDiameterDetails.SelectedItems.Count > 0 Then
                    Dim oSelectedListviewItem As ListViewItem = LVRodDiameterDetails.SelectedItems(0)
                    RodDiameter = Val(oSelectedListviewItem.SubItems(0).Text)
                    Dim strQuery1 As String
                    strQuery1 = "Select Hallite,ZMacro,Notes from RodCapDetails where BoreDiameter = " _
                        & Val(cmbBore.Text) & " and RodDiameter =" & RodDiameter & " and " & _
                        GetClevisCapDetails(True) & " and PortDimensions = '" & Trim(cmbRodCapPort.Text) & "'"    'ANUP 21-10-2010 START
                    Dim objDT As DataTable = oDataClass.GetDataTable(strQuery1)
                    If objDT.Rows.Count = 0 Then
                        If blnRevision = False Then
                            MessageBox.Show("Rod Seal Package is not available for the selected rod diameter!" + _
                            vbNewLine & "Please select the other rod diameter or configuration", _
                            "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                        mdiMonarch.btnNext.Visible = False
                        Exit Sub
                    Else
                        mdiMonarch.btnNext.Visible = True
                    End If
                End If
                RodCapPort = Trim(cmbRodCapPort.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Sub callTubeSealLogics()

        If Not cmbBore.Text Is Nothing AndAlso Not cmbSeries.Text Is Nothing Then
            Dim strQuery As String
            Try
                strQuery = ""
                'ANUP 12-10-2010 START
                strQuery = "select trz.TieRodSizes,trz.TieRodNutPartNumber,trz.Description,trz.DrawingNumber from " & _
                "TieRodSizes trz,BoreDiameter_TieRodSizes bdtr where Series ='" _
                & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & _
                "' and bdtr.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & _
                Val(cmbBore.Text) & ") and bdtr.PartNumberID=trz.TieRodPartNumber"
                'ANUP 12-10-2010 TILL HERE

                Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                If objDT.Rows.Count > 0 Then
                    ofrmTieRod2.txtTieRodSize.Text = objDT.Rows(0).Item(0).ToString()
                    ofrmTieRod2.txtTieRodNutSize.Text = objDT.Rows(0).Item(1).ToString()
                    strTieRodNutCodeNumber = ofrmTieRod2.txtTieRodNutSize.Text
                    strTieRodNutDescription = objDT.Rows(0).Item(2).ToString
                    strTieRodNutDrawingNumber = objDT.Rows(0).Item(3).ToString
                    mdiMonarch.mdiComponent.Items(5).SubItems.Add(strTieRodNutCodeNumber)
                    mdiMonarch.mdiComponent.Items(5).SubItems.Add(strTieRodNutDrawingNumber)
                    mdiMonarch.mdiComponent.Items(5).SubItems.Add(strTieRodNutDescription)
                    If Not Trim(SeriesForCosting).ToString.StartsWith("TX") Then
                        ofrmTieRod2.txtTieRodNutQty.Text = 8
                    Else
                        ofrmTieRod2.txtTieRodNutQty.Text = 4
                    End If
                End If
                strQuery = ""
                strQuery = "select Description from PackagingDetails where BoreDia_min <=" _
                & Val(cmbBore.Text) & " and BoreDia_Max >=" & Val(cmbBore.Text)
                objDT.Clear()
                objDT = oDataClass.GetDataTable(strQuery)
                If objDT.Rows.Count > 0 Then
                    ofrmTieRod2.txtPackaging.Text = objDT.Rows(0).Item(0).ToString()
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Try
                strQuery = ""
                '21_11_2012   RAGAVA
                If SeriesForCosting = "LN" Then
                    strQuery = "select DISTINCT DualSeal from cleviscapDetails where BoreDiameter = " & _
                                    Val(cmbBore.Text) & " and " & GetClevisCapDetails()
                Else
                    strQuery = "select DISTINCT ORingSeal,BackUpSeal from cleviscapDetails where BoreDiameter = " & _
                                   Val(cmbBore.Text) & " and " & GetClevisCapDetails()
                    strQuery = strQuery & " order by BackUpSeal desc"
                End If
                'strQuery = "select DISTINCT ORingSeal,BackUpSeal from cleviscapDetails where BoreDiameter = " & _
                'Val(cmbBore.Text) & " and " & GetClevisCapDetails()
                'strQuery = strQuery & " order by BackUpSeal desc"   
                'Till   Here   '21_11_2012   RAGAVA
                Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                If objDT.Rows.Count > 0 Then
                    Dim strRow As String = ""
                    '21_11_2012   RAGAVA
                    If SeriesForCosting = "LN" Then
                        If Trim(objDT.Rows(0).Item(0).ToString) <> "" Then
                            strRow = "DualSeal"
                        End If
                    Else
                        If Trim(objDT.Rows(0).Item(0).ToString) <> "" Then
                            strRow = "ORingSeal"
                        End If
                        If Trim(objDT.Rows(0).Item(1).ToString) <> "" Then
                            strRow = strRow & IIf(strRow <> "", "+", strRow) & "BackUpSeal"
                        End If
                    End If
                    'If Trim(objDT.Rows(0).Item(0).ToString) <> "" Then
                    '    strRow = "ORingSeal"
                    'End If
                    'If Trim(objDT.Rows(0).Item(1).ToString) <> "" Then
                    '    strRow = strRow & IIf(strRow <> "", "+", strRow) & "BackUpSeal"
                    'End If
                    'Till   Here   '21_11_2012   RAGAVA
                    ofrmTieRod2.txtTubeSeal1.Text = strRow
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub LoadTieRodSizes()

        Dim strQuery As String
        Dim oListViewItem As ListViewItem

        If Trim(cmbBore.Text) <> "" AndAlso Trim(cmbSeries.Text) <> "" _
        AndAlso Trim(cmbStyle.Text) <> "" AndAlso Val(LVRodDiameterDetails.SelectedItems.Count) > 0 AndAlso _
                        Trim(cmbRodCapPort.Text) <> "" Then
            Try
                If Trim(ofrmTieRod2.cmbRodSealPackage.Text) = "" Then
                    ofrmTieRod2.cmbRodSealPackage.Items.Clear()
                End If
                ofrmTieRod2.cmbRodSealPackage.Items.Clear()
                strQuery = ""
                Dim arrSeries As String()
                If SeriesForCosting.ToString.StartsWith("TP") Then
                    arrSeries = SeriesForCosting.ToString.Split("-")
                Else
                    arrSeries = SeriesForCosting.ToString.Split(" ")
                End If
                Dim strSeries As String = arrSeries(0)
                strQuery = "Select distinct Hallite,ZMacro,Notes,RU9 from RodCapDetails where BoreDiameter = " & _
                Val(cmbBore.Text) & " and RodDiameter =" & RodDiameter & " and " & _
                GetClevisCapDetails(True) & " and PortDimensions = '" & Trim(cmbRodCapPort.Text) & "'"     ''ANUP 21-10-2010 START
                Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                ofrmTieRod2.cmbRodSealPackage.Items.Add(" ")
                For Each dr As DataRow In objDT.Rows
                    Dim strRow As String = ""
                    If Trim(dr(0).ToString) <> "" Then
                        'strRow = "Hallite " & IIf(Trim(dr(2).ToString) = "", Trim(dr(2).ToString), "+" & Trim(dr(2).ToString))
                        strRow = "Hallite " & IIf(UCase(Trim(dr(2).ToString)) = "NO", "", "+" & Trim(dr(2).ToString))
                    ElseIf Trim(dr(1).ToString) <> "" Then
                        'strRow = "ZMacro" & Trim(dr(2).ToString)
                        strRow = "ZMacro" & IIf(UCase(Trim(dr(2).ToString)) = "NO", "", Trim(dr(2).ToString))         '27_10_2009  ragava
                        'ANUP 03-11-2010 START
                    ElseIf Trim(dr(3).ToString) <> "" Then
                        strRow = "RU9" & IIf(UCase(Trim(dr(2).ToString)) = "NO", "", "+" & Trim(dr(2).ToString))
                        'ANUP 03-11-2010 TILL HERE
                    End If
                    ofrmTieRod2.cmbRodSealPackage.Items.Add(strRow)
                Next
                '09_10_2009
                If ofrmTieRod2.cmbRodSealPackage.Items.Count = 2 Then
                    ofrmTieRod2.cmbRodSealPackage.SelectedIndex = 1
                    ofrmTieRod2.cmbRodSealPackage.Enabled = False
                Else
                    ofrmTieRod2.cmbRodSealPackage.Enabled = True
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

        Try
            If Trim(cmbBore.Text) <> "" AndAlso Trim(cmbSeries.Text) <> "" AndAlso _
            Val(LVRodDiameterDetails.SelectedItems.Count) > 0 Then
                ofrmTieRod2.cmbPistonSealPackage.Items.Clear()
                oListViewItem = ofrmTieRod1.LVRodDiameterDetails.SelectedItems(0)
                strQuery = ""
                Dim arrSeries As String()
                arrSeries = SeriesForCosting.ToString.Split(" ")
                Dim strSeries As String = arrSeries(0)

                'If ofrmTieRod1.LVNutSizeDetails.SelectedItems.Count > 0 Then
                '    oListViewItem = ofrmTieRod1.LVNutSizeDetails.SelectedItems(0)
                '    strQuery = "Select Oring,BackUpRing,PTFESeal,OringExpander,PSPSeal,WearRing1,WearRing2 from PistonSealDetails where BoreDiameter = " & _
                '    Val(cmbBore.SelectedItem) & " and PistonNutSize= " & Val(oListViewItem.SubItems(0).Text) & " and Series like '%" & strSeries & "%'"
                'Else
                '    strQuery = "Select Oring,BackUpRing,PTFESeal,OringExpander,PSPSeal,WearRing1,WearRing2 from PistonSealDetails where BoreDiameter = " & _
                '    Val(cmbBore.SelectedItem) & " and Series like '%" & strSeries & "%'"
                'End If

                'ANUP 02-11-2010 START
                If ofrmTieRod1.LVNutSizeDetails.SelectedItems.Count > 0 Then
                    oListViewItem = ofrmTieRod1.LVNutSizeDetails.SelectedItems(0)
                    strQuery = "Select Oring,BackUpRing,PTFESeal,OringExpander,PSPSeal,WearRing1,WearRing2,WynSeal,GlydP from PistonSealDetails where BoreDiameter = " & _
                    Val(cmbBore.Text) & " and PistonNutSize= " & Val(oListViewItem.SubItems(0).Text) _
                                                    & " and Series like '%" & strSeries & "%'"
                Else
                    strQuery = "Select Oring,BackUpRing,PTFESeal,OringExpander,PSPSeal,WearRing1,WearRing2,WynSeal,GlydP from PistonSealDetails where BoreDiameter = " & _
                    Val(cmbBore.Text) & " and Series like '%" & strSeries & "%'"
                End If
                'ANUP 02-11-2010 TILL HERE


                Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                If objDT.Rows.Count = 0 Then
                    If blnRevision = False Then
                        MessageBox.Show("Piston is not available for the selected configuration" + _
                        vbNewLine + "Change the Bore Diameter", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                    'If LVNutSizeDetails.Items.Count >= 1 Then
                    '    For Each oListViewItem1 As ListViewItem In Me.LVNutSizeDetails.Items
                    '        If oListViewItem1.Selected = True Then
                    '            oListViewItem1.Remove()
                    '        End If
                    '    Next
                    '    LVNutSizeDetails.FlushListViewData()
                    '    LVRodDiameterDetails.Refresh()
                    '    LVNutSizeDetails.Refresh()
                    'Else
                    '    For Each oListViewItem1 As ListViewItem In Me.LVRodDiameterDetails.Items
                    '        If oListViewItem1.Selected = True Then
                    '            oListViewItem1.Remove()
                    '        End If
                    '    Next
                    '    LVRodDiameterDetails.Refresh()
                    '    LVNutSizeDetails.Refresh()
                    '    If LVRodDiameterDetails.Items.Count = 0 Then
                    '        cmbBore.Focus()
                    '        cmbStrokeLength.SelectedIndex = -1
                    '        cmbRodMaterial.Items.Clear()
                    '        LVRodDiameterDetails.Clear()
                    '        LVNutSizeDetails.Clear()
                    '    End If
                    'End If
                    ofrmMdiMonarch.btnNext.Visible = False
                Else
                    If blnRodEndThreadSizeNotAvailable = False Then
                        ofrmMdiMonarch.btnNext.Visible = True
                        ofrmMdiMonarch.btnNext.Enabled = True
                    Else
                        ofrmMdiMonarch.btnNext.Visible = False
                        ofrmMdiMonarch.btnNext.Enabled = False
                    End If
                    ofrmTieRod2.cmbPistonSealPackage.Items.Add(" ")
                    For Each dr As DataRow In objDT.Rows
                        Dim strRow As String = ""

                        'ANUP 02-11-2010 START
                        Dim strRow2 As String = ""
                        If SeriesForCosting = "LN" Then
                            If Trim(dr(7).ToString) <> "" Then
                                strRow = "WynSeal"
                            End If
                            If Trim(dr(8).ToString) <> "" Then
                                strRow2 = "GlydP"
                            End If
                        Else
                            If Trim(dr(0).ToString) <> "" Then
                                strRow = "Oring"
                            End If
                            If Trim(dr(1).ToString) <> "" Then
                                If strRow <> "" Then
                                    strRow = strRow & "+"
                                End If
                                strRow = strRow & "BackUpRing"
                            End If
                            If Trim(dr(2).ToString) <> "" Then
                                If strRow <> "" Then
                                    strRow = strRow & "+"
                                End If
                                strRow = strRow & "PTFESeal"
                            End If
                            If Trim(dr(3).ToString) <> "" Then
                                If strRow <> "" Then
                                    strRow = strRow & "+"
                                End If
                                strRow = strRow & "OringExpander"
                            End If
                            If Trim(dr(4).ToString) <> "" Then
                                If strRow <> "" Then
                                    strRow = strRow & "+"
                                End If
                                strRow = strRow & "PSPSeal"
                            End If
                        End If

                        If Trim(dr(5).ToString) <> "" Then
                            If strRow <> "" Then
                                strRow = strRow & "+"
                            End If
                            If strRow2 <> "" Then
                                strRow2 = strRow2 & "+"
                            End If
                            strRow = strRow & "WearRing1"
                            strRow2 = strRow2 & "WearRing1"
                        End If
                        If Trim(dr(6).ToString) <> "" Then
                            If strRow <> "" Then
                                strRow = strRow & "+"
                            End If
                            If strRow2 <> "" Then
                                strRow2 = strRow2 & "+"
                            End If
                            strRow = strRow & "WearRing2"
                            strRow2 = strRow2 & "WearRing2"
                        End If
                        ofrmTieRod2.cmbPistonSealPackage.Items.Add(strRow)
                        If strRow2 <> "" Then
                            ofrmTieRod2.cmbPistonSealPackage.Items.Add(strRow2)
                        End If
                    Next
                    'ANUP 02-11-2010 TILL HERE

                    '09_10_2009
                    If ofrmTieRod2.cmbPistonSealPackage.Items.Count = 2 Then
                        ofrmTieRod2.cmbPistonSealPackage.SelectedIndex = 1
                        ofrmTieRod2.cmbPistonSealPackage.Enabled = False
                    Else
                        ofrmTieRod2.cmbPistonSealPackage.Enabled = True
                    End If
                End If
            End If
            LoadInformation()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub cmbRephasingPortPosition_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbRephasingPortPosition.SelectedIndexChanged

        Try
            If Trim(cmbRephasingPortPosition.Text) <> "" Then
                ComboBoxRephasingPortPosition()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Sub ComboBoxRephasingPortPosition()

        loadBoreDiameterValues(Trim(SeriesForCosting))

        If mdiMonarch.IsBtnmygClicked Or Module1.BtnBrowseClicked Then
            strRephasing = Module1.ReadValuesFromExcel.RephasingPortPosition
        Else
            strRephasing = Trim(cmbRephasingPortPosition.Text)
        End If

    End Sub

    Private Sub cmbStrokeLengthAdder_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As _
                        System.EventArgs) Handles cmbStrokeLengthAdder.SelectedIndexChanged

        loadStrokeControlDetails(Trim(cmbStrokeLengthAdder.Text))

    End Sub

    Public Sub loadStrokeControlDetails(ByVal stage As String)

        If LVRodDiameterDetails.GetCurrentIndex >= 0 Then
            Dim oSelectedListviewItem As ListViewItem = LVRodDiameterDetails.SelectedItems(0)
            RodDiameter = Val(oSelectedListviewItem.SubItems(0).Text)
            Select Case stage
                Case "2 Stage"
                    If RodDiameter = 1 Then
                        strstrokeControlCodeNumber = "490002"
                        strStrokeControlDrawingNumber = "N/A"
                        strStrokeControlDescription = "N/A"
                    ElseIf RodDiameter = 1.12 Then
                        strstrokeControlCodeNumber = "492637"
                        strStrokeControlDrawingNumber = "N/A"
                        strStrokeControlDescription = "N/A"
                    ElseIf RodDiameter = 1.25 Then
                        strstrokeControlCodeNumber = "292638"
                        strStrokeControlDrawingNumber = "N/A"
                        strStrokeControlDescription = "N/A"
                    ElseIf RodDiameter = 1.38 Then
                        strstrokeControlCodeNumber = "492639"
                        strStrokeControlDrawingNumber = "N/A"
                        strStrokeControlDescription = "N/A"
                    ElseIf RodDiameter = 1.5 Then
                        strstrokeControlCodeNumber = "492640"
                        strStrokeControlDrawingNumber = "N/A"
                        strStrokeControlDescription = "N/A"
                    End If
                Case "3 Stage"
                    If RodDiameter = 1.12 Then
                        strstrokeControlCodeNumber = "492636"
                        strStrokeControlDrawingNumber = "N/A"
                        strStrokeControlDescription = "N/A"
                    End If
            End Select
            ObjClsCostingDetails.AddCodeNumberToDataTable(strstrokeControlCodeNumber, "Stroke Control Code") 'anup 10-03-2011 
            LoadInformation()
        End If

    End Sub

    Private Sub optStrokeControlNo_CheckedChanged(ByVal sender As System.Object, ByVal e As _
                                System.EventArgs) Handles optStrokeControlNo.CheckedChanged

        If sender.checked = True Then
            cmbStrokeLengthAdder.Items.Clear()
            cmbStrokeLengthAdder.Enabled = False
        End If

    End Sub

    Private Sub cmbPortOrientation_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As _
                        System.EventArgs) Handles cmbPortOrientation.SelectedIndexChanged

        ClevisCapPortOrientation = Trim(cmbPortOrientation.Text)

    End Sub

    Private Sub cmbRodCapPort_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As _
                    System.EventArgs) Handles cmbRodCapPort.SelectedIndexChanged

        ComboBoxRodCapPort()

    End Sub

    Public Sub ComboBoxRodCapPort()

        Try
            Call CallRodCapPortLogics()
            LoadTieRodSizes()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub CheckRodEndThreadSizeLogics()
        Try
            Dim strQuery As String
            Dim strQuery1 As String
            ' Dim oListViewItem As ListViewItem

            If Trim(cmbBore.Text) <> "" AndAlso Trim(cmbSeries.Text) <> "" _
            AndAlso Trim(cmbStyle.Text) <> "" AndAlso Trim(RodMaterialForCosting) <> "" Then

                ofrmTieRod2.cmbRodEndThread.Items.Clear()
                ofrmTieRod2.LVRodClevis.Clear()
                'oListViewItem = ofrmTieRod1.LVRodDiameterDetails.SelectedItems(0)
                strQuery = ""
                Dim strMsg As String = ""
                'ANUP 12-10-2010 START
                If Not PistonThreadSize = "" Then
                    'Dim oListViewItemNutSize As ListViewItem = ofrmTieRod1.LVNutSizeDetails.SelectedItems(0)
                    strQuery = "select distinct rdd.RodThreadSize from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.BoreDIameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(cmbBore.Text) & ") and  bdrd.PartNumberID = rdd.PartNumber and RodDiameter = " & Val(RodDiameter) & " and MaterialType = '" & Trim(RodMaterialForCosting) & "' and IsASAE = '" & Trim(cmbStyle.Text) & "' and Series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "' and  PistonThreadSize = " & Val(PistonThreadSize)
                Else
                    strQuery = "select distinct rdd.RodThreadSize from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.BoreDIameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(cmbBore.Text) & ") and  bdrd.PartNumberID = rdd.PartNumber and RodDiameter = " & Val(RodDiameter) & " and MaterialType = '" & Trim(RodMaterialForCosting) & "' and IsASAE = '" & Trim(cmbStyle.Text) & "' and Series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "'"
                End If
                'ANUP 12-10-2010 TILL HERE
                If cmbStyle.Text = "ASAE" Then
                    strQuery = strQuery + " And StrokeLength=" & Val(StrokeLength)
                End If
                strQuery1 = "select ThreadSize from RodClevisDetails where pinHoleSize = " & PinSize
                strQuery = strQuery + " intersect " + strQuery1
                Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)

                If objDT.Rows.Count = 0 Then
                    If Not BtnBrowseClicked() Then
                        If blnRevision = False Then
                            strMsg = "Rod is not available for the selected Configuraion!" + vbNewLine
                            MessageBox.Show(strMsg, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                        ofrmMdiMonarch.btnNext.Visible = False
                        ofrmMdiMonarch.btnNext.Enabled = False
                        Exit Sub
                    End If
                Else
                    ofrmTieRod2.cmbRodEndThread.Items.Clear()
                    ofrmTieRod2.LVRodClevis.Clear()
                    ofrmMdiMonarch.btnNext.Visible = True
                    ofrmTieRod2.cmbRodEndThread.Items.Add(" ")
                    For Each dr As DataRow In objDT.Rows
                        ofrmTieRod2.cmbRodEndThread.Items.Add(dr(0).ToString)
                    Next
                    '09_10_2009
                    If objDT.Rows.Count = 1 Then
                        ofrmTieRod2.cmbRodEndThread.SelectedIndex = 1
                        ofrmTieRod2.cmbRodEndThread.Enabled = False
                    Else
                        ofrmTieRod2.cmbRodEndThread.Enabled = True
                    End If
                End If
                Try
                    strQuery = ""
                    objDT.Clear()
                    'anup 31-01-2011 start
                    If SeriesForCosting = "LN" Then
                        strQuery = "select Wiper_Description from RodWiperDetails where RodDiameter = " & _
                                                        Val(RodDiameter) & " and Series = 'LN'"
                    Else
                        strQuery = "select Wiper_Description from RodWiperDetails where RodDiameter = " & _
                                                                Val(RodDiameter) & " and Series != 'LN'"
                    End If

                    objDT = oDataClass.GetDataTable(strQuery)
                    If objDT.Rows.Count > 0 Then
                        ' ofrmTieRod2.txtRodWiper.Text = objDT.Rows(0).Item(0).ToString()
                        ofrmTieRod2.cmbRodWiper.Items.Clear()
                        ofrmTieRod2.cmbRodWiper.Enabled = True
                        For Each oDataRow As DataRow In objDT.Rows
                            ofrmTieRod2.cmbRodWiper.Items.Add(oDataRow("Wiper_Description"))
                        Next
                        If SeriesForCosting = "LN" Then
                            ofrmTieRod2.cmbRodWiper.SelectedIndex = 1
                        Else
                            ofrmTieRod2.cmbRodWiper.SelectedIndex = 0
                            ofrmTieRod2.cmbRodWiper.Enabled = False
                        End If
                        'anup 31-01-2011 till here
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub rdbStopTubeYes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdbStopTubeYes.Click

        RadioBtnStopTubeYesClick()

    End Sub

    Public Sub RadioBtnStopTubeYesClick()

        If txtStrokeLength.Text = "" AndAlso cmbStyle.Text <> "ASAE" Then
            rdbStopTubeYes.Checked = False
            rdbStopTubeNo.Checked = True
            MessageBox.Show("Please Enter Stroke Length Value", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtStrokeLength.Focus()
            Exit Sub
        End If

    End Sub

    Private Sub txtExtendedLength_TextChanged_1(ByVal sender As System.Object, ByVal e As _
                                    System.EventArgs) Handles txtExtendedLength.TextChanged

        Try
            ExtendedLength = Val(txtExtendedLength.Text)
        Catch ex As Exception
        End Try

    End Sub

    Private Sub LVRodDiameterDetails_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LVRodDiameterDetails.SelectedIndexChanged

        ListViewLVRodDiameterDetails_SelectedIndexChanged()

    End Sub

    Public Sub ListViewLVRodDiameterDetails_SelectedIndexChanged()

        Dim strQuery As String = ""
        Dim oSelectedListviewItem As ListViewItem
        Try
            If Not Module1.BtnBrowseClicked Then
                If LVRodDiameterDetails.SelectedItems.Count > 0 Then
                    oSelectedListviewItem = LVRodDiameterDetails.SelectedItems(0) '.BackColor = Color.BurlyWood
                    RodDiameter = Val(oSelectedListviewItem.SubItems(0).Text)
                Else
                    LVNutSizeDetails.Items.Clear()
                    PistonThreadSize = ""
                    Exit Sub
                End If
            ElseIf RodDiameter = 0 Then
                LVNutSizeDetails.Items.Clear()
                PistonThreadSize = ""
                Exit Sub
            End If


            If strRodMaterial = "Chrome" Then
                strRodMaterial = "BAR-RND-" & Format(RodDiameter, "0.00").ToString() & "CHROME PLATED"
            ElseIf strRodMaterial = "Nitro Steel" Then
                strRodMaterial = "BAR-RND-" & Format(RodDiameter, "0.00").ToString() & "NITRO"
            ElseIf strRodMaterial = "Induction Hardened" Then
                strRodMaterial = "ROD BLANK " & Format(RodDiameter, "0.00").ToString() & "-08-I"
            End If
            dblRodDiameter = RodDiameter.ToString
            strQuery = "select DISTINCT PortDimensions from RodCapDetails where  BoreDiameter=" & _
                BoreDiameter & " and " & GetClevisCapDetails(True) & " and RodDiameter = " & RodDiameter 'ANUP 21-10-2010 START
            Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery)
            cmbRodCapPort.Items.Clear()
            cmbRodCapPort.Items.Add(" ")
            For i As Integer = 0 To objDT1.Rows.Count - 1
                cmbRodCapPort.Items.Add(objDT1.Rows(i).Item("PortDimensions").ToString)
            Next
            '09_10_2009
            If objDT1.Rows.Count = 1 Then
                cmbRodCapPort.SelectedIndex = 1
                cmbRodCapPort.Enabled = False
            Else
                cmbRodCapPort.SelectedIndex = 1
                cmbRodCapPort.Enabled = True
            End If
            Dim strQuery1 As String
            strQuery1 = "Select Hallite,ZMacro,Notes from RodCapDetails where BoreDiameter = " & _
                Val(cmbBore.Text) & " and RodDiameter =" & RodDiameter & " and " & GetClevisCapDetails(True) 'ANUP 21-10-2010 START
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery1)
            If objDT.Rows.Count = 0 Then
                If blnRevision = False Then
                    MessageBox.Show("Rod Seal Package is not available for the selected rod diameter!" _
                        + vbNewLine & "Please select the other rod diameter or configuration", "Information", _
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
                Exit Sub
            End If
            Try
                _objNutSizeDT = Nothing
                If Not SeriesForCosting.StartsWith("TX") Then
                    strQuery = ""
                    strQuery = "SELECT distinct rd.PistonThreadSize  FROM " + vbNewLine
                    strQuery = strQuery & "RodDiameterDetails rd,BoreDiameter_RodDiameter bdrd ,PistonSealDetails p " + vbNewLine
                    strQuery = strQuery & "where bdrd.BoreDiameterID=(select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(cmbBore.Text) & ")" + vbNewLine
                    strQuery = strQuery & " and bdrd.PartNumberID = rd.PartNumber and RodDiameter = " & _
                        RodDiameter & " And IsASAE='" & cmbStyle.Text & "' And rd.Series <> 'TX'  and (rd.PistonThreadSize=p.PistonNutSize)"
                    strQuery = strQuery & "and (not p.Oring='' or not p.BackUpRing=''or   not p.PTFESeal='' or not p.OringExpander='' or not p.PSPSeal='' or not p.WearRing1='' or p.WearRing2='' or not p.WynSeal='' or not p.GlydP='')"  'ANUP 02-11-2010
                    If cmbStyle.Text = "ASAE" Then
                        strQuery = strQuery + " And StrokeLength=" & Val(cmbStrokeLength.Text)
                    End If
                    Dim aColumns As New ArrayList
                    LVNutSizeDetails.Clear()
                    _objNutSizeDT = Nothing
                    _objNutSizeDT = oDataClass.GetDataTable(strQuery)
                    aColumns.Add(New Object(2) {"PistonThreadSize", "Nut Size", True})
                    LVNutSizeDetails.DisplayHeaders = aColumns
                    LVNutSizeDetails.FlushListViewData()
                    LVNutSizeDetails.FullRowSelect = True
                    LVNutSizeDetails.SourceTable = _objNutSizeDT
                    LVNutSizeDetails.Populate()
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            If Module1.BtnBrowseClicked Then
                strColumnLoadDeratePressure = "N/A"
            Else
                If oSelectedListviewItem.SubItems(1).Text <> "N/A" Then
                    If Val(txtWorkingPressure.Text) > Val(oSelectedListviewItem.SubItems(1).Text) Then
                        strColumnLoadDeratePressure = oSelectedListviewItem.SubItems(1).Text
                        strColumnLoadDeratePressure = Format(Convert.ToDouble(strColumnLoadDeratePressure), "0")
                    End If
                Else
                    strColumnLoadDeratePressure = "N/A"
                End If
            End If

            Dim dblEffectlength As Double
            Dim dblCastingSumConstant As Double
            If BoreDiameter <= 4.5 Then
                dblCastingSumConstant = 10.25
            Else
                dblCastingSumConstant = 12.25
            End If
            dblEffectlength = StrokeLength * 2 + dblCastingSumConstant

            '20_04_2010    RAGAVA
            If RodDiameter = 1.12 Then
                txtColumnLoad.Text = Math.Round(Math.Pow(3.1416, 2) * 0.049 * Math.Pow(1.125, 4) * (30 * Math.Pow(10, 6) / Math.Pow(dblEffectlength, 2)), 2)
            ElseIf RodDiameter = 1.38 Then
                txtColumnLoad.Text = Math.Round(Math.Pow(3.1416, 2) * 0.049 * Math.Pow(1.375, 4) * (30 * Math.Pow(10, 6) / Math.Pow(dblEffectlength, 2)), 2)
            Else
                txtColumnLoad.Text = Math.Round(Math.Pow(3.1416, 2) * 0.049 * Math.Pow(RodDiameter, 4) * (30 * Math.Pow(10, 6) / Math.Pow(dblEffectlength, 2)), 2)
            End If
            'txtColumnLoad.Text = Math.Round(Math.Pow(3.1416, 2) * 0.049 * Math.Pow(RodDiameter, 4) * (30 * Math.Pow(10, 6) / Math.Pow(dblEffectlength, 2)), 2)
            '20_04_2010   RAGAVA  TILL  HERE

            ColumnLoad = txtColumnLoad.Text
            If Not _objNutSizeDT Is Nothing Then
                Dim oTable As DataTable
                oTable = _objNutSizeDT
                oTable.Columns.Add("Safety Factor")
                oTable.Columns.Add("Derate Pressure")
                For Each oRow As DataRow In oTable.Rows
                    If checkPistonSealPackage(oRow(0), RodDiameter, cmbRodMaterial.Text) = False Then
                        oRow.Delete()
                    Else
                        Dim dblCylinderForce As Double
                        Dim dblSafetyFactor As Double
                        Dim dblYieldStrength As Double
                        Dim dblDeratePressure As Double
                        Dim dblThreadtensileArea As Double
                        If cmbRodMaterial.Text = "Chrome" Then
                            dblYieldStrength = 83500
                        ElseIf cmbRodMaterial.Text = "Nitro Steel" Or cmbRodMaterial.Text = "Induction Hardened" Then
                            dblYieldStrength = 75000
                        End If
                        For Each oItem As Object In TensileArea
                            If Val(oRow("PistonThreadSize")) = Val(oItem(0)) Then
                                dblThreadtensileArea = Val(oItem(1))
                                Exit For
                            End If
                        Next

                        '03_11_2009   Ragava
                        Me.txtWorkingPressure.Text = dbltempWorkingPressureSeries
                        ' Me.cmbClevisCapPort.SelectedIndex = -1
                        dbltempWorkingPressureNut = 0
                        '03_11_2009   Ragava  Till  Here

                        '20_04_2010    RAGAVA
                        If RodDiameter = 1.12 Then
                            dblCylinderForce = Val(txtWorkingPressure.Text) * (3.1416 / 4) * _
                                                (Math.Pow(BoreDiameter, 2) - Math.Pow(1.125, 2))
                        ElseIf RodDiameter = 1.38 Then
                            dblCylinderForce = Val(txtWorkingPressure.Text) * (3.1416 / 4) * _
                                                    (Math.Pow(BoreDiameter, 2) - Math.Pow(1.375, 2))
                        Else
                            dblCylinderForce = Val(txtWorkingPressure.Text) * (3.1416 / 4) * _
                                                (Math.Pow(BoreDiameter, 2) - Math.Pow(RodDiameter, 2))
                        End If

                        dblSafetyFactor = dblYieldStrength / (dblCylinderForce / dblThreadtensileArea)
                        If dblSafetyFactor <= 2 Then
                            oRow("Safety Factor") = 2
                            dblCylinderForce = dblYieldStrength / (dblSafetyFactor / dblThreadtensileArea)
                            'dblDeratePressure = dblCylinderForce / (3.1416 / 4) * (Math.Pow(BoreDiameter, 2) - Math.Pow(RodDiameter, 2))
                            '14_10_2009
                            '20_04_2010    RAGAVA
                            If RodDiameter = 1.12 Then
                                dblDeratePressure = (dblYieldStrength / 2) * dblThreadtensileArea / _
                                        ((3.1416 / 4) * (Math.Pow(BoreDiameter, 2) - Math.Pow(1.125, 2)))
                            ElseIf RodDiameter = 1.38 Then
                                dblDeratePressure = (dblYieldStrength / 2) * dblThreadtensileArea / _
                                        ((3.1416 / 4) * (Math.Pow(BoreDiameter, 2) - Math.Pow(1.375, 2)))
                            Else
                                dblDeratePressure = (dblYieldStrength / 2) * dblThreadtensileArea / _
                                        ((3.1416 / 4) * (Math.Pow(BoreDiameter, 2) - Math.Pow(RodDiameter, 2)))
                            End If
                           
                            If Math.Round(dblDeratePressure, 2) > Val(txtWorkingPressure.Text) Then
                                oRow("Derate Pressure") = "N/A"
                            Else
                                oRow("Derate Pressure") = Math.Round(dblDeratePressure, 2)
                            End If
                        Else
                            oRow("Safety Factor") = Math.Round(dblSafetyFactor, 2)
                            oRow("Derate Pressure") = "N/A"
                        End If
                    End If
                Next
                oTable.AcceptChanges()
                LVNutSizeDetails.FlushListViewData()
                LVNutSizeDetails.SourceTable = oTable
                LVNutSizeDetails.Populate()
                Try
                    '09_10_2009
                    If oTable.Rows.Count = 1 Then
                        LVNutSizeDetails.Items(0).Selected = True
                        LVNutSizeDetails.Enabled = False
                    Else
                        LVNutSizeDetails.Enabled = True
                        LVNutSizeDetails.Items(0).Selected = True
                    End If
                Catch ex As Exception
                End Try
            End If
            
            Try
                LoadTieRodSizes()
                LoadPinSizeDetails()
            Catch ex As Exception
            End Try
        Catch ex As Exception
        End Try
        LoadInformation()
        For Each listviewItem As ListViewItem In LVRodDiameterDetails.Items
            Dim index As Integer = LVRodDiameterDetails.Items.IndexOf(listviewItem)
            LVRodDiameterDetails.Items(index).BackColor = Color.Ivory
            LVRodDiameterDetails.Items(index).ForeColor = Color.Black
        Next
        If Not Module1.BtnBrowseClicked Then
            For Each listviewItem As ListViewItem In LVRodDiameterDetails.SelectedItems
                Dim index As Integer = LVRodDiameterDetails.Items.IndexOf(listviewItem)
                LVRodDiameterDetails.Items(index).BackColor = Color.CornflowerBlue
                LVRodDiameterDetails.Items(index).ForeColor = Color.White
            Next
        End If

    End Sub

    Private Sub ChkRetractedLength_CheckedChanged(ByVal sender As System.Object, ByVal e As _
                System.EventArgs) Handles ChkRetractedLength.CheckedChanged
        '23_11_2009    Ragava
        If sender.Checked = True Then
            ChkPins.Checked = False
        End If
        '23_11_2009    Ragava  Till  Here
    End Sub

    Private Sub ChkPins_CheckedChanged(ByVal sender As System.Object, ByVal e As _
                        System.EventArgs) Handles ChkPins.CheckedChanged
        '23_11_2009    Ragava
        If sender.Checked = True Then
            ChkRetractedLength.Checked = False
        End If
        '23_11_2009    Ragava  Till  Here
    End Sub

    Private Sub txtStandardRunQty_TextChanged(ByVal sender As System.Object, ByVal e As _
                    System.EventArgs) Handles txtStandardRunQty.TextChanged

        If txtStandardRunQty.Text <> "" Then
            SRQ = txtStandardRunQty.Text
        End If

    End Sub

    Private Sub ColorTheForm()

        FunctionalClassObject.LabelGradient_GreenBorder_ColoringTheScreens(LabelGradient11, _
                                        LabelGradient8, LabelGradient10, LabelGradient12)
        FunctionalClassObject.LabelGradient_OrangeBorder_ColoringTheScreens(LabelGradient5)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient4)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient2)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient1)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient3)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient7)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient6)

    End Sub

    'ANUP 27-10-2010 START
    Private Sub chkReleaseCylinder_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                    As System.EventArgs) Handles chkReleaseCylinder.CheckedChanged

        Try
            If chkReleaseCylinder.Checked Then
                IsReleaseCylinderChecked = True
            Else
                IsReleaseCylinderChecked = False
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub CylinderReleasedFunctionality()

        Try
            If IsNew_Revision_Released = "Revision" Then
                If Not ReleasedRevisionFunctionality() Is Nothing Then
                    chkReleaseCylinder.Checked = True
                    chkReleaseCylinder.Enabled = False
                Else
                    chkReleaseCylinder.Checked = False
                    chkReleaseCylinder.Enabled = False
                End If
            ElseIf IsNew_Revision_Released = "Released" Then
                'ANUP 26-11-2010 START
                chkReleaseCylinder.Checked = True
                chkReleaseCylinder.Enabled = True
                chkReleaseCylinder.Visible = False
                GroupBox3.Visible = False
                'ANUP 26-11-2010 TILL HERE
            ElseIf IsNew_Revision_Released = "New" Then
                chkReleaseCylinder.Checked = False
                chkReleaseCylinder.Enabled = False
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Function ReleasedRevisionFunctionality() As String 'anup 23-12-2010 change private to public

        ReleasedRevisionFunctionality = Nothing
        Try
            Dim strQuery As String = String.Empty
            strQuery = "select ReleasedCylinderCodeNumber from dbo.ReleasedCylinderCodes where ReleasedCylinderCodeNumber =" & CylinderCodeNumber
            ReleasedRevisionFunctionality = IFLConnectionObject.GetValue(strQuery)
        Catch ex As Exception
            ReleasedRevisionFunctionality = Nothing
        End Try

    End Function

    Public Sub ClearAllFielsTieRod1()

        'cmbSeries.Items.Clear()
        'cmbStyle.Items.Clear()
        cmbBore.Items.Clear()
        txtStrokeLength.Text = ""
        cmbStrokeLength.Items.Clear()
        txtRodAdder.Text = ""
        cmbRodMaterial.Items.Clear()
        cmbClevisCapPinHole.Items.Clear()
        cmbRodClevisPinHole.Items.Clear()
        txtExtendedLength.Text = ""
        txtStandardRunQty.Text = ""
        LVRodDiameterDetails.Items.Clear()
        cmbStrokeLengthAdder.Items.Clear()
        cmbRodCapPort.Items.Clear()
        cmbClevisCapPort.Items.Clear()
        'cmbPortOrientation.Items.Clear()
        'cmbPortOrientationForRodCap.Items.Clear()
        LVNutSizeDetails.Items.Clear()
        PistonThreadSize = ""

    End Sub

    Public Sub LoadingDataFromExcelTieRod1(Optional ByVal _rowno As Integer = 0)    'SUGANDHI

        cmbSeries.Text = Module1.ReadValuesFromExcel.Series
        ComboBoxSeries()

        If cmbSeries.Text = "TP-High" Or cmbSeries.Text = "TP-Low" Then
            cmbRephasingPortPosition.Text = Module1.ReadValuesFromExcel.RephasingPortPosition
            ComboBoxRephasingPortPosition()
        End If

        cmbStyle.Text = Module1.ReadValuesFromExcel.Style
        If Not cmbStyle.Text = "" Then
            ComboBoxStyle()
            RetractedLengthCalculation()
            ComboBoxcmbStyleLeave(cmbStyle)
        End If

        Dim boolBore As Boolean = False
        For i As Integer = 0 To cmbBore.Items.Count - 1
            If cmbBore.Items(i) = Module1.ReadValuesFromExcel.Bore.ToString() Then
                cmbBore.Text = Module1.ReadValuesFromExcel.Bore
                RetractedLengthCalculation()
                ComboBoxBore(cmbBore)
                ComboBoxcmbStyleLeave(cmbBore)
                boolBore = True
                Exit For
            End If
        Next
        If Not boolBore Then
            IsErrorMessageTierod1 = True
            cmbBore.Text = ""
            Dim str As String
            For i As Integer = 0 To cmbBore.Items.Count - 1
                If i = cmbBore.Items.Count - 1 Then
                    str = str + cmbBore.Items(i)
                Else
                    str = str + cmbBore.Items(i) + ", "
                End If
            Next
            Module1.LogInfo.Add("Row Number :" + _rowno.ToString() + " Bore : Select the Bore value from the following " + "' " + str + " '")
        End If

        If cmbStyle.Text = "NON ASAE" Then

            Dim strStrokeLength As Double = Module1.ReadValuesFromExcel.StrokeLength.ToString()
            If strStrokeLength >= txtStrokeLength.MinimumValue Then
                If strStrokeLength <= txtStrokeLength.MaximumValue Then
                    txtStrokeLength.Text = Module1.ReadValuesFromExcel.StrokeLength.ToString()
                    If txtStrokeLength.Text <> "" Then
                        StrokeLength = Val(txtStrokeLength.Text)
                        dblStrokeLength = StrokeLength
                    End If
                    ComboBoxBore(txtStrokeLength)
                    RetractedLengthCalculation()
                    ComboBoxcmbStyleLeave(cmbStrokeLength)
                    If dblStrokeLengthModified <> Val(txtStrokeLength.Text) Then
                        StopTubeLength = 0
                        cmbStrokeLength.IFLDataTag = ""
                        dblStrokeLengthModified = Val(txtStrokeLength.Text)         '12_05_2010   RAGAVA
                    End If
                End If
            Else
                IsErrorMessageTierod1 = True
                Module1.LogInfo.Add("Row Number :" + _rowno.ToString() + " StrokeLength : Please enter StrokeLength value between " + txtStrokeLength.MinimumValue.ToString() + " to " + txtStrokeLength.MaximumValue.ToString())
            End If
        Else
            Dim boolStrokeLength As Boolean = False
            For i As Integer = 0 To cmbStrokeLength.Items.Count - 1
                If cmbStrokeLength.Items(i) = Module1.ReadValuesFromExcel.StrokeLength.ToString() Then
                    cmbStrokeLength.Text = Module1.ReadValuesFromExcel.StrokeLength
                    ComboBoxStrokeLength()
                    boolStrokeLength = True
                    Exit For
                End If
            Next
            If Not boolStrokeLength Then
                cmbStrokeLength.Text = "8.00"
            End If
        End If

        If txtRodAdder.Enabled = True Then
            Dim strRodAdder As String = Module1.ReadValuesFromExcel.RodAdder.ToString()
            If strRodAdder >= txtRodAdder.MinimumValue.ToString() Then
                If strRodAdder <= txtRodAdder.MaximumValue.ToString() Then
                    txtRodAdder.Text = Module1.ReadValuesFromExcel.RodAdder.ToString()
                    If txtRodAdder.Text <> "" Then
                        RodAdder = Val(txtRodAdder.Text)
                    End If
                    ComboBoxcmbStyleLeave(txtRodAdder)
                    If dblStrokeLengthModified <> Val(txtStrokeLength.Text) Then
                        StopTubeLength = 0
                        txtRodAdder.IFLDataTag = ""
                        dblStrokeLengthModified = Val(txtStrokeLength.Text)         '12_05_2010   RAGAVA
                    End If
                    RetractedLengthCalculation()
                End If
            Else
                IsErrorMessageTierod1 = True
                Module1.LogInfo.Add("Row Number :" + _rowno.ToString() + " RodAdder : Please enter RodAdder value between " + txtRodAdder.MinimumValue.ToString() + " to " + txtRodAdder.MaximumValue.ToString())
            End If
            ' txtRodAdder.Text = Module1.ReadValuesFromExcel.RodAdder.ToString()
        End If

        If rdbStopTubeYes.Enabled Then
            If Module1.ReadValuesFromExcel.StopTube Then
                rdbStopTubeYes.Checked = True
                txtStopTubeLength.Text = Convert.ToDouble(Module1.ReadValuesFromExcel.StopTubeLength)
                txtStopTubeLengthLeave()
                RadioBtnStopTubeYesClick()
            End If
        Else
            rdbStopTubeNo.Checked = True
            RadioBtnStopTubeNoCheckedChanged()
        End If

        If cmbStyle.Text = "ASAE" Then
            cmbClevisCapPinHole.Text = Module1.ReadValuesFromExcel.ClevisCapPinHole
            ComboBoxPinHole()
        Else
            cmbClevisCapPinHole.Text = "Standard"
            ComboBoxPinHole()
        End If

        cmbRodClevisPinHole.Text = Module1.ReadValuesFromExcel.RodClevisPinHole
        ComboBoxPinHole()

        Dim strStandardRunQty As String = Module1.ReadValuesFromExcel.StandardRunQty.ToString()
        If strStandardRunQty >= txtStandardRunQty.MinimumValue.ToString() Then
            If strStandardRunQty <= txtStandardRunQty.MaximumValue.ToString() Then
                txtStandardRunQty.Text = Module1.ReadValuesFromExcel.StandardRunQty.ToString()
                If txtStandardRunQty.Text <> "" Then
                    SRQ = txtStandardRunQty.Text
                End If
            End If
        Else
            IsErrorMessageTierod1 = True
            Module1.LogInfo.Add("Row Number : " + _rowno.ToString() + ", " + " StandardRunQty : Please enter StandardRunQty value between " + txtStandardRunQty.MinimumValue.ToString() + " to " + txtStandardRunQty.MaximumValue.ToString())
        End If

        If Not txtRetractedLength.Text = Module1.ReadValuesFromExcel.RetractedLength.ToString() Then
            IsErrorMessageTierod1 = True
            Module1.LogInfo.Add("Row Number : " + _rowno.ToString() + ", " + " Retracted Length : Entered Retracted Length value isn't matching with the System generated value. Retracted Length is " + txtRetractedLength.Text)
        End If

        If Not txtExtendedLength.Text = Module1.ReadValuesFromExcel.ExtendedLength.ToString() Then
            IsErrorMessageTierod1 = True
            Module1.LogInfo.Add("Row Number : " + _rowno.ToString() + ", " + "Extended Length : Entered Extended Length value isn't matching with the System generated value. Extended Length is " + txtExtendedLength.Text)
        End If
        Try
            ExtendedLength = Val(txtExtendedLength.Text)
        Catch ex As Exception
        End Try

        cmbRodMaterial.Text = Module1.ReadValuesFromExcel.RodMaterials
        If cmbRodMaterial.Text <> "" Then
            ComboBoxRodMaterial()
        End If
        Dim boolRodDiameter As Boolean = False
        If Not cmbBore.Text = "" Then

            For i As Integer = 0 To LVRodDiameterDetails.Items.Count - 1
                If LVRodDiameterDetails.Items(i).Text = Module1.ReadValuesFromExcel.RodDiameter.ToString() Then
                    LVRodDiameterDetails.Items(i).Selected = True
                    RodDiameter = Module1.ReadValuesFromExcel.RodDiameter
                    ListViewLVRodDiameterDetails_SelectedIndexChanged()
                    boolRodDiameter = True
                    Exit For
                Else
                    Dim s As String = Convert.ToDouble(Module1.ReadValuesFromExcel.RodDiameter.ToString())
                    Dim words As String() = s.Split(New Char() {"."c})

                    Dim word As String = words(1)
                    If Not word.Length = 2 Then
                        word = word + "0"
                        word = words(0) + "." + word
                    End If
                    If LVRodDiameterDetails.Items(i).Text = word Then
                        LVRodDiameterDetails.Items(i).Selected = True
                        RodDiameter = Convert.ToDouble(word)
                        ListViewLVRodDiameterDetails_SelectedIndexChanged()
                        boolRodDiameter = True
                        Exit For
                    End If
                End If
            Next
            If Not boolRodDiameter Then
                IsErrorMessageTierod1 = True
                RodDiameter = 0
                Dim str As String
                For i As Integer = 0 To LVRodDiameterDetails.Items().Count - 1
                    If i = LVRodDiameterDetails.Items().Count - 1 Then
                        str = str + LVRodDiameterDetails.Items(i).Text
                    Else
                        str = str + LVRodDiameterDetails.Items(i).Text + ", "
                    End If
                Next
                RodDiameter = 0
                Module1.LogInfo.Add("Row Number :" + _rowno.ToString() + " RodDiameter : Select the RodDiameter value from the following " + "' " + str + " '")
            End If
        End If
        'LVRodDiameterDetails.Items.Clear()
        'Dim listView As ListViewItem
        'listView = LVRodDiameterDetails.Items.Add(oReadValuesFromExcel.RodDiameter.ToString())
        'listView.SubItems.Add(oReadValuesFromExcel.RodDeratedPressureAtmaximumExtension.ToString())
        'LVRodDiameterDetails.Items(0).Selected = True

        cmbPortOrientation.Text = Module1.ReadValuesFromExcel.PortOrientationForClevisCap
        ClevisCapPortOrientation = Trim(cmbPortOrientation.Text)

        cmbPortOrientationForRodCap.Text = Module1.ReadValuesFromExcel.PortOrientationForRodCap

        If Not RodDiameter = 0 Then

            If cmbClevisCapPinHole.Text = "Standard" Then
                Dim boolClevisCapPort As Boolean = False
                For i As Integer = 0 To cmbClevisCapPort.Items.Count - 1
                    If cmbClevisCapPort.Items(i) = Module1.ReadValuesFromExcel.ClevisCapPort.ToString() Then
                        cmbClevisCapPort.Text = Module1.ReadValuesFromExcel.ClevisCapPort
                        ComboBoxClevisCapPort()
                        boolClevisCapPort = True
                        Exit For
                    End If
                Next
                If Not boolClevisCapPort Then
                    IsErrorMessageTierod1 = True
                    cmbClevisCapPort.Text = ""
                    Dim str As String
                    For i As Integer = 0 To cmbClevisCapPort.Items.Count - 1
                        If i = cmbClevisCapPort.Items.Count - 1 Then
                            str = str + cmbClevisCapPort.Items(i)
                        Else
                            str = str + cmbClevisCapPort.Items(i) + ", "
                        End If
                    Next
                    Module1.LogInfo.Add("Row Number :" + _rowno.ToString() + " ClevisCapPort : Select the ClevisCapPort value from the following " + "' " + str + " '")
                    'MessageBox.Show("Please enter correct value")
                End If
            End If

            Dim boolRodCapPort As Boolean = False
            If boolRodDiameter Then
                For i As Integer = 0 To cmbRodCapPort.Items.Count - 1
                    If cmbRodCapPort.Items(i) = Module1.ReadValuesFromExcel.RodCapPort.ToString() Then
                        cmbRodCapPort.Text = Module1.ReadValuesFromExcel.RodCapPort
                        ComboBoxRodCapPort()
                        boolRodCapPort = True
                        Exit For
                    End If
                Next

                If Not boolRodCapPort Then
                    IsErrorMessageTierod1 = True
                    cmbRodCapPort.Text = ""
                    Dim str As String = ""
                    For i As Integer = 0 To cmbRodCapPort.Items.Count - 1
                        If i = cmbRodCapPort.Items.Count - 1 Then
                            str = str + cmbRodCapPort.Items(i)
                        Else
                            str = str + cmbRodCapPort.Items(i) + ", "
                        End If
                    Next
                    Module1.LogInfo.Add("Row Number :" + _rowno.ToString() + " RodCapPort : Select the RodCapPort value from the following " + "' " + str + " '")
                    ' MessageBox.Show("Please enter correct value")
                End If
            End If
        End If

        If optStrokeControlYes.Enabled = True And optStrokeControlNo.Enabled = True Then
            If Module1.ReadValuesFromExcel.StrokeControl Then
                optStrokeControlYes.Checked = True
                RadioBtnStrokeControlYes()
                optStrokeControlNo.Checked = False
            Else
                optStrokeControlYes.Checked = False
                optStrokeControlNo.Checked = True
                If optStrokeControlNo.Checked = True Then
                    cmbStrokeLengthAdder.Items.Clear()
                    cmbStrokeLengthAdder.Enabled = False
                End If
            End If
        End If

        If LVNutSizeDetails.Enabled Then
            Dim boolNutSize As Boolean = False
            For i As Integer = 0 To LVNutSizeDetails.Items.Count - 1
                If LVNutSizeDetails.Items(i).Text = Module1.ReadValuesFromExcel.NutSize.ToString() Then
                    LVNutSizeDetails.Items(i).Selected = True
                    ListViewNutSizeDetails()
                    boolNutSize = True
                    Exit For
                Else
                    Dim s As String = Convert.ToDouble(Module1.ReadValuesFromExcel.NutSize.ToString())
                    Dim words As String() = s.Split(New Char() {"."c})

                    If words.Length = 2 Then
                        Dim word As String = words(1)
                        If Not word.Length = 2 Then
                            word = word + "0"
                            word = words(0) + "." + word
                        End If
                        If LVNutSizeDetails.Items(i).Text = word Then
                            LVNutSizeDetails.Items(i).Selected = True
                            ListViewNutSizeDetails()
                            boolNutSize = True
                            Exit For
                        End If
                    ElseIf words.Length = 1 Then

                        Dim word As String = "00"
                        word = words(0) + "." + word

                        If LVNutSizeDetails.Items(i).Text = word Then
                            LVNutSizeDetails.Items(i).Selected = True
                            ListViewNutSizeDetails()
                            boolNutSize = True
                            Exit For
                        End If
                    End If

                End If
            Next
            If Not boolNutSize Then
                IsErrorMessageTierod1 = True

                Dim str As String = ""
                For i As Integer = 0 To LVNutSizeDetails.Items().Count - 1
                    If i = LVNutSizeDetails.Items().Count - 1 Then
                        str = str + LVNutSizeDetails.Items(i).Text
                    Else
                        str = str + LVNutSizeDetails.Items(i).Text + ", "
                    End If
                Next
                RodDiameter = 0
                Module1.LogInfo.Add("Row Number :" + _rowno.ToString() + " Nut Size : Select the NutSize value from the following " + "' " + str + " '")
            End If
        Else
            ListViewNutSizeDetails()
        End If


        If cmbStrokeLengthAdder.Enabled = True Then
            cmbStrokeLengthAdder.Text = (Module1.ReadValuesFromExcel.StrokeControlStages).ToString() + " Stage"
            loadStrokeControlDetails(Trim(cmbStrokeLengthAdder.Text))
        End If

        If IsErrorMessageTierod1 Then

            ' ofrmContractDetails.ErrorMessages()

            If ModuleGeneratedModelNames.ArrayListModelName.Count = 0 Then
                Exit Sub
            Else
                Dim strMsg = "Models generated successfully" + vbNewLine
                For i As Integer = 0 To ModuleGeneratedModelNames.ArrayListModelName.Count - 1
                    If i = ModuleGeneratedModelNames.ArrayListModelName.Count - 1 Then
                        strMsg = strMsg & ModuleGeneratedModelNames.ArrayListModelName.Item(i).ToString()
                    Else
                        strMsg = strMsg & ModuleGeneratedModelNames.ArrayListModelName.Item(i).ToString() + " , "
                    End If
                Next

                MessageBox.Show(strMsg, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information, _
                                    MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Exit Sub
            End If

        End If

    End Sub

    Private Sub cmbStrokeLength_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                    Handles cmbStrokeLength.Leave

        ComboBoxcmbStyleLeave(sender)
        If dblStrokeLengthModified <> Val(txtStrokeLength.Text) AndAlso (sender.Name = "txtStrokeLength" Or sender.Name = "txtRodAdder") Then
            StopTubeLength = 0
            sender.IFLDataTag = ""
            dblStrokeLengthModified = Val(txtStrokeLength.Text)         '12_05_2010   RAGAVA
        End If

    End Sub

    Private Sub txtRodAdder_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                        Handles txtRodAdder.Leave

        ComboBoxcmbStyleLeave(sender)
        If dblStrokeLengthModified <> Val(txtStrokeLength.Text) AndAlso (sender.Name = "txtStrokeLength" Or sender.Name = "txtRodAdder") Then
            StopTubeLength = 0
            sender.IFLDataTag = ""
            dblStrokeLengthModified = Val(txtStrokeLength.Text)         '12_05_2010   RAGAVA
        End If
        RetractedLengthCalculation()

    End Sub

    Private Sub txtStrokeLength_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                Handles txtStrokeLength.Leave

        ComboBoxBore(sender)
        RetractedLengthCalculation()

    End Sub

    Private Sub cmbBore_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbBore.Leave

        ComboBoxcmbStyleLeave(sender)

    End Sub

    Private Sub cmbStrokeLength_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As _
                                        System.EventArgs) Handles cmbStrokeLength.SelectedIndexChanged
        ComboBoxBore(sender)

    End Sub

    Private Sub cmbClevisCapPinHole_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As _
                                    System.EventArgs) Handles cmbClevisCapPinHole.SelectedIndexChanged
        ComboBoxPinHole()

    End Sub

    Private Sub cmbRodClevisPinHole_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As _
                                System.EventArgs) Handles cmbRodClevisPinHole.SelectedIndexChanged
        ComboBoxPinHole()

    End Sub

    Private Sub cmbClevisCapPort_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As _
                                System.EventArgs) Handles cmbClevisCapPort.SelectedIndexChanged

        ComboBoxClevisCapPort()

    End Sub

    Public Sub ComboBoxClevisCapPort()

        If Trim(cmbClevisCapPort.Text) <> "" Then
            ClevisCapPort = Trim(cmbClevisCapPort.Text)
        End If
        Try
            LoadPinSizeDetails()
        Catch ex As Exception
        End Try

    End Sub
End Class