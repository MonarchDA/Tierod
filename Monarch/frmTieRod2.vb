Imports MonarchFunctionalLayer
Imports MonarchAPILayer
Public Class frmTieRod2
    Dim dblCylinderForce_PullForce As Double

    Private oBool As Boolean = False

    Public Property IsErrorMessageTierod2() As Boolean

        Get
            Return oBool
        End Get
        Set(ByVal value As Boolean)
            oBool = value
        End Set

    End Property

    Private Sub cmbPinMaterial_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                Handles cmbPinMaterial.SelectedIndexChanged

        ComboBoxPinMaterial()

    End Sub

    Private Sub ComboBoxPinMaterial()

        If Trim(cmbPinMaterial.Text) <> "" Then
            Try
                If optPinsYes.Checked = True Then    '06_04_2010    RAGAVA
                    Dim oListViewItem As ListViewItem
                    If Module1.BtnBrowseClicked Then
                        If Not LVPinSizeDetails.Items.Count = 0 Then
                            oListViewItem = LVPinSizeDetails.Items(0)
                        Else
                            Exit Sub
                        End If
                    Else
                        If LVPinSizeDetails.SelectedItems.Count > 0 Then
                            oListViewItem = LVPinSizeDetails.SelectedItems(0)
                        Else
                            Exit Sub
                        End If
                    End If

                    'oListViewItem = LVPinSizeDetails.Items(LVPinSizeDetails.GetCurrentIndex)
                    Dim strQuery As String
                    Dim objDT As DataTable
                    cmbClips.Items.Clear()
                    cmbClips.Items.Add(" ")
                    strQuery = ""
                    strQuery = "select distinct PinType from ClevisPinDetails where PinHoleSize = " & _
                            Val(oListViewItem.SubItems(0).Text) & " and PinMaterial = '" & _
                            Trim(cmbPinMaterial.Text) & "' And  " & Val(ofrmTieRod1.cmbBore.Text) & _
                            " >= BoreDiameterMinimum and " & Val(ofrmTieRod1.cmbBore.Text) & _
                            " < = BoreDiameterMaximum"
                    objDT = oDataClass.GetDataTable(strQuery)
                    For Each dr As DataRow In objDT.Rows
                        cmbClips.Items.Add(dr(0).ToString)
                    Next
                    '09_10_2009
                    If objDT.Rows.Count = 1 Then
                        cmbClips.SelectedIndex = 1
                        cmbClips.Enabled = False
                    Else
                        cmbClips.Enabled = True
                        cmbClips.Text = "Cotter Pins"                   '11_11_2009   Ragava
                    End If
                End If
                ' End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

    End Sub

    Private Sub cmbRodSealPackage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                    Handles cmbRodSealPackage.SelectedIndexChanged

        ComboBoxRodSealPackage()

    End Sub

    Private Sub ComboBoxRodSealPackage()

        If Trim(cmbRodSealPackage.Text) <> "" Then
            Dim strQuery As String
            ' Dim oListViewItem As ListViewItem        'sugandhi
            If Not RodDiameter = 0 Then
                ' oListViewItem = ofrmTieRod1.LVRodDiameterDetails.SelectedItems(0)        'sugandhi
                ' oListViewItem = ofrmTieRod1.LVRodDiameterDetails.Items(ofrmTieRod1.LVRodDiameterDetails.GetCurrentIndex)
                strQuery = ""
                Dim arrSeries As String()
                If SeriesForCosting.ToString.StartsWith("TP") Then
                    arrSeries = SeriesForCosting.ToString.Split("-")
                Else
                    arrSeries = SeriesForCosting.ToString.Split(" ")
                End If
                Dim strSeries As String = arrSeries(0)
                '27_10_2009  Ragava
                'strQuery = "Select PartNumber,RodCapDescription from RodCapDetails where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & " and RodDiameter =" & Val(oListViewItem.SubItems(0).Text) & " and " & GetClevisCapDetails() & " and PortDimensions = '" & Trim(ofrmTieRod1.cmbRodCapPort.Text) & "'"
                Dim strRod() As String = Trim(cmbRodSealPackage.Text).ToString.Split("+")
                Dim strNotes As String = Trim(strRod(UBound(strRod)))
                If strNotes = "Garlock" Or strNotes = "Glass Filled Nylon" Then
                    strQuery = "Select PartNumber,RodCapDescription from RodCapDetails where BoreDiameter = " & _
                        Val(BoreDiameter) & " and RodDiameter =" & RodDiameter & " and " & GetClevisCapDetails(True) _
                        & " and PortDimensions = '" & Trim(RodCapPort) & "' and Notes = '" & strNotes & "'"
                    For Each strCol As String In strRod
                        If strCol <> Trim(strRod(UBound(strRod))) Then
                            strQuery = strQuery & " and " & Trim(strCol) & " <> ''"
                        End If
                    Next
                Else
                    strQuery = "Select PartNumber,RodCapDescription from RodCapDetails where BoreDiameter = " _
                        & Val(BoreDiameter) & " and RodDiameter =" & RodDiameter & " and " & GetClevisCapDetails(True) _
                        & " and PortDimensions = '" & Trim(RodCapPort) & "'" ' and " & strNotes & " <> ''"'ANUP 21-10-2010 START
                    For Each strCol As String In strRod
                        strQuery = strQuery & " and " & Trim(strCol) & " <> ''"
                    Next
                    strQuery = strQuery & " and Notes ='No'"
                End If
                '27_10_2009  Ragava   Till  Here
                Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                If objDT.Rows.Count > 0 Then
                    txtRodCap.Text = objDT.Rows(0).Item(0).ToString
                    If txtRodCap.Text = "" Then
                        If blnRevision = False Then
                            MessageBox.Show("Rod Cap is not available for this configuration!" + vbNewLine + _
                            "Please select another configuration", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            cmbRodSealPackage.Focus()
                        End If
                    Else
                        strRodCapCodeNumber = objDT.Rows(0).Item("PartNumber").ToString
                        strRodCapDrawingNumber = "N/A"
                        strRodCapDescription = objDT.Rows(0).Item("RodCapDescription").ToString
                        mdiMonarch.mdiComponent.Items(7).SubItems.Add(strRodCapCodeNumber)
                        mdiMonarch.mdiComponent.Items(7).SubItems.Add(strRodCapDrawingNumber)
                        mdiMonarch.mdiComponent.Items(7).SubItems.Add(strRodCapDescription)
                    End If
                End If
            Else
                txtRodCap.Clear()
            End If
        End If
        LoadInformation()

    End Sub
    '06_04_2010   RAGAVA

    'Private Sub LoadClevis()
    '    Try
    '        Dim strQuery As String
    '        strQuery = ""
    '        dblCylinderForce_PullForce = Val(ofrmTieRod1.txtWorkingPressure.Text) * (3.1416 / 4) * (Math.Pow(BoreDiameter, 2) - Math.Pow(IIf(RodDiameter = 1.12, 1.125, RodDiameter), 2))
    '        dblCylinderPullForce = dblCylinderForce_PullForce
    '        strQuery = "select Description,PullForceMaximum from RodClevisDetails where "     '05_04_2010     RAGAVA
    '        strQuery = strQuery + " pinHoleSize=" & PinSize & " and pinHoleType='" & ofrmTieRod1.cmbRodClevisPinHole.Text & "'  and " & dblCylinderForce_PullForce & "<= pullforcemaximum and ThreadSize <= " & RodDiameter.ToString & " and Series = '" & IIf(Trim(ofrmTieRod1.cmbSeries.Text.ToString).StartsWith("TX"), "TX", "TL") & "'"      '20_04_2010   RAGAVA
    '        Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
    '        If objDT.Rows.Count = 0 Then
    '            If blnRevision = False Then
    '                MessageBox.Show("Rod Clevis is not available for the selected Configuration", "Information", MessageBoxButtons.OK)
    '            End If
    '            Me.rdbRodClevisNo.Checked = True
    '            Me.rdbRodClevisYes.Enabled = False
    '            Me.optPinsNo_Rod.Checked = True                   '13_04_2010   RAGAVA
    '            Me.optPinsYes_Rod.Enabled = False                 '13_04_2010   RAGAVA
    '        Else
    '            Me.rdbRodClevisNo.Checked = False
    '            Me.rdbRodClevisYes.Enabled = True
    '            Me.rdbRodClevisYes.Checked = True
    '            Me.optPinsNo_Rod.Checked = False                 '13_04_2010   RAGAVA
    '            Me.optPinsYes_Rod.Enabled = True                 '13_04_2010   RAGAVA
    '            Me.optPinsYes_Rod.Checked = True                 '13_04_2010   RAGAVA
    '        End If
    '        Try
    '            Dim aColumns As New ArrayList
    '            LVRodClevis.Columns.Clear()
    '            aColumns.Add(New Object(2) {"Description", "Description", True})
    '            aColumns.Add(New Object(2) {"PullForceMaximum", "Maximum" + vbNewLine + "Rated Pull Force", True})
    '            LVRodClevis.DisplayHeaders = aColumns
    '            Dim oTable As DataTable = objDT
    '            oTable.Columns.Add("Cost")
    '            LVRodClevis.FlushListViewData()
    '            LVRodClevis.SourceTable = objDT
    '            LVRodClevis.Populate()
    '            '09_10_2009
    '            If objDT.Rows.Count = 1 Then
    '                LVRodClevis.Items(0).Selected = True
    '                LVRodClevis.Enabled = False
    '            Else
    '                LVRodClevis.Enabled = True
    '            End If
    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        End Try
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub

    Private Sub LoadClevis()

        Try
            strRodClevis_Class = ""         '19_11_2012    RAGAVA
            dblDerateWorkingPressure = 0    '19_11_2012    RAGAVA
            Dim strQuery As String
            strQuery = ""
            dblCylinderForce_PullForce = Val(ofrmTieRod1.txtWorkingPressure.Text) * (3.1416 / 4) * _
            (Math.Pow(BoreDiameter, 2) - Math.Pow(IIf(RodDiameter = 1.12, 1.125, IIf(RodDiameter = 1.38, 1.375, RodDiameter)), 2))
            dblCylinderPullForce = dblCylinderForce_PullForce
            strQuery = "select PartNumber,Description,PullForceMaximum from RodClevisDetails where "     '05_04_2010     RAGAVA
            strQuery += "pinHoleSize=" & PinSize & " and pinHoleType='" & ofrmTieRod1.cmbRodClevisPinHole.Text _
             & "'  and " & dblCylinderForce_PullForce & "<= pullforcemaximum and ThreadSize IN "  '<= " & RodDiameter.ToString     
            'ANUP 12-10-2010 START
            strQuery += "(select rdd.RodThreadSize from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.PartNumberId = rdd.PartNumber and bdrd.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ")and rdd.series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and MaterialType = '" & Trim(RodMaterialForCosting) & "' and RodDiameter = " & RodDiameter
            'ANUP 12-10-2010 TILL HERE

            If Not SeriesForCosting.ToString.Contains("TX") Then
                strQuery += " and pistonthreadSize= " + dblPistonThreadSize.ToString
            End If
            If Trim(ofrmTieRod1.cmbStyle.Text) = "ASAE" Then
                strQuery += " And StrokeLength =" & Val(StrokeLength)
            End If
            strQuery += " ) order by PullForceMaximum Desc"

            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            If objDT.Rows.Count = 0 Then
                If blnRevision = False Then
                    If Not BtnBrowseClicked() Then
                        MessageBox.Show("Rod Clevis is not available for the selected Configuration", "Information", MessageBoxButtons.OK)
                    End If
                End If
                '19_11_2012   RAGAVA
                Try
                    strQuery = "select PartNumber,Description,Class1PullForce,PullForceMaximum from RodClevisDetails where "
                    strQuery += "pinHoleSize=" & PinSize & " and pinHoleType='" & ofrmTieRod1.cmbRodClevisPinHole.Text _
                     & "'  and " & dblCylinderForce_PullForce & "<= Class1PullForce and ThreadSize IN "
                    strQuery += "(select rdd.RodThreadSize from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.PartNumberId = rdd.PartNumber and bdrd.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ")and rdd.series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and MaterialType = '" & Trim(RodMaterialForCosting) & "' and RodDiameter = " & RodDiameter
                    If Not SeriesForCosting.ToString.Contains("TX") Then
                        strQuery += " and pistonthreadSize= " + dblPistonThreadSize.ToString
                    End If
                    If Trim(ofrmTieRod1.cmbStyle.Text) = "ASAE" Then
                        strQuery += " And StrokeLength =" & Val(StrokeLength)
                    End If
                    strQuery += " )  order by Class1PullForce Desc"
                    Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery)
                    If objDT1.Rows.Count = 0 Then
                        'DeRate
                        strQuery = "select PartNumber,Description,PullForceMaximum from RodClevisDetails where "
                        strQuery += "pinHoleSize=" & PinSize & " and pinHoleType='" & ofrmTieRod1.cmbRodClevisPinHole.Text _
                         & "'  and ThreadSize IN "
                        strQuery += "(select rdd.RodThreadSize from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.PartNumberId = rdd.PartNumber and bdrd.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ")and rdd.series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and MaterialType = '" & Trim(RodMaterialForCosting) & "' and RodDiameter = " & RodDiameter
                        If Not SeriesForCosting.ToString.Contains("TX") Then
                            strQuery += " and pistonthreadSize= " + dblPistonThreadSize.ToString
                        End If
                        If Trim(ofrmTieRod1.cmbStyle.Text) = "ASAE" Then
                            strQuery += " And StrokeLength =" & Val(StrokeLength)
                        End If
                        strQuery += " ) order by PullForceMaximum Desc"
                        Dim objDT2 As DataTable = oDataClass.GetDataTable(strQuery)
                        Dim dblDeratePullForce As Double = objDT2.Rows(0)("PullForceMaximum")
                        dblDerateWorkingPressure = dblDeratePullForce / ((3.1416 / 4) * _
                                    (Math.Pow(BoreDiameter, 2) - Math.Pow(IIf(RodDiameter = 1.12, 1.125, IIf(RodDiameter = 1.38, 1.375, RodDiameter)), 2)))
                        objDT = objDT2
                        strRodClevis_Class = ""
                        GoTo RodClevis_Class
                    Else
                        objDT = objDT1
                        strRodClevis_Class = "Class1"
                        GoTo RodClevis_Class
                    End If
                Catch ex As Exception

                End Try

                Me.rdbRodClevisNo.Checked = True
                Me.rdbRodClevisYes.Enabled = False
                Me.optPinsNo_Rod.Checked = True                   '13_04_2010   RAGAVA
                Me.optPinsYes_Rod.Enabled = False                 '13_04_2010   RAGAVA
            Else
RodClevis_Class:
                Me.rdbRodClevisNo.Checked = False
                Me.rdbRodClevisYes.Enabled = True
                Me.rdbRodClevisYes.Checked = True
                Me.optPinsNo_Rod.Checked = False                 '13_04_2010   RAGAVA
                Me.optPinsYes_Rod.Enabled = True                 '13_04_2010   RAGAVA
                Me.optPinsYes_Rod.Checked = True                 '13_04_2010   RAGAVA
            End If
            Try
                Dim aColumns As New ArrayList
                LVRodClevis.Columns.Clear()
                aColumns.Add(New Object(2) {"Description", "Description", True})
                aColumns.Add(New Object(2) {"PullForceMaximum", "Maximum" + vbNewLine + "Rated Pull Force", True})
                LVRodClevis.DisplayHeaders = aColumns

                Dim objDT1 As New DataTable
                objDT1.Columns.Add("Description")
                objDT1.Columns.Add("PullForceMaximum")
                objDT1.Columns.Add("Cost")
                For Each oDataRow As DataRow In objDT.Rows
                    Dim oRow As DataRow = objDT1.NewRow()
                    oRow("Description") = oDataRow("Description")
                    '19_11_2012   RAGAVA
                    If strRodClevis_Class = "Class1" Then
                        oRow("PullForceMaximum") = oDataRow("Class1PullForce")
                    Else
                        oRow("PullForceMaximum") = oDataRow("PullForceMaximum")
                    End If
                    'oRow("PullForceMaximum") = oDataRow("PullForceMaximum")
                    'Till  Here
                    Dim dblCost As Double = IFLConnectionObject.GetValue("Select Cost from CostingDetails where PartCode = '" _
                                                                        + oDataRow("PartNumber") + "'")
                    oRow("Cost") = dblCost
                    objDT1.Rows.Add(oRow)
                Next

                'oTable.Columns.Add("Cost")
                LVRodClevis.FlushListViewData()
                LVRodClevis.SourceTable = objDT1
                LVRodClevis.Populate()
                '09_10_2009

                'If Not LVRodClevis.Enabled Then
                If LVRodClevis.Items.Count > 0 Then
                    LVRodClevis.Items(0).Selected = True
                End If
                'End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub LVPinSizeDetails_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                    Handles LVPinSizeDetails.SelectedIndexChanged

        ListViewPinSizeDetails()

    End Sub

    Private Sub ListViewPinSizeDetails()

        Dim oListViewItem As ListViewItem
        Dim lVPinSize As Double

        Try
            'cmbPinMaterial.Items.Clear()
            'cmbClips.Items.Clear()
            'cmbRodClevisPinClips.Items.Clear()
            LVRodClevis.Clear()
            If Module1.BtnBrowseClicked Then
                If LVPinSizeDetails.Items.Count = 1 Then
                    lVPinSize = Convert.ToDouble(LVPinSizeDetails.Items(0).Text)
                Else
                    For i As Integer = 0 To LVPinSizeDetails.Items.Count - 1
                        Dim pinSizeNo As Double = LVPinSizeDetails.Items(i).Text

                        If pinSizeNo = Convert.ToDouble(Module1.ReadValuesFromExcel.PinSizeDetails) Then

                            lVPinSize = pinSizeNo

                            Exit For
                        End If
                    Next
                End If

            Else
                If LVPinSizeDetails.SelectedItems.Count > 0 Then
                    oListViewItem = LVPinSizeDetails.SelectedItems(0)
                    lVPinSize = Convert.ToDouble(oListViewItem.SubItems(0).Text)
                Else
                    cmbPinMaterial.Items.Clear()

                    WorkingPressure = Val(ofrmTieRod1.txtWorkingPressure.Text)
                    LoadInformation()
                    ListViewPinSizeDetailscontinuation()
                    Exit Sub
                End If
            End If
            'If LVPinSizeDetails.SelectedItems.Count > 0 Then
            cmbPinMaterial.Items.Add(" ")
            'cmbClips.Items.Add(" ")
            cmbRodClevisPinClips.Items.Add(" ")
            Dim strQuery As String
            PinSize = lVPinSize
            dblPinSize = PinSize       '10_09_2011   RAGAVA
            PinHoleSize = PinSize
            strQuery = ""
            Dim arrSeries As String()
            If SeriesForCosting.ToString.StartsWith("TP") Then
                arrSeries = SeriesForCosting.ToString.Split("-")
            Else
                arrSeries = SeriesForCosting.ToString.Split(" ")
            End If
            Dim strSeries As String = arrSeries(0)
            strQuery = "select * from ClevisCapDetails where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) _
                & " and " & GetClevisCapDetails() & " and port = '" & Trim(ClevisCapPort) & "' and PinHoleSize = " & _
                lVPinSize & " and PinHoleType = '" & Trim(ofrmTieRod1.cmbClevisCapPinHole.Text) & "'"

            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            txtClevisCap.Text = objDT.Rows(0).Item("PartNumber").ToString

            strClevisCapCodeNumber = objDT.Rows(0).Item("PartNumber").ToString
            ClevisCapCodeNumber = strClevisCapCodeNumber           '12_10_2009   ragava
            strClevisCapDrawingNumber = "N/A"
            strClevisCapDescription = objDT.Rows(0).Item("Description").ToString

            mdiMonarch.mdiComponent.Items(8).SubItems.Add(strClevisCapCodeNumber)
            mdiMonarch.mdiComponent.Items(8).SubItems.Add(strClevisCapDrawingNumber)
            mdiMonarch.mdiComponent.Items(8).SubItems.Add(strClevisCapDescription)
            Try
                strQuery = "Select distinct PinMaterial  from ClevisPinDetails where PinHoleSize=" & _
                    lVPinSize & " And  " & Val(ofrmTieRod1.cmbBore.Text) & " >= BoreDiameterMinimum and " & _
                    Val(ofrmTieRod1.cmbBore.Text) & " < = BoreDiameterMaximum"
                Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery)
                cmbPinMaterial.Items.Clear()
                cmbPinMaterial.Items.Add(" ")
                '  If Not mdiMonarch.isGeneratedBtnClicked Then
                For Each dr As DataRow In objDT1.Rows
                    cmbPinMaterial.Items.Add(dr(0).ToString)
                Next
                '09_10_2009
                If objDT1.Rows.Count = 1 Then
                    cmbPinMaterial.SelectedIndex = 1
                    cmbPinMaterial.Enabled = False
                Else
                    If optPinsNo.Checked = False Then     ''27_10_2009   Ragava
                        cmbPinMaterial.Enabled = True
                        cmbPinMaterial.Text = "Standard"                   '11_11_2009   Ragava
                    End If
                End If
                'Else
                'cmbPinMaterial.Text = Module1.ReadValuesFromExcel.PinMaterial.ToString()
                'End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            If Not SeriesForCosting.ToString.StartsWith("TX") Then
                setWorkingPressure(ofrmTieRod1.dbltempWorkingPressureNut)
            Else
                setWorkingPressure(ofrmTieRod1.dbltempWorkingPressureSeries)
            End If
            Try
                ofrmTieRod1.CheckRodEndThreadSizeLogics()
            Catch ex As Exception

            End Try
            'Else
            '    cmbPinMaterial.Items.Clear()
            ' End If
            WorkingPressure = Val(ofrmTieRod1.txtWorkingPressure.Text)
            LoadInformation()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        '06_04_2010   RAGAVA
        ListViewPinSizeDetailscontinuation()

    End Sub

    Private Sub ListViewPinSizeDetailscontinuation()

        If optPinsNo_Rod.Checked = True AndAlso optPinsNo.Checked = True Then
            cmbPinMaterial.SelectedIndex = -1
            cmbPinMaterial.Enabled = False
        End If
        If optPinsNo_Rod.Checked = True Then
            cmbRodClevisPinClips.SelectedIndex = -1
            cmbRodClevisPinClips.Enabled = False
        End If
        If optPinsNo.Checked = True Then
            'cmbPinMaterial.SelectedIndex = -1
            'cmbPinMaterial.Enabled = False
            cmbClips.SelectedIndex = -1
            cmbClips.Enabled = False
            'cmbRodClevisPinClips.SelectedIndex = -1
            'cmbRodClevisPinClips.Enabled = False
        End If
        '06_04_2010   RAGAVA   Till   Here
        For Each listviewItem As ListViewItem In LVPinSizeDetails.Items
            Dim index As Integer = LVPinSizeDetails.Items.IndexOf(listviewItem)
            LVPinSizeDetails.Items(index).BackColor = Color.Ivory
            LVPinSizeDetails.Items(index).ForeColor = Color.Black
        Next
        For Each listviewItem As ListViewItem In LVPinSizeDetails.SelectedItems
            Dim index As Integer = LVPinSizeDetails.Items.IndexOf(listviewItem)
            LVPinSizeDetails.Items(index).BackColor = Color.CornflowerBlue
            LVPinSizeDetails.Items(index).ForeColor = Color.White
        Next
        '06_04_2010   RAGAVA

        If mdiMonarch.IsBtnmygClicked Or Module1.BtnBrowseClicked Then
            LoadClevis()
        Else
            If LVPinSizeDetails.SelectedItems.Count > 0 Then
                LoadClevis()
            End If
        End If

    End Sub

    Private Sub setWorkingPressure(ByVal workingPressure As Double)

        Dim oSelectedListviewItem As ListViewItem

        If mdiMonarch.IsBtnmygClicked Or Module1.BtnBrowseClicked Then
            oSelectedListviewItem = LVPinSizeDetails.Items(0)
        Else
            oSelectedListviewItem = LVPinSizeDetails.SelectedItems(0)
        End If
        If oSelectedListviewItem.SubItems(2).Text <> "N/A" Then
            'If Val(ofrmTieRod1.txtWorkingPressure.Text) > Val(oSelectedListviewItem.SubItems(2).Text) Then
            If workingPressure > Val(oSelectedListviewItem.SubItems(2).Text) Then
                ofrmTieRod1.txtWorkingPressure.Text = oSelectedListviewItem.SubItems(2).Text
            End If
        Else
            '03_11_2009  Ragava
            If Val(ofrmTieRod1.txtWorkingPressure.Text) > ofrmTieRod1.dbltempWorkingPressureSeries Then
                ofrmTieRod1.txtWorkingPressure.Text = ofrmTieRod1.dbltempWorkingPressureSeries
            End If
            'ofrmTieRod1.txtWorkingPressure.Text = ofrmTieRod1.dbltempWorkingPressureSeries
            '03_11_2009  Ragava  Till  Here
        End If

    End Sub

    Private Sub optPinsNo_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                                 As System.EventArgs) Handles optPinsNo.CheckedChanged

        RadioBtnClevisCapPinsNo(sender)

    End Sub

    Private Sub RadioBtnClevisCapPinsNo(ByVal sender As System.Object)

        If sender.checked Then
            If optPinsNo_Rod.Checked = True AndAlso optPinsNo.Checked = True Then  '06_04_2010   RAGAVA
                cmbPinMaterial.SelectedIndex = -1
                cmbPinMaterial.Enabled = False
            End If
            If sender.Name = "optPinsNo" Then           '06_04_2010   RAGAVA
                ClevisPins = False
                cmbClips.SelectedIndex = -1
                cmbClips.Enabled = False
            ElseIf sender.Name = "optPinsNo_Rod" Then           '06_04_2010   RAGAVA
                RodClevisPins = False
                cmbRodClevisPinClips.SelectedIndex = -1
                cmbRodClevisPinClips.Enabled = False
            End If
            blnPins = False       '11_11_2009   Ragava
        Else
            'cmbPinMaterial.Enabled = True
            LVPinSizeDetails.Enabled = True
            If optPinsYes.Checked = True Then           '06_04_2010   RAGAVA
                cmbClips.Enabled = True
            End If
            If optPinsYes_Rod.Checked = True Then           '06_04_2010   RAGAVA
                cmbRodClevisPinClips.Enabled = True
            End If
        End If
        '06_04_2010   RAGAVA
        If optPinsYes_Rod.Checked = True Or optPinsYes.Checked = True Then
            cmbPinMaterial.Enabled = True
        End If

    End Sub

    Private Sub cmbRodEndThread_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                    Handles cmbRodEndThread.SelectedIndexChanged

        ComboBoxRodEndThread()

    End Sub

    Private Sub ComboBoxRodEndThread()

        strNewTableDrawingNumber = strNewTubeTableDrawingNumber          '02_12_2009  RAGAVA
        'LVRodClevis.Clear()                 '06_04_2010   RAGAVA   Commented
        If Trim(cmbRodEndThread.Text) <> "" Then
            dblRodThreadSize = Val(cmbRodEndThread.Text)
            Try

                Dim strQuery As String
                Try
                    strQuery = "select * from RodClevisDetails where " 'ThreadSize=" & Val(cmbRodEndThread.Text) & vbNewLine        '05_04_2010   RAGAVA
                    strQuery = strQuery + " pinHoleSize=" & PinSize & " and pinHoleType='" & _
                                        ofrmTieRod1.cmbRodClevisPinHole.Text & "' " 'and " & dblCylinderForce & " >= pullforcemaximum"
                    Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery)
                    If objDT1.Rows.Count > 0 Then
                        strRodClevisDrawingNumber = "N/A"
                        If strRodClevisCodeNumber Is Nothing Then
                            strRodClevisCodeNumber = objDT1.Rows(0).Item("PartNumber").ToString
                        End If
                        strRodClevisDescription = objDT1.Rows(0).Item("Description").ToString
                        mdiMonarch.mdiComponent.Items(9).SubItems.Add(strRodClevisCodeNumber)
                        mdiMonarch.mdiComponent.Items(9).SubItems.Add(strRodClevisDrawingNumber)
                        mdiMonarch.mdiComponent.Items(9).SubItems.Add(strRodClevisDescription)
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim StrSql As String
                Dim oListViewItem As ListViewItem                     '23_11_2009   Ragava
                Dim dblPistonNutSize As Double              '19_02_2010    RAGAVA
                If ofrmTieRod1.LVNutSizeDetails.SelectedItems.Count > 0 Then      '19_02_2010    RAGAVA
                    oListViewItem = ofrmTieRod1.LVNutSizeDetails.SelectedItems(0)                     '23_11_2009   Ragava
                    dblPistonNutSize = Val(oListViewItem.SubItems(0).Text)                     '23_11_2009   Ragava
                End If
                'ANUP 12-10-2010 START
                If ofrmTieRod1.LVNutSizeDetails.SelectedItems.Count < 1 Then
                    StrSql = "select rdd.DrawingPartNumber,rdd.PartNumber,rdd.Description,rdd.OverAllRodLength,rdd.StrokeLength from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.PartNumberId = rdd.PartNumber and bdrd.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " _
                    & Val(ofrmTieRod1.cmbBore.Text) & ")and rdd.series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and MaterialType = '" & Trim(ofrmTieRod1.cmbRodMaterial.Text) & "' and RodDiameter = " & RodDiameter & " and RodThreadSize = " & Val(ofrmTieRod2.cmbRodEndThread.Text) '02_12_2009   Ragava
                Else
                    StrSql = "select rdd.DrawingPartNumber,rdd.PartNumber,rdd.Description,rdd.OverAllRodLength,rdd.StrokeLength from RodDiameterDetails rdd,BoreDiameter_RodDiameter bdrd where bdrd.PartNumberId = rdd.PartNumber and bdrd.BoreDiameterID = (Select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) & ")and rdd.series = '" & IIf(Trim(SeriesForCosting.ToString).StartsWith("TX"), "TX", "TL/TH/TP/LN") & "' and IsASAE = '" & Trim(ofrmTieRod1.cmbStyle.Text) & "' and MaterialType = '" & Trim(ofrmTieRod1.cmbRodMaterial.Text) & "' and RodDiameter = " & RodDiameter & " and RodThreadSize = " & Val(ofrmTieRod2.cmbRodEndThread.Text) & " And pistonthreadSize= " & dblPistonNutSize  '02_12_2009   Ragava
                End If
                'ANUP 12-10-2010 TILL HERE
                If Trim(ofrmTieRod1.cmbStyle.Text) = "ASAE" Then
                    StrSql = StrSql + " And StrokeLength =" & Val(ofrmTieRod1.cmbStrokeLength.Text)
                Else
                    '02_12_2009  ragava
                    'StrSql = StrSql + " And StrokeLength =" & Val(ofrmTieRod1.txtStrokeLength.Text)             '23_11_2009    Ragava
                    Try
                        Dim objDT_temp As DataTable = oDataClass.GetDataTable(StrSql)
                        If objDT_temp.Rows.Count > 0 Then
                            RodStrokeDifference = Math.Round(Val(objDT_temp.Rows(0).Item("OverAllRodLength").ToString) _
                                                            - Val(objDT_temp.Rows(0).Item("StrokeLength").ToString), 2)
                        End If
                    Catch ex As Exception
                    End Try
                    Dim dblRodLength As Double = StrokeLength + RodStrokeDifference + StopTubeLength + RodAdder
                    StrSql = StrSql + " And OverAllRodLength =" & dblRodLength.ToString
                    '02_12_2009  ragava  Till  Here
                End If
                Dim objDT As DataTable = oDataClass.GetDataTable(StrSql)
                If objDT.Rows.Count > 0 Then
                    strRodDrawingNumber = objDT.Rows(0).Item("DrawingPartNumber").ToString
                    strRodCodeNumber = objDT.Rows(0).Item("PartNumber").ToString
                    strRodDescription = objDT.Rows(0).Item("Description").ToString
                    mdiMonarch.mdiComponent.Items(1).SubItems.Add(" ")
                    mdiMonarch.mdiComponent.Items(1).SubItems.Add(strRodDrawingNumber)
                    mdiMonarch.mdiComponent.Items(1).SubItems.Add(strRodDescription)
                    '02_12_2009   ragava
                Else
                    Dim strQuery As String = "Select CodeNumber from CodeNumberDetails where Type = 'ROD'"
                    Dim objDT3 As DataTable = oDataClass.GetDataTable(strQuery)
                    If strNewTableDrawingNumber <> "" Then
                        strNewTableDrawingNumber = (Val(strNewTableDrawingNumber) + 1).ToString
                    Else
                        strNewTableDrawingNumber = objDT3.Rows(0).Item("CodeNumber").ToString
                    End If
                    strRodCodeNumber = strNewTableDrawingNumber
                    '02_12_2009   ragava  Till  Here
                End If
                'End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            LoadInformation()
        End If

    End Sub

    Public Sub ActivatedCodeTieRod2()

        If optPinsNo.Checked = False AndAlso optPinsYes.Checked = False Then
            optPinsYes.Checked = True            '19_04_2010   RAGAVA
        End If
        If optPinsNo_Rod.Checked = False AndAlso optPinsYes_Rod.Checked = False Then
            optPinsYes_Rod.Checked = True            '20_04_2010   RAGAVA
        End If
        txtColumnLoad.Text = ColumnLoad
        txtWorkingPressure.Text = WorkingPressure
        LoadInformation()
        CustomerName = Trim(ofrmContractDetails.cmbCustomerName.Text)   '22_02_2010    Ragava
        AssemblyType = ofrmContractDetails.cmbAssemblyType.Text
        PartCode1 = ofrmContractDetails.txtlPartCode.Text
        If cmbThreadProtected.SelectedIndex = -1 Then
            cmbThreadProtected.Text = "Standard"                   '11_11_2009   Ragava
        End If
        'Try
        '    Dim strQuery As String
        '    If Trim(cmbPaint.Text) = "" And cmbPaint.Items.Count < 1 Then
        '        cmbPaint.Items.Clear()
        '        cmbPaint.Items.Add(" ")
        '        strQuery = ""
        '        strQuery = "select Distinct Color from TieRodPaintDetails"
        '        Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
        '        For Each dr As DataRow In objDT.Rows
        '            cmbPaint.Items.Add(dr(0).ToString)
        '        Next
        '        '09_10_2009
        '        If objDT.Rows.Count = 1 Then
        '            cmbPaint.Items(0).Selected = True
        '            cmbPaint.Enabled = False
        '        Else
        '            cmbPaint.Enabled = True
        '        End If
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

    End Sub

    Private Sub cmbClips_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                            Handles cmbClips.SelectedIndexChanged

        ComboBoxClips()

    End Sub

    Private Sub ComboBoxClips()

        If Trim(cmbClips.Text) <> "" Then
            Try
                Dim oListViewItem As ListViewItem
                Try
                    If Module1.BtnBrowseClicked Then
                        If Not LVPinSizeDetails.Items.Count = 0 Then
                            oListViewItem = LVPinSizeDetails.Items(0)
                        Else
                            Exit Try
                        End If
                    Else
                        If LVPinSizeDetails.SelectedItems.Count > 0 Then
                            oListViewItem = LVPinSizeDetails.SelectedItems(0)
                        Else
                            Exit Try
                        End If
                    End If

                    Dim strQuery As String
                    Dim objDT As DataTable
                    strQuery = ""
                    strQuery = "select *  from ClevisPinDetails where PinHoleSize = " & _
                        Val(oListViewItem.SubItems(0).Text) & " and PinMaterial = '" & Trim(cmbPinMaterial.Text) _
                        & "' And  " & Val(ofrmTieRod1.cmbBore.Text) & " >= BoreDiameterMinimum and " & _
                        Val(ofrmTieRod1.cmbBore.Text) & " < = BoreDiameterMaximum  and PinType = '" + cmbClips.Text + "'"
                    objDT = oDataClass.GetDataTable(strQuery)
                    strPinsCodeNumber = objDT.Rows(0).Item("PartNumber").ToString
                    strPinsDrawingNumber = "N/A"
                    strPinsDescription = "N/A"
                    strClevisCapPinCodeNumber = strPinsCodeNumber ' Todo:Sandeep 04-04-10
                    mdiMonarch.mdiComponent.Items(4).SubItems.Add(strPinsCodeNumber)
                    mdiMonarch.mdiComponent.Items(4).SubItems.Add(strPinsDrawingNumber)
                    mdiMonarch.mdiComponent.Items(4).SubItems.Add(strPinsDescription)
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
                If optPinsYes_Rod.Checked = True Then         '07_04_2010    RAGAVA
                    cmbRodClevisPinClips.Text = Trim(cmbClips.Text)      '11_11_2009  Ragava
                ElseIf optPinsNo_Rod.Checked = True Then      '07_04_2010    RAGAVA
                    cmbRodClevisPinClips.SelectedIndex = -1   '07_04_2010    RAGAVA
                End If
                ClevisPinClips = Trim(cmbClips.Text)   '11_11_2009  Ragava
                _strPinCodeBE = strPinsCodeNumber.ToString.Substring(0, 6)          '16_06_2011   RAGAVA
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            LoadInformation()
        End If

    End Sub

    Private Sub cmbPistonSealPackage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e _
                                    As System.EventArgs) Handles cmbPistonSealPackage.SelectedIndexChanged

        ComboBoxPistonSealPackage()

    End Sub

    Private Sub ComboBoxPistonSealPackage()

        If cmbPistonSealPackage.Text <> "" AndAlso cmbPistonSealPackage.Items.Count <> 0 Then
            strPistonSealPackage = Trim(cmbPistonSealPackage.Text)
            Try
                Dim oListViewItem As ListViewItem
                Dim strColumns() As String
                strColumns = (cmbPistonSealPackage.Text).Split("+")
                Dim StrSql As String = ""
                Dim arrSeries As String()
                arrSeries = SeriesForCosting.ToString.Split(" ")
                Dim strSeries As String = arrSeries(0)
                If SeriesForCosting.ToString.StartsWith("TX") = False Then
                    If ofrmTieRod1.LVNutSizeDetails.SelectedItems.Count > 0 Then
                        oListViewItem = ofrmTieRod1.LVNutSizeDetails.SelectedItems(0)
                        StrSql = "select * from PistonSealDetails where BoreDiameter = " & _
                            Val(ofrmTieRod1.cmbBore.Text) & " and PistonNutSize = " & Val(oListViewItem.SubItems(0).Text) _
                            & " and Series like '%" & strSeries & "%'"
                    End If
                Else
                    StrSql = "select * from PistonSealDetails where BoreDiameter = " & Val(ofrmTieRod1.cmbBore.Text) _
                        & " and Series like '%" & strSeries & "%'" '" and PistonNutSize = " & Val(oListViewItem.SubItems(0).Text)
                End If
                If Trim(StrSql) <> "" Then
                    For Each strCol As String In strColumns
                        StrSql = StrSql & " and " & Trim(strCol) & " <> ''"
                    Next
                    Dim objDT As DataTable = oDataClass.GetDataTable(StrSql)
                    If objDT.Rows.Count > 0 Then
                        strPistonCodeNumber = objDT.Rows(0).Item("PartNumber").ToString
                        strPistonDrawingNumber = "N/A"
                        strPistonDescription = objDT.Rows(0).Item("PistonDescription").ToString
                        mdiMonarch.mdiComponent.Items(2).SubItems.Add(strPistonCodeNumber)
                        mdiMonarch.mdiComponent.Items(2).SubItems.Add(strPistonDrawingNumber)
                        mdiMonarch.mdiComponent.Items(2).SubItems.Add(strPistonDescription)
                    End If
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            LoadInformation()
        End If

    End Sub

    Private Sub rdbRodClevisYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                    Handles rdbRodClevisYes.CheckedChanged

        RadioBtnRodClevisYes()

    End Sub

    Private Sub RadioBtnRodClevisYes()

        If rdbRodClevisYes.Checked Then
            rdbRodClevis = True        '19_02_2010     RAGAVA
            CaluculateLength(True)
            ofrmTieRod1.txtRetractedLength.Enabled = True
        End If

    End Sub

    Private Sub frmTieRod2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '02_11_2009   Ragava 
        'rdbRodClevisNo.Enabled = True
        'rdbRodClevisYes.Checked = True
        '02_11_2009   Ragava  Till  Here


        'Try
        '    Dim strQuery As String
        '    cmbPaint.Items.Clear()
        '    cmbPaint.Items.Add(" ")
        '    strQuery = ""
        '    strQuery = "select Distinct Color from TieRodPaintDetails"
        '    Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
        '    For Each dr As DataRow In objDT.Rows
        '        cmbPaint.Items.Add(dr(0).ToString)
        '    Next
        '    '09_10_2009
        '    If objDT.Rows.Count = 1 Then
        '        cmbPaint.Items(0).Selected = True
        '        cmbPaint.Enabled = False
        '    Else
        '        cmbPaint.Enabled = True
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try


        ofrmContractDetails.lVLogInformation.Items.Clear()

        ColorTheForm()

        'loadingdataFromExcelTieRod2()
    End Sub

    Private Sub rdbRodClevisNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                                                                Handles rdbRodClevisNo.CheckedChanged

        RadioBtnRodClevisNo()

    End Sub

    Private Sub RadioBtnRodClevisNo()

        If rdbRodClevisNo.Checked Then
            rdbRodClevis = False        '19_02_2010     RAGAVA
            'cmbRodClevis.SelectedIndex = -1
            'cmbRodClevis.Enabled = False
            CaluculateLength(False)
            ofrmTieRod1.txtRetractedLength.Enabled = True
            cmbRodClevisPinClips.SelectedIndex = -1
            cmbRodClevisPinClips.Enabled = False
            LVRodClevis.Enabled = False
            'strRodClevisCodeNumber = ""
            'strRodClevisDrawingNumber = ""
            'strRodClevisDescription = ""
            'mdiMonarch.mdiComponent.Items(9).SubItems.Add(strRodClevisCodeNumber)
            'mdiMonarch.mdiComponent.Items(9).SubItems.Add(strRodClevisDrawingNumber)
            'mdiMonarch.mdiComponent.Items(9).SubItems.Add(strRodClevisDescription)
        Else
            If optPinsNo.Checked = True Then
                cmbRodClevisPinClips.SelectedIndex = -1
                cmbRodClevisPinClips.Enabled = False
            Else
                cmbRodClevisPinClips.Enabled = True
            End If
            LVRodClevis.Enabled = True
        End If

    End Sub

    Private Sub LVRodClevis_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As _
                                                System.EventArgs) Handles LVRodClevis.SelectedIndexChanged

        ListViewRodClevis()

    End Sub

    Private Sub ListViewRodClevis()

        Dim strQuery As String
        Dim oListViewItem As ListViewItem

        Dim rodCleviesDesc As String
        Dim rodCleviesPullForceMaximum As String
        Try
            If Module1.BtnBrowseClicked Then
                If Not LVRodClevis.Items.Count = 0 Then
                    'oListViewItem = LVRodClevis.Items(0)

                    For i As Integer = 0 To LVRodClevis.Items.Count - 1
                        If LVRodClevis.Items(i).Text = Module1.ReadValuesFromExcel.RodClevis Then
                            ' LVRodClevis.Items(i).Selected = True

                            rodCleviesDesc = LVRodClevis.Items(i).SubItems(0).Text
                            rodCleviesPullForceMaximum = LVRodClevis.Items(i).SubItems(1).Text
                            Exit For
                        End If
                    Next
                Else
                    Exit Sub
                End If
            Else
                If LVRodClevis.SelectedItems.Count > 0 Then
                    oListViewItem = LVRodClevis.SelectedItems(0)
                    rodCleviesDesc = oListViewItem.SubItems(0).Text
                    rodCleviesPullForceMaximum = oListViewItem.SubItems(1).Text
                Else

                    Exit Sub
                End If
            End If

            '05_04_2010    RAGAVA
            '19_11_2012   RAGAVA
            If strRodClevis_Class = "Class1" Then
                strQuery = "select ThreadSize from RodClevisDetails where  Description ='" & Trim(rodCleviesDesc) _
                                                            & "' And Class1PullForce > = " & Val(rodCleviesPullForceMaximum)
            Else
                strQuery = "select ThreadSize from RodClevisDetails where  Description ='" & Trim(rodCleviesDesc) _
                                            & "' And PullForceMaximum > = " & Val(rodCleviesPullForceMaximum)
            End If
            'strQuery = "select ThreadSize from RodClevisDetails where  Description ='" & Trim(rodCleviesDesc) _
            '                                & "' And PullForceMaximum > = " & Val(rodCleviesPullForceMaximum)
            'Till   Here
            Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery)
            If objDT1.Rows.Count > 0 Then
                cmbRodEndThread.Items.Clear()
                For Each dr As DataRow In objDT1.Rows
                    cmbRodEndThread.Items.Add(dr(0).ToString)
                    cmbRodEndThread.Text = dr(0).ToString
                Next
            End If

            strQuery = ""
            '19_11_2012   RAGAVA
            If strRodClevis_Class = "Class1" Then
                strQuery = "select * from RodClevisDetails where  Description ='" & Trim(rodCleviesDesc) _
                                        & "' And Class1PullForce > = " & Val(rodCleviesPullForceMaximum)
            Else
                strQuery = "select * from RodClevisDetails where  Description ='" & Trim(rodCleviesDesc) _
                                            & "' And PullForceMaximum > = " & Val(rodCleviesPullForceMaximum)
            End If
            'strQuery = "select * from RodClevisDetails where  Description ='" & Trim(rodCleviesDesc) _
            '                            & "' And PullForceMaximum > = " & Val(rodCleviesPullForceMaximum)

            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            For Each dr As DataRow In objDT.Rows
                strRodClevisCodeNumber = objDT.Rows(0).Item("PartNumber").ToString
                strRodClevisDrawingNumber = "N/A"
                strRodClevisDescription = objDT.Rows(0).Item("Description").ToString
                mdiMonarch.mdiComponent.Items(9).SubItems.Add(strRodClevisCodeNumber)
                mdiMonarch.mdiComponent.Items(9).SubItems.Add(strRodClevisDrawingNumber)
                mdiMonarch.mdiComponent.Items(9).SubItems.Add(strRodClevisDescription)
                'cmbRodClevis.Items.Clear()
                'cmbRodClevis.Items.Add(" ")
                'cmbRodClevis.Items.Add(strRodClevisCodeNumber)
                'cmbRodClevis.Text = strRodClevisCodeNumber
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Try
            'If LVRodClevis.SelectedItems.Count > 0 Then
            'Dim strQuery As String
            'Dim oListViewItem As ListViewItem
            ' oListViewItem = LVRodClevis.SelectedItems(0)
            ' oListViewItem = LVRodClevis.Items(LVRodClevis.GetCurrentIndex)
            strQuery = ""
            '19_11_2012   RAGAVA
            If strRodClevis_Class = "Class1" Then
                strQuery = "select * from RodClevisDetails where  Description ='" & Trim(rodCleviesDesc) _
                                                       & "' And Class1PullForce = " & Val(rodCleviesPullForceMaximum)
            Else
                strQuery = "select * from RodClevisDetails where  Description ='" & Trim(rodCleviesDesc) _
                                       & "' And PullForceMaximum = " & Val(rodCleviesPullForceMaximum)
            End If
            'strQuery = "select * from RodClevisDetails where  Description ='" & Trim(rodCleviesDesc) _
            '                            & "' And PullForceMaximum = " & Val(rodCleviesPullForceMaximum)
            'Till  Here
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            If objDT.Rows.Count > 0 Then
                strRodClevisCodeNumber = objDT.Rows(0).Item("PartNumber").ToString
            End If
            strQuery = ""
            cmbRodClevisPinClips.Items.Clear()

            '11_11_2009  Ragava
            'strQuery = "select distinct  HairClip,CotterPin,R_Pin,RetainingRing from RodClevisPinDetails where   PartNumber = '" & strRodClevisCodeNumber & " ' And IsStandard =" & IIf(cmbPinMaterial.Text = "Standard", 1, 0)
            strQuery = "select distinct  [Hair pins],[Cotter Pins],[R - Style Pins],[Retaining rings] from RodClevisPinDetails where   PartNumber = '" & strRodClevisCodeNumber & " ' And IsStandard =" & IIf(cmbPinMaterial.Text = "Standard", 1, 0)
            '11_11_2009  Ragava   Till  Here

            Dim objDT1 As DataTable = oDataClass.GetDataTable(strQuery)
            cmbRodClevisPinClips.Items.Add(" ")
            For Each dr As DataRow In objDT1.Rows
                Dim strRow As String = ""
                If Trim(dr(0).ToString) <> "" Then
                    strRow = objDT1.Columns.Item(0).ToString
                    cmbRodClevisPinClips.Items.Add(strRow)
                End If
                If Trim(dr(1).ToString) <> "" Then
                    strRow = objDT1.Columns.Item(1).ToString
                    cmbRodClevisPinClips.Items.Add(strRow)
                End If
                If Trim(dr(2).ToString) <> "" Then
                    strRow = objDT1.Columns.Item(2).ToString
                    cmbRodClevisPinClips.Items.Add(strRow)
                End If
                If Trim(dr(3).ToString) <> "" Then
                    strRow = objDT1.Columns.Item(3).ToString
                    cmbRodClevisPinClips.Items.Add(strRow)
                End If
            Next
            '09_10_2009
            If optPinsNo.Checked = False Then
                If cmbRodClevisPinClips.Items.Count = 2 Then
                    cmbRodClevisPinClips.SelectedIndex = 1
                    cmbRodClevisPinClips.Enabled = False

                Else
                    If optPinsYes_Rod.Checked = True Then          '06_04_2010    RAGAVA
                        cmbRodClevisPinClips.Enabled = True
                        'cmbRodClevisPinClips.Text = "Cotter Pins"      '11_11_2009  Ragava
                        cmbRodClevisPinClips.Text = Trim(cmbClips.Text)      '11_11_2009  Ragava
                    End If
                End If
            End If
            '  End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        LoadInformation()
        For Each listviewItem As ListViewItem In LVRodClevis.Items
            Dim index As Integer = LVRodClevis.Items.IndexOf(listviewItem)
            LVRodClevis.Items(index).BackColor = Color.Ivory
            LVRodClevis.Items(index).ForeColor = Color.Black
        Next
        For Each listviewItem As ListViewItem In LVRodClevis.SelectedItems
            Dim index As Integer = LVRodClevis.Items.IndexOf(listviewItem)
            LVRodClevis.Items(index).BackColor = Color.CornflowerBlue
            LVRodClevis.Items(index).ForeColor = Color.White
        Next

    End Sub

    Private Sub CmbRodClevisPinClips_SelectedIndexChanged(ByVal sender As System.Object, ByVal e _
                                        As System.EventArgs) Handles cmbRodClevisPinClips.SelectedIndexChanged

        ComboBoxRodClevisPinClips()

    End Sub

    Private Sub ComboBoxRodClevisPinClips()

        Try
            If Trim(cmbRodClevisPinClips.Text) <> "" Then
                Dim strQuery As String = ""
                '11_11_2009  Ragava
                'strQuery = "Select " & sender.Text.ToString & " from RodClevisPinDetails  where PartNumber = '" & strRodClevisCodeNumber & " ' and IsStandard=" & IIf(cmbPinMaterial.Text = "Standard", 1, 0)
                strQuery = "Select [" & cmbRodClevisPinClips.Text.ToString & "] from RodClevisPinDetails  where PartNumber = '" _
                            & strRodClevisCodeNumber & " ' and IsStandard=" & IIf(cmbPinMaterial.Text = "Standard", 1, 0)
                '11_11_2009  Ragava   Till  Here
                Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                If objDT.Rows.Count > 0 Then
                    strRodClevisPinCodeNumber = objDT.Rows(0).Item(cmbRodClevisPinClips.Text.ToString).ToString
                End If
                RodPinClips = Trim(cmbRodClevisPinClips.Text)   '11_11_2009  Ragava

                _strPinCodeRE = strRodClevisPinCodeNumber.Substring(0, 6)          '16_06_2011   RAGAVA

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        LoadInformation()

    End Sub

    Private Sub optPinsYes_CheckedChanged(ByVal sender As System.Object, ByVal e _
                                    As System.EventArgs) Handles optPinsYes.CheckedChanged

        RadioBtnClevisCapPinsYes()

    End Sub

    Private Sub RadioBtnClevisCapPinsYes()

        If optPinsYes.Checked = True Then
            cmbPinMaterial.Enabled = True
            cmbClips.Enabled = True
            ClevisPins = True        '06_04_2010   RAGAVA
            blnPins = True       '11_11_2009   Ragava
        End If

    End Sub
    '06_04_2010    RAGAVA
    Private Sub optPinsYes_Rod_CheckedChanged(ByVal sender As System.Object, ByVal e _
                            As System.EventArgs) Handles optPinsYes_Rod.CheckedChanged

        If sender.Checked = True Then
            RodClevisPins = True        '06_04_2010   RAGAVA
            cmbPinMaterial.Enabled = True          '06_04_2010    RAGAVA
        End If

    End Sub


    Private Sub cmbThreadProtected_SelectedIndexChanged(ByVal sender As System.Object, _
                    ByVal e As System.EventArgs) Handles cmbThreadProtected.SelectedIndexChanged

        If cmbThreadProtected.Text <> "" Then
            ThreadProtected = cmbThreadProtected.Text
        End If

    End Sub

    Private Sub txtTieRodSize_TextChanged(ByVal sender As System.Object, _
                        ByVal e As System.EventArgs) Handles txtTieRodSize.TextChanged

        If Not txtTieRodSize.Text = "" Then
            TieRodSize = txtTieRodSize.Text
        End If

    End Sub

    Private Sub ColorTheForm()

        FunctionalClassObject.LabelGradient_GreenBorder_ColoringTheScreens(LabelGradient3, LabelGradient1, _
                                                            LabelGradient4, LabelGradient9)
        FunctionalClassObject.LabelGradient_OrangeBorder_ColoringTheScreens(LabelGradient5)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient2)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient6)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient8)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient7)
        FunctionalClassObject.subLabelGradient_Child_ColoringScreens(LabelGradient20)

    End Sub

    Public Sub ClearAllFielsTieRod2()

        cmbPaint.Items.Clear()
        cmbClips.Items.Clear()
        'cmbPinMaterial.Items.Clear()
        cmbPistonSealPackage.Items.Clear()
        cmbRodClevisPinClips.Items.Clear()
        cmbRodClevisPinClips.Items.Clear()
        cmbRodEndThread.Items.Clear()
        cmbRodSealPackage.Items.Clear()
        cmbRodWiper.Items.Clear()
        'cmbThreadProtected.Items.Clear()
        txtClevisCap.Text = ""

    End Sub

    Public Sub LoadingDataFromExcelTieRod2(Optional ByVal _rowno As Integer = 0)               'SUGANDHI

        LVPinSizeDetails.FlushListViewData()
        LVPinSizeDetails.SourceTable = Module1.PinSizeDetailsDataTable
        LVPinSizeDetails.Populate()

        If Module1.ReadValuesFromExcel.ClevisCapPins Then
            optPinsYes.Checked = True
            RadioBtnClevisCapPinsYes()
        Else
            optPinsNo.Checked = True
            RadioBtnClevisCapPinsNo(optPinsNo)
        End If

        If Module1.ReadValuesFromExcel.RodClevisPins Then
            optPinsYes_Rod.Checked = True
            If optPinsYes_Rod.Checked = True Then
                RodClevisPins = True        '06_04_2010   RAGAVA
                cmbPinMaterial.Enabled = True          '06_04_2010    RAGAVA
            End If
        Else
            optPinsNo_Rod.Checked = True
            RadioBtnClevisCapPinsNo(optPinsNo_Rod)
        End If

        Dim boolLVPinSizeDetails As Boolean = False
        If Not ofrmTieRod1.cmbBore.Text = "" Then

            If Not LVPinSizeDetails.Items.Count = 0 Then
                If LVPinSizeDetails.Items.Count = 1 Then
                    ListViewPinSizeDetails()
                    boolLVPinSizeDetails = True
                Else

                    For i As Integer = 0 To LVPinSizeDetails.Items.Count - 1
                        Dim pinSizeNo As Double = LVPinSizeDetails.Items(i).Text

                        If pinSizeNo = Convert.ToDouble(Module1.ReadValuesFromExcel.PinSizeDetails) Then
                            LVPinSizeDetails.Items(i).Selected = True
                            ListViewPinSizeDetails()
                            boolLVPinSizeDetails = True
                            Exit For
                        End If
                    Next
                End If

            End If
            If Not boolLVPinSizeDetails Then

                IsErrorMessageTierod2 = True

                Dim str As String
                For pinRowNo As Integer = 0 To LVPinSizeDetails.Items().Count - 1
                    If pinRowNo = LVPinSizeDetails.Items().Count - 1 Then
                        str = str + LVPinSizeDetails.Items(pinRowNo).Text
                    Else
                        str = str + LVPinSizeDetails.Items(pinRowNo).Text + ", "
                    End If
                Next

                Module1.LogInfo.Add("Row Number :" + _rowno.ToString() + " PinSize : Enter the PinSize from the following " _
                                                                        + "' " + str + " '")
            End If
        End If

        ' If optPinsNo.Checked = False Then
        If cmbPinMaterial.Enabled = True Then
            If Module1.ReadValuesFromExcel.PinMaterial = "Standard" Then
                cmbPinMaterial.Text = "Standard"
                ComboBoxPinMaterial()
            Else
                cmbPinMaterial.Text = "Hardend"
                ComboBoxPinMaterial()
            End If
        End If
        ' End If

        If cmbClips.Enabled = True Then
            cmbClips.Text = Module1.ReadValuesFromExcel.ClevisCapPinClips
            ComboBoxClips()
        End If

        If cmbThreadProtected.Enabled Then

            Dim boolcmbThreadProtected As Boolean = False
            For i As Integer = 0 To cmbThreadProtected.Items.Count - 1
                If cmbThreadProtected.Items(i).ToString() = Module1.ReadValuesFromExcel.ThreadProtected Then
                    cmbThreadProtected.Text = Module1.ReadValuesFromExcel.ThreadProtected
                    boolcmbThreadProtected = True
                    Exit For
                End If
            Next
            ' cmbThreadProtected.Text = Module1.ReadValuesFromExcel.ThreadProtected
            If Not boolcmbThreadProtected Then
                IsErrorMessageTierod2 = True
                cmbThreadProtected.Text = ""
                Dim str As String = ""
                For i As Integer = 0 To cmbThreadProtected.Items.Count - 1
                    If i = cmbThreadProtected.Items.Count - 1 Then
                        str = str + cmbThreadProtected.Items(i)
                    Else
                        str = str + cmbThreadProtected.Items(i) + ", "
                    End If
                Next
                Module1.LogInfo.Add("Row Number :" + _rowno.ToString() + " ThreadProtected : Select the ThreadProtected value from the following " + "' " + str + " '")

            End If
        End If
        If cmbThreadProtected.Text <> "" Then
            ThreadProtected = cmbThreadProtected.Text
        End If

        If cmbRodSealPackage.Enabled And cmbRodSealPackage.Items.Count <> 0 Then

            Dim boolRodSealPackage As Boolean = False
            For i As Integer = 0 To cmbRodSealPackage.Items.Count - 1

                If cmbRodSealPackage.Items(i).ToString() = Module1.ReadValuesFromExcel.RodSealPackage Then
                    cmbRodSealPackage.Text = Module1.ReadValuesFromExcel.RodSealPackage
                    ComboBoxRodSealPackage()
                    boolRodSealPackage = True
                End If
            Next

            If Not boolRodSealPackage Then
                IsErrorMessageTierod2 = True
                cmbRodSealPackage.Text = ""
                Dim str As String = ""
                For i As Integer = 0 To cmbRodSealPackage.Items.Count - 1
                    If i = cmbRodSealPackage.Items.Count - 1 Then
                        str = str + cmbRodSealPackage.Items(i)
                    Else
                        str = str + cmbRodSealPackage.Items(i) + ", "
                    End If
                Next
                Module1.LogInfo.Add("Row Number :" + _rowno.ToString() + " RodSealPackage : Select the RodSealPackage value from the following " + "' " + str + " '")
                ' MessageBox.Show("Please enter correct value")
            End If
        End If


        If Module1.ReadValuesFromExcel.RodClevisCheck Then
            rdbRodClevisYes.Checked = True
            RadioBtnRodClevisYes()
        Else
            rdbRodClevisYes.Checked = False
            rdbRodClevisNo.Checked = True
            ' RadioBtnRodClevisNo()
        End If

        If boolLVPinSizeDetails Then
            If Not rdbRodClevisNo.Checked Then

                Dim boolLVRodClevis As Boolean = False
                If Not LVRodClevis.Items.Count = 0 Then
                    If LVRodClevis.Enabled Then

                        For i As Integer = 0 To LVRodClevis.Items.Count - 1
                            If LVRodClevis.Items(i).Text = Module1.ReadValuesFromExcel.RodClevis Then
                                LVRodClevis.Items(i).Selected = True
                                ListViewRodClevis()
                                boolLVRodClevis = True
                                Exit For
                            End If
                        Next
                    ElseIf LVRodClevis.Enabled = False And LVRodClevis.Items.Count = 1 Then

                        If LVRodClevis.Items(0).Text = Module1.ReadValuesFromExcel.RodClevis Then
                            LVRodClevis.Items(0).Selected = True
                            ListViewRodClevis()
                            boolLVRodClevis = True

                        End If
                    End If
                End If
                If Not boolLVRodClevis Then
                    IsErrorMessageTierod2 = True

                    Dim str As String
                    For RodClevisNo As Integer = 0 To LVRodClevis.Items().Count - 1
                        If RodClevisNo = LVRodClevis.Items().Count - 1 Then
                            str = str + LVRodClevis.Items(RodClevisNo).Text
                        Else
                            str = str + LVRodClevis.Items(RodClevisNo).Text + ", "
                        End If
                    Next

                    Module1.LogInfo.Add("Row Number :" + _rowno.ToString() _
                                    + " RodClevis : Select the RodClevis from the following " + "' " + str + " '")
                End If
                ' ListViewRodClevis()

            End If
        End If

        If Not cmbRodEndThread.Text = "" Then
            If Not cmbRodEndThread.Items.Count = 0 AndAlso cmbRodEndThread.Items(0).ToString() <> "" Then
                Dim boolRodEndThreadSize As Boolean = False
                For i As Integer = 0 To cmbRodEndThread.Items.Count - 1
                    If cmbRodEndThread.Items(i) = Module1.ReadValuesFromExcel.RodEndThreadSize.ToString() Then
                        cmbRodEndThread.Text = Module1.ReadValuesFromExcel.RodEndThreadSize
                        ComboBoxRodEndThread()
                        boolRodEndThreadSize = True
                        Exit For
                    Else
                        Dim s As String = Convert.ToDouble(Module1.ReadValuesFromExcel.RodEndThreadSize.ToString())
                        Dim words As String() = s.Split(New Char() {"."c})

                        If words.Length = 2 Then
                            Dim word As String = words(1)
                            If Not word.Length = 2 Then
                                word = word + "0"
                                word = words(0) + "." + word
                            End If
                            If cmbRodEndThread.Items(i) = word Then
                                cmbRodEndThread.Text = word
                                ComboBoxRodEndThread()
                                boolRodEndThreadSize = True
                                Exit For
                            End If
                        ElseIf words.Length = 1 Then

                            Dim word As String = "00"
                            word = words(0) + "." + word

                            If cmbRodEndThread.Items(i) = word Then
                                cmbRodEndThread.Text = word
                                ComboBoxRodEndThread()
                                boolRodEndThreadSize = True
                                Exit For
                            End If
                        End If

                    End If
                Next
                If Not boolRodEndThreadSize Then
                    IsErrorMessageTierod2 = True
                    cmbRodEndThread.Text = ""
                    Dim str As String = ""
                    For i As Integer = 0 To cmbRodEndThread.Items.Count - 1
                        If i = cmbRodEndThread.Items.Count - 1 Then
                            str = str + cmbRodEndThread.Items(i)
                        Else
                            str = str + cmbRodEndThread.Items(i) + ", "
                        End If
                    Next
                    Module1.LogInfo.Add("Row Number :" + _rowno.ToString() _
                        + " RodEndThreadSize : Select the RodEndThreadSize value from the following " + "' " + str + " '")
                    ' MessageBox.Show("Please enter correct value")
                End If
            End If
        End If

        If cmbRodClevisPinClips.Enabled = True Then
            cmbRodClevisPinClips.Text = Module1.ReadValuesFromExcel.RodClevisPinClips
            ComboBoxRodClevisPinClips()
        End If

        If cmbPistonSealPackage.Enabled = True Then
            If cmbPistonSealPackage.Items.Count > 0 Then

                If Not cmbPistonSealPackage.Items.Count = 1 AndAlso cmbPistonSealPackage.Items(0).ToString() <> "" Then
                    Dim boolPistonSealPackage As Boolean = False
                    For i As Integer = 0 To cmbPistonSealPackage.Items.Count - 1
                        If cmbPistonSealPackage.Items(i) = Module1.ReadValuesFromExcel.PistonStealPackage.ToString() Then
                            cmbPistonSealPackage.Text = Module1.ReadValuesFromExcel.PistonStealPackage.ToString()
                            ComboBoxPistonSealPackage()
                            boolPistonSealPackage = True
                            Exit For
                        End If
                    Next
                    If Not boolPistonSealPackage Then
                        IsErrorMessageTierod2 = True
                        cmbPistonSealPackage.Text = ""
                        'MessageBox.Show("Please enter correct value")
                        IsErrorMessageTierod2 = True
                        Dim str As String = ""
                        For i As Integer = 0 To cmbPistonSealPackage.Items.Count - 1
                            If i = cmbPistonSealPackage.Items.Count - 1 Then
                                str = str + cmbPistonSealPackage.Items(i)
                            Else
                                str = str + cmbPistonSealPackage.Items(i) + ", "
                            End If
                        Next
                        Module1.LogInfo.Add("Row Number :" + _rowno.ToString() _
                                + " PistonSealPackage : Select the PistonSealPackage value from the following " + "' " + str + " '")
                        'cmbPistonSealPackage.SelectedIndex = 1
                    End If
                End If
            End If
        End If

        cmbPaint.Text = Module1.ReadValuesFromExcel.Paint

        If boolLVPinSizeDetails Then
            Dim boolRodWiper As Boolean = False
            If cmbRodWiper.Enabled Then
                If cmbRodWiper.Items.Count = 1 Then
                    If cmbRodWiper.Items(0) = Module1.ReadValuesFromExcel.RodWiper Then
                        cmbRodWiper.Text = Module1.ReadValuesFromExcel.RodWiper
                        boolRodWiper = True
                    End If
                ElseIf cmbRodWiper.Items.Count > 0 Then
                    If Not cmbRodWiper.Items.Count = 1 Then
                        For i As Integer = 0 To cmbRodWiper.Items.Count - 1

                            If cmbRodWiper.Items(i) = Module1.ReadValuesFromExcel.RodWiper Then
                                cmbRodWiper.Text = Module1.ReadValuesFromExcel.RodWiper
                                boolRodWiper = True
                            End If
                        Next
                    End If
                End If
            ElseIf cmbRodWiper.Enabled = False Then
                boolRodWiper = True
            End If
            If Not boolRodWiper Then
                IsErrorMessageTierod2 = True
                cmbRodWiper.Text = ""
                Dim str As String
                For rodWiperNo As Integer = 0 To cmbRodWiper.Items.Count - 1
                    If rodWiperNo = cmbRodWiper.Items.Count - 1 Then
                        str = str + cmbRodWiper.Items(rodWiperNo)
                    Else
                        str = str + cmbRodWiper.Items(rodWiperNo) + ", "
                    End If
                Next
                Module1.LogInfo.Add("Row Number :" + _rowno.ToString() _
                    + " Rod Wiper : Select the RodWiper from the following " + "' " + str + " '")
            End If
        End If

        If IsErrorMessageTierod2 Then

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
                ' strMsg = strMsg & ModuleGeneratedModelNames.ArrayListModelName.Item(0).ToString() + " , " + vbNewLine + ModuleGeneratedModelNames.ArrayListModelName.Item(1).ToString()
                MessageBox.Show(strMsg, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Exit Sub
            End If
        End If

    End Sub

    Private Sub optPinsNo_Rod_CheckedChanged(ByVal sender As System.Object, _
                    ByVal e As System.EventArgs) Handles optPinsNo_Rod.CheckedChanged

        RadioBtnClevisCapPinsNo(sender)

    End Sub

    
End Class