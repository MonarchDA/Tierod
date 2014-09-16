Imports MonarchFunctionalLayer
Public Class frmTieRod1

    Private Sub TieRod1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AddHandler btnCancel.Click, AddressOf cancelClick
    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub getNextFunctionality()
        If Not FunctionalClassObject.validateForm(Me) Is Nothing Then
            MessageBox.Show(FunctionalClassObject.ErrorMessage)
            FunctionalClassObject.validateForm(Me).Focus()
        Else
            Try
                FunctionalClassObject.PopulateFormscontrolsData(Me)
            Catch oException As Exception
            End Try
            captureImages(Me)
            Me.Hide()
            ofrmTieRod2.Show()
        End If
    End Sub

    Public Sub getBackFunctionality()

    End Sub


    Private Sub btnNextPage1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNextPage1.Click
        getNextFunctionality()
    End Sub

    Private Sub cmbPort_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPort.SelectedIndexChanged

    End Sub

    Private Sub cmbStyle_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbStyle.SelectedIndexChanged
        If sender.selecteditem = "ASAE" Then
            txtStrokeLength.Visible = False
            txtStrokeLength.Enabled = False
            cmbStrokeLength.Visible = True
            cmbStrokeLength.Enabled = True
            cmbStrokeLength.Items.Clear()
            cmbStrokeLength.Items.Add("")
            cmbStrokeLength.Items.Add("8")
            cmbStrokeLength.Items.Add("16")
        Else
            txtStrokeLength.Visible = True
            txtStrokeLength.Enabled = True
            cmbStrokeLength.Visible = False
            cmbStrokeLength.Enabled = False
        End If
    End Sub

    Private Sub cmbSeries_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSeries.SelectedIndexChanged
        If sender.selectedItem.startsWith("TP") Then
            cmbRephasingPortPosition.Enabled = True
        Else
            cmbRephasingPortPosition.Enabled = False
        End If
        If sender.selectedItem.startsWith("TX") Then
            LVNutSizeDetails.Enabled = False
        Else
            LVNutSizeDetails.Enabled = True
        End If
        loadBoreDiameterValues(sender.selectedItem)
    End Sub

    Private Sub cmbRodMaterial_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbRodMaterial.SelectedIndexChanged
        Try
            Dim strQuery As String = ""
            Dim aColumns As New ArrayList
            strQuery = "select distinct rd.RodDiameter from RodDiameterDetails rd,BoreDiameter_RodDiameter bdrd where bdrd.BoreDiameterID=(select BoreDiameterID from BoreDiameterMaster where BoreDiameter = " & Val(cmbBore.SelectedItem) & ") and bdrd.PartNumberID = rd.PartNumber and Series = '" & IIf(Trim(cmbSeries.SelectedItem.ToString).StartsWith("TX"), "TX", "TL/TH/TP") & "' and IsASAE = '" & cmbStyle.SelectedItem & "' and MaterialType='" & cmbRodMaterial.SelectedItem & "'"
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            LVRodDiameterDetails.Columns.Clear()
            aColumns.Add(New Object(2) {"RodDiameter", "Rod Diameter", True})
            'aColumns.Add(New Object(2) {"DeratePressure", "Derate Pressure", True})
            LVRodDiameterDetails.DisplayHeaders = aColumns
            LVRodDiameterDetails.FullRowSelect = True
            objDT.Columns.Add("Derate Pressure")
            'objDT.Columns.Item(1).AutoIncrement = True
            LVRodDiameterDetails.SourceTable = objDT
            LVRodDiameterDetails.Populate()
            LVRodDiameterDetails.Columns.Add("Derate Pressure")
        Catch ex As Exception
           
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
        ElseIf SeriesType.StartsWith("TL") OrElse SeriesType.StartsWith("TH") Then
            cmbBore.Items.Add("2")
            cmbBore.Items.Add("2.5")
            cmbBore.Items.Add("3")
            cmbBore.Items.Add("3.5")
            cmbBore.Items.Add("4")
            cmbBore.Items.Add("4.5")
            cmbBore.Items.Add("5")
        ElseIf SeriesType = "TP Low" Then
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
        Else : SeriesType = "TP High"
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

    End Sub

    Private Sub cmbPinHole_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPinHole.SelectedIndexChanged
        Try
            cmbPort.Items.Clear()
            cmbPort.Items.Add(" ")
            Dim strQuery As String = ""
            Dim aColumns As New ArrayList
            strQuery = "select Port from ClevisCapDetails where Series=" + vbNewLine
            strQuery = strQuery + "'" + IIf(Trim(cmbSeries.SelectedItem.ToString).StartsWith("TX"), "TX", "TL/TH/TP") + "'" + vbNewLine
            strQuery = strQuery + " and BoreDiameter=" + cmbBore.SelectedItem + vbNewLine
            strQuery = strQuery + " And PinHoleType='" + cmbPinHole.SelectedItem + "'"
            Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
            cmbPort.DisplayMember = "port"
            cmbPort.DataSource = objDT
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cmbBore_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbBore.SelectedIndexChanged
        Try

            If Not sender.selectedItem.startsWith("TX") AndAlso Trim(sender.selectedItem) <> "" Then
                Dim strQuery As String = ""
                strQuery = "Select distinct PistonNutSize from PistonSealDetails where BoreDiameter=" + cmbBore.SelectedItem
                Dim aColumns As New ArrayList
                Dim objDT As DataTable = oDataClass.GetDataTable(strQuery)
                LVNutSizeDetails.Columns.Clear()
                aColumns.Add(New Object(2) {"PistonNutSize", "Nut Size", True})
                LVNutSizeDetails.DisplayHeaders = aColumns
                LVNutSizeDetails.FullRowSelect = True
                LVNutSizeDetails.SourceTable = objDT
                LVNutSizeDetails.Populate()
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class