Imports MonarchFunctionalLayer
Public Class frmContractDetails

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
    '    getNextFunctionality()
    'End Sub
    Public Sub getNextFunctionality()
        If Not FunctionalClassObject.validateForm(Me) Is Nothing Then
            MessageBox.Show(FunctionalClassObject.ErrorMessage)
            FunctionalClassObject.validateForm(Me).Focus()
        Else
            CustomerName = txtCustomerName.Text
            ContractNumber = txtContractNumber.Text
            AssemblyType = cmbAssemblyType.SelectedItem
            PartCode = txtlPartCode.Text
            Try
                FunctionalClassObject.PopulateFormscontrolsData(Me)
            Catch oException As Exception

            End Try
            captureImages(Me)
            Me.Hide()
            ofrmTieRod1.Show()
        End If

    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        getNextFunctionality()
    End Sub

    Private Sub frmContractDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class