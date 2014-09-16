Imports MonarchFunctionalLayer
Public Class frmTieRod3
    Private Sub btnBackPage3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBackPage3.Click
        getBackFunctionality()
    End Sub

    Private Sub btnGenerateModel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerateModel.Click
        getNextFunctionality()
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

        End If
        Me.Hide()
    End Sub
    Public Sub getBackFunctionality()
        Me.Hide()
        ofrmTieRod2.Show()
    End Sub
End Class