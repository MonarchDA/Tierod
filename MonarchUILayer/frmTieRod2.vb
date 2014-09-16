Imports MonarchFunctionalLayer
Public Class frmTieRod2

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        getBackFunctionality()
    End Sub
    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
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
            ofrmTieRod3.Show()
        End If
        Me.Hide()
    End Sub
    Public Sub getBackFunctionality()
        Me.Hide()
        ofrmTieRod1.Show()
    End Sub
End Class