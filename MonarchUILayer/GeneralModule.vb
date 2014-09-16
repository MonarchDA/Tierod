Imports MonarchDatabaseLayer

Public Module GeneralModule
    Public ofrmContractDetails As New frmContractDetails
    Public ofrmTieRod1 As New frmTieRod1
    Public ofrmTieRod2 As New frmTieRod2
    Public ofrmTieRod3 As New frmTieRod3
    Public oclsimgcapture As New clsimgcapture
    Public oDataClass As New DataClass
    Public Sub cancelClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If MessageBox.Show("Do you really want to cancel this project run", "Confirmation for Cancel", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = Windows.Forms.DialogResult.OK Then
           
        End If
    End Sub
End Module
