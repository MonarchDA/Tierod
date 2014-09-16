Imports System.io
Imports System.Drawing.Imaging
Imports System.Threading
Imports System
Imports MonarchFunctionalLayer

Public Module MILFunctionalModule
    Dim fso As New Scripting.FileSystemObject

    Public Sub captureImages(ByVal formName As Form)
        oclsimgcapture.CaptureScreen()
        Dim pic As PictureBox
        Try
            pic = New PictureBox
            pic.Image = oclsimgcapture.Background
            pic.Name = formName.Name
            Dim alist As New ArrayList
            If alist.Contains(formName.Name) Then
                alist.Remove(formName.Name)
            End If
            alist.Add(New Object(1) {formName.Name, pic.Name})
            If fso.FolderExists(Application.StartupPath + "\InputImages\") = False Then
                fso.CreateFolder(Application.StartupPath + "\InputImages\")
            End If
            pic.Image.Save(Application.StartupPath + "\InputImages\" + formName.Name + ".jpg", ImageFormat.Jpeg)
        Catch oException As Exception
            MessageBox.Show(oException.Message)
        End Try
    End Sub
   
    Public Function GetDataToSave(ByVal oSaveButton As Object) As Byte()
        Dim oObject As Object = GetSaveFormList(CType(oSaveButton, Button), GetSaveButtonForms)
        Dim oGetSetUIClass As New IFLGetSetUI.IFLGetSetUIClass
        Dim oDataSet As New DataSet("MILSaveData")
        For Each oForm As Form In oObject
            Dim oTable As DataTable = oGetSetUIClass.StoreFormData(oForm)
            oDataSet.Tables.Add(oTable)
        Next
        oDataSet.WriteXml(Execution_Path1 + "\MIL.xml")
        GetDataToSave = GetByteArray(oDataSet)
    End Function

    Private ReadOnly Property GetContractDetails() As Object
        Get
            Return New Object(0) {frmContractDetails}
        End Get
    End Property

    Private ReadOnly Property GetTieRod1() As Object
        Get
            Return New Object(1) {frmContractDetails, frmTieRod1}
        End Get
    End Property

    Private ReadOnly Property GetTieRod2() As Object
        Get
            Return New Object(2) {frmContractDetails, frmTieRod1, frmTieRod2}
        End Get
    End Property

    Public ReadOnly Property GetTieRod3() As Object
        Get
            Return New Object(3) {ofrmContractDetails, ofrmTieRod1, ofrmTieRod2, ofrmTieRod3}
        End Get
    End Property

    Private ReadOnly Property GetSaveButtonForms() As ArrayList
        Get
            Dim aReturnData As New ArrayList
            aReturnData.Add(New Object(1) {"btnGenerate", GetTieRod3()})
            Return aReturnData
        End Get
    End Property

#Region "Functions"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="oSaveButton"></param>
    ''' <param name="aSaveFormList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSaveFormList(ByVal oSaveButton As Button, ByVal aSaveFormList As ArrayList) As Object
        Dim strButtonName As String = oSaveButton.Name
        Dim oData As Object
        Dim oReturnData As Object = Nothing

        For Each oData In aSaveFormList
            If strButtonName.ToUpper.Equals(oData(0).ToUpper) Then
                oReturnData = oData(1)
            End If
        Next
        Return oReturnData
    End Function

    ''' <summary>
    ''' Gets the byte array data.
    ''' </summary>
    ''' <param name="oDataSet"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetByteArray(ByVal oDataSet As DataSet) As Byte()
        oDataSet.WriteXml(Execution_Path1 + "\MIL.xml")
        Dim fsBLOBFile As New System.IO.FileStream(Execution_Path1 + "\MIL.xml", IO.FileMode.Open)
        Dim bytBLOBData(fsBLOBFile.Length() - 1) As Byte
        fsBLOBFile.Read(bytBLOBData, 0, bytBLOBData.Length)
        fsBLOBFile.Close()
        Return bytBLOBData
    End Function


#End Region

#Region "Enums"
    Private Enum Parameters
        Name
        Value
    End Enum
#End Region

End Module

