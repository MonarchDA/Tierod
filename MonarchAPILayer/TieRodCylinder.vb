Imports MonarchSolidworksLayer
Imports MonarchFunctionalLayer
Imports System.IO


Public Class TieRodCylinder

    Public Sub PerformTieRodCylinderFunctionalities(ByVal CylinderCodeNumber As String)
        Try
            updateTieRodCylinderDesignTables()
            oExcelClass.getExcelFiles("C:\DESIGN_TABLES")
            Try
                'IFLSolidWorksBaseClassObject.SetCurrentWorkingDirectory(DestinationFilePath)
                'IFLSolidWorksBaseClassObject.openDocument(DestinationFilePath + "\Bore")
                ProcessDirectory(DestinationFilePath + "\Bore")
                '02_09_2009   ragava
                ProcessDirectory(DestinationFilePath + "\ROD")

                ProcessDirectory(DestinationFilePath + "\Tie Rod")
                'If Condition Need to be Added
                ProcessDirectory(DestinationFilePath + "\Stoptube")
                'Opening Main ASSY
                Try
                    KillExcel()
                    oExcelClass.objApp = Nothing       '06_10_2009  ragava
                Catch ex As Exception

                End Try
                ProcessDirectory(DestinationFilePath + "\TIE_ROD_ASSEMBLY")

                '13_04_2010    RAGAVA
                Try
                    IFLSolidWorksBaseClassObject.ConnectSolidWorks()
                    System.Threading.Thread.Sleep(1000)
                    'IFLSolidWorksBaseClassObject.SolidWorksApplicationObject.CommandInProgress = False
                Catch ex As Exception
                End Try
                Dim blnRet As Boolean = IFLSolidWorksBaseClassObject.SolidWorksApplicationObject. _
                            SetCurrentWorkingDirectory(DestinationFilePath)
                'TILL    HERE

                '15_09_2009   ragava
                Try
                    DrawingUpdation()
                Catch ex As Exception
                    MsgBox("Error in Updating Drawings")
                End Try
                '22_09_2009   ragava
                Try
                    UpdateTableDrawing()
                Catch ex As Exception

                End Try
                'Dim strMsg = "Model generated successfully" + vbNewLine
                'strMsg = strMsg & "Location--------" & DestinationFilePath
                '01_10_2009    ragava
                'Try
                '    Directory.Delete("C:\MONARCH_TESTING", True)
                'Catch ex As Exception
                '    MsgBox(ex.Message)
                'End Try

                '' Sugandhi
                'Dim strMsg = "Model generated successfully" + vbNewLine
                'strMsg = strMsg & "Assembly Drawing is Located at W:\" & PartCode.ToString() & ".SLDDRW"
                'MessageBox.Show(strMsg, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                ''     sugandhi

                If IsGenerateBtnClicked Then
                    Dim strMsg = "Model generated successfully" + vbNewLine
                    strMsg = strMsg & "Assembly Drawing is Located at W:\" & PartCode.ToString() & ".SLDDRW"
                    MessageBox.Show(strMsg, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information, _
                                MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)

                Else
                    Dim strMsg As String = "Assembly Drawing is Located at W:\" & PartCode.ToString() & ".SLDDRW"
                    ModuleGeneratedModelNames.ArrayListModelName.Add(strMsg)
                End If

            Catch oException As Exception
                ErrorMessage = oException.Message
                ErrorMessage += "Error in updating Tie Rod Cylinder"
            End Try
        Catch ex As Exception
        End Try

    End Sub

    Public Sub updateTieRodCylinderDesignTables()

        Try
            oExcelClass.checkExcelInstance()
            oExcelClass.objBook = oExcelClass.objApp.Workbooks.Open("C:\DESIGN_TABLES\GUI_PARAMETERS.xls")
            oExcelClass.objSheet = oExcelClass.objBook.Worksheets("Sheet1")
            oExcelClass.updateDesign_Parameters("Sheet1")
        Catch ex As Exception
        End Try

    End Sub
End Class