Imports System.Diagnostics
Imports System.io
Imports Microsoft.Office.Interop.Excel

Public Class ExcelClass

    Public IDEnumerator As IDictionaryEnumerator
    Public objApp As Microsoft.Office.Interop.Excel.Application
    Public objBook As Microsoft.Office.Interop.Excel.Workbook
    Public objBooks As Microsoft.Office.Interop.Excel.Workbooks
    Public objSheets As Microsoft.Office.Interop.Excel.Sheets
    Public objSheet As Microsoft.Office.Interop.Excel.Worksheet
    Public objrange As Microsoft.Office.Interop.Excel.Range
    Dim _dirObj_Directory As Directory
    Dim _strSubdirectory As String
    Dim _strFileName As String
    Dim strFileNameSave As String
    Dim intIinput As Integer = 3

    Public Sub connectToExcel()

        Try
            'objApp = New Excel.Application
            objApp = New Microsoft.Office.Interop.Excel.Application
            If objApp Is Nothing Then
                MsgBox("EXCEL 2003 is not available in the System")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub checkExcelInstance()

        Try
            If objApp Is Nothing Then
                connectToExcel()
            End If
        Catch ex As Exception
        End Try

    End Sub

    'THE BELOW CODE IS COMMON FOR UPDATING THE MAIN EXCEL SHEET AND INDIVIDUAL EXCEL SHEETS.
    Public Sub updateDesign_Parameters(ByVal strType As String)

        Dim intI As Integer = 0
        Dim intStart As Integer = 0
        Dim intEnd As Integer = 0
        Dim intJ As Integer = 0
        Try
            intI = 1
            intStart = 0

            Do While UCase(objSheet.Range("D" & intI).Value) = UCase("Start")
                intStart = intStart + 1
                intI = intI + 1
            Loop

            intEnd = 0
            intI = 1

            Do While UCase(objSheet.Range("D" & intI).Value) <> UCase("End")
                intEnd = intEnd + 1
                intI = intI + 1
            Loop

            Dim storeVal() As String
            Dim StorePos() As Integer
            ReDim Preserve storeVal(0)
            ReDim Preserve StorePos(0)
            intJ = 0
            For intI = intStart + 1 To intEnd
                If Not objSheet.Range("D" & intI).Value Is Nothing And UCase(objSheet.Range("D" & intI).Value) <> UCase("No") Then
                    ReDim Preserve storeVal(intJ)
                    ReDim Preserve StorePos(intJ)
                    storeVal(intJ) = objSheet.Range("B" & intI).Value
                    StorePos(intJ) = intI
                    intJ = intJ + 1
                End If
            Next intI

            For intI = 0 To storeVal.Length - 1
                While IDEnumerator.MoveNext()
                    If Trim(storeVal(intI)).Equals(IDEnumerator.Key) Then
                        objSheet.Range("C" & StorePos(intI)).Value() = IDEnumerator.Value()
                    End If
                End While
                IDEnumerator.Reset()
            Next intI
            objBook.Save()
            objBook.Close()
            objApp.Quit()
            objApp = Nothing
            Exit Sub
        Catch ex As Exception
        End Try

    End Sub

    Public Sub updateGUIParameters(ByVal key As String, ByVal value As String)
        'Dim intI As Integer = 3
        Try
            'objSheet.Range("B2").Value = "Input"
            'objSheet.Range("C2").Value = "Value"
            'While IDEnumerator.MoveNext()
            objSheet.Range("B" & intIinput).Value() = key
            objSheet.Range("C" & intIinput).Value() = value
            intIinput += 1
            'End While
            Try
                objSheet.Rows.AutoFit()
            Catch ex As Exception

            End Try
            objBook.Save()

            Exit Sub
        Catch ex As Exception
        End Try

    End Sub

    'The below code is for updating the design part parameters of individual excel sheets.
    Public Sub getExcelFiles(ByVal strPath As String)

        ProcessDirectory(strPath)

    End Sub

    Public Sub ProcessDirectory(ByVal targetDirectory As String)

        Try
            If objApp Is Nothing Then
                connectToExcel()
            End If
            objApp.Visible = False       'True   commented on 25-09-08
            Dim arrFileEntries As String() = Directory.GetFiles(targetDirectory)
            Dim arrFileEntriesSave As String() = Directory.GetFiles(targetDirectory)
            ' Process the list of files found in the directory.
            For Each _strFileName In arrFileEntries
                objApp.DisplayAlerts = False
                Try
                    objApp.AskToUpdateLinks = False
                    objBook = objApp.Workbooks.Open(_strFileName)
                    objBook.UpdateLinks = Microsoft.Office.Interop.Excel.XlUpdateLinks.xlUpdateLinksAlways
                    objBook.SaveLinkValues = True

                Catch ex As Exception
                    'MsgBox(ex.Message())
                    'getErrorLog(ex)
                End Try
                objBook.Save()
                'clsobjExcelVariables.objBook.Close()   'uncommented by ragava on 25-09-08
            Next
            Dim arrSubdirectoryEntries As String() = Directory.GetDirectories(targetDirectory)
            ' Recurse into subdirectories of this directory.
            For Each _strSubdirectory In arrSubdirectoryEntries
                ProcessDirectory(_strSubdirectory)
            Next _strSubdirectory
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try

    End Sub 'ProcessDirectory

End Class
