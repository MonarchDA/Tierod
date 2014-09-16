Imports System.Diagnostics
Imports System.io
Imports Microsoft.Office.Interop.Excel

''' <summary>
''' ExcelUtil Class is the utility class for excel operations.
''' </summary>
''' <remarks></remarks>
Public Class ExcelUtil

    Private _objApp As Microsoft.Office.Interop.Excel.Application
    Private _objBook As Microsoft.Office.Interop.Excel.Workbook
    Private _objBooks As Microsoft.Office.Interop.Excel.Workbooks
    Private _objSheets As Microsoft.Office.Interop.Excel.Worksheets
    Private _objSheet As Microsoft.Office.Interop.Excel.Worksheet
    Private _objrange As Microsoft.Office.Interop.Excel.Range

    Private _strErrorMessage As String
    Private _blnError As Boolean

    Public ReadOnly Property ErrorMessage() As String
        Get
            Return _strErrorMessage
        End Get
    End Property

    Public Sub New()
        connectToExcel()
    End Sub

    Public Sub connectToExcel()
        Try
            'objApp = New Excel.Application
            '   objApp = New Microsoft.Office.Interop.Excel.Application
            _objApp = CreateObject("Excel.Application")
            If _objApp Is Nothing Then
                MsgBox("EXCEL 2003 is not available in the System")
            End If
        Catch ex As Exception
            _strErrorMessage = String.Empty
        End Try
    End Sub

    Public Function OpenWorkBook(ByVal strFileName As String) As Boolean
        Return OpenWorkBook(strFileName, True)
    End Function

    Public Function OpenWorkBook(ByVal strFileName As String, ByVal isActive As Boolean) As Boolean
        Try
            _objBook = _objApp.Workbooks.Open(strFileName, isActive)
            Return True
        Catch ex As Exception
            _strErrorMessage = "Unable to OpenWorkBook for given file."
            Return False
        End Try
    End Function

    Public Function CreateWorkSheet(ByVal strSheetName As String, ByVal isReplace As Boolean) As Boolean
        Try
            'Dim intCount As Integer
            If isReplace Then
                If OpenWorksheet(strSheetName) Then
                    _objSheet.Cells.Clear()
                    _objSheet.Activate()
                    'intCount = _objBook.Worksheets.Count
                    '_objSheet.Move(_objBook.Worksheets(intCount))
                    'OpenWorksheet(strSheetName)
                    Return True
                End If
            End If
            'intCount = _objBook.Worksheets.Count
            _objSheet = _objBook.Worksheets.Add
            _objSheet.Name = strSheetName
            _objSheet.Activate()
            'intCount = _objBook.Worksheets.Count
            '_objSheet.Move(_objBook.Worksheets(intCount))
            'OpenWorksheet(strSheetName)

            Return True
        Catch ex As Exception
            _strErrorMessage = "Unable to Create Sheet for given Workbook."
            Return False
        End Try
    End Function

    Public Function ClearAll(ByVal strSheetName As String) As Boolean
        Try
            _objSheet = _objBook.Worksheets(strSheetName)
            _objSheet.Activate()
            _objSheet.Cells.Clear()
            Return True
        Catch ex As Exception
            _strErrorMessage = "Unable to Open Sheet for given Workbook."
            Return False
        End Try
    End Function

    Public Function OpenWorksheet(ByVal strSheetName As String) As Boolean
        Try
            _objSheet = _objBook.Worksheets(strSheetName)
            _objSheet.Activate()
            Return True
        Catch ex As Exception
            _strErrorMessage = "Unable to Open Sheet for given Workbook."
            Return False
        End Try
    End Function

    Public Function GetExcelColumn(ByVal index As Integer) As String

        Dim quotient As Integer = index \ 26 ''//Truncate  
        If quotient > 0 Then
            Return GetExcelColumn(quotient - 1) & Chr((index Mod 26) + 65).ToString

        Else
            Return Chr(index + 65).ToString
        End If
    End Function

    Public Sub Write(ByVal key As String, ByVal value As String, ByVal cellParameterValue As String, ByVal cellValue As String)

        _objSheet.Range(cellValue).Value = value

    End Sub

    Public Sub Write(ByVal rowValue As Integer, ByVal colValue As Integer, ByVal Value As String)
        Dim cellValue As String
        cellValue = GetExcelColumn(colValue) & rowValue
        _objSheet.Range(cellValue).Value = Value
    End Sub

    Public Sub Write(ByVal cellValue As String, ByVal value As String)

        _objSheet.Range(cellValue).Value = value
    End Sub

    Public Function Read(ByVal CellValue As String) As Object
        Return _objSheet.Range(CellValue).Value
    End Function

    Public Function Read(ByVal rowValue As Integer, ByVal colValue As Integer) As Object
        Dim cellValue As String

        cellValue = GetExcelColumn(colValue) & rowValue
        Return _objSheet.Range(CellValue).Value
    End Function

    Public Function Save() As Boolean
        Try
            _objBook.Save()
            Return True
        Catch ex As Exception
            _strErrorMessage = "Unable to save the book"
            Return False
        End Try
    End Function

    Public Function Close() As Boolean
        Try
            _objBook.Close()
            Return True
        Catch ex As Exception
            _strErrorMessage = "Unable to Close the book."""
            Return False
        End Try
    End Function

    Public Function SaveAndClose() As Boolean
        Return (Save() AndAlso Close())
    End Function

    Public Function Quit() As Boolean
        Try
            _objApp.Quit()
            _objApp = Nothing
            Return True
        Catch ex As Exception
            _strErrorMessage = "Unable to Quit the Excel Application"
            Return False
        End Try
    End Function
End Class
