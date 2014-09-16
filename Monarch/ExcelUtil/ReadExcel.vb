Imports System.Text.RegularExpressions
Imports System.Text

''' <summary>
''' Reads the HookUP information from the given Excel
''' </summary>
''' <remarks></remarks>
Public Class ReadExcel

    Private _strErrorMessage As StringBuilder
    Private _isError As Boolean
    Private oExcel As ExcelUtil

    Public ReadOnly Property ErrorMessage() As String
        Get
            Return _strErrorMessage.ToString
        End Get
    End Property
    Public ReadOnly Property isError() As Boolean
        Get
            Return _isError
        End Get
    End Property

    Private Sub Init()
        _isError = False
        _strErrorMessage = New StringBuilder
    End Sub

    Public Function Open(ByVal strFileName As String) As Boolean
        Init()

        oExcel = OpenSheet(strFileName)
        If oExcel Is Nothing Then
            Return False
        End If
        Return True
    End Function

    Public Function Read(ByVal CellValue As String) As String
        Read = oExcel.Read(CellValue)
    End Function

    Public Sub Close()
        Try
            oExcel.SaveAndClose()
            oExcel.Quit()
        Catch ex As Exception

        End Try
    End Sub

    Private Function OpenSheet(ByVal strFileName As String) As ExcelUtil
        Dim oExcel As New ExcelUtil

        If Not oExcel.OpenWorkBook(strFileName, False) Then
            _isError = True

            _strErrorMessage.Append(oExcel.ErrorMessage)
            Return Nothing
        End If

        If Not oExcel.OpenWorksheet("Sheet1") Then
            _isError = True
            _strErrorMessage.Append(oExcel.ErrorMessage)
            Return Nothing
        End If

        Return oExcel
    End Function


End Class
