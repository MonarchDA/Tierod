Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports MonarchFunctionalLayer

Public Class DataClass
#Region "Variables"
    '    Dim _strConnectionString As String
    '#End Region
    '#Region "Properties"
    '    Public Property strConnectionString() As String
    '        Get
    '            Return _strConnectionString
    '        End Get
    '        Set(ByVal value As String)
    '            _strConnectionString = ConfigurationSettings.AppSettings.Item("ConnectionString").ToString()
    '        End Set
    '    End Property
    '#End Region
    '#Region "Constructor"
    Public Sub New()
    End Sub
    'Public Sub New(ByVal dbConnectionString As String)
    '    _strConnectionString = dbConnectionString
    'End Sub
#End Region
#Region "Functions"
    Public Function GetTableData(ByVal _strTableName As String) As DataTable

        Dim _dt As New DataTable
        'Dim _da As New SqlDataAdapter("Select * from " & _strTableName, strConnectionString)
        Dim _da As New SqlDataAdapter("Select * from " & _strTableName, IFLConnectionObject.ConnectionString)
        Try
            _da.Fill(_dt)
        Catch ex As Exception
            MsgBox("Error in Filling DataTable " & _strTableName)
        End Try
        Return _dt

    End Function

    Public Function GetDataTable(ByVal _strQueryString As String) As DataTable
        Dim _dt As New DataTable
        'Dim _da As New SqlDataAdapter(_strQueryString, strConnectionString)
        Dim _da As New SqlDataAdapter(_strQueryString, IFLConnectionObject.ConnectionString)
        Dim _sqlCommand As New SqlCommand()
        Try
            _da.Fill(_dt)
        Catch ex As Exception
            MsgBox("Error in Processing Below Query " & Environment.NewLine & _strQueryString)
        End Try
        Return _dt
    End Function

    'anup 23-12-2010 start
    Public Sub UpdateRevision_Details(ByVal strContractNumber As String, ByVal revisionNumber As Integer, Optional ByVal IsReleased As String = "")
        Try
            Dim StrSql As String
            Dim objDT As DataTable
            StrSql = "Delete from RevisionTable where ContractNumber='" & strContractNumber & "' and RevisionNumber=" & revisionNumber
            objDT = GetDataTable(StrSql)
            objDT.Clear()
            If IsReleased = "Release" Then
                StrSql = "Insert into RevisionTable(ContractNumber,Description,Date,RevisionNumber) values('" & strContractNumber & "','" & IsReleased & "','" & Format(Date.Today, "dMMMyy") & "'," & revisionNumber & ")"
            Else
                StrSql = "Insert into RevisionTable(ContractNumber,Date,RevisionNumber) values('" & strContractNumber & "','" & Format(Date.Today, "dMMMyy") & "'," & revisionNumber & ")"
            End If
            objDT = GetDataTable(StrSql)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    'anup 23-12-2010 till here

    Public Function DisplayEmptyDescription() As DataTable
        DisplayEmptyDescription = Nothing
        Try
            Dim StrSql As String
            Dim objDT As DataTable
            StrSql = "Select top 7 ContractNumber,ECR_Number,Description,RevisedBy,Date,RevisionNumber from RevisionTable where ContractNumber = '" & ContractNumber & "' order by RevisionNumber Desc" 'and Description is Null  or CompiledBy is Null or ApprovedBy is Null"
            objDT = GetDataTable(StrSql)
            Return objDT
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    

#End Region
End Class
