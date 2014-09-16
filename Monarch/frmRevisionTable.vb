Imports System.Data.SqlClient
Imports System.Data
Imports MonarchFunctionalLayer
Imports System.io
Imports Microsoft.Win32.Registry
Imports Microsoft.Win32.RegistryKey
Imports System.Diagnostics.Process
Imports Microsoft.Office.Interop
Imports Microsoft.Win32

Public Class frmRevisionTable
    Dim Objdt As DataTable
    Dim strSql As String
    Dim intCount As Integer = 0
    Private m_SelectedStyle As DataGridViewCellStyle
    Private m_SelectedRow As Integer = -1
    Private Sub btnUpdateRevision_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateRevision.Click
        Try
            Objdt = Objdt.GetChanges(DataRowState.Modified)
            For Each dr As DataRow In Objdt.Rows
                'If Trim(dr("Description")) <> "" Then
                strSql = "Update RevisionTable Set ContractNumber = '" & dr("ContractNumber") & "' , ECR_Number = '" _
                & dr("ECR_Number") & "' , Description = '" & dr("Description") & "',RevisedBy='" & dr("RevisedBy") _
                & "', Date='" & Format(Date.Today, "dMMMyy") & "',RevisionNumber = " & dr("RevisionNumber") _
                & "Where ContractNumber = '" & dr("ContractNumber") & "' and revisionNumber=" & dr("revisionNumber")
                Dim objDT1 As DataTable = oDataClass.GetDataTable(strSql)
                'End If
            Next
            Me.Dispose()
        Catch ex As Exception
            Me.Dispose()
        End Try
      
    End Sub

    Private Sub frmRevisionTable_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            ColorTheForm()
            ' oDataClass.UpdateRevision_Details()
            Objdt = oDataClass.DisplayEmptyDescription()
            'ANUP 26-11-2010 START
            Try
                'anup 23-12-2010 start
                Dim oFrmTieRod1 As New frmTieRod1
                If IsNew_Revision_Released = "Released" OrElse (Not oFrmTieRod1.ReleasedRevisionFunctionality() _
                                                        Is Nothing AndAlso IsNew_Revision_Released = "Revision") Then
                    'anup 23-12-2010 till here
                    For Each oDataRow As DataRow In Objdt.Rows
                        'If oDataRow("ECR_Number") Is Nothing OrElse oDataRow("ECR_Number") = "" OrElse IsDBNull(oDataRow("ECR_Number")) Then
                        If IsDBNull(oDataRow("ECR_Number")) Then
                            oDataRow("ECR_Number") = GetECR_Number()
                        End If
                    Next
                End If
            Catch ex As Exception
            End Try
          
            'ANUP 26-11-2010 TILL HERE
            Objdt.Columns("ContractNumber").ReadOnly = True
            Objdt.Columns("ECR_Number").ReadOnly = False
            Objdt.Columns("Description").ReadOnly = False
            Objdt.Columns("RevisedBy").ReadOnly = False
            Objdt.Columns("Date").ReadOnly = True
            Objdt.Columns("RevisionNumber").ReadOnly = True
            dgvRevisionTable.DataSource = Objdt
            dgvRevisionTable.AllowUserToAddRows = False

            ' dgvRevisionTable.DefaultCellStyle.BackColor = Color.Blue
            'SetGridColors()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub SetGridColors()
        ' Initialize basic DataGridView properties.
        dgvRevisionTable.Dock = DockStyle.Fill
        dgvRevisionTable.BackgroundColor = Color.Black
        dgvRevisionTable.BorderStyle = BorderStyle.Fixed3D

        ' Set property values appropriate for read-only display and 
        ' limited interactivity. 
        dgvRevisionTable.AllowUserToAddRows = False
        dgvRevisionTable.AllowUserToDeleteRows = False
        dgvRevisionTable.AllowUserToOrderColumns = True
        '  dgvRevisionTable.ReadOnly = True
        'dgvRevisionTable.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        ' dgvRevisionTable.MultiSelect = False
        dgvRevisionTable.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        dgvRevisionTable.AllowUserToResizeColumns = False
        dgvRevisionTable.ColumnHeadersHeightSizeMode = _
            DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dgvRevisionTable.AllowUserToResizeRows = False
        dgvRevisionTable.RowHeadersWidthSizeMode = _
            DataGridViewRowHeadersWidthSizeMode.DisableResizing

        ' Set the selection background color for all the cells.
        dgvRevisionTable.DefaultCellStyle.SelectionBackColor = Color.White
        dgvRevisionTable.DefaultCellStyle.SelectionForeColor = Color.Black

        ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default
        ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
        dgvRevisionTable.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

        ' Set the background color for all rows and for alternating rows. 
        ' The value for alternating rows overrides the value for all rows. 
        dgvRevisionTable.RowsDefaultCellStyle.BackColor = Color.Black
        dgvRevisionTable.AlternatingRowsDefaultCellStyle.BackColor = Color.Black

        ' Set the row and column header styles.
        dgvRevisionTable.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvRevisionTable.ColumnHeadersDefaultCellStyle.BackColor = Color.Black
        dgvRevisionTable.RowHeadersDefaultCellStyle.BackColor = Color.Black

        ' Set the Format property on the "Last Prepared" column to cause
        ' the DateTime to be formatted as "Month, Year".
        'dgvRevisionTable.Columns("ECR_Number").DefaultCellStyle.Format = "y"

        ' Specify a larger font for the "Ratings" column. 
        Dim font As New Font( _
            dgvRevisionTable.DefaultCellStyle.Font.FontFamily, 10, FontStyle.Bold)
        Try
            dgvRevisionTable.Columns("ECR_Number").DefaultCellStyle.Font = font
            dgvRevisionTable.Columns("Description").DefaultCellStyle.Font = font
            dgvRevisionTable.Columns("RevisedBy").DefaultCellStyle.Font = font
        Finally
            font.Dispose()
        End Try

    End Sub

    Private Sub dgvRevisionTable_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) _
                                                Handles dgvRevisionTable.SelectionChanged
        If m_SelectedRow >= 0 Then
            dgvRevisionTable.Rows(m_SelectedRow).DefaultCellStyle = Nothing
        End If
        m_SelectedRow = dgvRevisionTable.CurrentRow.Index
        Dim dt As DataGridViewTextBoxColumn = TryCast(Me.dgvRevisionTable.Columns(3), DataGridViewTextBoxColumn)
        dt.MaxInputLength = 2
        dgvRevisionTable.CurrentRow.DefaultCellStyle = m_SelectedStyle

    End Sub

    Private Sub ColorTheForm()

        FunctionalClassObject.LabelGradient_GreenBorder_ColoringTheScreens(LabelGradient3, _
                                                LabelGradient5, LabelGradient4, LabelGradient2)
    End Sub

    'ANUP 26-11-2010 START
    Private Function GetECR_Number() As String

        Try
            Dim TempExcelApp As Excel.Application
            Dim TempWorkBook As Excel.Workbook
            Dim tempWorkSheet As Excel.Worksheet
            TempExcelApp = New Excel.Application
            TempExcelApp.Visible = False

            'anup 04-02-2011 start
            'If File.Exists("W:\ECR_TieRod\ECR_Codes.xls") Then
            '    TempWorkBook = TempExcelApp.Workbooks.Open("W:\ECR_TieRod\ECR_Codes.xls")
            If File.Exists("W:\ECR\ECR_Codes.xls") Then
                TempWorkBook = TempExcelApp.Workbooks.Open("W:\ECR\ECR_Codes.xls")
                'anup 04-02-2011 till here

                tempWorkSheet = TempExcelApp.Sheets(1)
                Dim intTotalCostExcelRange As Integer = 2
                Do
                    Try
                        If IsNothing(tempWorkSheet.Range("A" + intTotalCostExcelRange.ToString).Value) Then
                            GetECR_Number = tempWorkSheet.Range("A" + Val(intTotalCostExcelRange - 1).ToString).Value
                            GetECR_Number += tempWorkSheet.Range("B" + Val(intTotalCostExcelRange - 1).ToString).Value
                            GetECR_Number += ((tempWorkSheet.Range("C" + Val(intTotalCostExcelRange - 1).ToString).Value) + 1).ToString
                            Exit Do
                        End If
                    Catch ex As Exception
                        GetECR_Number = "11IFL-1"       'anup 04-02-2011

                    End Try
                    intTotalCostExcelRange += 1
                Loop
            Else
                GetECR_Number = "11IFL-1"      'anup 04-02-2011
            End If


        Catch ex As Exception

        End Try

    End Function
    'ANUP 26-11-2010 TILL HERE

End Class