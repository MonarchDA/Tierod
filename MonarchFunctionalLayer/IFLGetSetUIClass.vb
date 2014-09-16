Imports System.IO
Imports System.Data.SqlClient
Namespace IFLGetSetUI
    ''' -----------------------------------------------------------------------------
    ''' Project	 : IFLGetSetUIProject
    ''' Class	 : IFLGetSetUI.IFLGetSetUIClass
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' It will store and retrieve the respective form data.
    ''' </summary>
    ''' This class is used to deal with the getting and setting the UI.
    ''' <remarks>
    ''' This Class IFLGetSetUIClass is used to set the form values into database and retrive them when required.
    ''' </remarks>
    ''' <history>
    ''' 	[swamy.reddy]	8/4/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class IFLGetSetUIClass

#Region "Class Variables"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' This variable is used to get the Datatable from the database.
        ''' </summary>
        ''' <remarks>
        ''' Using this _oTable we can store the current table name..
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/5/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private _oTable As DataTable

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''  This variable is used to get the form.
        ''' </summary>
        ''' <remarks>
        ''' This is used to assign the current form.
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/5/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private _SourceForm As Form

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''  This variable is used to get the form.
        ''' </summary>
        ''' <remarks>
        '''  The below variable is used to create a row.
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/5/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private odataRow As DataRow

        Private oDataRowView As DataRowView

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''  The below variable is used to give the Error message when exception raises.
        ''' </summary>
        ''' <remarks>
        '''  The below variable is used to give the Error message when exception raises.
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/5/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private _strErrorMessage As String

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' The below variable is used to give the Error object when exception raises.
        ''' </summary>
        ''' <remarks>
        ''' The below variable is used to give the Error object when exception raises.
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/5/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private _oErrorObject As Exception

#End Region

#Region "Properties"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' This property is used to store the current form.
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/4/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property SourceForm()
            Get
                Return _SourceForm
            End Get
            Set(ByVal Value)
                _SourceForm = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' This property is used as an arraylist to store the specified fields.
        ''' </summary>
        ''' <value>Arraylist is the datatype, used to store values.</value>
        ''' <remarks>
        ''' This is used to store the items.
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/4/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private ReadOnly Property aTableStructure() As ArrayList
            Get
                aTableStructure = New ArrayList
                aTableStructure.Add("TabIndex")
                aTableStructure.Add("ConName")
                aTableStructure.Add("Value")
            End Get
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' This property is used to assign the error message.
        ''' </summary>
        ''' <value>
        ''' _strErrorMessage which stores the Error message its an exception type.
        ''' </value>
        ''' <remarks>
        ''' This property is used to store the error message and displays that message to the user.
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/5/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property ErrorMessage()
            Get
                Return _strErrorMessage
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' This property is used to assign the error object.
        ''' </summary>
        ''' <value>
        ''' Exception type.
        ''' </value>
        ''' <remarks>
        ''' Used to store the object which will be catch in the catch method.
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/5/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property ErrorObject()
            Get
                Return _oErrorObject
            End Get
        End Property

#End Region

#Region "Procedures"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' This procedure is used to create a source table 
        ''' </summary>
        ''' <param name="strFormName">Form name as a string.</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' It will return a table object that can be created with specified structure which contains the tab index, control name and control value
        ''' </remarks>
        ''' <history>
        ''' 	[swamy.reddy]	8/4/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Function CreateSourceTable(ByVal strFormName As String) As Boolean
            CreateSourceTable = True
            Try
                _oTable = New DataTable(strFormName)
                _oTable.Columns.Add("TabIndex", System.Type.GetType("System.Int32"))
                _oTable.Columns.Add("ConName", System.Type.GetType("System.String"))
                _oTable.Columns.Add("Value", System.Type.GetType("System.String"))
            Catch oException As Exception
                CreateSourceTable = False
                _strErrorMessage = "Source Table is not Created !!" + vbCrLf
                _strErrorMessage += "System generated error:-" + vbCrLf + oException.Message
                _oErrorObject = oException
            End Try
           
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' This procedure is used to add the Control information into the Table
        ''' </summary>
        ''' <param name="ocontrol">ocontrol is the control of the respective form.</param>
        ''' <remarks>
        ''' It will check the controls and with tabindex it will store the respective control value into the database.
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/4/2008	Created
        '''     [Pradeep Sen]	    8/8/2008	Modified : Cause given in the line of change
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Sub AddDataToTable(ByVal oControl As Control)
            Dim oCheckBox As CheckBox
            Dim oRadioButton As RadioButton
            Dim olistView As IFLCustomUILayer.IFLListView

            Try
                'If ocontrol.Visible = True And ocontrol.Enabled = True Then
                ' Visible is not a valid option to check, as cases will be there if the form which
                ' needs to be fetched back is not visible then this fails
                ' TODO : Needs a new function to check for the valid control types to be taken
                'If oControl.Enabled = True Then

                'If TypeOf (oControl) Is TextBox Or TypeOf (oControl) Is ComboBox Then
                If TypeOf (oControl) Is TextBox Or TypeOf (oControl) Is ComboBox Or TypeOf (oControl) Is RichTextBox Then            '04_11_2009    Ragava
                    If Not oControl.Text = "" Then
                        odataRow = _oTable.NewRow
                        odataRow.Item("TabIndex") = oControl.TabIndex
                        odataRow.Item("ConName") = oControl.Name
                        odataRow.Item("Value") = oControl.Text
                        _oTable.Rows.Add(odataRow)
                    End If
                ElseIf TypeOf (oControl) Is CheckBox Then
                    oCheckBox = CType(oControl, CheckBox)
                    odataRow = _oTable.NewRow
                    odataRow.Item("TabIndex") = oCheckBox.TabIndex
                    odataRow.Item("ConName") = oCheckBox.Name
                    odataRow.Item("Value") = oCheckBox.Checked
                    _oTable.Rows.Add(odataRow)
                ElseIf TypeOf (oControl) Is RadioButton Then
                    oRadioButton = CType(oControl, RadioButton)
                    odataRow = _oTable.NewRow
                    odataRow.Item("TabIndex") = oRadioButton.TabIndex
                    odataRow.Item("ConName") = oRadioButton.Name
                    odataRow.Item("Value") = oRadioButton.Checked
                    _oTable.Rows.Add(odataRow)
                ElseIf TypeOf (oControl) Is ListView Then
                    olistView = CType(oControl, IFLCustomUILayer.IFLListView)
                    odataRow = _oTable.NewRow
                    odataRow.Item("TabIndex") = olistView.TabIndex
                    odataRow.Item("ConName") = olistView.Name
                    Dim strLVitem As String = ""
                    'For Each itemI As ListViewItem In olistView.SelectedItems
                    '    strLVitem = strLVitem + "-" + itemI.SubItems(0).Text
                    'Next
                    Dim itemI As ListViewItem = olistView.SelectedItems(0)
                    strLVitem = itemI.Text
                    For i As Integer = 1 To itemI.SubItems.Count - 1
                        strLVitem += "-"
                        strLVitem += itemI.SubItems(i).Text
                    Next
                    odataRow.Item("Value") = strLVitem
                    _oTable.Rows.Add(odataRow)
                End If
                ' End If
            Catch oException As Exception
                _strErrorMessage = "UNABLE TO ADD THE DATA TO THE TABLE !!" + vbCrLf
                _strErrorMessage += "System generated error:-" + vbCrLf + oException.Message
                _oErrorObject = oException
            End Try
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' This procedure is used to find  the control.
        ''' </summary>
        ''' <param name="oFormControl">oFormControl is the control of the respective form.</param>
        ''' <remarks>
        ''' This procedure will check the whether the control is radiobutton, checkbox, textbox, like this it will check the form controls.
        ''' 
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/5/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Sub FindControl(ByVal oFormControl As Control)
            Try
                For Each ocontrol As Control In oFormControl.Controls
                    If TypeOf (ocontrol) Is GroupBox Or TypeOf (ocontrol) Is Panel Or TypeOf (ocontrol) Is TabControl Then
                        FindControl(ocontrol)
                    Else
                        SetControlValue(ocontrol)
                    End If
                Next
            Catch oException As Exception
                _strErrorMessage = "UNABLE TO FIND THE CONTROL !!" + vbCrLf
                _strErrorMessage += "System generated error:-" + vbCrLf + oException.Message
                _oErrorObject = oException
            End Try
            'Try
            '    Dim conCount As Integer = 0
            '    Dim oControl As Control = Nothing
            '    conCount = oFormControl.Controls.Count
            '    Dim intCount As Integer = 0

            '    For intCount = 1 To conCount

            '        oControl = oFormControl.Controls.Item(conCount - intCount)
            '        If TypeOf (oControl) Is GroupBox Or TypeOf (oControl) Is Panel Or TypeOf (oControl) Is TabControl Then
            '            FindControl(oControl)
            '        Else
            '            SetControlValue(oControl)
            '        End If
            '    Next
            'Catch oException As Exception
            '    _strErrorMessage = "UNABLE TO FIND THE CONTROL !!" + vbCrLf
            '    _strErrorMessage += "System generated error:-" + vbCrLf + oException.Message
            '    _oErrorObject = oException
            'End Try

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''  This procedure is used to set the control value.
        ''' </summary>
        ''' <param name="ocontrol">ocontrol is the control of the respective form.</param>
        ''' <remarks>
        ''' This will find the control of the form , and it will store the control value .
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/5/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Sub SetControlValue(ByVal ocontrol As Control)
            Dim oCheckBox As CheckBox
            Dim oRadioButton As RadioButton
            Dim oComboBox As IFLCustomUILayer.IFLComboBox
            Dim oListView As IFLCustomUILayer.IFLListView
            Try
                If Not ocontrol.Name = "txtRetractedLength" Then
                    If ocontrol.Name = oDataRowView.Item("ConName") Then
                        'If TypeOf (ocontrol) Is TextBox Or TypeOf (ocontrol) Is ComboBox Then
                        '    ocontrol.Text = odataRow.Item("Value")
                        If TypeOf (ocontrol) Is TextBox Then
                            ocontrol.Text = oDataRowView.Item("Value")
                        ElseIf TypeOf (ocontrol) Is ComboBox Then
                            oComboBox = CType(ocontrol, IFLCustomUILayer.IFLComboBox)
                            'oComboBox.Items.Add(oDataRowView.Item("Value"))
                            oComboBox.Text = oDataRowView.Item("Value")
                        ElseIf TypeOf (ocontrol) Is CheckBox Then
                            oCheckBox = CType(ocontrol, CheckBox)
                            oCheckBox.Checked = oDataRowView.Item("Value")
                        ElseIf TypeOf (ocontrol) Is RadioButton Then
                            oRadioButton = CType(ocontrol, RadioButton)
                            oRadioButton.Checked = oDataRowView.Item("Value")
                        ElseIf TypeOf (ocontrol) Is ListView Then
                            oListView = CType(ocontrol, IFLCustomUILayer.IFLListView)
                            Dim oListviewItem As ListViewItem
                            Dim items() As String
                            items = Split(oDataRowView.Item("Value"), "-")
                            If ocontrol.Name = "LVRodClevis" Then
                                oListviewItem = oListView.FindItemWithText(ListViewString(items))
                            Else
                                oListviewItem = oListView.FindItemWithText(items(0))
                            End If

                            Dim Position As Integer = oListView.Items.IndexOf(oListviewItem)
                            If Position >= 0 Then
                                oListView.Items(Position).Selected = True
                            End If

                            '04_11_2009   Ragava
                        ElseIf TypeOf (ocontrol) Is RichTextBox Then
                            ocontrol.Text = oDataRowView.Item("Value")
                            '04_11_2009   Ragava   Till   Here
                        End If
                        Exit Sub
                    End If
                End If
            Catch oException As Exception
                _strErrorMessage = "UNABLE TO SET THE CONTROL VALUE !!" + vbCrLf
                _strErrorMessage += "System generated error:-" + vbCrLf + oException.Message
                _oErrorObject = oException
            End Try

        End Sub

        Private Function ListViewString(ByVal items() As String) As String
            ListViewString = String.Empty
            Try
                Dim intItemsCount As Integer = items.Length - 2
                Return items(intItemsCount)
            Catch ex As Exception

            End Try
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' This procedue is used to get the controls of the form
        ''' </summary>
        ''' <param name="oFormControl">This is used to store the form control</param>
        ''' <remarks>
        ''' This procedue is used to get the controls of the form
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/5/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        ''' 

        Private Sub GetControls(ByVal oFormControl As Control)

            Dim oTabControl As TabControl
            Dim oTabPage As TabPage

            For Each oControl As Control In oFormControl.Controls
                If TypeOf (oControl) Is TabPage Then
                    oTabPage = CType(oControl, TabPage)
                    oTabControl = oTabPage.Parent
                    oTabControl.SelectedTab = oTabPage
                    GetControls(oControl)
                ElseIf TypeOf (oControl) Is GroupBox OrElse _
                        TypeOf (oControl) Is Panel OrElse _
                        TypeOf (oControl) Is TabControl Then
                    GetControls(oControl)
                Else
                    AddDataToTable(oControl)
                End If
            Next
            'Dim intCount As Integer = 0
            'Dim conCount As Integer = oFormControl.Controls.Count
            'Dim oControl As Control
            'For intCount = 1 To conCount
            '    oControl = oFormControl.Controls.Item(conCount - intCount)

            '    If TypeOf (oControl) Is TabPage Then
            '        oTabPage = CType(oControl, TabPage)
            '        oTabControl = oTabPage.Parent
            '        oTabControl.SelectedTab = oTabPage
            '        GetControls(oControl)
            '    ElseIf TypeOf (oControl) Is GroupBox OrElse _
            '            TypeOf (oControl) Is Panel OrElse _
            '            TypeOf (oControl) Is TabControl Then
            '        GetControls(oControl)
            '    Else
            '            AddDataToTable(oControl)
            '    End If
            'Next


        End Sub

#End Region

#Region "Functions"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' This function is used to check the table struture.
        ''' </summary>
        ''' <returns>Boolean is the return type, if the table doesnt have structure it will return false else true.</returns>
        ''' <remarks>
        ''' It will equate the table name with the form if they are equal it will check whether the array items exist in the 
        ''' table columns if they are it will return true else false.
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/4/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Function CheckTableStructure() As Boolean
            CheckTableStructure = True
            Try
                If _oTable.TableName = _SourceForm.Name Then
                    For Each strColumnName As String In aTableStructure
                        If Not _oTable.Columns.Contains(strColumnName) Then
                            CheckTableStructure = False
                            _strErrorMessage = "The Tablename doen't contain the specified structure like" + vbCrLf
                            _strErrorMessage += "TabIndex , ConName and Value"
                            Exit For
                        End If
                    Next
                Else
                    _strErrorMessage = "The Tablename and the Form are not Coinciding"
                    CheckTableStructure = False
                End If
            Catch oException As Exception
                _strErrorMessage = "THE TABLE STRUCTURE IS INCORRECT !!" + vbCrLf
                _strErrorMessage += "System generated error:-" + vbCrLf + oException.Message
                _oErrorObject = oException
            End Try
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' This function is used to set the form with data.
        ''' </summary>
        ''' <returns>Boolean is the return type, if the data is not set to the form it will return false else true.</returns>
        ''' <remarks>
        ''' This function is used to set the data to the form with datatable and form as an parameters.
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/4/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function SetDataToForm(ByVal oTable As DataTable, ByVal oSourceForm As Form) As Boolean
            Dim odataview As DataView
            SetDataToForm = False
            _oTable = oTable
            _SourceForm = oSourceForm
            Try
                If CheckTableStructure() Then
                    odataview = SortTable()
                    For Each oDataRowView In odataview
                        FindControl(_SourceForm)
                    Next
                    SetDataToForm = True
                End If
            Catch oException As Exception
                _strErrorMessage = "THE DATA IS NOT BEEN SET TO THE FORM !!" + vbCrLf
                _strErrorMessage += "System generated error:-" + vbCrLf + oException.Message
                _oErrorObject = oException
            End Try

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' This function is used to sort the table.
        ''' </summary>
        ''' <returns>dataview is the return type, which is stored into the SortTable.</returns>
        ''' <remarks>
        ''' This function is used to sort the table with the tabindex.
        ''' </remarks>
        ''' <history>
        ''' 	[satishkumar.hv]	8/5/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Function SortTable() As DataView
            Dim oDataTable As DataTable
            Dim oRow As DataRow
            SortTable = Nothing
            Try
                oDataTable = _oTable
                CreateSourceTable(_oTable.TableName)
                For Each odataRow In oDataTable.Rows
                    oRow = _oTable.NewRow
                    oRow("TabIndex") = CType(odataRow.Item("TabIndex"), Int32)
                    oRow("ConName") = odataRow.Item("ConName")
                    oRow("Value") = odataRow.Item("Value")
                    _oTable.Rows.Add(oRow)
                Next
                SortTable = _oTable.DefaultView
                SortTable.Sort = "TabIndex"
                'SortTable.Sort = "TabIndex"
            Catch oException As Exception
                _strErrorMessage = "THE DATA IS NOT BEEN SET TO THE FORM !!" + vbCrLf
                _strErrorMessage += "System generated error:-" + vbCrLf + oException.Message
                _oErrorObject = oException
            End Try
           
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' This function is used to retreive and store all controls information.
        ''' </summary>
        ''' <param name="oFormControl">This is used to store the form control. </param>
        ''' <returns>Datatable is the return type</returns>
        ''' <remarks>
        ''' This function will check the each tab page and in that it will check the tab control.
        ''' </remarks>
        ''' <history>
        ''' 	[swamy.reddy]	8/4/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function StoreFormData(ByVal oFormControl As Control) As DataTable

            StoreFormData = Nothing
            Try
                If CreateSourceTable(oFormControl.Name) Then
                    GetControls(oFormControl)
                    Return _oTable
                End If
            Catch oException As Exception
                _strErrorMessage = "THE DATA IS NOT RETREIVED FROM THE SPECIFIED CONTROL FROM THE FORM !!" + vbCrLf
                _strErrorMessage += "System generated error:-" + vbCrLf + oException.Message
                _oErrorObject = oException
            End Try

        End Function



#End Region

    End Class
End Namespace
