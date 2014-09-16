Imports IFLCustomUILayer
Public Class FunctionalClass

#Region "Class Variables"
    Dim _strErrorMessage As String
    Private _alInputTable As New ArrayList
    Private _alUserInputsMinMax As New ArrayList
    Private _iNoofCoreDetailsRowsEntered As Integer
    Private _iCminValue As Double
    Private _iCoreLastStep As Double

#End Region

#Region "Properties"
    Public Property ErrorMessage() As String
        Get
            Return _strErrorMessage
        End Get
        Set(ByVal value As String)
            _strErrorMessage = value
        End Set
    End Property
    Public Property CMinValue() As Double
        Get
            Return _iCminValue
        End Get
        Set(ByVal value As Double)
            _iCminValue = value
        End Set
    End Property
    Public Property NoofCoreDetailsRowsEntered() As Integer '29-05
        Get
            Return _iNoofCoreDetailsRowsEntered
        End Get
        Set(ByVal value As Integer)
            _iNoofCoreDetailsRowsEntered = value
        End Set
    End Property
    Public Property CoreLastStep() As Double
        Get
            Return _iCoreLastStep
        End Get
        Set(ByVal value As Double)
            _iCoreLastStep = value
        End Set
    End Property
    Public ReadOnly Property UserInputsMinMaxImages() As ArrayList  '26-05
        Get
            If _alUserInputsMinMax.Count = 0 Then
                _alUserInputsMinMax.Add(New Object(5) {"Dfe", 415, 624, "Image1", "Enter the value in the scope (415 to 624)", MessageBoxButtons.OKCancel})  'Parameter-minvalue-maxvalue-imagename--ErrorMessg
                _alUserInputsMinMax.Add(New Object(5) {"A", 1000, 2000, "Image2", "Check the stability of the assembly and if required increase cross flat length(1000 to 2000)", MessageBoxButtons.OK})
                _alUserInputsMinMax.Add(New Object(5) {"L.C.", 775, 1500, "Image2", "Leg centre entered is not in the scope of program, Enter the value within the scope(775 to1500)", MessageBoxButtons.OKCancel})
                ' _alUserInputsMinMax.Add(New Object(5) {"C&CWT", 7500, 99999.99, "Image3", "Check the value of inner most winding diameter D1. Enter the value with in the scope of program(>7500) ", MessageBoxButtons.OK})

                _alUserInputsMinMax.Add(New Object(5) {"H", 0, 99999.99, "Image2", "Enter value between 0 and 99999.99", MessageBoxButtons.OK})
                _alUserInputsMinMax.Add(New Object(5) {"B", 0, 99999.99, "Image2", "Enter value between 0 and 99999.99", MessageBoxButtons.OK})
                _alUserInputsMinMax.Add(New Object(5) {"D1", 1, 99999.99, "Image3", "Enter value between 1 and 99999.99", MessageBoxButtons.OK})
                _alUserInputsMinMax.Add(New Object(5) {"WL1", 0, 99999.99, "Image3", "Enter value between 0 and 99999.99", MessageBoxButtons.OK})
                _alUserInputsMinMax.Add(New Object(5) {"M2", 0, 99999.99, "Image3", "Enter value between 0 and 99999.99", MessageBoxButtons.OK})
                _alUserInputsMinMax.Add(New Object(5) {"D3", 0, 99999.99, "Image3", "Enter value between 0 and 99999.99", MessageBoxButtons.OK})
                _alUserInputsMinMax.Add(New Object(5) {"D4", 0, 99999.99, "Image3", "Enter value between 0 and 99999.99", MessageBoxButtons.OK})
                _alUserInputsMinMax.Add(New Object(5) {"WL2", 0, 99999.99, "Image3", "Enter value between 0 and 99999.99", MessageBoxButtons.OK})
                _alUserInputsMinMax.Add(New Object(5) {"M4", 0, 99999.99, "Image3", "Enter value between 0 and 99999.99", MessageBoxButtons.OK})

            End If
            UserInputsMinMaxImages = _alUserInputsMinMax

        End Get
    End Property

    Private ReadOnly Property InputTable() As ArrayList
        Get
            If _alInputTable.Count = 0 Then
                _alInputTable.Add(New Object(2) {"Core Dia in mm (415 to 624)", "Dfe", 0})
                _alInputTable.Add(New Object(2) {"Window Width in mm", "B", 0})
                _alInputTable.Add(New Object(2) {"Leg Center in mm", "L.C.", 0})
                _alInputTable.Add(New Object(2) {"Window Height in mm", "A", 0})
                _alInputTable.Add(New Object(2) {"Total Stack Height including ducts in mm", "H", 0})
                _alInputTable.Add(New Object(2) {"Innermost Winding I.D.in mm", "D1", 0})
                _alInputTable.Add(New Object(2) {"Innermost Winding O.D.in mm", "D2", 0})
                _alInputTable.Add(New Object(2) {"Innermost Winding axial height in mm", "WL1", 0})
                _alInputTable.Add(New Object(2) {"Innermost Winding position above bottom yoke top in mm ", "M2", 0})
                _alInputTable.Add(New Object(2) {"Outermost Winding I.D.in mm", "D3", 0})
                _alInputTable.Add(New Object(2) {"Outermost Winding O.D.in mm", "D4", 0})
                _alInputTable.Add(New Object(2) {"Outermost Winding axial height in mm ", "WL2", 0})
                _alInputTable.Add(New Object(2) {"Outermost Winding position above bottom yoke top in mm", "M4", 0})
                _alInputTable.Add(New Object(2) {"Core and Winding Weight in kg", "C&CWT", 0})

            End If
            Return _alInputTable
        End Get
    End Property
    Private _alCoreDetails As New ArrayList

    Private ReadOnly Property CoreDetailsTable() As ArrayList '26-05
        Get
            If _alCoreDetails.Count = 0 Then
                _alCoreDetails.Add(New Object(3) {"C1", "   -   ", "H1", "   -   "})
                _alCoreDetails.Add(New Object(3) {"C2", "   -   ", "H2", "   -   "})
                _alCoreDetails.Add(New Object(3) {"C3", "   -   ", "H3", "   -   "})
                _alCoreDetails.Add(New Object(3) {"C4", "   -   ", "H4", "   -   "})

                _alCoreDetails.Add(New Object(3) {"C5", "   -   ", "H5", "   -   "})
                _alCoreDetails.Add(New Object(3) {"C6", "   -   ", "H6", "   -   "})
                _alCoreDetails.Add(New Object(3) {"C7", "   -   ", "H7", "   -   "})
                _alCoreDetails.Add(New Object(3) {"C8", "   -   ", "H8", "   -   "})

                _alCoreDetails.Add(New Object(3) {"C9", "   -   ", "H9", "   -   "})
                _alCoreDetails.Add(New Object(3) {"C10", "   -   ", "H10", "   -   "})
                _alCoreDetails.Add(New Object(3) {"C11", "   -   ", "H11", "   -   "})
                _alCoreDetails.Add(New Object(3) {"C12", "   -   ", "H12", "   -   "})


                _alCoreDetails.Add(New Object(3) {"C13", "   -   ", "H13", "   -   "})
                _alCoreDetails.Add(New Object(3) {"C14", "   -   ", "H14", "   -   "})
                _alCoreDetails.Add(New Object(3) {"C15", "   -   ", "H15", "   -   "})
                _alCoreDetails.Add(New Object(3) {"C16", "   -   ", "H16", "   -   "})


                _alCoreDetails.Add(New Object(3) {"C17", "   -   ", "H17", "   -   "})
                _alCoreDetails.Add(New Object(3) {"C18", "   -   ", "H18", "   -   "})
                _alCoreDetails.Add(New Object(3) {"C19", "   -   ", "H19", "   -   "})
                _alCoreDetails.Add(New Object(3) {"C20", "   -   ", "H20", "   -   "})

            End If
            CoreDetailsTable = _alCoreDetails
        End Get
    End Property

#End Region

#Region "Functions"
    ''' <summary>
    ''' Checks the validation for the empty fields.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 
    Public Function validateForm(ByVal FormName As Form) As Control
        validateForm = Nothing
        Dim ctl As GroupBox
        Dim CtrlGrp As GroupBox        '20_10_2009    ragava
        'FormName.Visible = True
        Try
            '19_10_2009  ragava
            If FormName.Name.IndexOf("frmTieRod3") <> -1 Then
                For Each objControl As Object In FormName.Controls
                    If TypeOf (objControl) Is GroupBox Then
                        ctl = CType(objControl, GroupBox)
                        '20_10_2009  ragava
                        Dim AL_RtxBxNumbering As New ArrayList
                        For Each ctlRichtxtBox1 As Control In ctl.Controls
                            If TypeOf (ctlRichtxtBox1) Is GroupBox Then
                                CtrlGrp = CType(ctlRichtxtBox1, GroupBox)
                                For Each ctlRichtxtBox As Control In CtrlGrp.Controls
                                    If TypeOf (ctlRichtxtBox) Is RichTextBox Then
                                        Dim CtrlRich As RichTextBox = CType(ctlRichtxtBox, RichTextBox)
                                        For Each str As String In CtrlRich.Lines
                                            '01_12_2009  Ragava
                                            If Trim(str).IndexOf("}") < 1 Then
                                                ErrorMessage = "Numbering is not done... Please number the notes to proceed" + vbNewLine
                                                validateForm = CtrlRich
                                                Return validateForm
                                            End If
                                            '01_12_2009  Ragava   Till  Here
                                            '20_10_2009  ragava
                                            If AL_RtxBxNumbering.Count > 0 Then          '28_10_2009  ragava
                                                If AL_RtxBxNumbering.Contains(str.Substring(0, str.IndexOf("}"))) = True Then
                                                    ErrorMessage = "Numbering sequence is not unique" + vbNewLine
                                                    validateForm = ctlRichtxtBox
                                                    Return validateForm
                                                End If
                                            End If
                                            '20_10_2009  ragava   Till  Here
                                            AL_RtxBxNumbering.Add(str.Substring(0, str.IndexOf("}")))
                                        Next
                                    End If
                                Next
                            End If
                        Next
                        '20_10_2009  ragava    Till  Here
                        For Each ctl1 As Control In ctl.Controls
                            Dim ctlObject As Control
                            If TypeOf (ctl1) Is TextBox Then
                                ctlObject = CType(ctl1, TextBox)
                                If Trim(ctlObject.Text) <> "" Then
                                    Dim ctlObject1 As Control
                                    For Each ctl2 As Control In ctl.Controls
                                        If TypeOf (ctl2) Is TextBox Then
                                            ctlObject1 = CType(ctl2, TextBox)
                                            If Trim(ctlObject.Name) <> Trim(ctlObject1.Name) Then
                                                If (Trim(ctlObject.Text) = Trim(ctlObject1.Text)) Or (AL_RtxBxNumbering.Contains(Trim(ctlObject.Text))) Then           '20_10_2009  ragava
                                                    ErrorMessage = "Numbering sequence is not unique" + vbNewLine
                                                    validateForm = ctl2
                                                    Return validateForm
                                                End If
                                            End If
                                        End If
                                    Next
                                Else
                                    '28_10_2009  ragava
                                    If ctlObject.Enabled = True Then
                                        ErrorMessage = "Numbering is not done... Please number the notes to proceed" + vbNewLine
                                        validateForm = ctlObject
                                        Return validateForm
                                    End If
                                    'If ctlObject.Name = "txtRetractedLength" Or ctlObject.Name = "txtExtenedLength" Or ctlObject.Name = "txtRodDiameter" Then
                                    '    ErrorMessage = "Sequence Numbering is Not Done" + vbNewLine
                                    '    validateForm = ctlObject
                                    '    Return validateForm
                                    'End If
                                    '28_10_2009  ragava   Till  Here
                                End If
                            End If
                        Next
                    End If
                Next
                Return validateForm
            End If
            '19_10_2009  ragava     Till  Here

        Catch ex As Exception
        End Try
        Try
            For Each objControl As Object In FormName.Controls
                If TypeOf (objControl) Is GroupBox Then
                    ctl = CType(objControl, GroupBox)
                    For Each ctl1 As Control In ctl.Controls
                        If Not checkControl(ctl1) Then
                            validateForm = ctl1
                            Return validateForm
                        End If
                        'If TypeOf (ctl1) Is TextBox Then
                        '    ctlObject = CType(ctl1, TextBox)
                        '    If ctlObject.Enabled = True Then
                        '        If ctlObject.Text = "" Then
                        '            ErrorMessage = "Please Enter value in the " + vbNewLine
                        '            ErrorMessage += (ctl1.Name).Replace("txt", "") + vbNewLine
                        '            ErrorMessage += "Application Generated Error"
                        '            validateForm = ctl1
                        '            Return validateForm
                        '        End If
                        '    End If
                        'ElseIf TypeOf (ctl1) Is ComboBox Then
                        '    If CType(ctl1, ComboBox).Enabled = True Then
                        '        If CType(ctl1, ComboBox).SelectedItem Is Nothing Then
                        '            ErrorMessage = "Select Item from the " + vbNewLine
                        '            ErrorMessage += (ctl1.Name).Replace("cmb", "") + vbNewLine
                        '            ErrorMessage += "Application Generated Error"
                        '            validateForm = ctl1
                        '            Return validateForm
                        '        End If
                        '    End If
                        'End If
                        'checkControl(ctl1)
                    Next
                Else
                    If Not checkControl(objControl) Then
                        validateForm = objControl
                        Return validateForm
                    End If
                End If
            Next
        Catch ex As Exception
        End Try
    End Function

    Public Function checkControl(ByVal ctl1 As Control) As Boolean
        Dim ctlObject As Control
        checkControl = True
        If TypeOf (ctl1) Is TextBox Then
            ctlObject = CType(ctl1, TextBox)
            If ctlObject.Visible = True Then
                If ctlObject.Enabled = True Then
                    If ctlObject.Name = "txtlPartCode" Then
                        Return checkControl
                    ElseIf Trim(ctlObject.Text) = "" Then         '29_10_2009   Ragava
                        ErrorMessage = "Please enter "
                        ErrorMessage += (ctl1.Name).Replace("txt", "")
                        'ErrorMessage += "Application Generated Error"
                        checkControl = False
                        Return checkControl
                    End If
                End If
            End If
        ElseIf TypeOf (ctl1) Is ComboBox Then
            If CType(ctl1, ComboBox).Visible = True Then
                If CType(ctl1, ComboBox).Enabled = True Then
                    If CType(ctl1, ComboBox).SelectedItem Is Nothing Or Trim(CType(ctl1, ComboBox).SelectedItem) = "" Then         '29_10_2009   Ragava
                        'ErrorMessage = "Select item from the " + vbNewLine
                        'ErrorMessage += (ctl1.Name).Replace("cmb", "") + vbNewLine

                        ErrorMessage = "Select " & (ctl1.Name).Replace("cmb", "") & " from the list"

                        'ErrorMessage += "Application Generated Error"
                        checkControl = False
                        Return checkControl
                    End If
                End If
            End If



            '29_10_2009   Ragava
        ElseIf TypeOf (ctl1) Is ListView Then
            If CType(ctl1, ListView).Enabled = True Then
                If CType(ctl1, ListView).Items.Count > 1 Then
                    If CType(ctl1, ListView).SelectedItems.Count < 1 Then
                        ErrorMessage = "Select " & (ctl1.Name).Replace("LV", "") & " from the List"
                        checkControl = False
                        Return checkControl
                    End If
                End If
            End If
            '29_10_2009   Ragava   Till   Here
        End If
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="strFormName"></param>
    ''' <param name="strControlName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetControlValue(ByVal strFormName As String, ByVal strControlName As String) As Object
        GetControlValue = Nothing
        For Each oItem As Object In arAllFormsCotrolValues
            If oItem(FormObjects.FormName) = strFormName Then
                GetControlValue = GetCurrentControlValue(oItem(FormObjects.ControlName), strControlName)
                Exit For
            End If
        Next
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="oHashTable"></param>
    ''' <param name="strcontrolname"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCurrentControlValue(ByVal oHashTable As Hashtable, ByVal strcontrolname As String) As Object
        GetCurrentControlValue = oHashTable(strcontrolname)
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="key"></param>
    ''' <param name="alist"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getValue(ByVal key As String, ByVal alist As ArrayList) 'As String
        getValue = ""
        For Each aData As Object In alist
            If aData(0).Equals(key) Then
                getValue = aData(1) '.ToString
                Exit For
            End If
        Next
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="e"></param>
    ''' <param name="oControl"></param>
    ''' <param name="strMessage"></param>
    ''' <remarks></remarks>
    Public Sub ValidateAlphabetWithDecimal(ByRef e As System.Windows.Forms.KeyPressEventArgs, ByRef oControl As IFLCustomUILayer.IFLTextBox, ByVal strMessage As String)
        Dim KeyAscii As Short = Asc(e.KeyChar)

        If oControl.Text.Length = 0 And KeyAscii = 46 Then
            MessageBox.Show("First Character Should not be .", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            e.Handled = True
            oControl.Clear()
            Exit Sub
        End If
        If (KeyAscii >= 65 And KeyAscii <= 90) Then
            e.Handled = False
        ElseIf (KeyAscii >= 97 And KeyAscii <= 122) Then
            e.Handled = False
        ElseIf (KeyAscii = 13) Or (KeyAscii = 46) Or (KeyAscii = 8) Then   ' Enter or . or Backspace
            e.Handled = False
        Else
            e.Handled = True
            MessageBox.Show("Please enter only Alphabets", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        If KeyAscii = 13 Then
            If oControl.Text.Length = 0 Then
                MessageBox.Show("Enter " & strMessage, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                e.Handled = True
                Exit Sub
            End If
        End If

    End Sub

    Public Sub ValidateAplhaNumerics(ByRef e As System.Windows.Forms.KeyPressEventArgs, ByRef oControl As IFLCustomUILayer.IFLTextBox, ByVal strMessage As String)
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Dim bolIsAllowed As Boolean

        bolIsAllowed = (KeyAscii >= 65 And KeyAscii <= 90) OrElse (KeyAscii >= 97 And KeyAscii <= 122) _
                    OrElse ((KeyAscii = 13) Or (KeyAscii = 46) Or (KeyAscii = 8)) OrElse (KeyAscii >= 48 And KeyAscii <= 57)


        If bolIsAllowed Then
            e.Handled = False
        Else
            e.Handled = True
            MessageBox.Show("Please enter only Alphabets", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If KeyAscii = 13 And oControl.Text.Length = 0 Then
            MessageBox.Show("Enter " & strMessage, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            e.Handled = True
            Exit Sub
        End If

    End Sub
    Public Sub PopulateFormscontrolsData(ByVal oForm As Form)
        Dim oHastTable As New Hashtable
        For Each oControl As Control In oForm.Controls
            If TypeOf oControl Is TextBox OrElse TypeOf oControl Is ComboBox Then
                If Not oHastTable.ContainsKey(oControl.Name) Then
                    'For Removing string like 'txt','cmb', taken substring
                    oHastTable.Add(oControl.Name.Substring(3), oControl.Text)
                End If
            ElseIf TypeOf oControl Is DataGridView Then
                Dim oDataGridControl As DataGridView
                oDataGridControl = CType(oControl, DataGridView)
                For Each oControlView As Control In oDataGridControl.Controls
                    oHastTable.Add(oControl.Name.Substring(3), oControl.Text)
                Next
            ElseIf TypeOf oControl Is GroupBox Then
                Dim ctl As GroupBox
                ctl = CType(oControl, GroupBox)
                For Each ctl1 As Control In ctl.Controls
                    If TypeOf ctl1 Is TextBox OrElse TypeOf ctl1 Is ComboBox Then

                        If Not oHastTable.ContainsKey(ctl1.Name) Then
                            If ctl1.Name = "txtStrokeLength" Then
                                oHastTable.Add("StrokeLength1", ctl1.Text)
                            Else
                                oHastTable.Add(ctl1.Name.Substring(3), ctl1.Text)
                            End If
                        End If

                        '04_11_2009  Ragava
                    ElseIf TypeOf ctl1 Is GroupBox Then
                        Dim ctlGrp As GroupBox
                        ctlGrp = CType(ctl1, GroupBox)
                        For Each ctrl As Control In ctlGrp.Controls
                            If TypeOf ctrl Is RichTextBox Then
                                If Not oHastTable.ContainsKey(ctrl.Name) Then
                                    oHastTable.Add(ctrl.Name, ctrl.Text)
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        Next
        arAllFormsCotrolValues.Add(New Object(1) {oForm.Name, oHastTable})
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="formName"></param>
    ''' <param name="controlName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getValue(ByVal formName As String, ByVal controlName As String) As Object
        getValue = GetControlValue(formName, controlName)
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Key"></param>
    ''' <param name="alist"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getValueGridValues(ByVal Key As String, ByVal alist As ArrayList) As Double
        getValueGridValues = Val(getValue(Key, alist))
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PopulateInputTable() As DataTable

        Dim oDatatable As New DataTable("InputTable")
        Dim oDataRow As DataRow
        oDatatable.Columns.Add("Parameter", System.Type.GetType("System.String"))
        oDatatable.Columns.Add("Nomenclature", System.Type.GetType("System.String"))
        oDatatable.Columns.Add("ParamValue", System.Type.GetType("System.Double"))

        For Each oItem As Object In InputTable
            oDataRow = oDatatable.NewRow()
            oDataRow("Parameter") = oItem(InputTableColumns.Parameter)
            oDataRow("Nomenclature") = oItem(InputTableColumns.Nomenclature)
            oDataRow("ParamValue") = oItem(InputTableColumns.Value)

            oDatatable.Rows.Add(oDataRow)
        Next

        PopulateInputTable = oDatatable
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="e"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsCharNumbersBackspaceDecimal(ByVal e As KeyPressEventArgs) As Boolean '26-05
        If (Asc(e.KeyChar) >= 48 AndAlso Asc(e.KeyChar) <= 57) OrElse Asc(e.KeyChar) = 8 Then
            IsCharNumbersBackspaceDecimal = True
        ElseIf e.KeyChar = "." Then
            IsCharNumbersBackspaceDecimal = True
        End If
    End Function
#End Region

#Region "Procedures"
    ''' <summary>
    ''' Provides the help of the controls in the Status strip.
    ''' </summary>
    ''' <param name="textBoxControlName"></param>
    ''' <param name="strInfo"></param>
    ''' <remarks></remarks>
    Public Sub setUserInformation(ByVal textBoxControlName As IFLCustomUILayer.IFLNumericBox, ByVal strInfo As String, ByVal ToolStripStatusLabel1 As ToolStripStatusLabel)
        textBoxControlName.StatusMessage = strInfo
        ToolStripStatusLabel1.BackColor = Color.Coral
        textBoxControlName.StatusObject = ToolStripStatusLabel1
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub NumericTextboxKeyPressEvent(ByVal sender As Object, ByVal e As KeyPressEventArgs) '26-05
        Dim nonNumberEntered As [Boolean]
        Dim iDecimalIndex As Integer
        Dim bolIsAllowed As Boolean = False

        nonNumberEntered = True

        Dim strTextEntered As String = DirectCast((sender), System.Windows.Forms.DataGridViewTextBoxEditingControl).Text
        Dim strText As String = strTextEntered + e.KeyChar.ToString()

        If strText.Contains(".") Then
            iDecimalIndex = strText.IndexOf(".")
            If iDecimalIndex <= 5 Then
                bolIsAllowed = IsCharNumbersBackspaceDecimal(e)
            Else
                bolIsAllowed = False
            End If

        Else
            If strText.Length <= 5 Then
                bolIsAllowed = IsCharNumbersBackspaceDecimal(e)
            Else
                bolIsAllowed = False
            End If
        End If

        'to restrict digits(to 2) after decimal point
        If strTextEntered.Contains(".") Then
            iDecimalIndex = strTextEntered.IndexOf(".")

            If strText.Length > iDecimalIndex + 3 Then
                bolIsAllowed = False
                'e.Handled = True
            Else
                bolIsAllowed = IsCharNumbersBackspaceDecimal(e)
            End If
        End If

        e.Handled = Not (bolIsAllowed)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="e"></param>
    ''' <param name="oParameter"></param>
    ''' <remarks></remarks>
    Private Sub ShowDialogOnInvalidValues(ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs, ByRef oParameter As Object)  '26-05
        Dim dlgResult As DialogResult
        dlgResult = MessageBox.Show(oParameter(Inputs.ErrorMessage).ToString, "Information", oParameter(Inputs.MsgButtons), MessageBoxIcon.Information)

        If dlgResult = DialogResult.OK Then
            '  e.Cancel = True
        ElseIf dlgResult = DialogResult.Cancel Then
            Application.Exit()
        End If
    End Sub

#End Region

#Region "Enums"
    ''' <summary>
    ''' This enum is to describe the items in the arraylist CoreDetailsTable
    ''' </summary>
    ''' <remarks></remarks>
    Private Enum CoreTableColumns '26-05
        StepWidth
        StepWidthValue
        StepThickness
        StepThicknessValue
    End Enum
    Private Enum InputTableColumns
        Parameter
        Nomenclature
        Value
    End Enum
    ''' <summary>
    ''' This enum is to describe the items in the arraylist UserInputsMinMaxImages
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum Inputs '26-05
        Parameter
        MinValue
        MaxValue
        ImageName
        ErrorMessage
        MsgButtons
    End Enum

#End Region

    'TODO: ANUP 27-07-2010
    Public Sub LabelGradient_GreenBorder_ColoringTheScreens(ByVal LabelGradient1 As LabelGradient.LabelGradient, ByVal LabelGradient2 As LabelGradient.LabelGradient, ByVal LabelGradient3 As LabelGradient.LabelGradient, ByVal LabelGradient4 As LabelGradient.LabelGradient)
        LabelGradient1.GradientColorOne = Color.Black
        LabelGradient1.GradientColorTwo = Color.Black

        LabelGradient2.GradientColorOne = Color.Black
        LabelGradient2.GradientColorTwo = Color.FromArgb(255, 47, 23)
        LabelGradient2.GradientMode = Drawing2D.LinearGradientMode.Horizontal

        LabelGradient3.GradientColorOne = Color.FromArgb(255, 47, 23)
        LabelGradient3.GradientColorTwo = Color.FromArgb(255, 47, 23)

        LabelGradient4.GradientColorOne = Color.Black
        LabelGradient4.GradientColorTwo = Color.FromArgb(255, 47, 23)
        LabelGradient4.GradientMode = Drawing2D.LinearGradientMode.Horizontal
    End Sub

    Public Sub LabelGradient_OrangeBorder_ColoringTheScreens(ByVal LabelGradient As LabelGradient.LabelGradient)
        LabelGradient.GradientColorOne = Color.Black
        LabelGradient.GradientColorTwo = Color.White
        LabelGradient.GradientMode = Drawing2D.LinearGradientMode.Horizontal
    End Sub

    Public Sub subLabelGradient_Child_ColoringScreens(ByVal LabelGradient As LabelGradient.LabelGradient)
        LabelGradient.GradientColorOne = Color.Black
        LabelGradient.GradientColorTwo = Color.FromArgb(255, 47, 23)
        LabelGradient.GradientMode = Drawing2D.LinearGradientMode.Horizontal
    End Sub

End Class
