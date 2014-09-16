Imports SldWorks
Imports System.IO
Imports System.Threading
'Imports SolidWorks

Public Class IFLSolidWorksBaseClass
    'Inherits IFLSolidWorksClass ' BBL.IFLSolidWorks
#Region "Variables"
    Public _htDocumentInstances As New Hashtable()     '27_07_2009  ragava
    Public blnVisibleSolidWorks As Boolean = True
    Private _strErrorMessage As String
    Private _oErrorObject As Exception
    Private _oCurrentSolidWorksObject As Object
    Dim oSolidWorksApplication As SldWorks.SldWorks
    Dim oSolidWorksModel As SldWorks.ModelDoc2
    Dim oSolidWorksAssembly As SldWorks.AssemblyDoc
    Dim oSolidWorksPartDocument As SldWorks.PartDoc
    Dim oSolidWorksDrawingDocument As SldWorks.DrawingDoc
    Dim oSolidWorksView As SldWorks.View       '16_09_2009    ragava
    Dim nWarnings As Long
    Dim Errors As Long

    Public swModelExt As SldWorks.ModelDocExtension         '22_09_2009  ragava
    Public swOleObj As SldWorks.SwOLEObject                 '22_09_2009  ragava
#End Region
    Public Sub New(ByVal blnVisibleSolidworks As Boolean)
        blnVisibleSolidworks = blnVisibleSolidworks
        If SolidWorksApplicationObject() Is Nothing Then
            If Not ConnectSolidWorks() Then
                _strErrorMessage = "Unable to Connect To SolidWorks" + vbNewLine
                _strErrorMessage += "System Generated Error"
            End If
        End If
    End Sub
#Region "Properties[Dont Delete this line]"

    Public Property SolidWorksApplicationObject() As SldWorks.SldWorks
        Get
            Return oSolidWorksApplication
        End Get
        Set(ByVal value As SldWorks.SldWorks)
            oSolidWorksApplication = value
        End Set
    End Property

    Public Property SolidWorksModel() As SldWorks.ModelDoc2
        Get
            Return oSolidWorksModel
        End Get
        Set(ByVal value As SldWorks.ModelDoc2)
            oSolidWorksModel = value
        End Set
    End Property

    Public Property SolidWorksAssembly() As SldWorks.AssemblyDoc
        Get
            Return oSolidWorksAssembly
        End Get
        Set(ByVal value As SldWorks.AssemblyDoc)
            oSolidWorksAssembly = value
        End Set
    End Property

    Public Property SolidWorksPartDocument() As SldWorks.PartDoc
        Get
            Return oSolidWorksPartDocument
        End Get
        Set(ByVal value As SldWorks.PartDoc)
            oSolidWorksPartDocument = value
        End Set
    End Property

    Public Property SolidWorksDrawingDocument() As SldWorks.DrawingDoc
        Get
            Return oSolidWorksDrawingDocument
        End Get
        Set(ByVal value As SldWorks.DrawingDoc)
            oSolidWorksDrawingDocument = value
        End Set
    End Property
    '16_09_2009   ragava
    Public Property SolidWorksView() As SldWorks.View
        Get
            oSolidWorksView.BreakLineGap = 0.00254        '29_09_2009  ragava
            Return oSolidWorksView
        End Get
        Set(ByVal value As SldWorks.View)
            oSolidWorksView = value
            oSolidWorksView.BreakLineGap = 0.00254        '29_09_2009  ragava
        End Set
    End Property

    '24_09_2009  ragava
    Public ReadOnly Property selectOLE() As Boolean
        Get
            Dim bret As Boolean = False
            bret = SolidWorksModel.Extension.SelectByID2("Worksheet", "OLEITEM", 0.05817688726539, 0.2160797219023, 0, False, 0, Nothing, 0)
            If bret = False Then
                SolidWorksModel.Extension.SelectByID2("Worksheet", "OLEITEM", 0.07228921126611, 0.2439971402624, 0, False, 0, Nothing, 0)
            End If
        End Get
    End Property

    '27_07_2009  ragava  Commented
    'Public ReadOnly Property ErrorMessage()
    '    Get
    '        Return _strErrorMessage
    '    End Get
    'End Property

    'Public ReadOnly Property ErrorObject()
    '    Get
    '        Return _oErrorObject
    '    End Get
    'End Property

    'Public ReadOnly Property IsCurrentSolidWorks() As SldWorks.SldWorks
    '    Get
    '        Return _oCurrentSolidWorksObject
    '    End Get
    'End Property
    '27_07_2009  ragava   Till  Here
#End Region
#Region "Enums"
    Public Enum SolidworksDocumentType
        PartDocument = 1
        AssemblyDocument = 2
        DrawingDocument = 3
        None = 4
    End Enum
#End Region
#Region "SolidWorks Connection "
    ''' <summary>
    ''' Opens the solidworks Session.
    ''' </summary>
    ''' <remarks></remarks>
    Public Function ConnectSolidWorks() As Boolean
        ConnectSolidWorks = False
        Try
            SolidWorksApplicationObject = CreateObject("SldWorks.Application")
            If Not SolidWorksApplicationObject Is Nothing Then
                SolidWorksApplicationObject.Visible = True 'blnVisibleSolidWorks
                _oCurrentSolidWorksObject = SolidWorksApplicationObject
                ConnectSolidWorks = True
            End If
        Catch oException As Exception
            _strErrorMessage = "UNABLE TO OPEN THE SOLIDWORKS!!" + vbCrLf
            _strErrorMessage += "System generated error:-" + vbCrLf + oException.Message
            _oErrorObject = oException
        End Try
    End Function
    '13_04_2010   RAGAVA
    Public Function SaveAndKill() As Boolean
        Try
            SolidWorksModel = SolidWorksApplicationObject.ActiveDoc
            System.Threading.Thread.Sleep(1000)
            SolidWorksModel.SaveSilent()
            System.Threading.Thread.Sleep(2000)
            KillAllSolidWorksServices()
        Catch oException As Exception
            _strErrorMessage = "Unable to Close Documents!" + vbNewLine
            _strErrorMessage += "Some files are not saved properly" + vbNewLine
            _strErrorMessage += "System generated error:-" + vbNewLine + oException.Message
            _oErrorObject = oException
        End Try
    End Function
    ''' <summary>
    ''' </summary>
    ''' <remarks></remarks>
    Public Function SaveAndCloseAllDocuments() As Boolean
        Try
            SolidWorksModel = SolidWorksApplicationObject.ActiveDoc
            System.Threading.Thread.Sleep(1000)
            SolidWorksModel.SaveSilent()
            System.Threading.Thread.Sleep(2000)        '09_04_2010   RAGAVA
            SolidWorksApplicationObject.CloseAllDocuments(False)
            System.Threading.Thread.Sleep(1000)        '09_04_2010   RAGAVA
            'SolidWorksModel = SolidWorksApplicationObject.ActiveDoc
            'If _htDocumentInstances.Count > 0 Then
            '    For Each strFile As String In _htDocumentInstances.Keys
            '        SaveAndCloseAllDocuments = SaveAndCloseDocument(strFile)
            '    Next
            '    _htDocumentInstances.Clear()
            'End If
        Catch oException As Exception
            _strErrorMessage = "Unable to Close Documents!" + vbNewLine
            _strErrorMessage += "Some files are not saved properly" + vbNewLine
            _strErrorMessage += "System generated error:-" + vbNewLine + oException.Message
            _oErrorObject = oException
        End Try
    End Function
    Public Function CloseAllDocuments() As Boolean
        Try
            SolidWorksApplicationObject.CloseAllDocuments(False)
            'SolidWorksModel = SolidWorksApplicationObject.ActiveDoc
            'If _htDocumentInstances.Count > 0 Then
            '    For Each strFile As String In _htDocumentInstances.Keys
            'CloseAllDocuments = CloseDocument(strFile)
            '    Next
            '_htDocumentInstances.Clear()
            'End If
        Catch oException As Exception
            _strErrorMessage = "Unable to Close Documents!" + vbNewLine
            _strErrorMessage += "Some files are not saved properly" + vbNewLine
            _strErrorMessage += "System generated error:-" + vbNewLine + oException.Message
            _oErrorObject = oException
        End Try
    End Function
    ''' <summary>
    ''' Closes the specified document.
    ''' </summary>
    ''' <param name="docName"></param>
    ''' <remarks></remarks>
    Public Function SaveAndCloseDocument(ByVal docName As String) As Boolean
        Try
            'If _htDocumentInstances.ContainsKey(docName) = True Then
            'SaveDocument(_htDocumentInstances(docName))
            SolidWorksModel = SolidWorksApplicationObject.ActivateDoc(docName)
            SolidWorksApplicationObject.IActiveDoc.SaveSilent()
            'SolidWorksModel.SaveSilent()
            SolidWorksApplicationObject.Close()
            SaveAndCloseDocument = CloseDocument(docName)
            'End If
        Catch oException As Exception
            _strErrorMessage = "Unable to Close Document!" + vbNewLine
            _strErrorMessage += "file is not saved properly" + vbNewLine
            _strErrorMessage += "or file is not updated properly" + vbNewLine
            _strErrorMessage += "System generated error:-" + vbNewLine + oException.Message
            _oErrorObject = oException
        End Try
    End Function
    Public Function CloseDocument(ByVal _strDocName As String) As Boolean
        Try
            'If _htDocumentInstances.ContainsKey(_strDocName) = True Then

            SolidWorksApplicationObject.CloseDoc(_strDocName)
            CloseDocument = True
            '27_07_2009  ragava Below
            'If _htDocumentInstances.ContainsKey(_strDocName) = True Then
            '    _htDocumentInstances.Remove(_strDocName)
            'End If
            'End If
        Catch oException As Exception
            _strErrorMessage = "Unable to Close Document!" + vbNewLine
            _strErrorMessage += "file is not saved properly" + vbNewLine
            _strErrorMessage += "or file is not updated properly" + vbNewLine
            _strErrorMessage += "System generated error:-" + vbNewLine + oException.Message
            _oErrorObject = oException
        End Try
    End Function
    ''' <summary>
    ''' Opens the document.
    ''' </summary>
    ''' <param name="_strDocName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function openDocument(ByVal _strDocName As String) As Boolean
        Dim oSwModel As SldWorks.ModelDoc2
        'Dim fileData As FileInfo = My.Computer.FileSystem.GetFileInfo(docName)
        Dim lnDoc As Long = CheckFileType(_strDocName)
        openDocument = False
        '27_07_2009  ragava Below
        'If _htDocumentInstances.ContainsKey(_strDocName) = True Then
        '    SolidWorksModel = _htDocumentInstances(_strDocName)
        '    openDocument = True
        'Else
        Try
            CheckCommandInProgress()       '04_03_2011    RAGAVA
            SolidWorksModel = SolidWorksApplicationObject.OpenDoc6(_strDocName, lnDoc, SwConst.swOpenDocOptions_e.swOpenDocOptions_Silent, "", Errors, nWarnings)
            oSwModel = SolidWorksModel
            '_htDocumentInstances.Add(_strDocName, oSwModel)
            openDocument = True
        Catch oException As Exception
            _strErrorMessage = "Unable to open file" + vbNewLine
            _strErrorMessage += "Check file format" + vbNewLine
            _strErrorMessage += "System generated error" + oException.Message
            _oErrorObject = oException
            openDocument = False
        End Try
        'End If
    End Function

    '04_03_2011   RAGAVA
    Public Sub CheckCommandInProgress()
        Try
            SolidWorksApplicationObject.CommandInProgress = False
            If SolidWorksApplicationObject.CommandInProgress = True Then
                Dim iCount As Integer = 0
                While (iCount < 10)
                    SolidWorksApplicationObject.CommandInProgress = False
                    If SolidWorksApplicationObject.CommandInProgress = False Then
                        Exit Try
                    End If
                    iCount = iCount + 1
                End While
                KillAllSolidWorksServices()
                SolidWorksApplicationObject = Nothing
                ConnectSolidWorks()
                System.Threading.Thread.Sleep(1000)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Function Saveas_detached(ByVal strdrwpath As String) As Boolean
        Try
            Saveas_detached = False
            Dim bret As Boolean = openDocument(strdrwpath)
            SolidWorksModel.SaveAs2(strdrwpath, 3, False, True)
            SolidWorksModel.EditRebuild3()
            SolidWorksModel.SaveSilent()
        Catch ex As Exception

        End Try
    End Function

    '07_09_2012   RAGAVA
    Public Function DeleteDanglingDimension(ByVal strDrawingFile As String) As Boolean
        Try
            Dim SwAnn As SldWorks.Annotation
            Dim boolstatus As Boolean = False
            Dim bret As Boolean = openDocument(strDrawingFile)
            oSolidWorksDrawingDocument = SolidWorksModel
            Dim swView As SldWorks.View = oSolidWorksDrawingDocument.GetFirstView
            Dim AL_Annotations As New ArrayList
            While Not swView Is Nothing
                boolstatus = SolidWorksDrawingDocument.ActivateView(swView.Name)
                SwAnn = swView.GetFirstAnnotation()
                If Not SwAnn Is Nothing Then
                    Dim SwAnn_Type As String = SwAnn.GetType
                    While Not SwAnn Is Nothing
                        If SwAnn.IsDangling Then
                            SwAnn.Select(True)
                        End If
                        SwAnn = SwAnn.GetNext3
                    End While
                    SolidWorksModel.EditDelete()
                    'If SwAnn_Type = "" Then
                    '    SwAnn_Type = "DIMENSION"
                    'End If
                    'For Each Ann As Object In AL_Annotations
                    '    boolstatus = SolidWorksModel.Extension.SelectByID2(Ann, SwAnn_Type, 0, 0, 0, False, 0, Nothing, 0)
                    '    SolidWorksModel.EditDelete()
                    'Next
                End If
                swView = swView.GetNextView
            End While
        Catch ex As Exception

        End Try
    End Function

    Public Function openAssemblyDrawingDocument(ByVal _strDocName As String) As Boolean
        Dim oSwModel As SldWorks.ModelDoc2
        Dim lnDoc As Long = CheckFileType(_strDocName)
        openAssemblyDrawingDocument = False
        Try
            CheckCommandInProgress()         '04_03_2011   RAGAVA
            SolidWorksModel = SolidWorksApplicationObject.OpenDoc6(_strDocName, lnDoc, SwConst.swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", Errors, nWarnings)
            oSwModel = SolidWorksModel
            openAssemblyDrawingDocument = True
        Catch oException As Exception
            _strErrorMessage = "Unable to open file" + vbNewLine
            _strErrorMessage += "Check file format" + vbNewLine
            _strErrorMessage += "System generated error" + oException.Message
            _oErrorObject = oException
            openAssemblyDrawingDocument = False
        End Try
        'End If
    End Function

    ''' <summary>
    ''' Checks the type of the documet[i.e , part or assembly or drawing]
    ''' </summary>
    ''' <param name="_strPartAssemblyFileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckFileType(ByVal _strPartAssemblyFileName As String) As Long
        Try
            Dim fileData As FileInfo = My.Computer.FileSystem.GetFileInfo(_strPartAssemblyFileName)
            If fileData.Extension = ".SLDASM" Then
                CheckFileType = SwConst.swDocumentTypes_e.swDocASSEMBLY
            ElseIf fileData.Extension = ".SLDPRT" Then
                CheckFileType = SwConst.swDocumentTypes_e.swDocPART
            ElseIf fileData.Extension = ".SLDDRW" Then
                CheckFileType = SwConst.swDocumentTypes_e.swDocDRAWING
            Else
                CheckFileType = SwConst.swDocumentTypes_e.swDocNONE
            End If
            Return CheckFileType
        Catch oException As Exception
            _strErrorMessage = "Access Denied" + vbNewLine
            _strErrorMessage += "File Name:" + _strPartAssemblyFileName + vbNewLine
            _strErrorMessage += "System Generated Error:" + oException.Message
            _oErrorObject = oException
        End Try
    End Function
    ''' <summary>
    ''' SolidWorks session process will be killed in case of unhandled exception case.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub killSolidWorks(ByVal _strProcessName As String)
        Dim proc As System.Diagnostics.Process
        Try
            For Each proc In System.Diagnostics.Process.GetProcessesByName(_strProcessName)
                If proc.HasExited = False Then
                    proc.Kill()
                End If
            Next
        Catch oException As Exception
            _strErrorMessage = "Unable to kill the Service" + vbNewLine
            _strErrorMessage += "System Generated Error" + oException.Message
            _oErrorObject = oException
        End Try
    End Sub
    Public Sub KillAllSolidWorksServices()
        killSolidWorks("SLDWORKS")
        killSolidWorks("SolidWorksLicTemp.0001")
        killSolidWorks("SolidWorksLicensing")
        killSolidWorks("swvbaserver")
    End Sub

    Private Function SaveDocument(ByVal oSolidWorksModel As SldWorks.ModelDoc2) As Boolean
        Try
            SaveDocument = oSolidWorksModel.SaveSilent()
        Catch oException As Exception
            _strErrorMessage = "Unable to kill the Service" + vbNewLine
            _strErrorMessage += "System Generated Error" + oException.Message
            _oErrorObject = oException
        End Try
    End Function
    Public Function SaveDocument(ByVal _strDocName As String) As Boolean
        Try
            'If _htDocumentInstances.ContainsKey(_strDocName) = True Then
            '    SaveDocument = SaveDocument(_htDocumentInstances(_strDocName))
            'End If
            SaveDocument = oSolidWorksModel.SaveSilent()
        Catch oException As Exception
            _strErrorMessage = "Unable to kill the Service" + vbNewLine
            _strErrorMessage += "System Generated Error" + oException.Message
            _oErrorObject = oException
        End Try
    End Function
    '27_07_2009  ragava
    'Public Function Save() As Boolean
    '    Save = False
    '    Try
    '        Save = SolidWorksApplicationObject.IActiveDoc2.SaveSilent()
    '    Catch oException As Exception
    '        _strErrorMessage = "UNABLE TO SAVE THE PART OR ASSEMBLY !!" + vbCrLf
    '        _strErrorMessage += "System generated error:-" + vbCrLf + oException.Message
    '        _oErrorObject = oException
    '    End Try
    'End Function

    Public Function deleteFile(ByVal strfilename As String) As Boolean
        deleteFile = False
        Try
            deleteFile = SolidWorksModel.Extension.SelectByID2(strfilename, "COMPONENT", 0, 0, 0, True, 0, Nothing, 0)
            Return deleteFile = SolidWorksModel.DeleteSelection(False)
        Catch oException As Exception
            _strErrorMessage = "Unable to perform Delete Operation" + vbNewLine
            _strErrorMessage += "System Generated Error:" + vbNewLine + oException.Message
            _oErrorObject = oException
        End Try
    End Function
    Public Function ActivateDocument(ByVal fileName As String) As Boolean
        ActivateDocument = False
        Try
            SolidWorksApplicationObject.ActivateDoc2(fileName, True, 0)
        Catch oException As Exception
            _strErrorMessage = "Unable to Activate the Document" + vbNewLine
            _strErrorMessage += "System generated Error" + vbNewLine + oException.Message
            _oErrorObject = oException
        End Try
    End Function
#End Region
End Class
