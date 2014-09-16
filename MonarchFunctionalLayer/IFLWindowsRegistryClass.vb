Imports Microsoft.Win32.Registry
Imports Microsoft.Win32.RegistryKey
Imports System

''' <summary>
''' This class handles Registry edit funtionality.
''' </summary>
''' <remarks>
''' This class consists functions like IsRegistryExists(),IsRegistryValueExists(), CheckAttributesExists() etc.
''' </remarks>
Public Class IFLWindowsRegistryClass

#Region "Class Variables"

    ''' <summary>
    ''' This variable is an instance of Exception
    ''' </summary>
    ''' <remarks>
    '''  This variable Holds error object araised in Catch block ,and scope of this Variable is Class level
    ''' </remarks>
    Private _oErrorObject As Exception

    ''' <summary>
    ''' This variable stores the error message 
    ''' </summary>
    ''' <remarks>
    ''' The scope of this variable is class level.
    ''' </remarks>
    Private _strErrorMessage As String

#End Region

#Region "Properties"

    ''' <summary>
    ''' This property allows user to access the data holded by variable _strErrorMessage, through out the application
    ''' </summary>
    ''' <returns>
    ''' This property returns value holded by variable _strErrorMessage
    ''' </returns>
    ''' <remarks>
    ''' The Variable _strErrorMessage holds Error Message As string
    ''' </remarks>
    Public ReadOnly Property ErrorMessage() As String
        Get
            Return _strErrorMessage
        End Get
    End Property

    ''' <summary>
    ''' This property Allows user to access the data holded by the variable _oErrorObject through out the appplication
    ''' </summary>
    ''' <returns>
    ''' This Property Returns value holded by variable _oErrorObject
    ''' </returns>
    ''' <remarks>
    ''' This property holds the error object as Exception
    ''' </remarks>
    Public ReadOnly Property ErrorObject() As Exception
        Get
            Return _oErrorObject
        End Get
    End Property

#End Region

#Region "Functions"

    ''' <summary>
    ''' This Function is used to check the resgistry key exists or not in windows Registry.
    ''' </summary>
    ''' <param name="aKeys">
    ''' Parameter akeys as arraylist 
    ''' </param>
    ''' <returns>
    ''' This function Returns True if key exists otherwise Returns False
    ''' </returns>
    ''' <remarks>
    ''' This Function is used by IsRegistryValueExists
    ''' </remarks>
    Public Function IsRegistryExists(ByVal aKeys As ArrayList) As Boolean

        Dim oRegistryKeyObject As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine
        IsRegistryExists = False
        Try
            For Each strKeyname As String In aKeys
                oRegistryKeyObject = oRegistryKeyObject.OpenSubKey(strKeyname, True)
                IsRegistryExists = (Not IsNothing(oRegistryKeyObject))
                If Not IsRegistryExists Then
                    Exit For
                End If
            Next
        Catch ex As Exception
            _strErrorMessage = ex.Message
            _oErrorObject = ex
            IsRegistryExists = False
        End Try
    End Function

    ''' <summary>
    ''' This Function is used to check whether registry value is exists or not in registry edit
    ''' </summary>
    ''' <param name="aKeys">
    ''' Parameter aKeys as Arraylist
    ''' </param>
    ''' <param name="aAttributes">
    ''' Parameter aAttributes as Arraylist which holds Attributes like server,database,userid,password as parameters
    ''' </param>
    ''' <returns>
    ''' Returns True if Value exists in Registry Edit otherwise Return False
    ''' </returns>
    ''' <remarks>
    ''' This Function Calls IsRegistryExists,CheckAttributesExists Functions
    ''' </remarks>
    Public Function IsRegistryValueExists(ByVal aKeys As ArrayList, ByVal aAttributes As ArrayList) As Boolean

        IsRegistryValueExists = IsRegistryExists(aKeys)
        If IsRegistryValueExists Then
            Dim oRegistryKeyObject As Microsoft.Win32.RegistryKey = GetRegistryKey(aKeys)
            If Not IsNothing(oRegistryKeyObject) Then
                IsRegistryValueExists = CheckAttributesExists(aAttributes, oRegistryKeyObject)
            Else
                IsRegistryValueExists = False
            End If
        End If

    End Function

    ''' <summary>
    '''  This Function is used to check whether Attribute exists or not in Registry Edit.
    ''' </summary>
    ''' <param name="aAttributes">
    ''' Parameter aAttributes as array list holds attributes like server,database,userid,password as parameters
    ''' </param>
    ''' <param name="oRegistryKeyObject">
    ''' Parameter oRegistryKeyObject as arraylist 
    ''' </param>
    ''' <returns>
    ''' This function Returns true if Attribute not found otherwise return False
    ''' </returns>
    ''' <remarks>
    ''' This Function  is called by IsRegistryValueExists 
    ''' </remarks>
    Private Function CheckAttributesExists(ByVal aAttributes As ArrayList, ByVal oRegistryKeyObject As Microsoft.Win32.RegistryKey) As Boolean

        CheckAttributesExists = False
        Dim IsArrtibuteFound As Boolean = False
        Dim aRegistryValueNames() As String = oRegistryKeyObject.GetValueNames
        If aRegistryValueNames.Length > 0 Then
            For Each strAttributeName As String In aAttributes
                IsArrtibuteFound = False
                For Each strRegName As String In aRegistryValueNames
                    If strAttributeName.Trim.ToUpper.Equals(strRegName.Trim.ToUpper) Then
                        IsArrtibuteFound = True
                        Exit For
                    End If
                Next
                If Not IsArrtibuteFound Then
                    Exit For
                End If
            Next
        End If

        CheckAttributesExists = IsArrtibuteFound

    End Function

    ''' <summary>
    '''  This Function is used to get the value of the respective attribute.
    ''' </summary>
    ''' <param name="aKeys">
    ''' Parameter aKeys as arraylist
    ''' </param>
    ''' <param name="strAttributeName">
    ''' Parameter strAttributeName as string
    ''' </param>
    ''' <returns>
    ''' This function Returns Value of Attribute Name
    ''' </returns>
    ''' <remarks>
    ''' This Function is used in frmMainFormto Retrive the Value of the Attribute 
    ''' </remarks>
    Public Function GetValues(ByVal aKeys As ArrayList, ByVal strAttributeName As String) As Object

        GetValues = Nothing
        If IsRegistryExists(aKeys) Then
            Dim oRegistryKeyObject As Microsoft.Win32.RegistryKey = GetRegistryKey(aKeys)
            GetValues = oRegistryKeyObject.GetValue(strAttributeName)
        End If

    End Function

    ''' <summary>
    ''' This Function get the registry key 
    ''' </summary>
    ''' <param name="aKeys">
    ''' Parameter aKeys as ArrayList
    ''' </param>
    ''' <returns>
    ''' This function Returns Microsoft.Win32.RegistryKey
    ''' </returns>
    ''' <remarks>
    ''' This Function is used by IsRegistryValueExists,GetValues
    '''  </remarks>
    Private Function GetRegistryKey(ByVal aKeys As ArrayList) As Microsoft.Win32.RegistryKey

        GetRegistryKey = Microsoft.Win32.Registry.LocalMachine
        Dim IsRegistryExists As Boolean = False
        Try
            For Each strKeyname As String In aKeys
                GetRegistryKey = GetRegistryKey.OpenSubKey(strKeyname, True)
                IsRegistryExists = (Not IsNothing(GetRegistryKey))
                If Not IsRegistryExists Then
                    Exit For
                End If
            Next
        Catch ex As Exception
            _strErrorMessage = ex.Message
            _oErrorObject = ex
            IsRegistryExists = False
        End Try

    End Function

    ''' <summary>
    ''' This Function is save's the registry in the windows register
    ''' </summary>
    ''' <param name="strKey">
    ''' Parameter strKey as String 
    ''' </param>
    ''' <param name="strSubKey">
    ''' Parameter strSubKey as String
    ''' </param>
    ''' <param name="aArributeNamesValues">
    ''' Parameter aArributeNamesValues as string
    ''' </param>
    ''' <returns>
    '''This function Returns True if data is saved otherwise Returns False
    ''' </returns>
    ''' <remarks>
    ''' This Function is used by frmBDLogin form
    ''' </remarks>
    Public Function RegistryDataSave(ByVal strKey As String, ByVal strSubKey As String, ByVal aArributeNamesValues As ArrayList) As Boolean
        RegistryDataSave = False
        Try
            Dim newkey As Microsoft.Win32.RegistryKey
            newkey = CreateRegisterKey(strKey, strSubKey)
            For Each oAttribueNameValue As Object In aArributeNamesValues
                newkey.SetValue(oAttribueNameValue(0), oAttribueNameValue(1))
            Next
            RegistryDataSave = True
        Catch ex As Exception
            _strErrorMessage = ex.Message
            _oErrorObject = ex
        End Try

    End Function

    ''' <summary>
    ''' This function is used to create a registry key in the windows registry.
    ''' </summary>
    ''' <param name="strKey">
    ''' Parameter strKey as Parameter
    ''' </param>
    ''' <param name="strSubKey">
    ''' Parameter strSubKey as Parameter
    ''' </param>
    ''' <returns>
    ''' This function Returns Key as object
    ''' </returns>
    ''' <remarks>
    ''' This function is used by RegistryDataSave
    ''' </remarks>
    Public Function CreateRegisterKey(ByVal strKey As String, ByVal strSubKey As String) As Object
        Dim oRegistry As Microsoft.Win32.Registry = Nothing
        CreateRegisterKey = False
        Try
            Dim okey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(strKey, True)
            CreateRegisterKey = okey.CreateSubKey(strSubKey)
        Catch ex As Exception
            _strErrorMessage = ex.Message
            _oErrorObject = ex
        End Try
    End Function
#End Region

End Class

