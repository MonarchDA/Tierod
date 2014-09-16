Public Class clsMilLicense
    Private intNoOfDaysFromFile As Integer = 10
    Private dtExpDate As Date

    Private ReadOnly Property MacIDs() As ArrayList
        Get
            MacIDs = New ArrayList
            MacIDs.Add(New Object(1) {"00:19:BB:49:9A:66", "Sandeep"})
            MacIDs.Add(New Object(1) {"00:18:71:6B:D7:AF", "Raghavendra"})
            MacIDs.Add(New Object(1) {"00:19:BB:49:9C:6B", "Purna"})
            MacIDs.Add(New Object(1) {"00:18:FE:6A:6E:5B", "Sadik"})
            MacIDs.Add(New Object(1) {"02:00:54:55:4E:01", "Vista"})
            MacIDs.Add(New Object(1) {"00:19:BB:49:9C:22", "Jinka"})
            Return MacIDs
        End Get
    End Property

    Public Property ExpDate() As Date
        Get
            Return dtExpDate
        End Get
        Set(ByVal value As Date)
            dtExpDate = value
        End Set
    End Property

    Public Function LicenseValidation() As Boolean
        Dim CheckMACId As Boolean
        Dim CheckDays As Boolean
        CheckMACId = GetMAC()
        CheckDays = ISValidDate()
        If CheckDays = True And CheckMACId = True Then
            LicenseValidation = True
        Else
            LicenseValidation = False
        End If
    End Function

    Private Function GetMAC() As Boolean
        Dim moReturn As Management.ManagementObjectCollection
        Dim moSearch As Management.ManagementObjectSearcher
        Dim mo As Management.ManagementObject
        Dim MacId As String
        moSearch = New Management.ManagementObjectSearcher("Select * from Win32_NetworkAdapter where AdapterTypeID = 0")
        Debug.WriteLine("Network")
        moReturn = moSearch.Get

        For Each mo In moReturn
            Try
                MacId = mo("MACaddress").ToString
                For Each oItem As Object In MacIDs
                    If oItem(0) = MacId Then
                        GetMAC = True
                        Exit Function
                    Else
                        GetMAC = False
                    End If
                Next
            Catch
            End Try
        Next
    End Function

    Private Function ISValidDate() As Boolean
        'The purpose of this module is to allow you to place a time
        'limit on the unregistered use of your shareware application.
        'This module can not be defeated by rolling back the system clock.
        'Simply call the DateGood function when your application is first
        'loading, passing it the number of days it can be used without
        'registering.
        '
        'Register Parameters:
        ' CRD: Current Run Date
        ' LRD: Last Run Date
        ' FRD: First Run Date

        Dim TmpCRD As New Date
        Dim TmpLRD As Date
        Dim TmpFRD As Date

        ' TmpCRD = Format(Now, "d/m/yyyy")
        'this parameter will be registered in following registry
        'which we can view by run -------> regedit
        'HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ FolderWithEXEName \ LicParam

        TmpCRD = Date.Today
        Dim ExeName As String

        ExeName = System.Reflection.Assembly.GetExecutingAssembly.Location.Replace(Application.StartupPath & "\", "")
        TmpLRD = GetSetting(ExeName, "MILLicParam", "LRD", "12/12/2006")
        TmpFRD = GetSetting(ExeName, "MILLicParam", "FRD", "12/12/2006")

        'If this is the applications first load, write initial settings
        'to the register
        If TmpLRD = "#12/12/2006#" Then
            SaveSetting(ExeName, "MILLicParam", "LRD", TmpCRD)
            SaveSetting(ExeName, "MILLicParam", "FRD", TmpCRD)
        End If

        'Read LRD and FRD from register
        TmpLRD = GetSetting(ExeName, "MILLicParam", "LRD", "12/12/2006")
        TmpFRD = GetSetting(ExeName, "MILLicParam", "FRD", "12/12/2006")
        Dim ExpDate As Date
        ExpDate = DateAdd("d", intNoOfDaysFromFile, TmpFRD)
        dtExpDate = ExpDate.AddDays(1)
        If TmpFRD > TmpCRD Then 'System clock rolled back
            ISValidDate = False
        ElseIf TmpCRD > ExpDate Then 'Expiration expired
            ISValidDate = False
        ElseIf TmpCRD > TmpLRD Then 'Everything OK write New LRD date
            SaveSetting(System.Reflection.Assembly.GetExecutingAssembly.Location, "CDALicParam", "LRD", TmpCRD)
            ISValidDate = True
        ElseIf TmpCRD = TmpLRD Then
            ISValidDate = True
        Else
            ISValidDate = False
        End If

    End Function

End Class
