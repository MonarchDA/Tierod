Public Module Module1            'SUGANDHI

    Private oReadValuesFromExcel As ReadValuesFromExcel
    Private oBtnBrowseClicked As Boolean = False
    Private oPinSizeDetailsDataTable As New DataTable

    Private oLogInfo As New ArrayList
    Private oArraList1 As New ArrayList
    Private oArraList2 As New ArrayList
    Private oArraList3 As New ArrayList
    Private oArraList4 As New ArrayList
    Private oArraList5 As New ArrayList
    Private oArraList6 As New ArrayList
    Private oArraList7 As New ArrayList
    Private oArraList8 As New ArrayList
    Private oArraList9 As New ArrayList
    Private oArraList10 As New ArrayList
    Private oArraList11 As New ArrayList
    Private oArraList12 As New ArrayList
    Private oArraList13 As New ArrayList
    Private oArraList14 As New ArrayList
    Private oArraList15 As New ArrayList
    Private oArraList16 As New ArrayList
    Private oArraList17 As New ArrayList
    Private oArraList18 As New ArrayList
    Private oArraList19 As New ArrayList
    Private oArraList20 As New ArrayList
    Private oArraList21 As New ArrayList
    Private oArraList22 As New ArrayList
    Private oArraList23 As New ArrayList
    Private oArraList24 As New ArrayList
    Private oArraList25 As New ArrayList
    Private oArraList26 As New ArrayList
    Private oArraList27 As New ArrayList
    Private oArraList28 As New ArrayList
    Private oArraList29 As New ArrayList
    Private oArraList30 As New ArrayList
    Private oArraList31 As New ArrayList
    Private oArraList32 As New ArrayList
    Private oArraList33 As New ArrayList
    Private oArraList34 As New ArrayList
    Private oArraList35 As New ArrayList
    Private oArraList36 As New ArrayList
    Private oArraList37 As New ArrayList
    Private oArraList38 As New ArrayList
    Private oArraList39 As New ArrayList
    Private oArraList40 As New ArrayList
    Private oArraList41 As New ArrayList
    Private oArraList42 As New ArrayList
    Private oArraList43 As New ArrayList
    Private oArraList44 As New ArrayList
    Private oArraList45 As New ArrayList
    Private oArraList46 As New ArrayList
    Private oArraList47 As New ArrayList
    Private oArraList48 As New ArrayList
    Private oArraList49 As New ArrayList
    Private oArraList50 As New ArrayList
    Private oArraList51 As New ArrayList
    Private oArraList52 As New ArrayList
    Private oArraList53 As New ArrayList
    Private oArraList54 As New ArrayList
    Private oArraList55 As New ArrayList
    Private oArraList56 As New ArrayList
    Private oArraList57 As New ArrayList
    Private oArraList58 As New ArrayList
    Private oArraList59 As New ArrayList
    Private oArraList60 As New ArrayList
    Private oArraList61 As New ArrayList
    Private oArraList62 As New ArrayList
    Private oArraList63 As New ArrayList
    Private oArraList64 As New ArrayList
    Private oArraList65 As New ArrayList
    Private oArraList66 As New ArrayList
    Private oArraList67 As New ArrayList
    Private oArraList68 As New ArrayList
    Private oArraList69 As New ArrayList
    Private oArraList70 As New ArrayList
    Private oArraList71 As New ArrayList
    Private oArraList72 As New ArrayList
    Private oArraList73 As New ArrayList
    Private oArraList74 As New ArrayList
    Private oArraList75 As New ArrayList
    Private oArraList76 As New ArrayList
    Private oArraList77 As New ArrayList
    Private oArraList78 As New ArrayList
    Private oArraList79 As New ArrayList
    Private oArraList80 As New ArrayList
    Private oArraList81 As New ArrayList
    Private oArraList82 As New ArrayList
    Private oArraList83 As New ArrayList
    Private oArraList84 As New ArrayList
    Private oArraList85 As New ArrayList
    Private oArraList86 As New ArrayList
    Private oArraList87 As New ArrayList
    Private oArraList88 As New ArrayList
    Private oArraList89 As New ArrayList
    Private oArraList90 As New ArrayList
    Private oArraList91 As New ArrayList
    Private oArraList92 As New ArrayList
    Private oArraList93 As New ArrayList
    Private oArraList94 As New ArrayList
    Private oArraList95 As New ArrayList
    Private oArraList96 As New ArrayList
    Private oArraList97 As New ArrayList
    Private oArraList98 As New ArrayList
    Private oArraList99 As New ArrayList
    Private oArraList100 As New ArrayList

    Private _RowCount As Integer = 0

    Public Property BtnBrowseClicked() As Boolean
        Get
            Return oBtnBrowseClicked
        End Get
        Set(ByVal value As Boolean)
            oBtnBrowseClicked = value
        End Set
    End Property

    Public Property PinSizeDetailsDataTable() As DataTable
        Get
            Return oPinSizeDetailsDataTable
        End Get
        Set(ByVal value As DataTable)
            oPinSizeDetailsDataTable = value
        End Set
    End Property

    Public Property ReadValuesFromExcel() As ReadValuesFromExcel
        Get
            Return oReadValuesFromExcel
        End Get
        Set(ByVal value As ReadValuesFromExcel)
            oReadValuesFromExcel = value
        End Set
    End Property

    Public Property LogInfo() As ArrayList
        Get
            Return oLogInfo
        End Get
        Set(ByVal value As ArrayList)
            oLogInfo = value
        End Set
    End Property

    Public Property RowCount() As Integer
        Get
            Return _RowCount
        End Get
        Set(ByVal value As Integer)
            _RowCount = value
        End Set
    End Property

    Public Property ArraList1() As ArrayList
        Get
            Return oArraList1
        End Get
        Set(ByVal value As ArrayList)
            oArraList1 = value
        End Set
    End Property

    Public Property ArraList2() As ArrayList
        Get
            Return oArraList2
        End Get
        Set(ByVal value As ArrayList)
            oArraList2 = value
        End Set
    End Property

    Public Property ArraList3() As ArrayList
        Get
            Return oArraList3
        End Get
        Set(ByVal value As ArrayList)
            oArraList3 = value
        End Set
    End Property

    Public Property ArraList4() As ArrayList
        Get
            Return oArraList4
        End Get
        Set(ByVal value As ArrayList)
            oArraList4 = value
        End Set
    End Property

    Public Property ArraList5() As ArrayList
        Get
            Return oArraList5
        End Get
        Set(ByVal value As ArrayList)
            oArraList5 = value
        End Set
    End Property

    Public Property ArraList6() As ArrayList
        Get
            Return oArraList6
        End Get
        Set(ByVal value As ArrayList)
            oArraList6 = value
        End Set
    End Property

    Public Property ArraList7() As ArrayList
        Get
            Return oArraList7
        End Get
        Set(ByVal value As ArrayList)
            oArraList7 = value
        End Set
    End Property

    Public Property ArraList8() As ArrayList
        Get
            Return oArraList8
        End Get
        Set(ByVal value As ArrayList)
            oArraList8 = value
        End Set
    End Property

    Public Property ArraList9() As ArrayList
        Get
            Return oArraList9
        End Get
        Set(ByVal value As ArrayList)
            oArraList9 = value
        End Set
    End Property

    Public Property ArraList10() As ArrayList
        Get
            Return oArraList10
        End Get
        Set(ByVal value As ArrayList)
            oArraList10 = value
        End Set
    End Property

    Public Property ArraList11() As ArrayList
        Get
            Return oArraList11
        End Get
        Set(ByVal value As ArrayList)
            oArraList11 = value
        End Set
    End Property

    Public Property ArraList12() As ArrayList
        Get
            Return oArraList12
        End Get
        Set(ByVal value As ArrayList)
            oArraList12 = value
        End Set
    End Property

    Public Property ArraList13() As ArrayList
        Get
            Return oArraList13
        End Get
        Set(ByVal value As ArrayList)
            oArraList13 = value
        End Set
    End Property

    Public Property ArraList14() As ArrayList
        Get
            Return oArraList14
        End Get
        Set(ByVal value As ArrayList)
            oArraList14 = value
        End Set
    End Property

    Public Property ArraList15() As ArrayList
        Get
            Return oArraList15
        End Get
        Set(ByVal value As ArrayList)
            oArraList15 = value
        End Set
    End Property

    Public Property ArraList16() As ArrayList
        Get
            Return oArraList16
        End Get
        Set(ByVal value As ArrayList)
            oArraList16 = value
        End Set
    End Property

    Public Property ArraList17() As ArrayList
        Get
            Return oArraList17
        End Get
        Set(ByVal value As ArrayList)
            oArraList17 = value
        End Set
    End Property

    Public Property ArraList18() As ArrayList
        Get
            Return oArraList18
        End Get
        Set(ByVal value As ArrayList)
            oArraList18 = value
        End Set
    End Property

    Public Property ArraList19() As ArrayList
        Get
            Return oArraList19
        End Get
        Set(ByVal value As ArrayList)
            oArraList19 = value
        End Set
    End Property

    Public Property ArraList20() As ArrayList
        Get
            Return oArraList20
        End Get
        Set(ByVal value As ArrayList)
            oArraList20 = value
        End Set
    End Property

    Public Property ArraList21() As ArrayList
        Get
            Return oArraList21
        End Get
        Set(ByVal value As ArrayList)
            oArraList21 = value
        End Set
    End Property

    Public Property ArraList22() As ArrayList
        Get
            Return oArraList22
        End Get
        Set(ByVal value As ArrayList)
            oArraList22 = value
        End Set
    End Property

    Public Property ArraList23() As ArrayList
        Get
            Return oArraList23
        End Get
        Set(ByVal value As ArrayList)
            oArraList23 = value
        End Set
    End Property

    Public Property ArraList24() As ArrayList
        Get
            Return oArraList24
        End Get
        Set(ByVal value As ArrayList)
            oArraList24 = value
        End Set
    End Property

    Public Property ArraList25() As ArrayList
        Get
            Return oArraList25
        End Get
        Set(ByVal value As ArrayList)
            oArraList25 = value
        End Set
    End Property

    Public Property ArraList26() As ArrayList
        Get
            Return oArraList26
        End Get
        Set(ByVal value As ArrayList)
            oArraList26 = value
        End Set
    End Property

    Public Property ArraList27() As ArrayList
        Get
            Return oArraList27
        End Get
        Set(ByVal value As ArrayList)
            oArraList27 = value
        End Set
    End Property

    Public Property ArraList28() As ArrayList
        Get
            Return oArraList28
        End Get
        Set(ByVal value As ArrayList)
            oArraList28 = value
        End Set
    End Property

    Public Property ArraList29() As ArrayList
        Get
            Return oArraList29
        End Get
        Set(ByVal value As ArrayList)
            oArraList29 = value
        End Set
    End Property

    Public Property ArraList30() As ArrayList
        Get
            Return oArraList30
        End Get
        Set(ByVal value As ArrayList)
            oArraList30 = value
        End Set
    End Property

    Public Property ArraList31() As ArrayList
        Get
            Return oArraList31
        End Get
        Set(ByVal value As ArrayList)
            oArraList31 = value
        End Set
    End Property

    Public Property ArraList32() As ArrayList
        Get
            Return oArraList32
        End Get
        Set(ByVal value As ArrayList)
            oArraList32 = value
        End Set
    End Property

    Public Property ArraList33() As ArrayList
        Get
            Return oArraList33
        End Get
        Set(ByVal value As ArrayList)
            oArraList33 = value
        End Set
    End Property

    Public Property ArraList34() As ArrayList
        Get
            Return oArraList34
        End Get
        Set(ByVal value As ArrayList)
            oArraList34 = value
        End Set
    End Property

    Public Property ArraList35() As ArrayList
        Get
            Return oArraList35
        End Get
        Set(ByVal value As ArrayList)
            oArraList35 = value
        End Set
    End Property

    Public Property ArraList36() As ArrayList
        Get
            Return oArraList36
        End Get
        Set(ByVal value As ArrayList)
            oArraList36 = value
        End Set
    End Property

    Public Property ArraList37() As ArrayList
        Get
            Return oArraList37
        End Get
        Set(ByVal value As ArrayList)
            oArraList37 = value
        End Set
    End Property

    Public Property ArraList38() As ArrayList
        Get
            Return oArraList38
        End Get
        Set(ByVal value As ArrayList)
            oArraList38 = value
        End Set
    End Property

    Public Property ArraList39() As ArrayList
        Get
            Return oArraList39
        End Get
        Set(ByVal value As ArrayList)
            oArraList39 = value
        End Set
    End Property

    Public Property ArraList40() As ArrayList
        Get
            Return oArraList40
        End Get
        Set(ByVal value As ArrayList)
            oArraList40 = value
        End Set
    End Property

    Public Property ArraList41() As ArrayList
        Get
            Return oArraList41
        End Get
        Set(ByVal value As ArrayList)
            oArraList41 = value
        End Set
    End Property

    Public Property ArraList42() As ArrayList
        Get
            Return oArraList42
        End Get
        Set(ByVal value As ArrayList)
            oArraList42 = value
        End Set
    End Property

    Public Property ArraList43() As ArrayList
        Get
            Return oArraList43
        End Get
        Set(ByVal value As ArrayList)
            oArraList43 = value
        End Set
    End Property

    Public Property ArraList44() As ArrayList
        Get
            Return oArraList44
        End Get
        Set(ByVal value As ArrayList)
            oArraList44 = value
        End Set
    End Property

    Public Property ArraList45() As ArrayList
        Get
            Return oArraList45
        End Get
        Set(ByVal value As ArrayList)
            oArraList45 = value
        End Set
    End Property

    Public Property ArraList46() As ArrayList
        Get
            Return oArraList46
        End Get
        Set(ByVal value As ArrayList)
            oArraList46 = value
        End Set
    End Property

    Public Property ArraList47() As ArrayList
        Get
            Return oArraList47
        End Get
        Set(ByVal value As ArrayList)
            oArraList47 = value
        End Set
    End Property

    Public Property ArraList48() As ArrayList
        Get
            Return oArraList48
        End Get
        Set(ByVal value As ArrayList)
            oArraList48 = value
        End Set
    End Property

    Public Property ArraList49() As ArrayList
        Get
            Return oArraList49
        End Get
        Set(ByVal value As ArrayList)
            oArraList49 = value
        End Set
    End Property

    Public Property ArraList50() As ArrayList
        Get
            Return oArraList50
        End Get
        Set(ByVal value As ArrayList)
            oArraList50 = value
        End Set
    End Property

    Public Property ArraList51() As ArrayList
        Get
            Return oArraList51
        End Get
        Set(ByVal value As ArrayList)
            oArraList51 = value
        End Set
    End Property

    Public Property ArraList52() As ArrayList
        Get
            Return oArraList52
        End Get
        Set(ByVal value As ArrayList)
            oArraList52 = value
        End Set
    End Property

    Public Property ArraList53() As ArrayList
        Get
            Return oArraList53
        End Get
        Set(ByVal value As ArrayList)
            oArraList53 = value
        End Set
    End Property

    Public Property ArraList54() As ArrayList
        Get
            Return oArraList54
        End Get
        Set(ByVal value As ArrayList)
            oArraList54 = value
        End Set
    End Property

    Public Property ArraList55() As ArrayList
        Get
            Return oArraList55
        End Get
        Set(ByVal value As ArrayList)
            oArraList55 = value
        End Set
    End Property

    Public Property ArraList56() As ArrayList
        Get
            Return oArraList56
        End Get
        Set(ByVal value As ArrayList)
            oArraList56 = value
        End Set
    End Property

    Public Property ArraList57() As ArrayList
        Get
            Return oArraList57
        End Get
        Set(ByVal value As ArrayList)
            oArraList57 = value
        End Set
    End Property

    Public Property ArraList58() As ArrayList
        Get
            Return oArraList58
        End Get
        Set(ByVal value As ArrayList)
            oArraList58 = value
        End Set
    End Property

    Public Property ArraList59() As ArrayList
        Get
            Return oArraList59
        End Get
        Set(ByVal value As ArrayList)
            oArraList59 = value
        End Set
    End Property

    Public Property ArraList60() As ArrayList
        Get
            Return oArraList60
        End Get
        Set(ByVal value As ArrayList)
            oArraList60 = value
        End Set
    End Property

    Public Property ArraList61() As ArrayList
        Get
            Return oArraList61
        End Get
        Set(ByVal value As ArrayList)
            oArraList61 = value
        End Set
    End Property

    Public Property ArraList62() As ArrayList
        Get
            Return oArraList62
        End Get
        Set(ByVal value As ArrayList)
            oArraList62 = value
        End Set
    End Property

    Public Property ArraList63() As ArrayList
        Get
            Return oArraList63
        End Get
        Set(ByVal value As ArrayList)
            oArraList63 = value
        End Set
    End Property

    Public Property ArraList64() As ArrayList
        Get
            Return oArraList64
        End Get
        Set(ByVal value As ArrayList)
            oArraList64 = value
        End Set
    End Property

    Public Property ArraList65() As ArrayList
        Get
            Return oArraList65
        End Get
        Set(ByVal value As ArrayList)
            oArraList65 = value
        End Set
    End Property

    Public Property ArraList66() As ArrayList
        Get
            Return oArraList66
        End Get
        Set(ByVal value As ArrayList)
            oArraList66 = value
        End Set
    End Property

    Public Property ArraList67() As ArrayList
        Get
            Return oArraList67
        End Get
        Set(ByVal value As ArrayList)
            oArraList67 = value
        End Set
    End Property

    Public Property ArraList68() As ArrayList
        Get
            Return oArraList68
        End Get
        Set(ByVal value As ArrayList)
            oArraList68 = value
        End Set
    End Property

    Public Property ArraList69() As ArrayList
        Get
            Return oArraList69
        End Get
        Set(ByVal value As ArrayList)
            oArraList69 = value
        End Set
    End Property

    Public Property ArraList70() As ArrayList
        Get
            Return oArraList70
        End Get
        Set(ByVal value As ArrayList)
            oArraList70 = value
        End Set
    End Property

    Public Property ArraList71() As ArrayList
        Get
            Return oArraList71
        End Get
        Set(ByVal value As ArrayList)
            oArraList71 = value
        End Set
    End Property

    Public Property ArraList72() As ArrayList
        Get
            Return oArraList72
        End Get
        Set(ByVal value As ArrayList)
            oArraList72 = value
        End Set
    End Property

    Public Property ArraList73() As ArrayList
        Get
            Return oArraList73
        End Get
        Set(ByVal value As ArrayList)
            oArraList73 = value
        End Set
    End Property

    Public Property ArraList74() As ArrayList
        Get
            Return oArraList74
        End Get
        Set(ByVal value As ArrayList)
            oArraList74 = value
        End Set
    End Property

    Public Property ArraList75() As ArrayList
        Get
            Return oArraList75
        End Get
        Set(ByVal value As ArrayList)
            oArraList75 = value
        End Set
    End Property

    Public Property ArraList76() As ArrayList
        Get
            Return oArraList76
        End Get
        Set(ByVal value As ArrayList)
            oArraList76 = value
        End Set
    End Property

    Public Property ArraList77() As ArrayList
        Get
            Return oArraList77
        End Get
        Set(ByVal value As ArrayList)
            oArraList77 = value
        End Set
    End Property

    Public Property ArraList78() As ArrayList
        Get
            Return oArraList78
        End Get
        Set(ByVal value As ArrayList)
            oArraList78 = value
        End Set
    End Property

    Public Property ArraList79() As ArrayList
        Get
            Return oArraList79
        End Get
        Set(ByVal value As ArrayList)
            oArraList79 = value
        End Set
    End Property

    Public Property ArraList80() As ArrayList
        Get
            Return oArraList80
        End Get
        Set(ByVal value As ArrayList)
            oArraList80 = value
        End Set
    End Property

    Public Property ArraList81() As ArrayList
        Get
            Return oArraList81
        End Get
        Set(ByVal value As ArrayList)
            oArraList81 = value
        End Set
    End Property

    Public Property ArraList82() As ArrayList
        Get
            Return oArraList82
        End Get
        Set(ByVal value As ArrayList)
            oArraList82 = value
        End Set
    End Property

    Public Property ArraList83() As ArrayList
        Get
            Return oArraList83
        End Get
        Set(ByVal value As ArrayList)
            oArraList83 = value
        End Set
    End Property

    Public Property ArraList84() As ArrayList
        Get
            Return oArraList84
        End Get
        Set(ByVal value As ArrayList)
            oArraList84 = value
        End Set
    End Property

    Public Property ArraList85() As ArrayList
        Get
            Return oArraList85
        End Get
        Set(ByVal value As ArrayList)
            oArraList85 = value
        End Set
    End Property

    Public Property ArraList86() As ArrayList
        Get
            Return oArraList86
        End Get
        Set(ByVal value As ArrayList)
            oArraList86 = value
        End Set
    End Property

    Public Property ArraList87() As ArrayList
        Get
            Return oArraList87
        End Get
        Set(ByVal value As ArrayList)
            oArraList87 = value
        End Set
    End Property

    Public Property ArraList88() As ArrayList
        Get
            Return oArraList88
        End Get
        Set(ByVal value As ArrayList)
            oArraList88 = value
        End Set
    End Property

    Public Property ArraList89() As ArrayList
        Get
            Return oArraList89
        End Get
        Set(ByVal value As ArrayList)
            oArraList89 = value
        End Set
    End Property

    Public Property ArraList90() As ArrayList
        Get
            Return oArraList90
        End Get
        Set(ByVal value As ArrayList)
            oArraList90 = value
        End Set
    End Property

    Public Property ArraList91() As ArrayList
        Get
            Return oArraList91
        End Get
        Set(ByVal value As ArrayList)
            oArraList91 = value
        End Set
    End Property

    Public Property ArraList92() As ArrayList
        Get
            Return oArraList92
        End Get
        Set(ByVal value As ArrayList)
            oArraList92 = value
        End Set
    End Property

    Public Property ArraList93() As ArrayList
        Get
            Return oArraList93
        End Get
        Set(ByVal value As ArrayList)
            oArraList93 = value
        End Set
    End Property

    Public Property ArraList94() As ArrayList
        Get
            Return oArraList94
        End Get
        Set(ByVal value As ArrayList)
            oArraList94 = value
        End Set
    End Property

    Public Property ArraList95() As ArrayList
        Get
            Return oArraList95
        End Get
        Set(ByVal value As ArrayList)
            oArraList95 = value
        End Set
    End Property

    Public Property ArraList96() As ArrayList
        Get
            Return oArraList96
        End Get
        Set(ByVal value As ArrayList)
            oArraList96 = value
        End Set
    End Property

    Public Property ArraList97() As ArrayList
        Get
            Return oArraList97
        End Get
        Set(ByVal value As ArrayList)
            oArraList97 = value
        End Set
    End Property

    Public Property ArraList98() As ArrayList
        Get
            Return oArraList98
        End Get
        Set(ByVal value As ArrayList)
            oArraList98 = value
        End Set
    End Property

    Public Property ArraList99() As ArrayList
        Get
            Return oArraList99
        End Get
        Set(ByVal value As ArrayList)
            oArraList99 = value
        End Set
    End Property

    Public Property ArraList100() As ArrayList
        Get
            Return oArraList100
        End Get
        Set(ByVal value As ArrayList)
            oArraList100 = value
        End Set
    End Property

End Module
