Public Module ModuleGeneratedModelNames

    Private _ArrayListModelName As New ArrayList
    Private _IsGenerateBtnClicked As Boolean = False

#Region "Properties"

    Public Property ArrayListModelName() As ArrayList
        Get
            Return _ArrayListModelName
        End Get
        Set(ByVal value As ArrayList)
            _ArrayListModelName = value
        End Set
    End Property

    Public Property IsGenerateBtnClicked() As Boolean
        Get
            Return _IsGenerateBtnClicked
        End Get
        Set(ByVal value As Boolean)
            _IsGenerateBtnClicked = value
        End Set
    End Property
#End Region

End Module
