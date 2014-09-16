Public Class ThreadingSpeeds
    Private dblThreadDia As Double
    Private dblRPM As Double
    Private strTOOLS As String
    Private dblSecondThread As Double
    Private strODSecondTools As String
    Private dblTh_Per_In As Double
    Private strSpecs As String

    Public Property ThreadDia() As Double
        Get
            Return dblThreadDia
        End Get
        Set(ByVal value As Double)
            dblThreadDia = value
        End Set
    End Property

    Public Property RPM() As Double
        Get
            Return dblRPM
        End Get
        Set(ByVal value As Double)
            dblRPM = value
        End Set
    End Property

    Public Property SecondThread() As Double
        Get
            Return dblSecondThread
        End Get
        Set(ByVal value As Double)
            dblSecondThread = value
        End Set
    End Property

    Public Property Th_Per_In() As Double
        Get
            Return dblTh_Per_In
        End Get
        Set(ByVal value As Double)
            dblTh_Per_In = value
        End Set
    End Property

    Public Property TOOLS() As String
        Get
            Return strTOOLS
        End Get
        Set(ByVal value As String)
            strTOOLS = value
        End Set
    End Property

    Public Property ODSecondTools() As String
        Get
            Return strODSecondTools
        End Get
        Set(ByVal value As String)
            strODSecondTools = value
        End Set
    End Property

    Public Property Specs() As String
        Get
            Return strSpecs
        End Get
        Set(ByVal value As String)
            strSpecs = value
        End Set
    End Property

    Public Sub New(ByVal ThreadDia As Double, ByVal RPM As Double, ByVal TOOLS As String, ByVal SecondThread As Double, ByVal ODSecondTools As String, ByVal Th_Per_In As Double, ByVal Specs As String)
        dblThreadDia = ThreadDia
        dblRPM = RPM
        strTOOLS = TOOLS
        dblSecondThread = SecondThread
        strODSecondTools = ODSecondTools
        dblTh_Per_In = Th_Per_In
        strSpecs = Specs
    End Sub

End Class
