Public Class CYL_Rod
    Private _LargeDia As Double
    Private _SmallDia As Double
    Private _Length As Double
    Private _Th_Length As Double
    Private _TH_Per_IN As Double
    Private _2ndthreadnum As Double
    Private _2ndthreadlength As Double
    Private _ShoulderType As String
    Private _RodType As String
    Private _PartNo As String
    Private _ProgNo As String
    Private _ByName As String
    Private _Description As String
    Private _Drawing_Num As Double
    Private _Xhome As Double
    Private _Zhome As Double
    Private _NominalThreadDia As Double
    Private _Operation As Double
    Private _WorkCenter As Double
    Private _AutoDoor As Boolean
    Private _output2ndop As Boolean
    Private _2ndthreaddia As Double
    Private _Drawing_Rev As Double
    Private _2ndoptype As String
    Private _2ndopzzero As Double
    Private _2ndthreadcornerrad As Double
    Private _2ndshoulder As Double
    Private _2ndchamfer As Double
    Private _skimlength As Double
    Private _skimdiameter As Double
    Private _chamferdepthofcut As Double


    Public Property LargeDia() As Double
        Get
            Return _LargeDia
        End Get
        Set(ByVal value As Double)
            _LargeDia = value
        End Set
    End Property

    Public Property SmallDia() As Double
        Get
            Return _SmallDia
        End Get
        Set(ByVal value As Double)
            _SmallDia = value
        End Set
    End Property

    Public Property Length() As Double
        Get
            Return _Length
        End Get
        Set(ByVal value As Double)
            _Length = value
        End Set
    End Property

    Public Property Th_Length() As Double
        Get
            Return _Th_Length
        End Get
        Set(ByVal value As Double)
            _Th_Length = value
        End Set
    End Property

    Public Property TH_Per_IN() As Double
        Get
            Return _TH_Per_IN
        End Get
        Set(ByVal value As Double)
            _TH_Per_IN = value
        End Set
    End Property

    Public Property Secondthreadnum() As Double
        Get
            Return _2ndthreadnum
        End Get
        Set(ByVal value As Double)
            _2ndthreadnum = value
        End Set
    End Property

    Public Property Secondthreadlength() As Double
        Get
            Return _2ndthreadlength
        End Get
        Set(ByVal value As Double)
            _2ndthreadlength = value
        End Set
    End Property

    Public Property Drawing_Num() As Double
        Get
            Return _Drawing_Num
        End Get
        Set(ByVal value As Double)
            _Drawing_Num = value
        End Set
    End Property

    Public Property Xhome() As Double
        Get
            Return _Xhome
        End Get
        Set(ByVal value As Double)
            _Xhome = value
        End Set
    End Property

    Public Property Zhome() As Double
        Get
            Return _Zhome
        End Get
        Set(ByVal value As Double)
            _Zhome = value
        End Set
    End Property

    Public Property NominalThreadDia() As Double
        Get
            Return _NominalThreadDia
        End Get
        Set(ByVal value As Double)
            _NominalThreadDia = value
        End Set
    End Property

    Public Property Operation() As Double
        Get
            Return _Operation
        End Get
        Set(ByVal value As Double)
            _Operation = value
        End Set
    End Property

    Public Property WorkCenter() As Double
        Get
            Return _WorkCenter
        End Get
        Set(ByVal value As Double)
            _WorkCenter = value
        End Set
    End Property

    Public Property Secondthreaddia() As Double
        Get
            Return _2ndthreaddia
        End Get
        Set(ByVal value As Double)
            _2ndthreaddia = value
        End Set
    End Property

    Public Property Drawing_Rev() As Double
        Get
            Return _Drawing_Rev
        End Get
        Set(ByVal value As Double)
            _Drawing_Rev = value
        End Set
    End Property

    Public Property Secondopzzero() As Double
        Get
            Return _2ndopzzero
        End Get
        Set(ByVal value As Double)
            _2ndopzzero = value
        End Set
    End Property

    Public Property Secondthreadcornerrad() As Double
        Get
            Return _2ndthreadcornerrad
        End Get
        Set(ByVal value As Double)
            _2ndthreadcornerrad = value
        End Set
    End Property

    Public Property Secondshoulder() As Double
        Get
            Return _2ndshoulder
        End Get
        Set(ByVal value As Double)
            _2ndshoulder = value
        End Set
    End Property

    Public Property Secondchamfer() As Double
        Get
            Return _2ndchamfer
        End Get
        Set(ByVal value As Double)
            _2ndchamfer = value
        End Set
    End Property

    Public Property skimlength() As Double
        Get
            Return _skimlength
        End Get
        Set(ByVal value As Double)
            _skimlength = value
        End Set
    End Property

    Public Property skimdiameter() As Double
        Get
            Return _skimdiameter
        End Get
        Set(ByVal value As Double)
            _skimdiameter = value
        End Set
    End Property

    Public Property chamferdepthofcut() As Double
        Get
            Return _chamferdepthofcut
        End Get
        Set(ByVal value As Double)
            _chamferdepthofcut = value
        End Set
    End Property

    Public Property AutoDoor() As Boolean
        Get
            Return _AutoDoor
        End Get
        Set(ByVal value As Boolean)
            _AutoDoor = value
        End Set
    End Property

    Public Property output2ndop() As Boolean
        Get
            Return _output2ndop
        End Get
        Set(ByVal value As Boolean)
            _output2ndop = value
        End Set
    End Property

    Public Property ShoulderType() As String
        Get
            Return _ShoulderType
        End Get
        Set(ByVal value As String)
            _ShoulderType = value
        End Set
    End Property

    Public Property RodType() As String
        Get
            Return _RodType
        End Get
        Set(ByVal value As String)
            _RodType = value
        End Set
    End Property

    Public Property PartNo() As String
        Get
            Return _PartNo
        End Get
        Set(ByVal value As String)
            _PartNo = value
            _ProgNo = "0" & _PartNo & "1.MIN"
        End Set
    End Property

    Public ReadOnly Property ProgNo() As String
        Get
            Return _ProgNo
        End Get
    End Property

    Public Property ByName() As String
        Get
            Return _ByName
        End Get
        Set(ByVal value As String)
            _ByName = value
        End Set
    End Property

    Public ReadOnly Property Description() As String
        Get
            Return _Description
        End Get
    End Property

    Public Property Secondoptype() As String
        Get
            Return _2ndoptype
        End Get
        Set(ByVal value As String)
            _2ndoptype = value
        End Set
    End Property

    Public Function GetYesOrNo(ByVal IsTrue As Boolean) As String
        If IsTrue Then
            Return "Yes"
        Else
            Return "No"
        End If
    End Function

    Public Sub setDescription(ByVal PistonNutSize As Double, ByVal Stroke As Double, ByVal ThreadSize As Double, ByVal ThreadLength As Double)
        _Description = "ROD WELDT " & PistonNutSize & "-" & Stroke & "-" & ThreadSize & "-" & ThreadLength
    End Sub

    Public Sub setDescription(ByVal PistonNutSize As Double, ByVal Stroke As Double, ByVal WeldSize As Double)
        _Description = "ROD WELDT " & PistonNutSize & "-" & Stroke & "-" & WeldSize
    End Sub

    Public Sub setDescription(ByVal PistonNutSize As Double, ByVal Stroke As Double)
        _Description = "ROD WELDT " & PistonNutSize & "-" & Stroke
    End Sub

    Public Sub setDescription(ByVal strDescription As String)
        _Description = strDescription
    End Sub

    Public Overrides Function ToString() As String
        Dim strFormat As String
        strFormat = "'{0}','{1}','{2}','{3}',{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},'{15}','{16}',{17},{18},{19},'{20}',{21},{22}," & _
                    "{23},{24},{25},{26},{27},{28},{29},{30}"
        'strFormat = GetTestString()

        Return String.Format(strFormat, PartNo, ProgNo, ByName, Description, _
                Drawing_Num, Drawing_Rev, Operation, WorkCenter, Math.Abs(CInt(AutoDoor)), LargeDia, SmallDia, _
                 NominalThreadDia, Length, Th_Length, TH_Per_IN, _
                ShoulderType, RodType, Xhome, Zhome, Math.Abs(CInt(output2ndop)), Secondoptype, _
                Secondthreaddia, Secondshoulder, Secondthreadlength, Secondthreadnum, _
                Secondthreadcornerrad, Secondchamfer, skimlength, _
                skimdiameter, chamferdepthofcut, Secondopzzero)
    End Function

    Private Function GetTestString() As String
        Dim strFormat As New System.Text.StringBuilder

        For intCount As Integer = 0 To 30
            strFormat.Append("{")
            strFormat.Append(intCount)
            strFormat.Append("},")
        Next
        Return strFormat.ToString
    End Function

End Class


