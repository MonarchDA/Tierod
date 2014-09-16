Imports IFLBaseDataLayer
Imports IFLCommonLayer
Imports MonarchFunctionalLayer

Public Class CNCDataBaseClass

    Private _strQuery As String
    Private _strErrorMessage As String
    Private _oMILConnectionObject As IFLConnectionClass


    Public Property MIL_WeldedConnectionObject() As IFLConnectionClass
        Get
            Return _oMILConnectionObject
        End Get
        Set(ByVal value As IFLConnectionClass)
            _oMILConnectionObject = value
        End Set
    End Property

    Public Sub New()
        Try
            Dim strXMLFilePath As String = System.Environment.CurrentDirectory + "\MILWeldedConnection.xml"
            Dim oDataSet As New DataSet
            oDataSet.ReadXml(strXMLFilePath)
            If Not oDataSet.Tables.Count <= 0 Then
                Dim strServer As String = oDataSet.Tables(0).Rows(0).Item(0).ToString
                Dim strDataBase As String = oDataSet.Tables(0).Rows(0).Item(1).ToString
                _oMILConnectionObject = IFLConnectionClass.GetConnectionObject(strServer, strDataBase, "System.Data.SqlClient", , , True)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Function GetTH_Per_In(ByVal Nut_ThreadSize As Double) As Double
        _strQuery = "select TH_Per_IN from Nut_Thread_Num where Nut_ThreadSize like '" & Nut_ThreadSize & "%'"
        GetTH_Per_In = MIL_WeldedConnectionObject.GetValue(_strQuery)
        If IsNothing(GetTH_Per_In) Then
            GetTH_Per_In = 0
            _strErrorMessage = "Data not retrieved from Nut_Thread_Num table" + vbCrLf
        End If
    End Function

    Public Function GetThreadingSpeeds(ByVal ThreadDia As Double) As ThreadingSpeeds
        _strQuery = "SELECT ThreadDia,RPM,TOOLS,SecondThreadNum,OdSecondTools,Th_Per_In,Specs FROM ThreadingSpeeds  where ThreadDia = " + ThreadDia.ToString
        Dim oDataRow As DataRow = MIL_WeldedConnectionObject.GetDataRow(_strQuery)
        If IsNothing(oDataRow) Then
            GetThreadingSpeeds = Nothing
            _strErrorMessage = "Data not retrieved from ThreadingSpeeds table" + vbCrLf
            Return GetThreadingSpeeds
        End If
        GetThreadingSpeeds = New ThreadingSpeeds(oDataRow("ThreadDia"), oDataRow("RPM"), oDataRow("TOOLS"), oDataRow("SecondThreadNum"), oDataRow("OdSecondTools"), oDataRow("Th_Per_In"), oDataRow("Specs"))
    End Function

    Public Function InsertCyl_RodData(ByVal oCyl_Rod As CYL_Rod) As Boolean
        Try
            _strQuery = "INSERT INTO CYL_Rod " & _
               "(PartNo,ProgNo,ByName,Description,Drawing_Num,Drawing_Rev,Operation" & _
               ",WorkCenter,AutoDoor,LargeDia,SmallDia,NominalThreadDia" & _
               ",Length,Th_Length,Th_Per_In,ShoulderType,RodType" & _
              " ,Xhome ,Zhome ,output2ndop,Secondoptype ,Secondthreaddia" & _
               ",Secondshoulder,Secondthreadlength" & _
              " ,Secondthreadnum ,Secondthreadcornerrad" & _
               ",Secondchamfer,skimlength,skimdiameter,chamferdepthofcut" & _
              " ,Secondopzzero,DateAndTime) " & _
            "VALUES( " & oCyl_Rod.ToString & ", " & Date.Today & ")"
            InsertCyl_RodData = MIL_WeldedConnectionObject.ExecuteQuery(_strQuery)
            If Not InsertCyl_RodData Then
                Return False
                _strErrorMessage = "Data not inserted in CYL_Rod table" + vbCrLf
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Function DoesPartCodeExist(ByVal oCyl_Rod As CYL_Rod) As Boolean
        Try
            _strQuery = "select PartNo from CYL_Rod where PartNo ='" & oCyl_Rod.PartNo & "'"
            DoesPartCodeExist = MIL_WeldedConnectionObject.GetValue(_strQuery)
            If Not DoesPartCodeExist Then
                Return False
                _strErrorMessage = "Data not retrieved in CYL_Rod table" + vbCrLf
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Function UpdateCyl_RodData(ByVal oCyl_Rod As CYL_Rod) As Boolean
        Try

            _strQuery = "Update CYL_Rod SET ByName ='" & oCyl_Rod.ByName & "' , Description ='" & oCyl_Rod.Description & "' , Drawing_Num =" & oCyl_Rod.Drawing_Num.ToString
            _strQuery += " , Drawing_Rev =" & oCyl_Rod.Drawing_Rev.ToString & " , Operation =" & oCyl_Rod.Operation.ToString & " , WorkCenter =" & oCyl_Rod.WorkCenter.ToString
            _strQuery += " , AutoDoor ='" & oCyl_Rod.AutoDoor.ToString & "' , LargeDia =" & oCyl_Rod.LargeDia.ToString & " , SmallDia =" & oCyl_Rod.SmallDia.ToString
            _strQuery += " , NominalThreadDia =" & oCyl_Rod.NominalThreadDia.ToString & " , Length =" & oCyl_Rod.Length.ToString & " , Th_Length =" & oCyl_Rod.Th_Length.ToString
            _strQuery += " , Th_Per_In =" & oCyl_Rod.TH_Per_IN.ToString & " , ShoulderType ='" & oCyl_Rod.ShoulderType & "' , RodType ='" & oCyl_Rod.RodType
            _strQuery += "' , Xhome =" & oCyl_Rod.Xhome.ToString & " , Zhome =" & oCyl_Rod.Zhome.ToString & " , output2ndop ='" & oCyl_Rod.output2ndop.ToString
            _strQuery += "' , Secondoptype ='" & oCyl_Rod.Secondoptype & "' , Secondthreaddia =" & oCyl_Rod.Secondthreaddia.ToString & " , Secondshoulder =" & oCyl_Rod.Secondshoulder.ToString
            _strQuery += " , Secondthreadlength =" & oCyl_Rod.Secondthreadlength.ToString & " , Secondthreadnum =" & oCyl_Rod.Secondthreadnum.ToString & " , Secondthreadcornerrad =" & oCyl_Rod.Secondthreadcornerrad.ToString
            _strQuery += " , Secondchamfer =" & oCyl_Rod.Secondchamfer.ToString & " , skimlength =" & oCyl_Rod.skimlength.ToString & " , skimdiameter =" & oCyl_Rod.skimdiameter.ToString
            _strQuery += " , chamferdepthofcut =" & oCyl_Rod.chamferdepthofcut.ToString & " , Secondopzzero =" & oCyl_Rod.Secondopzzero.ToString
            _strQuery += " where PartNo ='" & oCyl_Rod.PartNo & "'"
            UpdateCyl_RodData = MIL_WeldedConnectionObject.ExecuteQuery(_strQuery)
            If Not UpdateCyl_RodData Then
                Return False
                _strErrorMessage = "Data not updated in CYL_Rod table" + vbCrLf
            End If
        Catch ex As Exception

        End Try
    End Function
End Class
