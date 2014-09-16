Public Class CYL_RodService

    'Private LargeDia As String
    'Private SmallDia As String

    'Private Length As String
    'Private Th_Length As String
    'Private TH_Per_IN As String
    'Private _2ndthreadnum As String
    'Private _2ndthreadlength As String
    'Private ShoulderType As String
    'Private RodType As String
    'Private PartNo As String

    'Private ProgNo As String
    'Private By As String
    'Private Description As String
    'Private Drawing_Num As String
    'Private Xhome As String
    'Private Zhome As String
    'Private NominalThreadDia As String
    'Private Operation As String
    'Private WorkCenter As String
    'Private AutoDoor As String
    'Private output2ndop As String
    'Private _2ndthreaddia As String
    'Private Drawing_Rev As String
    'Private _2ndoptype As String
    'Private _2ndopzzero As String
    'Private _2ndthreadcornerrad As String
    'Private _2ndshoulder As String
    'Private _2ndchamfer As String
    'Private skimlength As String
    'Private skimdiameter As String
    'Private chamferdepthofcut As String


    Private _message As String
    Private _oCYLRos As CYL_Rod

    Public ReadOnly Property Message() As String
        Get
            Return _message
        End Get
    End Property

    Public Function MyFormat(ByVal dblValue As Double, ByVal strMyFormat As String) As String
        Dim intValue As Integer = dblValue


        If strMyFormat.Length > 1 AndAlso intValue = dblValue Then
            Return intValue.ToString & "."
        End If

        Return Format(dblValue, strMyFormat)

    End Function

    Public Function MyFormat(ByVal dblValue As Double) As String
        Return Format(dblValue)
    End Function

    Public Function Start(ByVal FilePath As String, ByVal oCYLRos As CYL_Rod) As Boolean
        Try
            _oCYLRos = oCYLRos
            _message = String.Empty
            Dim strValue As String = Generate_Code_Click()

            If CVSFileUtil.WriteCNC(strValue, FilePath) Then
                Return True
            Else
                _message = CVSFileUtil.Message
                Return False
            End If

        Catch ex As Exception
            _message = "CNC Code Generation Failed." & vbLf & ex.Message
            Return False
        End Try
    End Function

    Private Function Generate_Code_Click() As String
        Try

        
            Dim ld, sd, l, th_l, th_in, xh, zh, nth As Double
            Dim temprpm, stock, rstock, ap, apx, temp As Double
            Dim i, num_cuts, th_speed As Integer
            Dim partnum, prognum, stype, rtype, By, desc, dwg, dwg_rev, op, wc As String
            Dim path, outputfile As String
            Dim today, opendoor, closedoor As String
            Dim speedtable As DataTable
            Dim selectspeedtable As DataTable
            Dim door, op2 As Boolean
            Dim ddformat As String
            Dim od2ndthread, length2ndthread, threadnum2nd, len2ndshoulder As Double
            Dim chamfer, skimlen, skimdia, cornerrad, zzero, zstart As Double
            Dim op2type, th_spec As String
            Dim oThreaded As ThreadingSpeeds
            Dim TH_Per_IN As String
            Dim _2ndthreadnum, th_tool, thread_tool As String

            ddformat = "#.####"
            today = CStr(Date.Today)
            ap = 0.22


            ld = _oCYLRos.LargeDia
            sd = _oCYLRos.SmallDia
            l = -1 * (_oCYLRos.Length)
            th_l = -1 * _oCYLRos.Th_Length
            th_in = _oCYLRos.TH_Per_IN
            threadnum2nd = _oCYLRos.Secondthreadnum
            length2ndthread = _oCYLRos.Secondthreadlength
            stype = _oCYLRos.ShoulderType
            rtype = _oCYLRos.RodType
            partnum = _oCYLRos.PartNo
            prognum = _oCYLRos.ProgNo
            By = _oCYLRos.ByName
            desc = _oCYLRos.Description
            dwg = _oCYLRos.Drawing_Num
            dwg_rev = _oCYLRos.Drawing_Rev
            xh = _oCYLRos.Xhome
            zh = _oCYLRos.Zhome
            nth = _oCYLRos.NominalThreadDia
            op = _oCYLRos.Operation
            wc = _oCYLRos.WorkCenter
            door = _oCYLRos.AutoDoor
            op2 = _oCYLRos.output2ndop
            od2ndthread = _oCYLRos.Secondthreaddia

            If door Then
                opendoor = "M91"
                closedoor = "M90"
            Else
                opendoor = ""
                closedoor = ""
            End If

            '''''FROM DB
            '        speedtable = CurrentDb.OpenRecordset("SELECT ALL * FROM ThreadingSpeeds" _
            '& " WHERE ThreadDia = " & nth)

            ''''' Must Create Data Row
            'speedtable.MoveLast()
            'th_speed = speedtable!RPM
            'th_spec = speedtable!Specs
            'th_tool = speedtable!Tools
            'TH_Per_IN = speedtable!TH_Per_IN

            Dim oDataBase As New CNCDataBaseClass
            oThreaded = oDataBase.GetThreadingSpeeds(nth)
            th_speed = oThreaded.RPM
            th_spec = oThreaded.Specs
            th_tool = oThreaded.TOOLS
            TH_Per_IN = oThreaded.Th_Per_In


            outputfile = prognum

            '''' outputfile is the File Name
            'Open outputfile For Output As #1
            Dim outputString As New System.Text.StringBuilder

            '-------------Header Section
            outputString.AppendLine("(PROG #" + prognum + ")")
            outputString.AppendLine("(PROG REV #0 " + today + " BY " + By + ")")
            outputString.AppendLine("(" + desc + ")")
            outputString.AppendLine("(#" + dwg + " DWG" + " REV#" + dwg_rev + ")")
            outputString.AppendLine("(OP. " + op + ")")
            outputString.AppendLine("(W/C# " + wc + ")")
            outputString.AppendLine(String.Empty)

            '--------------Face Part
            outputString.AppendLine("(RGH FACE)")
            outputString.AppendLine(String.Empty)
            outputString.AppendLine(closedoor)
            outputString.AppendLine(String.Empty)
            outputString.AppendLine("N005G00" + "X" + MyFormat(xh, ddformat) + "Z" + MyFormat(zh, ddformat) + "T0303M42")
            outputString.AppendLine("G50S2000")
            outputString.AppendLine("G96S600M03")
            outputString.AppendLine("X" + MyFormat(ld + 0.1, ddformat) + "Z.015M8")
            outputString.AppendLine("G01X-0.3F.01")
            temprpm = (600 * 12) / (3.14159 * (ld + 0.1))
            If temprpm > "2500" Then temprpm = "2500"

            outputString.AppendLine("G97S" + MyFormat(temprpm, "#") + "M03")
            outputString.AppendLine("G00" + "X" + MyFormat(xh, ddformat) + "Z" + MyFormat(zh, ddformat))
            outputString.AppendLine("M01")
            outputString.AppendLine(String.Empty)

            '--------------Finish face & turn
            outputString.AppendLine("(FIN FACE & TURN TO " + MyFormat(sd, ddformat) + " DIA)")
            outputString.AppendLine(String.Empty)
            outputString.AppendLine("N0100G00T0202M42")
            outputString.AppendLine("G50S2500")

            If rtype = "Nitro" Then
                outputString.AppendLine("X" + MyFormat(ld, ddformat) + "Z.10M08")
                outputString.AppendLine("G96S800M03")

                outputString.AppendLine("G00" + "X" + MyFormat(ld - 0.15, ddformat) + "Z.1")
                outputString.AppendLine("G01" + "X" + MyFormat(ld - 0.1, ddformat) + "Z" + MyFormat(l + 0.005, ddformat) + "F.014")
                outputString.AppendLine("G00" + "X" + MyFormat(ld, ddformat) + "Z.1")

                stock = ld - sd - 0.035 - 0.1
                num_cuts = CInt(stock / ap)
                If ((stock / ap) / num_cuts) > 1 Then num_cuts = num_cuts + 1
                ap = stock / num_cuts
                For i = 1 To num_cuts
                    outputString.AppendLine("G00" + "X" + MyFormat(ld - (i * ap) - 0.1, ddformat) + "Z.1")
                    outputString.AppendLine("G01" + "Z" + MyFormat(l + 0.005, ddformat) + "F.018")
                    If i < num_cuts Then
                        outputString.AppendLine("G00" + "X" + MyFormat(ld - (i * ap) + ap, ddformat) + "Z.1")
                    End If
                Next i
            End If

            If rtype = "Chrome" Then

                outputString.AppendLine("X" + MyFormat(ld, ddformat) + "Z.1M08")
                outputString.AppendLine("G96S800M03")


                stock = ld - sd - 0.035
                num_cuts = CInt(stock / ap)

                If ((stock / ap) / num_cuts) > 1 Then num_cuts = num_cuts + 1
                '   outputString.AppendLine( "Stock" + MyFormat(stock, ddformat) + "Cuts" + MyFormat(num_cuts, "#") + "Actual" + MyFormat(stock / ap, ddformat)

                ap = stock / num_cuts
                '   outputString.AppendLine( "ap=" + MyFormat(ap, ddformat)

                For i = 1 To num_cuts
                    outputString.AppendLine("G00" + "X" + MyFormat(ld - (i * ap), ddformat) + "Z.1")
                    outputString.AppendLine("G01" + "Z" + MyFormat(l + 0.005, ddformat) + "F.018")
                    If i < num_cuts Then
                        outputString.AppendLine("G00" + "X" + MyFormat(ld - (i * ap) + ap, ddformat) + "Z.1")
                    End If
                Next i
            End If

            outputString.AppendLine("X" + MyFormat(ld + 0.05, ddformat))
            outputString.AppendLine("G00Z0")
            outputString.AppendLine("X" + MyFormat(nth + 0.1, ddformat))
            outputString.AppendLine("G01 X-.06 F.014")
            outputString.AppendLine("G00X" + MyFormat(nth - 0.225 - 0.0366, ddformat) + "Z.05")
            outputString.AppendLine("G01" + "X" + MyFormat(nth - 0.005, ddformat) + "Z" + MyFormat(-0.0783, ddformat) + "F.012")
            outputString.AppendLine("Z" + MyFormat(th_l - 0.1, ddformat) + "F.018")
            outputString.AppendLine("X" + MyFormat(sd, ddformat) + "F.012")
            outputString.AppendLine("Z" + MyFormat(l, ddformat))

            If stype = "Rad" Then
                outputString.AppendLine("X" + MyFormat(ld - 0.128, ddformat))
                outputString.AppendLine("G03" + "X" + MyFormat(ld, ddformat) + "Z" + MyFormat(l - 0.064, ddformat) + "K-.064F.012")
                outputString.AppendLine("G01" + "X" + MyFormat(ld + 0.01, ddformat) + "Z" + MyFormat(l - 0.064 - 0.024, ddformat))
                temprpm = (400 * 12) / (3.14159 * sd)
                outputString.AppendLine("G97S" + MyFormat(th_speed, "#") + "M03")
                outputString.AppendLine("G00" + "X" + MyFormat(xh, ddformat) + "Z" + MyFormat(zh, ddformat))
                outputString.AppendLine("M01")
            End If
            If stype = "Chamfer" Then
                outputString.AppendLine("X" + MyFormat(ld - 0.1628, ddformat))
                outputString.AppendLine("G03" + "X" + MyFormat(ld - 0.0638, ddformat) + "Z" + MyFormat(l - 0.0379, ddformat) + "I0K-.0512")
                outputString.AppendLine("G01" + "X" + MyFormat(ld, ddformat) + "Z" + MyFormat(l - 0.1571, ddformat))
                outputString.AppendLine("X" + MyFormat(ld + 0.01, ddformat) + "Z" + MyFormat(l - 0.2, ddformat))
                outputString.AppendLine("G97S" + MyFormat(th_speed, "#") + "M03")
                outputString.AppendLine("G00" + "X" + MyFormat(xh, ddformat) + "Z" + MyFormat(zh, ddformat))
                outputString.AppendLine("M01")
            End If
            outputString.AppendLine(String.Empty)


            outputString.AppendLine("(THREAD " + th_spec + ")")
            outputString.AppendLine(String.Empty)
            'outputString.AppendLine("N0200G00" + MyFormat(th_tool, "#") + "M42")
            outputString.AppendLine("N0200G00" + th_tool + "M42")
            outputString.AppendLine("G97S" + MyFormat(th_speed, "#") + "M03" + "M08")


            If nth = 0.375 Then
                outputString.AppendLine("X" + MyFormat(nth + 0.3, ddformat) + "Z0.4M22")
                outputString.AppendLine("G71X0.318" + "Z" + MyFormat(th_l - 0.1, ddformat) + "B60D.028H.0554F.0416M33M75")
            End If

            If nth = 0.437 Then
                outputString.AppendLine("X" + MyFormat(nth + 0.3, ddformat) + "Z0.4M22")
                outputString.AppendLine("G71X0.370" + "Z" + MyFormat(th_l - 0.1, ddformat) + "B60D.028H.065F.05M33M75")
            End If

            If nth = 0.5 Then
                outputString.AppendLine("X" + MyFormat(nth + 0.3, ddformat) + "Z0.4M22")
                outputString.AppendLine("G71X0.433" + "Z" + MyFormat(th_l - 0.1, ddformat) + "B60D.028H.065F.05M33M75")
            End If

            If nth = 0.625 Then
                outputString.AppendLine("X" + MyFormat(nth + 0.3, ddformat) + "Z0.4M22")
                outputString.AppendLine("G71X0.553" + "Z" + MyFormat(th_l - 0.1, ddformat) + "B60D.028H.068F.0555M33M75")
            End If


            If nth = 0.75 Then
                outputString.AppendLine("X" + MyFormat(nth + 0.3, ddformat) + "Z0.4M23")
                outputString.AppendLine("G33" + "X0.7213" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3930")
                outputString.AppendLine("G33" + "X0.6966" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3866")
                outputString.AppendLine("G33" + "X0.6738" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3860")
                outputString.AppendLine("G33" + "X0.6718" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
            End If

            If nth = 0.875 Then
                outputString.AppendLine("X" + MyFormat(nth + 0.3, ddformat) + "Z0.4M23")
                outputString.AppendLine("G33" + "X0.8426" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3922")
                outputString.AppendLine("G33" + "X0.8152" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3845")
                outputString.AppendLine("G33" + "X0.7878" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3839")
                outputString.AppendLine("G33" + "X0.7858" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
            End If

            If nth = 1 Then
                outputString.AppendLine("X" + MyFormat(nth + 0.3, ddformat) + "Z0.4M23")
                outputString.AppendLine("G33" + "X0.9676" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3922")
                outputString.AppendLine("G33" + "X0.9402" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3845")
                outputString.AppendLine("G33" + "X0.9127" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3839")
                outputString.AppendLine("G33" + "X0.9107" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
            End If


            If nth = 1.125 Then
                outputString.AppendLine("X" + MyFormat(nth + 0.3, ddformat) + "Z0.4M23")
                outputString.AppendLine("G33" + "X1.0910" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3904")
                outputString.AppendLine("G33" + "X1.0570" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3808")
                outputString.AppendLine("G33" + "X1.0230" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3802")
                outputString.AppendLine("G33" + "X1.0210" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
            End If

            If nth = 1.25 Then
                outputString.AppendLine("X" + MyFormat(nth + 0.3, ddformat) + "Z0.4M23")
                outputString.AppendLine("G33" + "X1.2160" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3904")
                outputString.AppendLine("G33" + "X1.1820" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3808")
                outputString.AppendLine("G33" + "X1.1480" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3802")
                outputString.AppendLine("G33" + "X1.1460" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
            End If

            If nth = 1.375 Then
                outputString.AppendLine("X" + MyFormat(nth + 0.3, ddformat) + "Z0.4M23")
                outputString.AppendLine("G33" + "X1.3410" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3904")
                outputString.AppendLine("G33" + "X1.3070" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3808")
                outputString.AppendLine("G33" + "X1.2729" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3802")
                outputString.AppendLine("G33" + "X1.2709" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
            End If

            If nth = 1.5 Then
                outputString.AppendLine("X" + MyFormat(nth + 0.3, ddformat) + "Z0.4M23")
                outputString.AppendLine("G33" + "X1.4660" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3904")
                outputString.AppendLine("G33" + "X1.4320" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3808")
                outputString.AppendLine("G33" + "X1.3979" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
                outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3802")
                outputString.AppendLine("G33" + "X1.3959" + "Z" + MyFormat(th_l - 0.1, ddformat) + "F" + MyFormat(1 / th_in, ddformat) + "L.015")
            End If

            outputString.AppendLine("G00" + "X" + MyFormat(xh, ddformat) + "Z" + MyFormat(zh, ddformat) + "M5M9")
            outputString.AppendLine(opendoor)

            If op2 = False Then
                outputString.AppendLine("M02")

                'CLOSE THE FILE
                'Close #1
                'MsgBox("Done 1st Op. Only")
                Return outputString.ToString
            End If



            If op2 = True Then
                op2type = _oCYLRos.Secondoptype
                If op2type = "Threaded" Then
                    od2ndthread = _oCYLRos.Secondthreaddia
                    length2ndthread = _oCYLRos.Secondthreadlength
                    _2ndthreadnum = _oCYLRos.Secondthreadnum
                    zzero = _oCYLRos.Secondopzzero
                    zstart = zzero + 0.05
                    If od2ndthread = ld Then
                        outputString.AppendLine("M00")
                        outputString.AppendLine(String.Empty)
                        outputString.AppendLine("(ROTATE PART IN JAWS)")
                        outputString.AppendLine("(CLOSE DOOR USING MID AUTO MANUAL)")
                        outputString.AppendLine(String.Empty)
                        outputString.AppendLine(closedoor)
                        outputString.AppendLine(String.Empty)
                        outputString.AppendLine("(FACE & CHAMFER 2ND SIDE)")
                        outputString.AppendLine(String.Empty)
                        outputString.AppendLine("N0300G00T0202M42")
                        outputString.AppendLine("G96S800M03")
                        outputString.AppendLine("X" + MyFormat(od2ndthread + 0.1, ddformat) + "Z-0.03M8")
                        outputString.AppendLine("G01X-0.06F0.012")
                        outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread - 0.2566, ddformat) + "Z0.02")
                        outputString.AppendLine("G01" + "X" + MyFormat(od2ndthread, ddformat) + "Z-0.1083F0.012")
                        outputString.AppendLine("X" + MyFormat(od2ndthread + 0.01, ddformat) + "Z-0.15")

                        'TABLE 
                        'speedtable = CurrentDb.OpenRecordset("SELECT ALL * FROM ThreadingSpeeds" _
                        '                & " WHERE ThreadDia = " & od2ndthread)
                        'speedtable.MoveLast()
                        'th_speed = speedtable!RPM
                        'th_spec = speedtable!Specs
                        'thread_tool = speedtable!od2ndtools
                        '            [2ndthreadnum] = speedtable![2ndthreadnum]

                        oThreaded = oDataBase.GetThreadingSpeeds(od2ndthread)
                        th_speed = oThreaded.RPM
                        th_spec = oThreaded.Specs
                        thread_tool = oThreaded.ODSecondTools
                        _2ndthreadnum = oThreaded.SecondThread


                        outputString.AppendLine("G97S" + MyFormat(th_speed, "#") + "M03")
                        outputString.AppendLine("G00" + "X" + MyFormat(xh, ddformat) + "Z" + MyFormat(zh, ddformat))
                        outputString.AppendLine("M01")
                        outputString.AppendLine(String.Empty)
                        outputString.AppendLine("(THREAD " + th_spec + ")")
                        outputString.AppendLine(String.Empty)
                        ' outputString.AppendLine("N0400G00" + MyFormat(thread_tool, "#") + "M42")
                        outputString.AppendLine("N0400G00" + thread_tool + "M42")
                        outputString.AppendLine("G97S" + MyFormat(th_speed, "#") + "M03M08")


                        If od2ndthread = 0.375 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M22")
                            outputString.AppendLine("G71X0.318" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "B60D.028H.0554F.0416M33M75")
                        End If

                        If od2ndthread = 0.437 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M22")
                            outputString.AppendLine("G71X0.370" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "B60D.028H.065F.05M33M75")
                        End If

                        If od2ndthread = 0.5 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M22")
                            outputString.AppendLine("G71X0.433" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "B60D.028H.065F.05M33M75")
                        End If

                        If od2ndthread = 0.625 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M22")
                            outputString.AppendLine("G71X0.553" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "B60D.028H.068F.0555M33M75")
                        End If

                        If od2ndthread = 0.75 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X0.7213" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3930")
                            outputString.AppendLine("G33" + "X0.6966" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3866")
                            outputString.AppendLine("G33" + "X0.6738" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3860")
                            outputString.AppendLine("G33" + "X0.6718" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        If od2ndthread = 0.875 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X0.8426" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3922")
                            outputString.AppendLine("G33" + "X0.8152" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3845")
                            outputString.AppendLine("G33" + "X0.7878" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3839")
                            outputString.AppendLine("G33" + "X0.7858" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        If od2ndthread = 1 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X0.9676" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3922")
                            outputString.AppendLine("G33" + "X0.9402" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3845")
                            outputString.AppendLine("G33" + "X0.9127" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3839")
                            outputString.AppendLine("G33" + "X0.9107" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        If od2ndthread = 1.125 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X1.0910" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3904")
                            outputString.AppendLine("G33" + "X1.0570" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3813")
                            outputString.AppendLine("G33" + "X1.0248" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3807")
                            outputString.AppendLine("G33" + "X1.0228" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        If od2ndthread = 1.25 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X1.2160" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3904")
                            outputString.AppendLine("G33" + "X1.1820" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3813")
                            outputString.AppendLine("G33" + "X1.1498" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3807")
                            outputString.AppendLine("G33" + "X1.1478" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        If od2ndthread = 1.375 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X1.3410" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3904")
                            outputString.AppendLine("G33" + "X1.3070" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3813")
                            outputString.AppendLine("G33" + "X1.2748" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3807")
                            outputString.AppendLine("G33" + "X1.2728" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        If od2ndthread = 1.5 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X1.4660" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3904")
                            outputString.AppendLine("G33" + "X1.4320" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3813")
                            outputString.AppendLine("G33" + "X1.3998" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3807")
                            outputString.AppendLine("G33" + "X1.3978" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        If od2ndthread = 1.75 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X1.7160" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3904")
                            outputString.AppendLine("G33" + "X1.6820" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3813")
                            outputString.AppendLine("G33" + "X1.6498" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3807")
                            outputString.AppendLine("G33" + "X1.6478" + "Z" + MyFormat(zzero - length2ndthread, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        outputString.AppendLine("G00" + "X" + MyFormat(xh, ddformat) + "Z" + MyFormat(zh, ddformat) + "M5M9")
                        outputString.AppendLine(opendoor)
                    End If

                    If od2ndthread < ld Then
                        len2ndshoulder = _oCYLRos.Secondshoulder
                        If length2ndthread > len2ndshoulder Then
                            'CLOSE THE FILE
                            ' Close #1
                            Throw New Exception("Please Check Shoulder & Thread Length")
                            '  MsgBox("Please Check Shoulder & Thread Length")
                            Return outputString.ToString
                        End If
                        outputString.AppendLine("M00")
                        outputString.AppendLine(String.Empty)
                        outputString.AppendLine("(ROTATE PART IN JAWS)")
                        outputString.AppendLine("(CLOSE DOOR USING MID AUTO MANUAL)")
                        outputString.AppendLine(String.Empty)
                        outputString.AppendLine(closedoor)
                        outputString.AppendLine(String.Empty)
                        outputString.AppendLine("(FACE TURN & CHAMFER 2ND SIDE)")
                        outputString.AppendLine(String.Empty)
                        outputString.AppendLine("N0300G00T0202M42")
                        outputString.AppendLine("G96S800M03")
                        outputString.AppendLine("X" + MyFormat(ld + 0.1, ddformat) + "Z" + MyFormat(zzero, ddformat) + "M8")
                        outputString.AppendLine("G01X-0.06F0.012")

                        If rtype = "Nitro" Then
                            outputString.AppendLine("G00" + "X" + MyFormat(ld - 0.1, ddformat) + "Z" + MyFormat(zstart, ddformat) + "M8")
                            outputString.AppendLine("G01" + "X" + MyFormat(ld - 0.06, ddformat) + "Z" + MyFormat(zzero - len2ndshoulder + 0.005, ddformat) + "F.014")
                            outputString.AppendLine("G00" + "X" + MyFormat(ld, ddformat) + "Z" + MyFormat(zstart, ddformat))

                            cornerrad = _oCYLRos.Secondthreadcornerrad
                            stock = ld - od2ndthread - 0.1
                            ap = 0.2
                            num_cuts = CInt(stock / ap)
                            If ((stock / ap) / num_cuts) > 1 Then num_cuts = num_cuts + 1
                            '   outputString.AppendLine( "Stock" + MyFormat(stock, ddformat) + "Cuts" + MyFormat(num_cuts, "#") + "Actual" + MyFormat(stock / ap, ddformat)
                            ap = stock / num_cuts
                            '   outputString.AppendLine( "ap=" +MyFormat(ap,ddformat)
                            For i = 1 To num_cuts
                                If i < num_cuts Then
                                    outputString.AppendLine("G00" + "X" + MyFormat(ld - (i * ap) - 0.1, ddformat) + "Z" + MyFormat(zstart, ddformat))
                                    outputString.AppendLine("G01" + "Z" + MyFormat(zzero - len2ndshoulder + 0.005, ddformat) + "F0.018")
                                    outputString.AppendLine("G00" + "X" + MyFormat(ld - (i * ap) + ap, ddformat) + "Z" + MyFormat(zstart, ddformat))
                                End If
                                If i = num_cuts Then

                                    outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.03, ddformat) + "Z" + MyFormat(zstart, ddformat))
                                    outputString.AppendLine("G01" + "Z" + MyFormat(zzero - len2ndshoulder + 0.005 + cornerrad - 0.0312, ddformat) + "F0.018")
                                    outputString.AppendLine("G02" + "X" + MyFormat(od2ndthread + 0.03 + (2 * (cornerrad - 0.0312)), ddformat) + "Z" + MyFormat(zzero - len2ndshoulder + 0.005, ddformat) + "I" + MyFormat(cornerrad - 0.0312, ddformat) + "F0.012")
                                    outputString.AppendLine("G01" + "X" + MyFormat(ld + 0.05, ddformat))
                                    outputString.AppendLine("G00" + "Z" + MyFormat(zstart, ddformat))

                                    outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread - 0.2566, ddformat) + "Z" + MyFormat(zstart, ddformat))
                                    outputString.AppendLine("G01" + "X" + MyFormat(od2ndthread - 0.005, ddformat) + "Z" + MyFormat(zstart - 0.1283, ddformat) + "F0.012")
                                    outputString.AppendLine("G01" + "Z" + MyFormat(zzero - len2ndshoulder + cornerrad - 0.0312, ddformat) + "F0.018")
                                    outputString.AppendLine("G02" + "X" + MyFormat(od2ndthread - 0.005 + (2 * (cornerrad - 0.0312)), ddformat) + "Z" + MyFormat(zzero - len2ndshoulder, ddformat) + "I" + MyFormat(cornerrad - 0.0312, ddformat) + "F0.012")
                                    outputString.AppendLine("G01" + "X" + MyFormat(ld - 0.128, ddformat))
                                    outputString.AppendLine("G03" + "X" + MyFormat(ld, ddformat) + "Z" + MyFormat(zzero - len2ndshoulder - 0.064) + "K-0.064")
                                    outputString.AppendLine("G01" + "X" + MyFormat(ld + 0.01, ddformat) + "Z" + MyFormat(zzero - len2ndshoulder - 0.064 - 0.044, ddformat) + "F0.018")
                                End If
                            Next i
                        End If

                        If rtype = "Chrome" Then
                            cornerrad = _oCYLRos.Secondthreadcornerrad
                            stock = ld - od2ndthread
                            ap = 0.2
                            num_cuts = CInt(stock / ap)

                            If ((stock / ap) / num_cuts) > 1 Then num_cuts = num_cuts + 1
                            '   outputString.AppendLine( "Stock" + MyFormat(stock, ddformat) + "Cuts" + MyFormat(num_cuts, "#") + "Actual" + MyFormat(stock / ap, ddformat)
                            ap = stock / num_cuts
                            '   outputString.AppendLine( "ap=" +MyFormat(ap,ddformat)
                            For i = 1 To num_cuts

                                If i < num_cuts Then
                                    outputString.AppendLine("G00" + "X" + MyFormat(ld - (i * ap), ddformat) + "Z" + MyFormat(zstart, ddformat))
                                    outputString.AppendLine("G01" + "Z" + MyFormat(zzero - len2ndshoulder + 0.005, ddformat) + "F0.018")
                                    outputString.AppendLine("G00" + "X" + MyFormat(ld - (i * ap) + ap, ddformat) + "Z" + MyFormat(zstart, ddformat))
                                End If

                                If i = num_cuts Then
                                    outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.03, ddformat) + "Z" + MyFormat(zstart, ddformat))
                                    outputString.AppendLine("G01" + "Z" + MyFormat(zzero - len2ndshoulder + 0.005 + cornerrad - 0.0312, ddformat) + "F0.018")
                                    outputString.AppendLine("G02" + "X" + MyFormat(od2ndthread + 0.03 + (2 * (cornerrad - 0.0312)), ddformat) + "Z" + MyFormat(zzero - len2ndshoulder + 0.005, ddformat) + "I" + MyFormat(cornerrad - 0.0312, ddformat) + "F0.012")
                                    outputString.AppendLine("G01" + "X" + MyFormat(ld + 0.05, ddformat))
                                    outputString.AppendLine("G00" + "Z" + MyFormat(zstart, ddformat))

                                    outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread - 0.2566, ddformat) + "Z" + MyFormat(zstart, ddformat))
                                    outputString.AppendLine("G01" + "X" + MyFormat(od2ndthread - 0.005, ddformat) + "Z" + MyFormat(zstart - 0.1283, ddformat) + "F0.012")
                                    outputString.AppendLine("G01" + "Z" + MyFormat(zzero - len2ndshoulder + cornerrad - 0.0312, ddformat) + "F0.018")
                                    outputString.AppendLine("G02" + "X" + MyFormat(od2ndthread - 0.005 + (2 * (cornerrad - 0.0312)), ddformat) + "Z" + MyFormat(zzero - len2ndshoulder, ddformat) + "I" + MyFormat(cornerrad - 0.0312, ddformat) + "F0.012")
                                    outputString.AppendLine("G01" + "X" + MyFormat(ld - 0.128, ddformat))
                                    outputString.AppendLine("G03" + "X" + MyFormat(ld, ddformat) + "Z" + MyFormat(zzero - len2ndshoulder - 0.064) + "K-0.064")
                                    outputString.AppendLine("G01" + "X" + MyFormat(ld + 0.01, ddformat) + "Z" + MyFormat(zzero - len2ndshoulder - 0.064 - 0.044, ddformat) + "F0.018")
                                End If
                            Next i
                        End If

                        'TABLE 
                        '        speedtable = CurrentDb.OpenRecordset("SELECT ALL * FROM ThreadingSpeeds" _
                        '                           & " WHERE ThreadDia = " & od2ndthread)
                        '        speedtable.MoveLast()
                        '        th_speed = speedtable!RPM
                        '        th_spec = speedtable!Specs
                        '        thread_tool = speedtable!od2ndtools
                        '[2ndthreadnum] = speedtable![2ndthreadnum]

                        oThreaded = oDataBase.GetThreadingSpeeds(od2ndthread)
                        th_speed = oThreaded.RPM
                        th_spec = oThreaded.Specs
                        thread_tool = oThreaded.ODSecondTools
                        _2ndthreadnum = oThreaded.SecondThread


                        outputString.AppendLine("G97S" + MyFormat(th_speed, "#") + "M03")
                        outputString.AppendLine("G00" + "X" + MyFormat(xh, ddformat) + "Z" + MyFormat(zh, ddformat))
                        outputString.AppendLine("M01")
                        outputString.AppendLine(String.Empty)

                        outputString.AppendLine("(THREAD " + th_spec + ")")
                        outputString.AppendLine(String.Empty)
                        '    outputString.AppendLine("N0400G00" + MyFormat(thread_tool, "#") + "M42")
                        outputString.AppendLine("N0400G00" + thread_tool + "M42")
                        outputString.AppendLine("G97S" + MyFormat(th_speed, "#") + "M03M08")


                        If od2ndthread = 0.375 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M22")
                            outputString.AppendLine("G71X0.318" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "B60D.028H.0554F.0416M33M75")
                        End If

                        If od2ndthread = 0.437 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M22")
                            outputString.AppendLine("G71X0.370" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "B60D.028H.065F.05M33M75")
                        End If

                        If od2ndthread = 0.5 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M22")
                            outputString.AppendLine("G71X0.433" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "B60D.028H.065F.05M33M75")
                        End If

                        If od2ndthread = 0.625 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M22")
                            outputString.AppendLine("G71X0.553" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "B60D.028H.068F.0555M33M75")
                        End If


                        If od2ndthread = 0.75 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X0.7213" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3930")
                            outputString.AppendLine("G33" + "X0.6966" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3866")
                            outputString.AppendLine("G33" + "X0.6738" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3860")
                            outputString.AppendLine("G33" + "X0.6718" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        If od2ndthread = 0.875 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X0.8426" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3922")
                            outputString.AppendLine("G33" + "X0.8152" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3845")
                            outputString.AppendLine("G33" + "X0.7878" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3839")
                            outputString.AppendLine("G33" + "X0.7858" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        If od2ndthread = 1 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X0.9676" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3922")
                            outputString.AppendLine("G33" + "X0.9402" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3845")
                            outputString.AppendLine("G33" + "X0.9127" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(nth + 0.3, ddformat) + "Z.3839")
                            outputString.AppendLine("G33" + "X0.9107" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        If od2ndthread = 1.125 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X1.0910" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3904")
                            outputString.AppendLine("G33" + "X1.0570" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3813")
                            outputString.AppendLine("G33" + "X1.0248" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3807")
                            outputString.AppendLine("G33" + "X1.0228" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        If od2ndthread = 1.25 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X1.2160" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3904")
                            outputString.AppendLine("G33" + "X1.1820" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3813")
                            outputString.AppendLine("G33" + "X1.1498" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3807")
                            outputString.AppendLine("G33" + "X1.1478" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        If od2ndthread = 1.375 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X1.3410" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3904")
                            outputString.AppendLine("G33" + "X1.3070" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3813")
                            outputString.AppendLine("G33" + "X1.2748" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3807")
                            outputString.AppendLine("G33" + "X1.2728" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        If od2ndthread = 1.5 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X1.4660" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3904")
                            outputString.AppendLine("G33" + "X1.4320" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3813")
                            outputString.AppendLine("G33" + "X1.3998" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3807")
                            outputString.AppendLine("G33" + "X1.3978" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        If od2ndthread = 1.75 Then
                            outputString.AppendLine("X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z0.4M23")
                            outputString.AppendLine("G33" + "X1.7160" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3904")
                            outputString.AppendLine("G33" + "X1.6820" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3813")
                            outputString.AppendLine("G33" + "X1.6498" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                            outputString.AppendLine("G00" + "X" + MyFormat(od2ndthread + 0.3, ddformat) + "Z.3807")
                            outputString.AppendLine("G33" + "X1.6478" + "Z" + MyFormat(zzero - length2ndthread - 0.1, ddformat) + "F" + MyFormat(1 / [_2ndthreadnum], ddformat) + "L.015")
                        End If

                        '   outputString.AppendLine( "(Total removed" + MyFormat(apx, ddformat) + ")"
                        outputString.AppendLine("G00" + "X" + MyFormat(xh, ddformat) + "Z" + MyFormat(zh, ddformat) + "M5M9")
                        outputString.AppendLine(opendoor)
                    End If
                End If
                If op2type = "Face & Chamfer" Then
                    outputString.AppendLine("M00")
                    outputString.AppendLine(String.Empty)
                    outputString.AppendLine("(ROTATE PART IN JAWS)")
                    outputString.AppendLine("(CLOSE DOOR USING MID AUTO MANUAL)")
                    outputString.AppendLine(String.Empty)
                    outputString.AppendLine(closedoor)
                    outputString.AppendLine(String.Empty)
                    outputString.AppendLine("(FACE & CHAMFER 2ND SIDE)")
                    outputString.AppendLine(String.Empty)
                    outputString.AppendLine("N0300G00T0202M42")
                    outputString.AppendLine("G96S800M03")
                    outputString.AppendLine("X" + MyFormat(ld + 0.1, ddformat) + "Z-0.03M8")
                    outputString.AppendLine("G01X-0.06F0.012")
                    outputString.AppendLine("G00" + "X" + MyFormat(ld - 0.1966, ddformat) + "Z0.02")
                    outputString.AppendLine("G01" + "X" + MyFormat(ld, ddformat) + "Z-0.0783F0.012")
                    outputString.AppendLine("X" + MyFormat(ld + 0.01, ddformat) + "Z-0.15")
                    outputString.AppendLine("G00" + "X" + MyFormat(xh, ddformat) + "Z" + MyFormat(zh, ddformat) + "M5M9")
                    outputString.AppendLine(opendoor)
                End If
                If op2type = "Chamfer & Skim" Then
                    chamfer = 2 * _oCYLRos.Secondchamfer
                    skimlen = _oCYLRos.skimlength
                    skimdia = _oCYLRos.skimdiameter
                    zzero = _oCYLRos.Secondopzzero
                    zstart = zzero + 0.05
                    outputString.AppendLine("M00")
                    outputString.AppendLine(String.Empty)
                    outputString.AppendLine("(ROTATE PART IN JAWS)")
                    outputString.AppendLine("(CLOSE DOOR USING MID AUTO MANUAL)")
                    outputString.AppendLine(String.Empty)
                    outputString.AppendLine(closedoor)
                    outputString.AppendLine(String.Empty)
                    outputString.AppendLine("(FACE & CHAMFER 2ND SIDE)")
                    outputString.AppendLine(String.Empty)
                    outputString.AppendLine("N0300G00T0202M42")
                    outputString.AppendLine("G50S2500")
                    outputString.AppendLine("G96S800M03")
                    outputString.AppendLine("X" + MyFormat(ld + 0.1, ddformat) + "Z" + MyFormat(zzero, ddformat) + "M8")
                    outputString.AppendLine("G01X-0.06F0.012")
                    ap = _oCYLRos.chamferdepthofcut
                    num_cuts = CInt(chamfer / ap)

                    If ((chamfer / ap) / num_cuts) > 1 Then num_cuts = num_cuts + 1
                    '   outputString.AppendLine( "Stock" + MyFormat(stock, ddformat) + "Cuts" + MyFormat(num_cuts, "#") + "Actual" + MyFormat(stock / ap, ddformat)

                    ap = chamfer / num_cuts
                    '   outputString.AppendLine( "ap=" + MyFormat(ap, ddformat)
                    For i = 1 To num_cuts
                        outputString.AppendLine("G00" + "X" + MyFormat(ld - (i * ap) - 0.1 - 0.0366 - (ld - skimdia), ddformat) + "Z" + MyFormat(zstart, ddformat))
                        If i < num_cuts Then
                            outputString.AppendLine("G01" + "X" + MyFormat(ld, ddformat) + "Z" + MyFormat(zstart - (i * ap) / 2 - 0.05 - 0.0183, ddformat) + "F0.012")
                            outputString.AppendLine("G00" + "Z" + MyFormat(zzero + 0.05, ddformat))
                        End If
                        If i = num_cuts Then
                            outputString.AppendLine("G01" + "X" + MyFormat(skimdia, ddformat) + "Z" + MyFormat(zstart - (i * ap) / 2 - 0.05 - 0.0183, ddformat) + "F0.012")
                            outputString.AppendLine("Z" + MyFormat((zzero - (skimlen + 0.0183)), ddformat))
                            outputString.AppendLine("X" + MyFormat(ld, ddformat) + "Z" + MyFormat((zzero - (skimlen + 0.0183) - (ld - skimdia) / 2), ddformat))
                        End If
                    Next i

                    outputString.AppendLine("G00" + "X" + MyFormat(xh, ddformat) + "Z" + MyFormat(zh, ddformat) + "M5M9")
                    outputString.AppendLine(opendoor)
                End If
                If op2type = "Chamfer Only" Then
                    chamfer = 2 * _oCYLRos.Secondchamfer
                    zzero = _oCYLRos.Secondopzzero
                    zstart = zzero + 0.05
                    outputString.AppendLine("M00")
                    outputString.AppendLine(String.Empty)
                    outputString.AppendLine("(ROTATE PART IN JAWS)")
                    outputString.AppendLine("(CLOSE DOOR USING MID AUTO MANUAL)")
                    outputString.AppendLine(String.Empty)
                    outputString.AppendLine(closedoor)
                    outputString.AppendLine(String.Empty)
                    outputString.AppendLine("(FACE & CHAMFER 2ND SIDE)")
                    outputString.AppendLine(String.Empty)
                    outputString.AppendLine("N0300G00T0202M42")
                    outputString.AppendLine("G50S2500")
                    outputString.AppendLine("G96S800M03")
                    outputString.AppendLine("X" + MyFormat(ld + 0.1, ddformat) + "Z" + MyFormat(zzero, ddformat) + "M8")
                    outputString.AppendLine("G01X-0.06F0.012")
                    ap = _oCYLRos.chamferdepthofcut
                    num_cuts = CInt(chamfer / ap)

                    If ((chamfer / ap) / num_cuts) > 1 Then num_cuts = num_cuts + 1
                    '   outputString.AppendLine( "Stock" + MyFormat(chamfer, ddformat) + "Cuts" + MyFormat(num_cuts, "#") + "Actual" + MyFormat(chamfer / ap, ddformat)

                    ap = chamfer / num_cuts
                    '   outputString.AppendLine( "ap=" + MyFormat(ap, ddformat)
                    For i = 1 To num_cuts
                        outputString.AppendLine("G00" + "X" + MyFormat(ld - (i * ap) - 0.1 - 0.0366, ddformat) + "Z" + MyFormat(zstart, ddformat))
                        outputString.AppendLine("G01" + "X" + MyFormat(ld, ddformat) + "Z" + MyFormat(zstart - (i * ap) / 2 - 0.05 - 0.0183, ddformat) + "F0.012")

                        If i < num_cuts Then
                            outputString.AppendLine("G00" + "Z" + MyFormat(zzero + 0.05, ddformat))
                        End If


                        If i = num_cuts Then
                            outputString.AppendLine("X" + MyFormat(ld + 0.01, ddformat) + "Z" + MyFormat(zstart - (i * ap) / 2 - 0.05 - 0.0183 - 0.0417, ddformat))
                        End If
                    Next i
                    outputString.AppendLine("G00" + "X" + MyFormat(xh, ddformat) + "Z" + MyFormat(zh, ddformat) + "M5M9")
                    outputString.AppendLine(opendoor)
                End If
            End If
            outputString.AppendLine("M02")
            'CLOSE THE FILE
            'Close #1

            'MsgBox("Done")

            Return outputString.ToString
        Catch ex As Exception

        End Try
    End Function

End Class
