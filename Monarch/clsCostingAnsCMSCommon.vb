Public Class clsCostingAnsCMSCommon

    Public Function GetChromeMachiningTableName() As String
        GetChromeMachiningTableName = Nothing
        Try
            If strStyleModified.Equals("NON ASAE") Then
                If GetSeries() = "TL/TH/TP" Then
                    GetChromeMachiningTableName = "Costing_ChromeROd_TH_TL_TP_NASAE"
                ElseIf GetSeries() = "TX" Then
                    GetChromeMachiningTableName = "Costing_ChromeROd_TX_NASAE"
                End If
            ElseIf strStyleModified.Equals("ASAE") Then
                If GetSeries() = "TL/TH/TP" Then
                    GetChromeMachiningTableName = "Costing_ChromeROd_TH_TL_TP_ASAE"
                ElseIf GetSeries() = "TX" Then
                    GetChromeMachiningTableName = "Costing_ChromeROd_TX_ASAE"
                End If
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Function GetNitroRodMachiningTableName() As String
        GetNitroRodMachiningTableName = Nothing
        Try
            If strStyleModified.Equals("NON ASAE") Then
                If GetSeries() = "TL/TH/TP" Then
                    GetNitroRodMachiningTableName = "Costing_NitroROd_TH_TL_TP_NASAE"
                End If
            ElseIf strStyleModified.Equals("ASAE") Then
                If GetSeries() = "TL/TH/TP" Then
                    GetNitroRodMachiningTableName = "Costing_NitroROd_TH_TL_TP_ASAE"
                End If
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Function GetInductionHBMachiningTableName(Optional ByVal strCalling As String = "") As String
        GetInductionHBMachiningTableName = Nothing
        Try
            If strCalling = "WCDetails" Then
                If strStyleModified.Equals("NON ASAE") Then
                    GetInductionHBMachiningTableName = "Costing_InductionHB_NASAE_WorkCentre"
                ElseIf strStyleModified.Equals("ASAE") Then
                    GetInductionHBMachiningTableName = "Costing_InductionHB_ASAE_WorkCentre"
                End If
            Else
                If strStyleModified.Equals("NON ASAE") Then
                    GetInductionHBMachiningTableName = "Costing_InductionHB_NASAE"
                ElseIf strStyleModified.Equals("ASAE") Then
                    GetInductionHBMachiningTableName = "Costing_InductionHB_ASAE"
                End If
            End If

        Catch ex As Exception

        End Try
    End Function

    Private Function GetSeries() As String
        GetSeries = Nothing
        Try
            If (SeriesForCosting.Equals("TL (TC)") OrElse SeriesForCosting.Equals("TH (TD)") OrElse SeriesForCosting.Equals("TP-High") _
                                     OrElse SeriesForCosting.Equals("TP-Low") OrElse SeriesForCosting.Equals("LN")) Then      '21_01_2011   RAGAVA  LN added
                GetSeries = "TL/TH/TP"
            ElseIf SeriesForCosting.Equals("TX (TXC)") Then
                GetSeries = "TX"
            End If
        Catch ex As Exception
            GetSeries = Nothing
        End Try
    End Function

End Class
