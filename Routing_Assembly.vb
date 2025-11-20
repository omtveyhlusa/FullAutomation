Imports SQLdb
Public Class Routing_Assembly
    Public dbECN As New SqlDatabase("omt-sql03", "ENG", "eng_uploadtool", "jj*7^++-PPrqW")
    Public dbNAV As New SqlDatabase("omt-sql04", "OMTLIVE2015", "eng_uploadtool", "jj*7^++-PPrqW")
    Public dbPDM As New SqlDatabase("omt-epdm-db22", "OMT-Sandbox", "dbadmin", "ePDMadmin")

    Public sqlstr As String

    Sub New(ByVal PartNumber As String, ByVal description As String, ByVal ItemCategory As String, ByVal ECNNumber As String)
        Dim dtasy As New DataTable
        Dim dtECN As New DataTable
        Dim dtwc As New DataTable

        Try
            PartNumber = Replace(Replace(PartNumber, "A", ""), "G", "")

            sqlstr = "SELECT TOP 1 ValueText
                    FROM Documents AS D WITH (NOLOCK) INNER JOIN
                    VariableValue AS VV WITH (NOLOCK) ON D.DocumentID = VV.DocumentID
                    WHERE Filename LIKE '" & PartNumber & ".SLD%' AND Filename NOT LIKE '%.SLDDRW' AND Deleted = 0 AND VariableID = 71 AND ValueText NOT LIKE ''
                    AND RevisionNo = (SELECT MAX(CAST(RevisionNo AS INT)) FROM VariableValue AS VV2 WITH (NOLOCK) WHERE VV.DocumentID = VV2.DocumentID AND VariableID = 71 AND ValueText NOT LIKE '')"
            dtECN = dbPDM.FillTableFromSql(sqlstr)
            If dtECN.Rows.Count <> 0 Then
                ECNNumber = dtECN.Rows(0).Item(0)
            End If

            sqlstr = "SELECT No_
                    FROM [OMT-Veyhl USA Corporation$Work Center] WITH (NOLOCK)
                    WHERE Blocked = 0"
            dtwc = dbNAV.FillTableFromSql(sqlstr)
            Dim wc As String

            'if exists, don't do anything
            dtasy.Columns.Add("WorkCenter")
            dtasy.Columns.Add("RunTime")

            description = description.ToUpper

            If (description.Contains("RENEW") Or description.Contains("RNW")) Then
                For Each r In dtwc.Rows
                    If r.item(0).ToString.Contains("830") Then
                        wc = r.item(0)
                        Exit For
                    End If
                Next
                If ItemCategory = "94" Then
                    If description.Contains(" SSE ") And description.Contains(" PIL ") Then
                        dtasy.Rows.Add(wc, 0)
                        GoTo ending
                    ElseIf description.Contains(" DSE ") Or description.Contains(" DSL ") Or description.Contains(" DSH ") Then
                        dtasy.Rows.Add(wc, 0)
                        GoTo ending
                    End If
                ElseIf ItemCategory = "99" Then
                    If description.Contains(" SSE ") And description.Contains(" PIL ") Then
                        dtasy.Rows.Add(wc, 12.1 / 60)
                        GoTo ending
                    ElseIf description.Contains(" SSE ") Then
                        dtasy.Rows.Add(wc, 10.4 / 60)
                        GoTo ending
                    ElseIf description.Contains(" DSE ") Or description.Contains(" DSL ") Or description.Contains(" DSH ") Then
                        dtasy.Rows.Add(wc, 12.1 / 60)
                        GoTo ending
                    End If
                End If

            ElseIf (description.Contains("HORIZON") Or description.Contains("HZ") Or description.Contains("HRZN")) Then
                For Each r In dtwc.Rows
                    If r.item(0).ToString.Contains("806") Then
                        wc = r.item(0)
                        Exit For
                    End If
                Next
                dtasy.Rows.Add(wc, 0)


            ElseIf ((description.Contains("CLEVER") Or description.Contains("CLVR")) And Not (description.Contains(" SSE ") Or description.Contains(" DSL ") Or description.Contains(" DSH ") Or description.Contains(" DSE "))) Then
                For Each r In dtwc.Rows
                    If r.item(0).ToString.Contains("814") Then
                        wc = r.item(0)
                        Exit For
                    End If
                Next
                If description.Contains(" FOOT ") Or description.Contains(" FT ") Then
                    dtasy.Rows.Add(wc, 1.3 / 60)
                    GoTo ending
                ElseIf description.Contains(" CHANNEL") Or description.Contains(" CHNNL ") Then
                    dtasy.Rows.Add(wc, 2 / 60)
                    GoTo ending
                End If

            ElseIf (description.Contains("CLEVER") Or description.Contains("CLVR") Or description.Contains("NEVI") Or description.Contains("NVI") Or (description.Contains("PLANES") And description.Contains("LITE"))) Then
                For Each r In dtwc.Rows
                    If r.item(0).ToString.Contains("813") Then
                        wc = r.item(0)
                        Exit For
                    End If
                Next
                If ItemCategory = "94" Then
                    If description.Contains(" SSE ") Then
                        dtasy.Rows.Add(wc, 0)
                        GoTo ending
                    ElseIf description.Contains(" DSE ") Or description.Contains(" DSL ") Or description.Contains(" DSH ") Then
                        dtasy.Rows.Add(wc, 0)
                        GoTo ending
                    End If
                ElseIf ItemCategory = "99" Then
                    If description.Contains(" SSE ") Then
                        dtasy.Rows.Add(wc, 7.8 / 60)
                        GoTo ending
                    ElseIf description.Contains(" DSE ") Or description.Contains(" DSL ") Or description.Contains(" DSH ") Then
                        dtasy.Rows.Add(wc, 8.7 / 60)
                        GoTo ending
                    End If
                End If

            ElseIf (description.Contains("ESSENTIAL") Or description.Contains("ESSNTL") Or description.Contains("ESNTL") Or description.Contains("LAMBDA") Or description.Contains("LMBDA") Or description.Contains("ENTYRE") Or description.Contains("ENTYR")) Then
                For Each r In dtwc.Rows
                    If r.item(0).ToString.Contains("810") Then
                        wc = r.item(0)
                        Exit For
                    End If
                Next
                If description.Contains("LAMBDA") Or description.Contains("LMBDA") Then
                    dtasy.Rows.Add(wc, 7.8 / 60)
                    GoTo ending
                ElseIf description.Contains("ENTYRE") Or description.Contains("ENTYR") Then
                    dtasy.Rows.Add(wc, 15 / 60)
                    GoTo ending
                ElseIf description.Contains(" SSE ") Then
                    dtasy.Rows.Add(wc, 12 / 60)
                    GoTo ending
                ElseIf description.Contains(" DSL ") Or description.Contains(" DSH ") Or description.Contains(" DSE ") Then
                    dtasy.Rows.Add(wc, 15 / 60)
                    GoTo ending
                End If

            ElseIf Not (description.Contains("CLEVER") Or description.Contains("CLVR") Or description.Contains("RENEW") Or description.Contains("RNW")) And (description.Contains(" SSE ") Or description.Contains(" DSE ") Or description.Contains(" DSL ") Or description.Contains(" DSH ") Or description.Contains(" CRANK ") Or description.Contains(" CRNK ")) Then
                For Each r In dtwc.Rows
                    If r.item(0).ToString.Contains("850") Then
                        wc = r.item(0)
                        Exit For
                    End If
                Next
                If ItemCategory = "94" Then
                    If description.Contains(" SSE ") And Not (description.Contains(" CRANK ") Or description.Contains(" CRNK ")) Then
                        dtasy.Rows.Add(wc, 0)
                        GoTo ending
                    ElseIf description.Contains(" SSE ") And (description.Contains(" CRANK ") Or description.Contains(" CRNK ")) Then
                        dtasy.Rows.Add(wc, 0)
                        GoTo ending
                    End If
                ElseIf ItemCategory = "99" Then
                    If description.Contains(" SSE ") And Not (description.Contains(" CRANK ") Or description.Contains(" CRNK ")) Then
                        dtasy.Rows.Add(wc, 11 / 60)
                        GoTo ending
                    ElseIf description.Contains(" SSE ") And (description.Contains(" CRANK ") Or description.Contains(" CRNK ")) Then
                        dtasy.Rows.Add(wc, 15 / 60)
                        GoTo ending
                    ElseIf (description.Contains(" DSE ") Or description.Contains(" DSL ") Or description.Contains(" DSH ")) Then
                        dtasy.Rows.Add(wc, 15 / 60)
                        GoTo ending
                    End If
                End If

            ElseIf (description.Contains("PARAMETER") Or description.Contains("HARDWARE") Or description.Contains("HRDWR") Or description.Contains("HDW")) Then
                For Each r In dtwc.Rows
                    If r.item(0).ToString.Contains("811") Then
                        wc = r.item(0)
                        Exit For
                    End If
                Next
                If description.Contains("PARAMETER") Then
                    dtasy.Rows.Add(wc, 1 / 60)
                    GoTo ending
                ElseIf description.Contains("HARDWARE") Or description.Contains("HRDWR") Or description.Contains("HDW") Then
                    dtasy.Rows.Add(wc, 2 / 60)
                    GoTo ending
                Else
                    dtasy.Rows.Add(wc, 2 / 60)
                    GoTo ending
                End If
            Else
                'dtasy.Rows.Add("808", 2.2 / 60)
                'GoTo ending
            End If

#Region "Old logic"
            ''830
            'If (description.Contains("RENEW") Or description.Contains("RNW")) And ItemCategory = "94" Then
            '    If description.Contains(" SSE ") Then
            '        dtasy.Rows.Add("830", 0)
            '        GoTo ending
            '    ElseIf description.Contains(" DSE ") Or description.Contains(" DSL ") Or description.Contains(" DSH ") Then
            '        dtasy.Rows.Add("830", 0)
            '        GoTo ending
            '    End If
            'End If
            'If (description.Contains("RENEW") Or description.Contains("RNW")) And ItemCategory = "99" Then
            '    If description.Contains(" SSE ") And description.Contains(" PIL ") Then
            '        dtasy.Rows.Add("830", 10.4)
            '        GoTo ending
            '    ElseIf description.Contains(" SSE ") Then
            '        dtasy.Rows.Add("830", 10.4)
            '        GoTo ending
            '    ElseIf description.Contains(" DSE ") Or description.Contains(" DSL ") Or description.Contains(" DSH ") Then
            '        dtasy.Rows.Add("830", 12.1)
            '        GoTo ending
            '    End If
            'End If

            ''813
            'If (description.Contains("CLEVER") Or description.Contains("CLVR")) And ItemCategory = "94" Then
            '    If description.Contains(" SSE ") Then
            '        dtasy.Rows.Add("813", 0)
            '        GoTo ending
            '    ElseIf description.Contains(" DSE ") Or description.Contains(" DSL ") Or description.Contains(" DSH ") Then
            '        dtasy.Rows.Add("813", 0)
            '        GoTo ending
            '    End If
            'End If
            'If (description.Contains("CLEVER") Or description.Contains("CLVR")) And ItemCategory = "99" Then
            '    If description.Contains(" SSE ") Then
            '        dtasy.Rows.Add("813", 7.8)
            '        GoTo ending
            '    ElseIf description.Contains(" DSE ") Or description.Contains(" DSL ") Or description.Contains(" DSH ") Then
            '        dtasy.Rows.Add("813", 8.7)
            '        GoTo ending
            '    End If
            'End If

            ''850
            'If Not (description.Contains("CLEVER") Or description.Contains("CLVR") Or description.Contains("RENEW") Or description.Contains("RNW")) And ItemCategory = "94" Then
            '    If description.Contains(" SSE ") And Not (description.Contains(" CRANK ") Or description.Contains(" CRNK ")) Then
            '        dtasy.Rows.Add("850", 0)
            '        GoTo ending
            '    ElseIf description.Contains(" SSE ") And (description.Contains(" CRANK ") Or description.Contains(" CRNK ")) Then
            '        dtasy.Rows.Add("850", 0)
            '        GoTo ending
            '    End If
            'End If
            'If Not (description.Contains("CLEVER") Or description.Contains("CLVR") Or description.Contains("RENEW") Or description.Contains("RNW")) And ItemCategory = "99" Then
            '    If description.Contains(" SSE ") And Not (description.Contains(" CRANK ") Or description.Contains(" CRNK ")) Then
            '        dtasy.Rows.Add("850", 11)
            '        GoTo ending
            '    ElseIf description.Contains(" SSE ") And (description.Contains(" CRANK ") Or description.Contains(" CRNK ")) Then
            '        dtasy.Rows.Add("850", 15)
            '        GoTo ending
            '    ElseIf (description.Contains(" DSE ") Or description.Contains(" DSL ") Or description.Contains(" DSH ")) Then
            '        dtasy.Rows.Add("850", 15)
            '        GoTo ending
            '    End If
            'End If

            ''810
            'If (description.Contains("ESSENTIAL") Or description.Contains("ESSNTL") Or description.Contains("ESNTL") Or description.Contains("LAMBDA") Or description.Contains("LMBDA") Or description.Contains("ENTYRE") Or description.Contains("ENTYR")) And ItemCategory = "94" Then

            'End If
            'If (description.Contains("ESSENTIAL") Or description.Contains("ESSNTL") Or description.Contains("ESNTL") Or description.Contains("LAMBDA") Or description.Contains("LMBDA") Or description.Contains("ENTYRE") Or description.Contains("ENTYR")) And ItemCategory = "99" Then
            '    If description.Contains("LAMBDA") Or description.Contains("LMBDA") Then
            '        dtasy.Rows.Add("810", 7.8)
            '        GoTo ending
            '    ElseIf description.Contains("ENTYRE") Or description.Contains("ENTYR") Then
            '        dtasy.Rows.Add("810", 15)
            '        GoTo ending
            '    ElseIf description.Contains(" SSE ") Then
            '        dtasy.Rows.Add("810", 12)
            '        GoTo ending
            '    ElseIf description.Contains(" DSL ") Or description.Contains(" DSH ") Or description.Contains(" DSE ") Then
            '        dtasy.Rows.Add("810", 15)
            '        GoTo ending
            '    End If
            'End If

            ''814
            'If ((description.Contains("CLEVER") Or description.Contains("CLVR")) And Not (description.Contains(" SSE ") Or description.Contains(" DSL ") Or description.Contains(" DSH ") Or description.Contains(" DSE "))) And ItemCategory = "94" Then

            'End If
            'If ((description.Contains("CLEVER") Or description.Contains("CLVR")) And Not (description.Contains(" SSE ") Or description.Contains(" DSL ") Or description.Contains(" DSH ") Or description.Contains(" DSE "))) And ItemCategory = "99" Then
            '    If description.Contains(" FOOT ") Or description.Contains(" FT ") Then
            '        dtasy.Rows.Add("814", 1.3)
            '        GoTo ending
            '    ElseIf description.Contains(" CHANNEL") Or description.Contains(" CHNNL ") Then
            '        dtasy.Rows.Add("814", 2)
            '        GoTo ending
            '    End If
            'End If

            ''811
            'If description.Contains("PARAMETER") Then
            '    dtasy.Rows.Add("811", 1)
            '    GoTo ending
            'ElseIf description.Contains("HARDWARE") Or description.Contains("HRDWR") Or description.Contains("HDW") Then
            '    dtasy.Rows.Add("811", 2)
            '    GoTo ending
            'Else
            '    dtasy.Rows.Add("811", 2)
            '    GoTo ending
            'End If
#End Region
ending:
            If dtasy.Rows.Count <> 0 Then
                'Add to ECN Routing
                If Not ECNNumber.Contains("ECN") Then
                    If ECNNumber.Length = 4 Then
                        ECNNumber = "ECN-0" & ECNNumber
                    Else
                        ECNNumber = "ECN-" & ECNNumber
                    End If
                End If

                sqlstr = "DELETE FROM [dbo].[ECN ME Assembly] WHERE [ECN No_] LIKE '" & ECNNumber & "' AND [Part No_] LIKE '" & PartNumber & "'"
                dbECN.ExecuteNonQueryForSql(sqlstr)

                For Each r As DataRow In dtasy.Rows
                    sqlstr = "INSERT INTO [dbo].[ECN ME Assembly]
                                   ([ECN No_]
                                   ,[ECN Rev]
                                   ,[Part No_]
                                   ,[Run Time]
                                   ,[Work Center No_]
                                   ,[Is Changing]
                                   ,[Work Instructions]
                                   ,[Comments]
                                   ,[Line])
                           VALUES
                                 ('" & ECNNumber & "'
                                 ,'0'
                                 ,'" & PartNumber & "'
                                 ,'" & r.Item(1) & "'
                                 ,'" & r.Item(0) & "'
                                 ,'1'
                                 ,''
                                 ,'Automated Value'
                                 ,'10')"
                    dbECN.ExecuteNonQueryForSql(sqlstr)
                Next

                '---Form Stuff---
                'Form1.DgvRouting.DataSource = dtasy
            End If
        Catch ex As Exception
            Debug.Print("Error found!")


        End Try
    End Sub
End Class
