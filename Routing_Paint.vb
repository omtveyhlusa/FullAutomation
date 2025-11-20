Imports SolidWorks.Interop.swconst
'Imports SolidWorks.Interop.swdocumentmgr
Imports SolidWorks.Interop.sldworks
Imports EPDM.Interop.epdm
Public Class Routing_Paint
    'Sql
    Public dbECN As New SQLdb.SqlDatabase("omt-sql03", "ENG", "eng_uploadtool", "jj*7^++-PPrqW")
    Public dbPDM As New SQLdb.SqlDatabase("omt-epdm-db22", "OMT-Sandbox", "dbadmin", "ePDMadmin")
    Public dbNAV As New SQLdb.SqlDatabase("omt-sql04", "OMTLIVE2015", "eng_uploadtool", "jj*7^++-PPrqW")
    Public sqlstr As String

    'Datatables
    Public dtColor As DataTable
    Public dtMatThickness As DataTable

    Public vault1 As EdmVault5

    Sub New(ByVal PartNumber As String, ByVal Revision As String, ByVal ItemCategory As String, ByVal swModel As ModelDoc2, ByVal swPart As PartDoc, ByVal swAssembly As AssemblyDoc, ByVal swApp As SldWorks)
        Dim swSelMgr As SelectionMgr
        Dim swSelData As SelectData
        Dim swBody As Body2
        Dim swFace As Face2
        Dim vBodies As Object
        Dim SurfArea As Double
        Dim PowderLB As Double
        Dim Weight As Double
        Dim outsource As Boolean
        Dim TestWeight As Object
        Dim ecnnum As String

        Try
            swModel.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)
            Dim swCustPropMgr = swModel.Extension.CustomPropertyManager("")
            ecnnum = swCustPropMgr.Get("ECN Number")
            If ecnnum.Contains("fromparent+") Then
                ecnnum = Replace(ecnnum, "fromparent+", "")
            End If
            'TestWeight = swModel.Extension.GetMassProperties2(1, 0, False)
            'Weight = TestWeight(5) * 2.205 'convert from Kilogram to Pound
            'pull basic information tables
            If ecnnum = "" Then
                Dim dtecn As New DataTable
                sqlstr = "SELECT TOP 1 ValueText
                        FROM Documents AS D WITH (NOLOCK) INNER JOIN
                        VariableValue AS VV WITH (NOLOCK) ON D.DocumentID = VV.DocumentID
                        WHERE Filename LIKE '" & PartNumber & ".SLD%' AND Filename NOT LIKE '%.SLDDRW' AND Deleted = 0 AND VariableID = 71 AND ValueText NOT LIKE ''
                        AND RevisionNo = (SELECT MAX(CAST(RevisionNo AS INT)) FROM VariableValue AS VV2 WITH (NOLOCK) WHERE VV.DocumentID = VV2.DocumentID AND VariableID = 71 AND ValueText NOT LIKE '')"
                dtecn = dbPDM.FillTableFromSql(sqlstr)
                If dtecn.Rows.Count <> 0 Then
                    ecnnum = dtecn.Rows(0).Item(0)
                End If
            End If
            sqlstr = "SELECT *
                    FROM PaintSpecs
                    WHERE Code LIKE '0B07'"
            dtColor = dbECN.FillTableFromSql(sqlstr)

            If ItemCategory = "96" Then
                TestWeight = swModel.Extension.GetMassProperties2(1, 0, False)
                Weight = TestWeight(5) * 2.205 'convert from Kilogram to Pound
                sqlstr = "SELECT Size, SizeMM
                    FROM FlatBedCutTime
                    ORDER BY SizeInch DESC"
                dtMatThickness = dbECN.FillTableFromSql(sqlstr)

#Region "Determine MIL usage based on thickness"
                'Determine MIL usage based on thickness
                dtMatThickness.Columns.Add("MIL")
                Dim minMIL As Double = dtColor.Rows(0).Item("Min MIL")
                Dim maxMIL As Double = dtColor.Rows(0).Item("Max MIL")
                Dim Gravity As Double = dtColor.Rows(0).Item("Gravity")
                Dim Coverage As Double = dtColor.Rows(0).Item("Coverage")

                Dim y1 As Double = minMIL
                Dim y2 As Double = maxMIL
                Dim x1 As Double
                Dim x2 As Double

                For Each r As DataRow In dtMatThickness.Rows
                    If dtMatThickness.Rows(0).Item("SizeMM") = r.Item("SizeMM") Then
                        r.Item("MIL") = maxMIL
                        x2 = r.Item("SizeMM")
                    ElseIf dtMatThickness.Rows(dtMatThickness.Rows.Count - 1).Item("SizeMM") = r.Item("SizeMM") Then
                        r.Item("MIL") = minMIL
                        x1 = r.Item("SizeMM")
                    End If
                Next

                Dim mSlope As Double = (Math.Abs(y1 - y2) / Math.Abs(x1 - x2))
                Dim bInt As Double = y2 - mSlope * x2

                For Each r As DataRow In dtMatThickness.Rows
                    If r.Item("MIL").ToString = "" Then
                        r.Item("MIL") = Math.Round(mSlope * r.Item("SizeMM") + bInt, 2)
                    End If
                Next
#End Region

                swSelMgr = swModel.SelectionManager
                swSelData = swSelMgr.CreateSelectData
                vBodies = swPart.GetBodies2(swBodyType_e.swAllBodies, True)
                swBody = vBodies(0)
                swFace = swBody.GetFirstFace
                Do While Not swFace Is Nothing
                    SurfArea += swFace.GetArea * 1000000
                    swFace = swFace.GetNextFace
                Loop

            ElseIf ItemCategory = "86" Then
                TestWeight = swModel.Extension.GetMassProperties2(1, 0, False)
                Weight = TestWeight(5) * 2.205 'convert from Kilogram to Pound
                Debug.Print("Part Weight: " & Weight)
                swSelMgr = swModel.SelectionManager
                swSelData = swSelMgr.CreateSelectData
                vBodies = swPart.GetBodies2(swBodyType_e.swAllBodies, True)
                swBody = vBodies(0)
                swFace = swBody.GetFirstFace
                Do While Not swFace Is Nothing
                    SurfArea += swFace.GetArea * 1000000
                    swFace = swFace.GetNextFace
                Loop
                SurfArea = SurfArea / 2
                Debug.Print("Tube Surface Area(for Paint): " & SurfArea)
            ElseIf ItemCategory = "95" Then
                'Disabled for now until a different method comes to mind

                'TestWeight = swModel.Extension.GetMassProperties2(1, 0, False)
                'Weight = TestWeight(5) * 2.205 'convert from Kilogram to Pound

                'swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swSaveAssemblyAsPartOptions, swSaveAsmAsPartOptions_e.swSaveAsmAsPart_ExteriorFaces)
                'swModel.SaveAs3("C:\temp\TempFile.SLDPRT", 0, 0)
                'swApp.CloseAllDocuments(True)
                'swApp.OpenDoc6("C:\temp\TempFile.SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_ReadOnly, 0, 0, 0)
                'swPart = swApp.ActiveDoc
                'swModel = swPart

                'swSelMgr = swModel.SelectionManager
                'swSelData = swSelMgr.CreateSelectData
                'vBodies = swPart.GetBodies2(swBodyType_e.swAllBodies, True)
                'swBody = vBodies(0)
                'swFace = swBody.GetFirstFace
                'Do While Not swFace Is Nothing
                '    SurfArea += swFace.GetArea * 1000000
                '    swFace = swFace.GetNextFace
                'Loop
            End If

            If SurfArea <> 0 Then
                'Convert from mm2 to ft2
                SurfArea = SurfArea / 92903
                Debug.Print("Surface Area converted to ft2: " & SurfArea)

                'Powder calculation
                'PowderLB = (192.3 * 0.6) / (1.4 * 1.44) 'paint standard number X material utilization & reclaim % / mil thickness X specific gravity
                'PowderLB = PowderLB / SurfArea
                'PowderLB = 1 / PowderLB
                'PowderLB = PowderLB * 5.5 'NUMBER BUFFER
                PowderLB = (SurfArea / 25) * 1.3
                Debug.Print("Powder usage:" & PowderLB)

                If Weight >= 100 Then
                    outsource = True
                Else
                    outsource = False
                End If
                Debug.Print("Outsource status:" & outsource)
                'Routing information
                Debug.Print("ECN Number (before modification): " & ecnnum)
                If Not ecnnum = "" Then
                    If Not ecnnum.Contains("ECN") Then
                        If ecnnum.Length = 4 Then
                            ecnnum = "ECN-0" & ecnnum
                        Else
                            ecnnum = "ECN-" & ecnnum
                        End If
                    End If

                    Debug.Print("ECN Number:" & ecnnum)
                    Debug.Print("Part Number: " & PartNumber & "P")

                    sqlstr = "DELETE FROM [dbo].[ECN ME Fabrication] WHERE [ECN No_] LIKE '" & ecnnum & "' AND [Part No_] LIKE '" & PartNumber & "P'"
                    dbECN.ExecuteNonQueryForSql(sqlstr)

                    If outsource = True Then
                        sqlstr = "INSERT INTO [dbo].[ECN ME Fabrication]
                          ([ECN No_]
                          ,[Part No_]
                          ,[Operation No_]
                          ,[Work Center]
                          ,[Run Time]
                          ,[Move Time]
                          ,[Routing Link Code]
                          ,[Fixture No_]
                          ,[Work Instructions]
                          ,[Is Changing]
                          ,[Parts Per Rack]
                          ,[Powder Qty]
                          ,[Comments]
                          ,[ECN Rev])
                    VALUES
                          ('" & ecnnum & "'
                          ,'" & PartNumber & "P'
                          ,10
                          ,'OUTSOURCE'
                          ,0
                          ,0
                          ,'CONS'
                          ,''
                          ,''
                          ,1
                          ,0
                          ,'" & PowderLB & "'
                          ,'Automated Value'
                          ,0)"
                        dbECN.ExecuteNonQueryForSql(sqlstr)
                    Else
                        sqlstr = "INSERT INTO [dbo].[ECN ME Fabrication]
                              ([ECN No_]
                              ,[Part No_]
                              ,[Operation No_]
                              ,[Work Center]
                              ,[Run Time]
                              ,[Move Time]
                              ,[Routing Link Code]
                              ,[Fixture No_]
                              ,[Work Instructions]
                              ,[Is Changing]
                              ,[Parts Per Rack]
                              ,[Powder Qty]
                              ,[Comments]
                              ,[ECN Rev])
                        VALUES
                              ('" & ecnnum & "'
                              ,'" & PartNumber & "P'
                              ,10
                              ,'601 LOAD'
                              ,0
                              ,480
                              ,'CONS'
                              ,''
                              ,''
                              ,1
                              ,0
                              ,'" & PowderLB & "'
                              ,'Automated Value'
                              ,0)"
                        dbECN.ExecuteNonQueryForSql(sqlstr)
                    End If
                End If
            End If
        Catch ex As Exception
            Debug.Print("Error found!")
            If PDM_Class.errmsg = "" Then
                PDM_Class.errmsg = "errored out on Paint Information"
            Else
                PDM_Class.errmsg += ", errored out on Paint Information"
            End If
            GoTo PaintEnd
        End Try
PaintEnd:
        '---Form Stuff---
        'Form1.txtWeight.Text = Weight
        'Form1.txtPaintEst.Text = PowderLB

        'Form1.dtPaintResult.Rows.Add(PartNumber & "P", PowderLB)

    End Sub
End Class
