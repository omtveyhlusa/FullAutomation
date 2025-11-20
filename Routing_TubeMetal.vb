Imports SolidWorks.Interop.swconst
'Imports SolidWorks.Interop.swdocumentmgr
Imports SolidWorks.Interop.sldworks
Imports SQLdb
Public Class Routing_TubeMetal
    Public swPart As PartDoc
    Public swDim As Dimension
    Public swSketch As Sketch
    Public swSketchMgr As SketchManager
    Public swSketchSegment As SketchSegment
    Public swFeatureMgr As FeatureManager
    Public swFeature As Feature
    Public swFace As Face2
    Public swBody As Body2
    Public clFold As Feature
    Public CustPropMgr As CustomPropertyManager
    Public swSheetFeatData As SheetMetalFeatureData
    Public err As String
    Public warn As String
    Public smThickness As Double
    Public cuttimecalc As Double
    Public CuttingTimeNitrogen As Double
    Public Minholesize As Double
    Public piercetime As Double
    Public Pulsesize As Double
    Public HoleFormula As String
    Public PulseFormula As String
    Public HoleConversion As String
    Public HoleStuff As Double
    Public HoleCutTime As Double
    Public CounterSinkRemov As Double
    Public holetype As Integer
    Public HoleArea As Double
    Public PartBlankArea As Double
    Public PartOuterCut As Double
    Public PartInnerCut As Double
    Public endcuttime As Double
    Public Innercuttime As Double
    Public TotalCutTime As Double
    Public CutArea As Double
    Public pierceCount As Integer
    Public CutLength As Double
    Public cuttingLengthInner As Double
    Public Cutouts As Integer
    Public traveltime As Double
    Public DesignOwn As String
    Public StandardHole As String() = {0, 5, 22, 25}
    Public TappedHole As String() = {31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56}
    Public CounterHole As String() = {3, 23, 24, 26, 27, 28, 29, 30, 43, 44, 45}
    Public FlowHole As String() = {}
    Public i As Integer
    Dim SqlConn As New SqlDatabase("omt-sql03", "ENG", "eng_uploadtool", "jj*7^++-PPrqW")
    Dim sqlstr As String
    Sub New(ByVal PartNum As String, ByVal swModel As ModelDoc2, ByVal swApp As SldWorks, ByRef fileID As Integer, ByRef foldID As Integer)
        Try
            For Each config In swModel.GetConfigurationNames
                If config Like "*Default" Then
                    swModel.ShowConfiguration(config)
                End If
            Next

            Dim tubelength As Double
            Dim LastSketch As String
            Dim lastSketchFt As Feature
            Dim fccount As Integer
            Dim flatfeat As String
            Dim ext1 As Double
            Dim ext2 As Double
            Dim ext3 As Double
            Dim LineLength1 As Double
            Dim LineLength2 As Double
            Dim RadLength As Double
            Dim endLineLength As Double
            Dim TubeType As String
            Dim holecount As Integer
            Dim holelength As Double
            Dim tapCount As Integer
            Dim flowDrillCount As Integer
            Dim countersnkcount As Integer
            Dim FormPart1 As Double
            Dim FormPart2 As Double
            Dim FormPart3 As Double
            Dim VarPlace1 As Integer
            Dim VarPlace2 As Integer
            Dim ptnHoleList As New List(Of String)
            Dim ptnFlowList As New List(Of String)
            Dim IsSM As Boolean
            Dim ECNNumber As String

            Dim Revision As String
            Dim Description As String
            Dim Material As String

            Dim swCustPropMgr = swModel.Extension.CustomPropertyManager("")
            ECNNumber = swCustPropMgr.Get("ECN Number")
            Revision = swCustPropMgr.Get("Revision")
            Description = swCustPropMgr.Get("Description")
            Material = swCustPropMgr.Get("MatDesc")
            DesignOwn = swCustPropMgr.Get("Design Ownership")
            If ECNNumber.Contains("fromparent+") Then
                ECNNumber = Replace(ECNNumber, "fromparent+", "")
            End If
            Debug.Print("ECN Number:" & ECNNumber)
            Debug.Print("Revision:" & Revision)
            Debug.Print("Description:" & Description)
            Debug.Print("Material:" & Material)
            Debug.Print("Design Ownership:" & DesignOwn)

            '---Modify Part for data extraction---
            Dim ftMgr As FeatureManager = swModel.FeatureManager
            Dim fts As Object = ftMgr.GetFeatures(False)
            For Each ft As Feature In fts
                If ft.IsSuppressed = False Then
                    Dim name As String = ft.Name
                    Dim ltype As String = ft.GetTypeName
                    If ft.GetTypeName = "CutListFolder" Then
                        clFold = ft
                    End If
                    If ft.Name.Contains("Convert-Solid") Then
                        IsSM = True
                    End If

                    'If ft.Name = "Boss-Extrude1" Then
                    If ft.Name.Contains("Boss-Extrude") Then
                        If Material.Contains("REC") Or Description.Contains("REC") Then
                            swDim = swModel.Parameter("D1@" & ft.Name.ToString & "")
                            tubelength = Math.Round(swDim.GetValue2("Default"), 2)
                            swDim = swModel.Parameter("D2@Sketch1")
                            ext1 = swDim.GetValue2("Default")
                            swDim = swModel.Parameter("D4@Sketch1")
                            ext2 = swDim.GetValue2("Default")
                            swDim = swModel.Parameter("D1@Sketch1")
                            ext3 = swDim.GetValue2("Default")
                            LineLength1 = ext1 - (ext3 * 2)
                            LineLength2 = ext2 - (ext3 * 2)
                            RadLength = ((ext3 * 2) * Math.PI) / 4
                            endLineLength = Math.Round(((RadLength * 4) + (Math.Abs(LineLength1) * 2) + (Math.Abs(LineLength2) * 2)), 2)
                            'swDim = swModel.Parameter("D3@Sketch1")
                            'smThickness = swDim.GetValue2("Default")
                            TubeType = "Rec"
                        ElseIf Material.Contains("SQ") Then
                            swDim = swModel.Parameter("D1@" & ft.Name.ToString & "")
                            tubelength = swDim.GetValue2("Default")
                            swDim = swModel.Parameter("D2@Sketch1")
                            ext1 = swDim.GetValue2("Default")
                            swDim = swModel.Parameter("D2@Sketch1")
                            ext2 = swDim.GetValue2("Default")
                            swDim = swModel.Parameter("D1@Sketch1")
                            ext3 = swDim.GetValue2("Default")
                            LineLength1 = ext1 - (ext3 * 2)
                            LineLength2 = ext2 - (ext3 * 2)
                            RadLength = ((ext3 * 2) * Math.PI) / 4
                            endLineLength = Math.Round(((RadLength * 4) + (Math.Abs(LineLength1) * 2) + (Math.Abs(LineLength2) * 2)), 2)
                            'swDim = swModel.Parameter("D3@Sketch1")
                            'smThickness = swDim.GetValue2("Default")
                            TubeType = "Sq"
                        ElseIf Material.Contains("RND") Then
                            TubeType = "Rnd"
                            GoTo ECNRouting
                        ElseIf Material.Contains("TUBE D ") Then
                            TubeType = "DTube"
                            GoTo ECNRouting
                        End If
                    End If
                End If
                Dim config As String = ft.GetTypeName2
                If ft.GetTypeName2.Contains("Flat") Then
                    flatfeat = ft.Name.ToString
                End If
            Next
            Debug.Print("Is part Sheet Metal? " & IsSM)
            Debug.Print("Type: " & TubeType)

            For Each smft As Feature In fts
                If smft.GetTypeName.Equals("SheetMetal") And smThickness = 0 Then
                    swSheetFeatData = smft.GetDefinition
                    smThickness = swSheetFeatData.Thickness * 1000
                    smThickness = Math.Round(smThickness, 2)
                    Exit For
                End If
            Next
            Debug.Print("Sheet Thickness: " & smThickness)

            If IsSM = False Then
                Debug.Print("Not a sheetmetal part. Task ending.")
                Exit Sub
            End If

            CutTimeTable()

            '        '---Extract Hole Data---
            For Each ft As Feature In fts
                Dim name As String = ft.Name
                Dim type As String = ft.GetTypeName

                If ft.IsSuppressed = False Then
                    Dim children As Object = ft.GetChildren
                    If Not IsNothing(children) Then
                        For Each child As Feature In children
                            Dim chld As String = child.GetTypeName
                            If child.GetTypeName = "LPattern" And child.IsSuppressed = False Then
                                Dim lPat As LinearPatternFeatureData = child.GetDefinition
                                Dim lPatFtArray As Object = lPat.PatternFeatureArray
                                For Each patFt As Feature In lPatFtArray
                                    If patFt.Name = ft.Name Then GoTo nextFeature
                                Next
                            End If
                            If child.GetTypeName = "CirPattern" And child.IsSuppressed = False Then
                                Dim cPat As CircularPatternFeatureData = child.GetDefinition
                                Dim cPatFtArray As Object = cPat.PatternFeatureArray
                                For Each patFt As Feature In cPatFtArray
                                    If patFt.Name = ft.Name Then GoTo nextFeature
                                Next
                            End If
                        Next
                    End If
                    If ft.GetTypeName = "LPattern" Then '------------------------------------------------------------LINEAR PATTERN--------------------------------------------------------------------
                        ft.GetDefinition()
                        Dim lPtn As LinearPatternFeatureData
                        lPtn = ft.GetDefinition
                        Dim ptnFtArray As Object
                        ptnFtArray = lPtn.PatternFeatureArray
                        For Each ptnFt As Feature In ptnFtArray
                            Dim holeFt As WizardHoleFeatureData2
                            Try
                                holeFt = ptnFt.GetDefinition
                            Catch ex As Exception
                                GoTo nextFeature
                            End Try
                            Dim holeType As Integer = holeFt.Type
                            'Standard holes
                            If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("Clearance") Or ptnFt.Name.ToString.Contains("Diameter") Or ptnFt.Name.ToString.Contains("Dowel") Then
                                If StandardHole.Contains(holeType) Then
                                    ptnHoleList.Add(ptnFt.Name)
                                    fccount = holeFt.GetSketchPointCount * ((lPtn.D1TotalInstances * lPtn.D2TotalInstances) - lPtn.GetSkippedItemCount)
                                    holecount += fccount
                                    If holeFt.ThruHoleDiameter * 1000 <= Minholesize Then
                                        'do nothing
                                    ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= Pulsesize Then
                                        Cutouts += fccount
                                        holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTime += (fccount * HoleStuff)
                                    Else
                                        Cutouts += fccount
                                        holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTime += (fccount * HoleStuff)
                                    End If
                                End If
                            End If
                            'Tapped holes
                            If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("Tapped") AndAlso Not ptnFt.Name.ToString.Contains("Flow") Then
                                If TappedHole.Contains(holeType) Then
                                    ptnHoleList.Add(ptnFt.Name)
                                    fccount = holeFt.GetSketchPointCount * ((lPtn.D1TotalInstances * lPtn.D2TotalInstances) - lPtn.GetSkippedItemCount)
                                    tapCount += fccount
                                    If holeFt.ThruTapDrillDiameter * 1000 <= Minholesize Then
                                        'do nothing
                                    ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= Pulsesize Then
                                        Cutouts += fccount
                                        holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTime += (fccount * HoleStuff)
                                    Else
                                        Cutouts += fccount
                                        holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTime += (fccount * HoleStuff)
                                    End If
                                End If
                            End If
                            'Countersink Holes
                            If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("CSK") Then
                                If CounterHole.Contains(holeType) Then
                                    ptnHoleList.Add(ptnFt.Name)
                                    fccount = holeFt.GetSketchPointCount * ((lPtn.D1TotalInstances * lPtn.D2TotalInstances) - lPtn.GetSkippedItemCount)
                                    countersnkcount += fccount
                                    If holeFt.ThruHoleDiameter * 1000 <= Minholesize Then
                                        'do nothing
                                    ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= Pulsesize Then
                                        Cutouts += fccount
                                        holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTime += (fccount * HoleStuff)
                                        CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                                    Else
                                        Cutouts += fccount
                                        holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTime += (fccount * HoleStuff)
                                        CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                                    End If
                                End If
                            End If
                            'Flowdrill Holes
                            If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("Flow") Then
                                ptnFlowList.Add(ptnFt.Name)
                                fccount = holeFt.GetSketchPointCount * ((lPtn.D1TotalInstances * lPtn.D2TotalInstances) - lPtn.GetSkippedItemCount)
                                flowDrillCount += fccount
                                If holeFt.ThruTapDrillDiameter * 1000 <= Minholesize Then
                                    'do nothing
                                ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= Pulsesize Then
                                    Cutouts += fccount
                                    holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                    VarPlace1 = HoleFormula.IndexOf("*")
                                    VarPlace2 = HoleFormula.IndexOf("^")
                                    FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                    FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                    FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                    HoleCutTime += (fccount * HoleStuff)
                                Else
                                    Cutouts += fccount
                                    holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                    VarPlace1 = HoleFormula.IndexOf("*")
                                    VarPlace2 = HoleFormula.IndexOf("^")
                                    FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                    FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                    FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                    HoleCutTime += (fccount * HoleStuff)
                                End If
                            End If
                        Next
                    End If
                    If ft.GetTypeName = "CirPattern" Then '----------------------------------------------------------CIRCULAR PATTERN-------------------------------------------------------------------
                        ft.GetDefinition() '--Gets Feature Definition
                        Dim cPtn As CircularPatternFeatureData
                        cPtn = ft.GetDefinition
                        Dim ptnFtArray As Object
                        ptnFtArray = cPtn.PatternFeatureArray
                        For Each ptnFt As Feature In ptnFtArray
                            Dim holeFt As WizardHoleFeatureData2
                            holeFt = ptnFt.GetDefinition
                            Dim holeType As Integer = holeFt.Type
                            'Standard Holes
                            If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("Clearance") Or ptnFt.Name.ToString.Contains("Diameter") Or ptnFt.Name.ToString.Contains("Dowel") Then
                                If StandardHole.Contains(holeType) Then
                                    ptnHoleList.Add(ptnFt.Name)
                                    fccount = (holeFt.GetSketchPointCount) * ((cPtn.TotalInstances + cPtn.TotalInstances2) - cPtn.GetSkippedItemCount)
                                    holecount += fccount
                                    If holeFt.ThruHoleDiameter * 1000 <= Minholesize Then
                                        'do nothing
                                    ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= Pulsesize Then
                                        Cutouts += fccount
                                        holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTime += (fccount * HoleStuff)
                                    Else
                                        Cutouts += fccount
                                        holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTime += (fccount * HoleStuff)
                                    End If
                                End If
                            End If
                            'Tapped Holes
                            If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("Tapped") AndAlso Not ptnFt.Name.ToString.Contains("Flow") Then
                                If TappedHole.Contains(holeType) Then
                                    ptnHoleList.Add(ptnFt.Name)
                                    fccount = (holeFt.GetSketchPointCount) * ((cPtn.TotalInstances + cPtn.TotalInstances2) - cPtn.GetSkippedItemCount)
                                    tapCount += fccount
                                    If holeFt.ThruTapDrillDiameter * 1000 <= Minholesize Then
                                        'do nothing
                                    ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= Pulsesize Then
                                        Cutouts += fccount
                                        holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTime += (fccount * HoleStuff)
                                    Else
                                        Cutouts += fccount
                                        holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTime += (fccount * HoleStuff)
                                    End If
                                End If
                            End If
                            'Countersink Holes
                            If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("CSK") Then
                                If CounterHole.Contains(holeType) Then
                                    ptnHoleList.Add(ptnFt.Name)
                                    fccount = (holeFt.GetSketchPointCount) * ((cPtn.TotalInstances + cPtn.TotalInstances2) - cPtn.GetSkippedItemCount)
                                    countersnkcount += fccount
                                    If holeFt.ThruHoleDiameter * 1000 <= Minholesize Then
                                        'do nothing
                                    ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= Pulsesize Then
                                        Cutouts += fccount
                                        holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTime += (fccount * HoleStuff)
                                        CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                                    Else
                                        Cutouts += fccount
                                        holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTime += (fccount * HoleStuff)
                                        CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                                    End If
                                End If
                            End If
                            'Flowdrill Holes
                            If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("Flow") Then
                                ptnFlowList.Add(ptnFt.Name)
                                fccount = (holeFt.GetSketchPointCount) * ((cPtn.TotalInstances + cPtn.TotalInstances2) - cPtn.GetSkippedItemCount)
                                flowDrillCount += fccount
                                If holeFt.ThruTapDrillDiameter * 1000 <= Minholesize Then
                                    'do nothing
                                ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= Pulsesize Then
                                    Cutouts += fccount
                                    holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                    VarPlace1 = HoleFormula.IndexOf("*")
                                    VarPlace2 = HoleFormula.IndexOf("^")
                                    FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                    FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                    FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                    HoleCutTime += (fccount * HoleStuff)
                                Else
                                    Cutouts += fccount
                                    holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                    VarPlace1 = HoleFormula.IndexOf("*")
                                    VarPlace2 = HoleFormula.IndexOf("^")
                                    FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                    FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                    FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                    HoleCutTime += (fccount * HoleStuff)
                                End If
                            End If
                        Next
                    End If

                    If ft.GetTypeName.Contains("Mirror") Then '---------------------------------------------------------MIRROR FEATURE---------------------------------------------------------------------
line1:                  ft.GetDefinition()
                        Try
                            Dim sumvalue As String = ft.Name
                            Dim mPtn As MirrorPatternFeatureData = ft.GetDefinition
                            'mPtn = ft.GetDefinition
                            Dim ptnFtArray As Object = mPtn.PatternFeatureArray
                            'ptnFtArray = mPtn.PatternFeatureArray
                            For Each ptnFt As Feature In ptnFtArray
                                If ptnFt.GetTypeName = "HoleWzd" Then
                                    Dim holeFt As WizardHoleFeatureData2
                                    holeFt = ptnFt.GetDefinition
                                    Dim holeType As Integer = holeFt.Type
                                    'Standard Holes
                                    If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("Clearance") Or ptnFt.Name.ToString.Contains("Diameter") Or ptnFt.Name.ToString.Contains("Dowel") Then
                                        If StandardHole.Contains(holeType) Then
                                            ptnHoleList.Add(ptnFt.Name)
                                            fccount = ft.GetFaceCount
                                            holecount += fccount
                                            If ((holeFt.ThruHoleDiameter * 1000) * Math.PI) <= Minholesize Then
                                                'do nothing
                                            ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= Pulsesize Then
                                                Cutouts += fccount
                                                holelength += ((fccount * (((holeFt.ThruHoleDiameter * 1000) * Math.PI))))
                                                VarPlace1 = HoleFormula.IndexOf("*")
                                                VarPlace2 = HoleFormula.IndexOf("^")
                                                FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                                FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                HoleCutTime += (fccount * HoleStuff)
                                            Else
                                                Cutouts += fccount
                                                holelength += ((fccount * (((holeFt.ThruHoleDiameter * 1000) * Math.PI))))
                                                VarPlace1 = HoleFormula.IndexOf("*")
                                                VarPlace2 = HoleFormula.IndexOf("^")
                                                FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                                FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                                FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                HoleCutTime += (fccount * HoleStuff)
                                            End If
                                        End If
                                    End If
                                    'Tapped Holes
                                    If TappedHole.Contains(holeType) And Not ptnHoleList.Contains(ft.Name) And ptnFt.Name.ToString.Contains("Tapped") AndAlso Not ptnFt.Name.ToString.Contains("Flow") Then
                                        fccount = ft.GetFaceCount
                                        If ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) <= Minholesize Then
                                            'do nothing
                                        ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= Pulsesize Then
                                            Cutouts += fccount
                                            tapCount += fccount
                                            holelength += fccount * (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI))
                                            VarPlace1 = HoleFormula.IndexOf("*")
                                            VarPlace2 = HoleFormula.IndexOf("^")
                                            FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                            FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                            FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                            HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                            HoleCutTime += (fccount * HoleStuff)
                                        Else
                                            Cutouts += fccount
                                            tapCount += fccount
                                            holelength += fccount * (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI))
                                            VarPlace1 = HoleFormula.IndexOf("*")
                                            VarPlace2 = HoleFormula.IndexOf("^")
                                            FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                            FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                            FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                            HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                            HoleCutTime += (fccount * HoleStuff)
                                        End If
                                    End If
                                    'Countersink Holes
                                    If CounterHole.Contains(holeType) And Not ptnHoleList.Contains(ft.Name) And ptnFt.Name.ToString.Contains("CSK") Then
                                        fccount = ft.GetFaceCount
                                        If ((holeFt.ThruHoleDiameter * 1000) * Math.PI) <= Minholesize Then
                                            'do nothing
                                        ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= Pulsesize Then
                                            Cutouts += fccount
                                            countersnkcount += fccount
                                            holelength += fccount * (((holeFt.ThruHoleDiameter * 1000) * Math.PI))
                                            VarPlace1 = HoleFormula.IndexOf("*")
                                            VarPlace2 = HoleFormula.IndexOf("^")
                                            FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                            FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                            FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                            HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                            HoleCutTime += (fccount * HoleStuff)
                                            CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                                        Else
                                            Cutouts += fccount
                                            countersnkcount += fccount
                                            holelength += fccount * (((holeFt.ThruHoleDiameter * 1000) * Math.PI))
                                            VarPlace1 = HoleFormula.IndexOf("*")
                                            VarPlace2 = HoleFormula.IndexOf("^")
                                            FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                            FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                            FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                            HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                            HoleCutTime += (fccount * HoleStuff)
                                            CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                                        End If
                                    End If
                                    'Flowdrill Holes
                                    If ptnFt.Name.ToString.Contains("Flow") Then
                                        fccount = ft.GetFaceCount
                                        If ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) <= Minholesize Then
                                            'do nothing
                                        ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= Pulsesize Then
                                            Cutouts += fccount
                                            flowDrillCount += fccount
                                            holelength += fccount * (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI))
                                            VarPlace1 = HoleFormula.IndexOf("*")
                                            VarPlace2 = HoleFormula.IndexOf("^")
                                            FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                            FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                            FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                            HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                            HoleCutTime += (fccount * HoleStuff)
                                        Else
                                            Cutouts += fccount
                                            flowDrillCount += fccount
                                            holelength += fccount * (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI))
                                            VarPlace1 = HoleFormula.IndexOf("*")
                                            VarPlace2 = HoleFormula.IndexOf("^")
                                            FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                            FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                            FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                            HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                            HoleCutTime += (fccount * HoleStuff)
                                        End If
                                    End If
                                ElseIf ptnFt.GetTypeName.Contains("Mirror") Then
                                    'MsgBox("Mirror")
                                    ft = ptnFt
                                    GoTo line1
                                ElseIf ptnFt.GetTypeName = "CirPattern" Then '-------------------------------------------------MIRROR CIRCULAR PATTERN------------------------------------------------------
                                    ptnFt.GetDefinition()
                                    Dim cPtn As CircularPatternFeatureData
                                    cPtn = ptnFt.GetDefinition
                                    Dim ptFtArray As Object
                                    ptFtArray = cPtn.PatternFeatureArray
                                    For Each ptFt As Feature In ptFtArray
                                        Dim holeFt As WizardHoleFeatureData2
                                        holeFt = ptFt.GetDefinition
                                        Dim holeType As Integer = holeFt.Type
                                        'Standard Holes
                                        If ptFt.GetTypeName = "HoleWzd" And ptFt.Name.ToString.Contains("Clearance") Or ptFt.Name.ToString.Contains("Diameter") Or ptFt.Name.ToString.Contains("Dowel") Then
                                            If StandardHole.Contains(holeType) Then
                                                ptnHoleList.Add(ptnFt.Name)
                                                fccount = (holeFt.GetSketchPointCount) * ((cPtn.TotalInstances + cPtn.TotalInstances2) - cPtn.GetSkippedItemCount)
                                                holecount += fccount
                                                If holeFt.ThruHoleDiameter * 1000 <= Minholesize Then
                                                    'do nothing
                                                ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= Pulsesize Then
                                                    Cutouts += fccount
                                                    holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                                    VarPlace1 = HoleFormula.IndexOf("*")
                                                    VarPlace2 = HoleFormula.IndexOf("^")
                                                    FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                    FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                                    FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                    HoleCutTime += (fccount * HoleStuff)
                                                Else
                                                    Cutouts += fccount
                                                    holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                                    VarPlace1 = HoleFormula.IndexOf("*")
                                                    VarPlace2 = HoleFormula.IndexOf("^")
                                                    FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                                    FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                                    FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                    HoleCutTime += (fccount * HoleStuff)
                                                End If
                                            End If
                                        End If
                                        'Tapped Holes
                                        If ptFt.GetTypeName = "HoleWzd" And ptFt.Name.ToString.Contains("Tapped") AndAlso Not ptnFt.Name.ToString.Contains("Flow") Then
                                            If TappedHole.Contains(holeType) Then
                                                ptnHoleList.Add(ptnFt.Name)
                                                fccount = (holeFt.GetSketchPointCount) * ((cPtn.TotalInstances + cPtn.TotalInstances2) - cPtn.GetSkippedItemCount)
                                                tapCount += fccount
                                                If holeFt.ThruTapDrillDiameter * 1000 <= Minholesize Then
                                                    'do nothing
                                                ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= Pulsesize Then
                                                    Cutouts += fccount
                                                    holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                                    VarPlace1 = HoleFormula.IndexOf("*")
                                                    VarPlace2 = HoleFormula.IndexOf("^")
                                                    FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                    FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                                    FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                    HoleCutTime += (fccount * HoleStuff)
                                                Else
                                                    Cutouts += fccount
                                                    holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                                    VarPlace1 = HoleFormula.IndexOf("*")
                                                    VarPlace2 = HoleFormula.IndexOf("^")
                                                    FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                                    FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                                    FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                    HoleCutTime += (fccount * HoleStuff)
                                                End If
                                            End If
                                        End If
                                        'Countersink Holes
                                        If ptFt.GetTypeName = "HoleWzd" And ptFt.Name.ToString.Contains("CSK") Then
                                            If CounterHole.Contains(holeType) Then
                                                ptnHoleList.Add(ptnFt.Name)
                                                fccount = (holeFt.GetSketchPointCount) * ((cPtn.TotalInstances + cPtn.TotalInstances2) - cPtn.GetSkippedItemCount)
                                                countersnkcount += fccount
                                                If holeFt.ThruHoleDiameter * 1000 <= Minholesize Then
                                                    'do nothing
                                                ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= Pulsesize Then
                                                    Cutouts += fccount
                                                    holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                                    VarPlace1 = HoleFormula.IndexOf("*")
                                                    VarPlace2 = HoleFormula.IndexOf("^")
                                                    FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                    FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                                    FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                    HoleCutTime += (fccount * HoleStuff)
                                                    CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                                                Else
                                                    Cutouts += fccount
                                                    holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                                    VarPlace1 = HoleFormula.IndexOf("*")
                                                    VarPlace2 = HoleFormula.IndexOf("^")
                                                    FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                                    FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                                    FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                    HoleCutTime += (fccount * HoleStuff)
                                                    CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                                                End If
                                            End If
                                        End If
                                        'Flowdrill Holes
                                        If ptFt.GetTypeName = "HoleWzd" And ptFt.Name.ToString.Contains("Flow") Then
                                            ptnFlowList.Add(ptnFt.Name)
                                            fccount = (holeFt.GetSketchPointCount) * ((cPtn.TotalInstances + cPtn.TotalInstances2) - cPtn.GetSkippedItemCount)
                                            flowDrillCount += fccount
                                            If holeFt.ThruTapDrillDiameter * 1000 <= Minholesize Then
                                                'do nothing
                                            ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= Pulsesize Then
                                                Cutouts += fccount
                                                holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                                VarPlace1 = HoleFormula.IndexOf("*")
                                                VarPlace2 = HoleFormula.IndexOf("^")
                                                FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                                FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                HoleCutTime += (fccount * HoleStuff)
                                            Else
                                                Cutouts += fccount
                                                holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                                VarPlace1 = HoleFormula.IndexOf("*")
                                                VarPlace2 = HoleFormula.IndexOf("^")
                                                FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                                FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                                FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                HoleCutTime += (fccount * HoleStuff)
                                            End If
                                        End If
                                    Next

                                ElseIf ptnFt.GetTypeName = "LPattern" Then  '---------------------------------------------------MIRROR LINEAR PATTERN-------------------------------------------------------
                                    Dim lPtn As LinearPatternFeatureData
                                    lPtn = ptnFt.GetDefinition
                                    Dim ptFtArray As Object = lPtn.PatternFeatureArray
                                    For Each parent As Feature In ptFtArray
                                        If parent.GetTypeName = "HoleWzd" Then
                                            Dim holeFt As WizardHoleFeatureData2
                                            holeFt = parent.GetDefinition
                                            Dim holeType As Integer = holeFt.Type
                                            'Standard Holes
                                            If parent.GetTypeName = "HoleWzd" And parent.Name.ToString.Contains("Clearance") Or parent.Name.ToString.Contains("Diameter") Or parent.Name.ToString.Contains("Dowel") Then
                                                If StandardHole.Contains(holeType) Then
                                                    ptnHoleList.Add(ptnFt.Name)
                                                    fccount = holeFt.GetSketchPointCount * ((lPtn.D1TotalInstances * lPtn.D2TotalInstances) - lPtn.GetSkippedItemCount)
                                                    holecount += fccount
                                                    If holeFt.ThruHoleDiameter * 1000 <= Minholesize Then
                                                        'do nothing
                                                    ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= Pulsesize Then
                                                        Cutouts += fccount
                                                        holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                                        VarPlace1 = HoleFormula.IndexOf("*")
                                                        VarPlace2 = HoleFormula.IndexOf("^")
                                                        FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                        FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                                        FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                        HoleCutTime += (fccount * HoleStuff)
                                                    Else
                                                        Cutouts += fccount
                                                        holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                                        VarPlace1 = HoleFormula.IndexOf("*")
                                                        VarPlace2 = HoleFormula.IndexOf("^")
                                                        FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                                        FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                                        FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                        HoleCutTime += (fccount * HoleStuff)
                                                    End If
                                                End If
                                            End If
                                            'Tapped Holes
                                            If parent.GetTypeName = "HoleWzd" And parent.Name.ToString.Contains("Tapped") AndAlso Not parent.Name.ToString.Contains("Flow") Then
                                                If TappedHole.Contains(holeType) Then
                                                    ptnHoleList.Add(ptnFt.Name)
                                                    fccount = holeFt.GetSketchPointCount * ((lPtn.D1TotalInstances * lPtn.D2TotalInstances) - lPtn.GetSkippedItemCount)
                                                    tapCount += fccount
                                                    If holeFt.ThruTapDrillDiameter * 1000 <= Minholesize Then
                                                        'do nothing
                                                    ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= Pulsesize Then
                                                        Cutouts += fccount
                                                        holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                                        VarPlace1 = HoleFormula.IndexOf("*")
                                                        VarPlace2 = HoleFormula.IndexOf("^")
                                                        FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                        FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                                        FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                        HoleCutTime += (fccount * HoleStuff)
                                                    Else
                                                        Cutouts += fccount
                                                        holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                                        VarPlace1 = HoleFormula.IndexOf("*")
                                                        VarPlace2 = HoleFormula.IndexOf("^")
                                                        FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                                        FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                                        FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                        HoleCutTime += (fccount * HoleStuff)
                                                    End If
                                                End If
                                            End If
                                            'Countersink Holes
                                            If parent.GetTypeName = "HoleWzd" And parent.Name.ToString.Contains("CSK") Then
                                                If CounterHole.Contains(holeType) Then
                                                    ptnHoleList.Add(ptnFt.Name)
                                                    fccount = holeFt.GetSketchPointCount * ((lPtn.D1TotalInstances * lPtn.D2TotalInstances) - lPtn.GetSkippedItemCount)
                                                    countersnkcount += fccount
                                                    If holeFt.ThruHoleDiameter * 1000 <= Minholesize Then
                                                        'do nothing
                                                    ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= Pulsesize Then
                                                        Cutouts += fccount
                                                        holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                                        VarPlace1 = HoleFormula.IndexOf("*")
                                                        VarPlace2 = HoleFormula.IndexOf("^")
                                                        FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                        FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                                        FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                        HoleCutTime += (fccount * HoleStuff)
                                                        CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                                                    Else
                                                        Cutouts += fccount
                                                        holelength += (((holeFt.ThruHoleDiameter * 1000) * Math.PI)) * fccount
                                                        VarPlace1 = HoleFormula.IndexOf("*")
                                                        VarPlace2 = HoleFormula.IndexOf("^")
                                                        FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                                        FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                                        FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                        HoleCutTime += (fccount * HoleStuff)
                                                        CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                                                    End If
                                                End If
                                            End If
                                            'Flowdrill Holes
                                            If parent.GetTypeName = "HoleWzd" And parent.Name.ToString.Contains("Flow") Then
                                                ptnFlowList.Add(ptnFt.Name)
                                                fccount = holeFt.GetSketchPointCount * ((lPtn.D1TotalInstances * lPtn.D2TotalInstances) - lPtn.GetSkippedItemCount)
                                                flowDrillCount += fccount
                                                If holeFt.ThruTapDrillDiameter * 1000 <= Minholesize Then
                                                    'do nothing
                                                ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= Pulsesize Then
                                                    Cutouts += fccount
                                                    holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                                    VarPlace1 = HoleFormula.IndexOf("*")
                                                    VarPlace2 = HoleFormula.IndexOf("^")
                                                    FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                    FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                                    FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                    HoleCutTime += (fccount * HoleStuff)
                                                Else
                                                    Cutouts += fccount
                                                    holelength += (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI)) * fccount
                                                    VarPlace1 = HoleFormula.IndexOf("*")
                                                    VarPlace2 = HoleFormula.IndexOf("^")
                                                    FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                                    FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                                    FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                    HoleCutTime += (fccount * HoleStuff)
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            Next
                        Catch ex As Exception
                        End Try
                    End If
                    If ft.GetTypeName = "HoleWzd" Then '----------------------------------------------------------------HOLE WIZARD-----------------------------------------------------------------------
                        Dim holeFt As WizardHoleFeatureData2
                        holeFt = ft.GetDefinition
                        holetype = holeFt.Type
                        '--Normal Hole--
                        If StandardHole.Contains(holetype) And Not ptnHoleList.Contains(ft.Name) Then
                            fccount = ft.GetFaceCount
                            holecount += fccount
                            If ((holeFt.ThruHoleDiameter * 1000) * Math.PI) <= Minholesize Then
                                'do nothing
                            ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= Pulsesize Then
                                Cutouts += fccount
                                HoleArea += (fccount * (((Math.PI / 4) * ((holeFt.ThruHoleDiameter * 1000) ^ 2))))
                                holelength += ((fccount * (((holeFt.ThruHoleDiameter * 1000) * Math.PI))))
                                VarPlace1 = HoleFormula.IndexOf("*")
                                VarPlace2 = HoleFormula.IndexOf("^")
                                FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                HoleCutTime += (fccount * HoleStuff)
                            Else
                                Cutouts += fccount
                                HoleArea += (fccount * (((Math.PI / 4) * ((holeFt.ThruHoleDiameter * 1000) ^ 2))))
                                holelength += ((fccount * (((holeFt.ThruHoleDiameter * 1000) * Math.PI))))
                                VarPlace1 = HoleFormula.IndexOf("*")
                                VarPlace2 = HoleFormula.IndexOf("^")
                                FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                HoleCutTime += (fccount * HoleStuff)
                            End If
                        End If
                        '--Tapped Hole--
                        If TappedHole.Contains(holetype) And Not ptnHoleList.Contains(ft.Name) And ft.Name.ToString.Contains("Tapped") AndAlso Not ft.Name.ToString.Contains("Flow") Then
                            fccount = ft.GetFaceCount
                            If ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) <= Minholesize Then
                                'do nothing
                            ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= Pulsesize Then
                                Cutouts += fccount
                                tapCount += fccount
                                HoleArea += (fccount * (((Math.PI / 4) * ((holeFt.ThruTapDrillDiameter * 1000) ^ 2))))
                                holelength += fccount * (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI))
                                VarPlace1 = HoleFormula.IndexOf("*")
                                VarPlace2 = HoleFormula.IndexOf("^")
                                FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                HoleCutTime += (fccount * HoleStuff)
                            Else
                                Cutouts += fccount
                                tapCount += fccount
                                HoleArea += (fccount * (((Math.PI / 4) * ((holeFt.ThruTapDrillDiameter * 1000) ^ 2))))
                                holelength += fccount * (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI))
                                VarPlace1 = HoleFormula.IndexOf("*")
                                VarPlace2 = HoleFormula.IndexOf("^")
                                FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                HoleCutTime += (fccount * HoleStuff)
                            End If
                        End If
                        '--Countersink Hole--
                        If CounterHole.Contains(holetype) And Not ptnHoleList.Contains(ft.Name) And ft.Name.ToString.Contains("CSK") Then
                            fccount = ft.GetFaceCount
                            If ((holeFt.ThruHoleDiameter * 1000) * Math.PI) <= Minholesize Then
                                'do nothing
                            ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= Pulsesize Then
                                Cutouts += fccount
                                countersnkcount += fccount
                                HoleArea += (fccount * (((Math.PI / 4) * ((holeFt.ThruHoleDiameter * 1000) ^ 2))))
                                holelength += fccount * (((holeFt.ThruHoleDiameter * 1000) * Math.PI))
                                VarPlace1 = HoleFormula.IndexOf("*")
                                VarPlace2 = HoleFormula.IndexOf("^")
                                FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                HoleCutTime += (fccount * HoleStuff)
                                CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                            Else
                                Cutouts += fccount
                                countersnkcount += fccount
                                HoleArea += (fccount * (((Math.PI / 4) * ((holeFt.ThruHoleDiameter * 1000) ^ 2))))
                                holelength += fccount * (((holeFt.ThruHoleDiameter * 1000) * Math.PI))
                                VarPlace1 = HoleFormula.IndexOf("*")
                                VarPlace2 = HoleFormula.IndexOf("^")
                                FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                                FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                HoleCutTime += (fccount * HoleStuff)
                                CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                            End If
                        End If
                        '--Flowdrill--
                        If ft.Name.ToString.Contains("Flow") Then
                            fccount = ft.GetFaceCount
                            If ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) <= Minholesize Then
                                'do nothing
                            ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= Pulsesize Then
                                Cutouts += fccount
                                flowDrillCount += fccount
                                HoleArea += (fccount * (((Math.PI / 4) * ((holeFt.ThruTapDrillDiameter * 1000) ^ 2))))
                                holelength += fccount * (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI))
                                VarPlace1 = HoleFormula.IndexOf("*")
                                VarPlace2 = HoleFormula.IndexOf("^")
                                FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                HoleCutTime += (fccount * HoleStuff)
                            Else
                                Cutouts += fccount
                                flowDrillCount += fccount
                                HoleArea += (fccount * (((Math.PI / 4) * ((holeFt.ThruTapDrillDiameter * 1000) ^ 2))))
                                holelength += fccount * (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI))
                                VarPlace1 = HoleFormula.IndexOf("*")
                                VarPlace2 = HoleFormula.IndexOf("^")
                                FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                                FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                HoleCutTime += (fccount * HoleStuff)
                            End If
                        End If
                    End If
                End If
nextFeature: Next
            Debug.Print("Total hole cutouts: " & Cutouts)
            For Each ft As Feature In fts
                If ft.GetTypeName = "HoleWzd" And ft.IsSuppressed = False Then
                    swModel.Extension.SelectByID2(ft.Name, "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
                    swModel.EditSuppress2()
                End If
            Next
            Debug.Print("Holes suppressed")

            'Create Cut and Flatten part
            Dim cutface As Integer

            cutface = swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            swSketchMgr = swModel.SketchManager
            swSketchMgr.InsertSketch(True)
            lastSketchFt = swSketchMgr.ActiveSketch
            LastSketch = lastSketchFt.Name
            swModel.ClearSelection2(True)
            cutface = swModel.Extension.SelectByID2(LastSketch, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
            cutface = swModel.SelectedFeatureProperties(0, 0, 0, 0, 0, 0, 0, True, False, "AutoSketch")
            swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference, False)
            swSketchMgr.CreateCornerRectangle(-20, 0.0000005, 0#, 20, -0.0000005, 0#)
            swModel.Extension.SelectByID2("AutoSketch", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)

            swModel.ClearSelection2(True)
            ftMgr = swModel.FeatureManager
            'swFeature = ftMgr.FeatureCut3(True, False, True, swEndConditions_e.swEndCondThroughNext, 0, 0.01, 0.01, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, swStartConditions_e.swStartSketchPlane, 0, False)
            swFeature = ftMgr.FeatureCut4(True, False, True, swEndConditions_e.swEndCondThroughNext, 0, 0.01, 0.01, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, swStartConditions_e.swStartSketchPlane, 0, False, True)
            Debug.Print("Added cut to part")

            swPart = swModel

            Dim biggestarea As Double
            Dim currentarea As Double
            Dim bigface As String
            Dim vBodies As Object = swPart.GetBodies2(0, True)
            swBody = vBodies(0)
            swFace = swBody.GetFirstFace
            Do While Not swFace Is Nothing
                currentarea = swFace.GetArea
                If currentarea > biggestarea Then
                    biggestarea = currentarea
                    bigface = swPart.GetEntityName(swFace)
                End If
                swFace = swFace.GetNextFace
            Loop

            swModel.Extension.SelectByID2(bigface, "FACE", 0, 0, 0, True, 0, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend8", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend7", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend6", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend5", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend4", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend3", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend2", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend1", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
            swModel.ClearSelection2(True)
            swModel.Extension.SelectByID2(bigface, "FACE", 0, 0, 0, False, 1, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend8", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend7", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend6", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend5", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend4", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend3", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend2", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            swModel.Extension.SelectByID2("RoundBend1", "BODYFEATURE", 0, 0, 0, True, 4, Nothing, 0)
            swModel.InsertSheetMetalUnfold()
            Debug.Print("Part Unbent")
            Debug.Print("Create external box to treat the tube as a sheet metal component")
            '---Create External "box" to get more accurate cut list---
            Dim X1 As Double
            Dim Y1 As Double
            Dim Z1 As Double
            Dim X2 As Double
            Dim Y2 As Double
            Dim Z2 As Double
            Dim outbody As Integer
            Dim dbox(5) As Double
            Dim bodies As Object = swPart.GetBodies2(swBodyType_e.swSolidBody, True)
            If Not IsNothing(bodies) Then
                Dim i As Integer
                For i = 0 To UBound(bodies)
                    Dim swBody As Body2
                    swBody = bodies(i)
                    Dim x As Double
                    Dim y As Double
                    Dim z As Double
                    swBody.GetExtremePoint(1, 0, 0, x, y, z)
                    If i = 0 Or x > X2 Then
                        X2 = x
                    End If
                    swBody.GetExtremePoint(-1, 0, 0, x, y, z)
                    If i = 0 Or x < X1 Then
                        X1 = x
                    End If
                    swBody.GetExtremePoint(0, 1, 0, x, y, z)
                    If i = 0 Or y > Y2 Then
                        Y2 = y
                    End If
                    swBody.GetExtremePoint(0, -1, 0, x, y, z)
                    If i = 0 Or y < Y1 Then
                        Y1 = y
                    End If
                    swBody.GetExtremePoint(0, 0, 1, x, y, z)
                    If i = 0 Or z > Z2 Then
                        Z2 = z
                    End If
                    swBody.GetExtremePoint(0, 0, -1, x, y, z)
                    If i = 0 Or z < Z1 Then
                        Z1 = z
                    End If
                Next
            End If
            dbox(0) = X1 * 1000 : dbox(1) = Y1 * 1000 : dbox(2) = Z1 * 1000
            dbox(3) = X2 * 1000 : dbox(4) = Y2 * 1000 : dbox(5) = Z2 * 1000
            Dim xdiff As Double = Math.Round((X1 * -1000) + (X2 * 1000), 2)
            Dim ydiff As Double = Math.Round((Y1 * -1000) + (Y2 * 1000), 2)
            Dim zdiff As Double = Math.Round((Z1 * -1000) + (Z2 * 1000), 2)
            If Math.Round(xdiff, 1) = Math.Round(smThickness, 1) Then
                X1 = 0
                X2 = 0

            ElseIf Math.Round(ydiff, 1) = Math.Round(smThickness, 1) Then
                Y1 = 0
                Y2 = 0

            Else
                Z1 = 0
                Z2 = 0

            End If
            outbody = swModel.Extension.SelectByID2(bigface, "FACE", 0, 0, 0, True, 0, Nothing, 0)
            swSketchMgr = swModel.SketchManager
            swSketchMgr.InsertSketch(True)
            Threading.Thread.Sleep(500)
            lastSketchFt = swSketchMgr.ActiveSketch
            LastSketch = lastSketchFt.Name
            swModel.ClearSelection2(True)
            outbody = swModel.Extension.SelectByID2(LastSketch, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
            outbody = swModel.SelectedFeatureProperties(0, 0, 0, 0, 0, 0, 0, True, False, "OutBody")
            swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference, False)
            If X1 = 0 Then
                swSketchMgr.CreateCornerRectangle(Y1, Z1, X1, Y2, Z2, X2)
            ElseIf Y1 = 0 Then
                swSketchMgr.CreateCornerRectangle(X1, Z1, Y1, X2, Z2, Y2)
                swModel.Extension.SelectByID2("OutBody", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                swSketchMgr.CreateCornerRectangle(X1 - 0.000001, Z1 - 0.000001, Y1, X2 + 0.000001, Z2 + 0.000001, Y2)
                swModel.Extension.SelectByID2("OutBody", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                swModel.ClearSelection2(True)

            Else
                swSketchMgr.CreateCornerRectangle(X1, Y1, Z1, X2, Y2, Z2)
                swModel.Extension.SelectByID2("OutBody", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                swSketchMgr.CreateCornerRectangle(X1 - 0.000001, Y1 - 0.000001, Z1, X2 + 0.000001, Y2 + 0.000001, Z2)
                swModel.Extension.SelectByID2("OutBody", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                swModel.ClearSelection2(True)

            End If
            ftMgr = swModel.FeatureManager
            swFeature = ftMgr.InsertSheetMetalBaseFlange2(0, 0, 0, 0, 0, 0, 0, 0, 0, Nothing, 0, 0, 0, 0, 0, 0, True, 0, True)
            swPart.SetEntityName(swFeature, "SPECIALTAB")
            Debug.Print("outside box added to treat tube like sheet metal")
            'remove uneeded cut lines from inner profile of tube

            Dim fArea As Double

            Dim lArray = New Double() {xdiff, ydiff, zdiff}
            Array.Sort(lArray)

            Dim TubeLongSide As Double = lArray(2)
            Dim TubeShortside As Double = lArray(1)
            Dim LengthToRemove As Double

            Dim facearr As Object
            facearr = swFeature.GetFaces
            For Each oneFace In facearr
                swFace = oneFace
                fArea = Math.Round(swFace.GetArea * 1000000, 2)
                Dim tempval As Double = Math.Round(fArea / smThickness, 2)
                If tempval <> TubeLongSide And tempval <> TubeShortside And tempval < TubeLongSide Then
                    If tempval <> TubeLongSide + 0.01 And tempval <> TubeShortside + 0.01 Then
                        LengthToRemove += tempval
                    End If
                End If
            Next

            'remove cut information from inside cut for better accuracy.
            If Not IsNothing(clFold) Then
                CustPropMgr = clFold.CustomPropertyManager
                Dim vCustPropName As Object = CustPropMgr.GetNames
                Dim CustPropName As String
                i = 0
                For i = LBound(vCustPropName) To UBound(vCustPropName)
                    CustPropName = vCustPropName(i)
                    Dim custPropVal As String = ""
                    Dim custPropResolvedOut As String = ""
                    CustPropMgr.Get2(CustPropName, custPropVal, custPropResolvedOut)
                    Select Case CustPropName
                        Case "Cutting Length-Inner"
                            cuttingLengthInner = custPropResolvedOut
                        Case "Cut Outs"
                            pierceCount = custPropResolvedOut
                    End Select
                Next
            End If

            cuttingLengthInner = Math.Abs(cuttingLengthInner - LengthToRemove)

            'calculations
            endcuttime = (endLineLength * 2) * 0.0377 '0.025 'Cutting the ends of the tube
            traveltime = pierceCount * 0.016 'Travel time from cutout to cutout
            Innercuttime = CuttingTimeNitrogen * cuttingLengthInner
            TotalCutTime = endcuttime + traveltime + Innercuttime + HoleCutTime



ECNRouting:
#Region "Routing Information"
            Dim dt2 As New DataTable
            dt2.Columns.Add("OperationNo")
            dt2.Columns.Add("WorkCenter")
            dt2.Columns.Add("WorkCenterName")
            dt2.Columns.Add("RunTime")
            dt2.Columns.Add("MoveTime")
            dt2.Columns.Add("RoutingLinkCode")
            dt2.Columns.Add("Fixture")
            dt2.Columns.Add("WorkInstructions")
            dt2.Columns.Add("Comments")


            'Everything to 151
            'if item is less than 2mm thick then 3 seconds per hole
            'if item is more than 2mm thick then 6 seconds per hole




            Dim OPnum As Integer = 20
            'Lasers
            If ext1 > 101.6 Or ext2 > 101.6 Then
                dt2.Rows.Add(10, "156 TB LSR", "TUBE LASER (TRUMPF)", TotalCutTime / 60, 0, "", "", "", "")
            ElseIf DesignOwn.Contains("BUILT") Then
                dt2.Rows.Add(10, "151 TB LSR", "Tube Laser-North Adige", TotalCutTime / 60, 0, "", "", "", "")
            ElseIf (ext1 = 90 Or ext1 = 50) And (ext2 = 90 Or ext2 = 50) Then
                dt2.Rows.Add(10, "153 TB LSR", "Tube Lasers LT5 (SOUTH)", TotalCutTime / 60, 0, "", "", "", "")
            ElseIf (ext1 = 100 Or ext1 = 60) And (ext2 = 100 Or ext2 = 60) Then
                dt2.Rows.Add(10, "154 TB LSR", "Tube Laser LT5 (NORTH)", TotalCutTime / 60, 0, "", "", "", "")
            ElseIf Description.Contains("CLEVER") Or Description.Contains("CLVR") Or Description.Contains("TUBE FORK") Then 'FORK TUBES ARE FOR 3x6, anything bigger is SAW
                dt2.Rows.Add(10, "156 TB LSR", "TUBE LASER (TRUMPF)", TotalCutTime / 60, 0, "", "", "", "")
            Else
                dt2.Rows.Add(10, "151 TB LSR", "Tube Laser-North Adige", TotalCutTime / 60, 0, "", "", "", "")
            End If
            'SAW
            'QUESTION THE DOUBLE TUBE FORK CODE
            If Description.Contains("TUBE FORK") Then
                dt2.Rows.Add(10, "204 SAW", "Band Saw", TotalCutTime / 4, 0, "", "", "", "")
                OPnum += 10
            End If
            'Drill Presses
            If flowDrillCount <> 0 Then
                dt2.Rows.Add(OPnum, "HURCO", "Flo Drill", (flowDrillCount * 15) / 60, "", "", "", "", "")
                OPnum += 10
            End If
            If countersnkcount <> 0 Then
                dt2.Rows.Add(OPnum, "202 C.SINK", "Countersink", (countersnkcount * 20) / 60, 0, "", "", "", "")
                OPnum += 10
            End If
            If tapCount <> 0 Then
                dt2.Rows.Add(OPnum, "203 TAP", "Tap", (tapCount * 10) / 60, 0, "", "", "", "")
                OPnum += 10
            End If
            'MoveTime
            Dim rt As Integer = dt2.Rows.Count
            dt2.Rows(rt - 1).Item("MoveTime") = 630
#End Region

            sqlstr = "DELETE FROM [dbo].[ECN ME Fabrication]
                                                   WHERE [ECN No_] LIKE '" & ECNNumber & "' AND [Part No_] LIKE '" & PartNum & "'"
            SqlConn.ExecuteNonQueryForSql(sqlstr)

            sqlstr = "DELETE FROM [dbo].[ECN ME Fabrication Calculated]
                                                   WHERE [ECN No_] LIKE '" & ECNNumber & "' AND [Part No_] LIKE '" & PartNum & "'"
            SqlConn.ExecuteNonQueryForSql(sqlstr)

            For Each r As DataRow In dt2.Rows
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
                                 ('" & ECNNumber & "'
                                 ,'" & PartNum & "'
                                 ,'" & r.Item(0) & "'
                                 ,'" & r.Item(1) & "'
                                 ,'" & r.Item(3) & "'
                                 ,'" & r.Item(4) & "'
                                 ,'" & r.Item(5) & "'
                                 ,''
                                 ,''
                                 ,'0'
                                 ,'0'
                                 ,'0'
                                 ,'Automated Value'
                                 ,'')"
                SqlConn.ExecuteNonQueryForSql(sqlstr)

                sqlstr = "INSERT INTO [dbo].[ECN ME Fabrication Calculated]
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
                              ('" & ECNNumber & "'
                              ,'" & PartNum & "'
                              ,'" & r.Item(0) & "'
                              ,'" & r.Item(1) & "'
                              ,'" & r.Item(3) & "'
                              ,'" & r.Item(4) & "'
                              ,'" & r.Item(5) & "'
                              ,''
                              ,''
                              ,'0'
                              ,'0'
                              ,'0'
                              ,'Automated Value'
                              ,'')"
                SqlConn.ExecuteNonQueryForSql(sqlstr)
            Next
        Catch ex As Exception
            PDM_Class.errmsg = "Failed to extract tube information"
            GoTo tubeEnd
        End Try
tubeEnd:
#Region "Testing"
        ''Form1.txtMatDesc.Text = MaterialDesc
        ''Form1.txtMatPN.Text = MaterialPN
        'Form1.txtEcnNum.Text = ECNNumber
        ''Form1.TxtOutsideCut.Text = cuttingLengthOuter
        'Form1.TxtInsideCut.Text = cuttingLengthInner
        ''Form1.TxtTotalHoleCut.Text = totalHoleLength
        ''Form1.TxtActualCutLength.Text = ActualCutLength
        'Form1.TxtThickness.Text = smThickness
        ''Form1.TxtBends.Text = bendCount
        ''Form1.txtWeight.Text = partWeight
        ''Form1.TxtTotalHoles.Text = totalholes
        'Form1.TxtNormalHoles.Text = holecount
        'Form1.TxtTappedHoles.Text = tapCount
        'Form1.TxtCounterSink.Text = countersnkcount
        'Form1.TxtFlowDrill.Text = flowDrillCount
        'Form1.DgvRouting.DataSource = dt2
#End Region
    End Sub
    Private Sub CutTimeTable()
        Dim dttubes As DataTable
        sqlstr = "SELECT [WorkCenter] ,[PierceTime_Sec],[SecPerMM],[MMperSec],[TravelTime_Sec]
                        FROM [dbo].[TubeCutTime]
                        WHERE WorkCenter LIKE '151'"
        dtTubes = SqlConn.FillTableFromSql(sqlstr)
        CuttingTimeNitrogen = dtTubes.Rows(0).Item("SecPerMM")
        piercetime = dttubes.Rows(0).Item("PierceTime_Sec")
        HoleFormula = "(0.1074 * (X ^ -0.553))"
        Minholesize = 1
        Pulsesize = 1
    End Sub
End Class
