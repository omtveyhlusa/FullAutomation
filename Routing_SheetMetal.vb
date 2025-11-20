Imports SolidWorks.Interop.swconst
'Imports SolidWorks.Interop.swdocumentmgr
Imports SolidWorks.Interop.sldworks
Imports System.Windows.Forms.VisualStyles

Public Class Routing_SheetMetal
    'Solidworks
    Public swImportStepData As ImportStepData

    Public swModel As ModelDoc2
    Public swPart As PartDoc
    Public swAssembly As AssemblyDoc
    Public swDrawing As DrawingDoc
    Public dwgTable As TableAnnotation
    Public swAnn As Annotation
    Public swFeature As Feature
    Public swSketch As Sketch
    Public swSketchMgr As SketchManager
    Public swSketchSegment As SketchSegment
    Public swView As View
    Public swModelDocExt As ModelDocExtension
    Public swSheet As Sheet
    Public swSelMgr As SelectionMgr
    Public swSelData As SelectData
    Public swBody As Body2
    Public swFace As Face2
    Public swFaceFeature As Feature
    Public swEnt As Entity = Nothing
    Public swSafeEnt As Entity = Nothing
    Public ftMgr As FeatureManager
    Public clFold As Feature
    Public swSheetFeatData As SheetMetalFeatureData
    Public fts As Object

    Public errs As String = Nothing
    Public cfgNames As Object
    Public cfgName As String

    Public StandardHole As String() = {0, 5, 22, 25}
    Public TappedHole As String() = {31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56}
    Public CounterHole As String() = {3, 23, 24, 26, 27, 28, 29, 30, 43, 44, 45}
    Public FlowHole As String() = {}

    'Lists
    Public ptnHoleList As New List(Of String)
    Public ptnFlowList As New List(Of String)

    'Dim retErr As String
    'Part information
    Public parttype As String

    Public PartNum As String
    Public ECNNumber As String
    Public MaterialPN As String
    Public MaterialDesc As String
    Public Description As String
    Public Revision As String
    Public ItemCategory As String
    Public DesignOwn As String
    Public BoundingBoxL As Double
    Public BoundingBoxW As Double
    Public BoundingBoxT As Double
    Public tubesize1 As Double = 0
    Public tubesize2 As Double = 0
    Public blankSize As Double
    Public Retval As Integer
    Public err As String
    Public warn As String
    Public smThickness As Double
    Public smTubeThickness As Double
    Public fccount As Integer
    Public cuttimecalc As Double
    Public CuttingTimeNitrogen As Double
    Public CuttingTimeOxygen As Double
    Public piercetime As Double
    Public pulsesize As Double
    Public HoleFormula As String
    Public PulseFormula As String
    Public HoleConversion As String
    Public HoleStuff As Double
    Public totalholes As Integer
    Public holetype As Integer
    Public holeSize As Double
    Public totalHoleLength As Double
    Public holecount As Integer
    Public travelLength As Double
    Public flowDrillCount As Integer
    Public countersnkcount As Integer
    Public tapCount As Integer
    Public pierceCount As Integer
    Public cuttingLengthOuter As Double
    Public cuttingLengthInner As Double
    Public minholesize As Double
    Public ActualCutLength As Double
    Public HoleCutTimeOxy As Double
    Public HoleCutTimeNit As Double
    Public PerimCutTime As Double
    Public TotalCutTime As Double
    Public CounterSinkRemov As Double
    Public tpholesize As Double
    Public bendlength As Double
    Public maxbendlength As Double
    Public partWeight As Double
    Public bendCount As Integer
    Public SurfaceArea As Double
    Public PaintUsage As Double


    'Calculation variables
    Public FormPart1 As Double
    Public FormPart2 As Double
    Public FormPart3 As Double
    Public VarPlace1 As Integer
    Public VarPlace2 As Integer
    Public Dir1 As Integer
    Public Dir2 As Integer

    Public currentArea As Double
    Public biggestArea As Double
    Public vBodies As Object
    Public bRet As Boolean

    Public i As Integer
    Public mcc As Integer

    'datatables
    Public dttemp As DataTable
    Public dtTubes As DataTable

    'Time
    Public t1 As Date
    Public t2 As Date
    Public tt As Integer

    'SQL
    Public SqlConn As New SQLdb.SqlDatabase("omt-sql03", "ENG", "eng_uploadtool", "jj*7^++-PPrqW")
    Public DBpdm As New SQLdb.SqlDatabase("omt-epdm-db22", "OMT-Sandbox", "dbadmin", "ePDMadmin")
    Public dbNAV As New SQLdb.SqlDatabase("omt-sql04", "OMTLIVE2015", "eng_uploadtool", "jj*7^++-PPrqW")
    Public sqlstr As String
    Sub New(ByVal PartNum As String, ByRef swModel As ModelDoc2, ByRef fileID As Integer, ByRef foldID As Integer)
        'Converts Solidworks document to IPS
        Debug.Print("Get Basic Sheet Metal Part Information")
        t1 = Date.Now
        swModel.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)
        Dim swCustPropMgr = swModel.Extension.CustomPropertyManager("")
        ECNNumber = swCustPropMgr.Get("ECN Number")
        Description = swCustPropMgr.Get("Description")
        Revision = swCustPropMgr.Get("Revision")
        ItemCategory = swCustPropMgr.Get("ITEM GROUP")
        MaterialDesc = swCustPropMgr.Get("matDesc")
        MaterialPN = swCustPropMgr.Get("MATERIAL PN")
        DesignOwn = swCustPropMgr.Get("Design Ownership")
        If ECNNumber.Contains("fromparent+") Then
            ECNNumber = Replace(ECNNumber, "fromparent+", "")
        End If
        ftMgr = swModel.FeatureManager
        fts = ftMgr.GetFeatures(False)
        'Grab Cut List Data
        Dim custPropMgr As CustomPropertyManager
        For Each ft As Feature In fts
            If ft.GetTypeName = "CutListFolder" Then
                clFold = ft
                Exit For
            End If
        Next
        t2 = Date.Now
        Debug.Print("Basic information collected" & " (" & DateDiff(DateInterval.Second, t1, t2) & ")")
        t1 = Date.Now
        Debug.Print("Get Cut Table Information")
        If Not IsNothing(clFold) Then
            custPropMgr = clFold.CustomPropertyManager
            Dim vCustPropName As Object = custPropMgr.GetNames
            Dim CustPropName As String
            i = 0
            For i = LBound(vCustPropName) To UBound(vCustPropName)
                CustPropName = vCustPropName(i)
                Dim custPropVal As String = ""
                Dim custPropResolvedOut As String = ""
                custPropMgr.Get2(CustPropName, custPropVal, custPropResolvedOut)
                Select Case CustPropName
                    Case "Bounding Box Length"
                        BoundingBoxL = custPropResolvedOut
                    Case "Bounding Box Width"
                        BoundingBoxW = custPropResolvedOut
                    Case "Cutting Length-Outer"
                        cuttingLengthOuter = custPropResolvedOut
                    Case "Cutting Length-Inner"
                        cuttingLengthInner = custPropResolvedOut
                    Case "Bends"
                        'If parttype.ToUpper = "SHEET" Or parttype.ToUpper = "PLATE" Then
                        bendCount = custPropResolvedOut
                        'End If
                    Case "Cut Outs"
                        pierceCount = custPropResolvedOut
                    Case "Mass"
                        partWeight = custPropResolvedOut / 454
                End Select
            Next
        End If


        t2 = Date.Now
        Debug.Print("Cut Table information pulled" & " (" & DateDiff(DateInterval.Second, t1, t2) & ")")
        t1 = Date.Now
        Debug.Print("Get Thickness and Max Bend Length")
        'Converts Solidworks document to MM
        swModel.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)

        For Each smft As Feature In fts
            If smft.GetTypeName.Equals("SheetMetal") Then
                swSheetFeatData = smft.GetDefinition
                smThickness = swSheetFeatData.Thickness * 1000
                smThickness = Math.Round(smThickness, 2)
                Exit For
            End If
        Next

        CutTimeTable()

        For Each ft As Feature In fts
            'Get max bend length
            If ft.GetTypeName2 = "ProfileFeature" And ft.Name.Contains("Bend-Lines") Then
                Try
                    Dim swSketch As Sketch
                    swSketch = ft.GetSpecificFeature2
                    Dim SketchSegs As Object
                    SketchSegs = swSketch.GetSketchSegments
                    Dim swSketchSeg As SketchSegment
                    For j = LBound(SketchSegs) To UBound(SketchSegs)
                        swSketchSeg = SketchSegs(j)
                        If swSketchSeg.ConstructionGeometry = True Then
                            bendlength = Math.Round(swSketchSeg.GetLength * 1000, 2)
                            If bendlength > maxbendlength Then
                                maxbendlength = bendlength
                            End If
                        End If
                    Next j
                Catch ex As Exception
                End Try
            End If
        Next
        t2 = Date.Now
        Debug.Print("Finished" & " (" & DateDiff(DateInterval.Second, t1, t2) & ")")
        t1 = Date.Now
        Debug.Print("Get Hole Information")
        For Each ft As Feature In fts
            Dim name As String = ft.Name
            Dim type As String = ft.GetTypeName


            If ft.IsSuppressed = False And (ft.GetTypeName = "LPattern" Or ft.GetTypeName = "CirPattern" Or ft.GetTypeName = "MirrorPattern" Or ft.GetTypeName = "HoleWzd") Then
                FormPart1 = Nothing
                FormPart2 = Nothing
                FormPart3 = Nothing
                Dim children As Object = ft.GetChildren
                If Not IsNothing(children) Then
                    For Each child As Feature In children
                        Dim chld As String = child.GetTypeName
                        If child.GetTypeName = "LPattern" And child.IsSuppressed = False Then
                            Dim lPat As LinearPatternFeatureData = child.GetDefinition
                            Dim lPatFtArray As Object = lPat.PatternFeatureArray
                            For Each patFt As Feature In lPatFtArray
                                If patFt.Name = ft.Name Then GoTo NextFeature
                            Next
                        End If
                        If child.GetTypeName = "CirPattern" And child.IsSuppressed = False Then
                            Dim cPat As CircularPatternFeatureData = child.GetDefinition
                            Dim cPatFtArray As Object = cPat.PatternFeatureArray
                            For Each patFt As Feature In cPatFtArray
                                If patFt.Name = ft.Name Then GoTo NextFeature
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
                        If ptnFt.GetTypeName = "HoleWzd" Then
                            Dim holeFt As WizardHoleFeatureData2
                            holeFt = ptnFt.GetDefinition
                            Dim holeType As Integer = holeFt.Type
                            'Standard holes
                            If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("Clearance") Or ptnFt.Name.ToString.Contains("Diameter") Or ptnFt.Name.ToString.Contains("Dowel") Then
                                If StandardHole.Contains(holeType) Then
                                    ptnHoleList.Add(ptnFt.Name)
                                    If lPtn.D1TotalInstances = 1 Then
                                        Dir1 = 1
                                    Else
                                        Dir1 = lPtn.D1TotalInstances - 1
                                    End If
                                    If lPtn.D2TotalInstances = 1 Then
                                        Dir2 = 1
                                    Else
                                        Dir2 = lPtn.D2TotalInstances - 1
                                    End If
                                    fccount = holeFt.GetSketchPointCount * ((Dir1 + Dir2) - lPtn.GetSkippedItemCount)
                                    holeSize = ((holeFt.ThruHoleDiameter * 1000) * Math.PI)
                                    If holeFt.ThruHoleDiameter * 1000 <= minholesize Then
                                        'do nothing
                                    ElseIf (holeFt.ThruHoleDiameter * 1000) >= pulsesize Then
                                        holecount += fccount
                                        HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                        HoleCutTimeOxy += (fccount * HoleStuff)

                                        HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                        HoleCutTimeNit += (fccount * HoleStuff)
                                    Else
                                        holecount += fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTimeOxy += (fccount * HoleStuff)
                                    End If
                                End If
                            End If
                            'Tapped holes
                            If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("Tapped") AndAlso Not ptnFt.Name.ToString.Contains("Flow") Then
                                If TappedHole.Contains(holeType) Then
                                    ptnHoleList.Add(ptnFt.Name)
                                    If lPtn.D1TotalInstances = 1 Then
                                        Dir1 = 1
                                    Else
                                        Dir1 = lPtn.D1TotalInstances - 1
                                    End If
                                    If lPtn.D2TotalInstances = 1 Then
                                        Dir2 = 1
                                    Else
                                        Dir2 = lPtn.D2TotalInstances - 1
                                    End If
                                    fccount = holeFt.GetSketchPointCount * ((Dir1 + Dir2) - lPtn.GetSkippedItemCount)
                                    If holeFt.ThruTapDrillDiameter <> 0 Then
                                        tpholesize = holeFt.ThruTapDrillDiameter
                                    ElseIf holeFt.TapDrillDiameter <> 0 Then
                                        tpholesize = holeFt.TapDrillDiameter
                                    End If
                                    holeSize = ((tpholesize * 1000) * Math.PI)
                                    If holeFt.ThruTapDrillDiameter * 1000 <= minholesize Then
                                        'do nothing
                                    ElseIf holeSize >= pulsesize Then
                                        tapCount += fccount
                                        HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                        HoleCutTimeOxy += (fccount * HoleStuff)

                                        HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                        HoleCutTimeNit += (fccount * HoleStuff)
                                    Else
                                        tapCount += fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTimeOxy += (fccount * HoleStuff)
                                    End If
                                End If
                            End If
                            'Countersink Holes
                            If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("CSK") Then
                                If CounterHole.Contains(holeType) Then
                                    ptnHoleList.Add(ptnFt.Name)
                                    If lPtn.D1TotalInstances = 1 Then
                                        Dir1 = 1
                                    Else
                                        Dir1 = lPtn.D1TotalInstances - 1
                                    End If
                                    If lPtn.D2TotalInstances = 1 Then
                                        Dir2 = 1
                                    Else
                                        Dir2 = lPtn.D2TotalInstances - 1
                                    End If
                                    fccount = holeFt.GetSketchPointCount * ((Dir1 + Dir2) - lPtn.GetSkippedItemCount)
                                    If holeFt.ThruHoleDiameter * 1000 <= minholesize Then
                                        'do nothing
                                    ElseIf (holeFt.ThruHoleDiameter * 1000) >= pulsesize Then
                                        countersnkcount += fccount
                                        HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                        HoleCutTimeOxy += (fccount * HoleStuff)

                                        HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                        HoleCutTimeNit += (fccount * HoleStuff)
                                    Else
                                        countersnkcount += fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTimeOxy += (fccount * HoleStuff)
                                        CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                                    End If
                                End If
                            End If
                            'Flowdrill Holes
                            If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("Flow") Then
                                ptnFlowList.Add(ptnFt.Name)
                                If lPtn.D1TotalInstances = 1 Then
                                    Dir1 = 1
                                Else
                                    Dir1 = lPtn.D1TotalInstances - 1
                                End If
                                If lPtn.D2TotalInstances = 1 Then
                                    Dir2 = 1
                                Else
                                    Dir2 = lPtn.D2TotalInstances - 1
                                End If
                                fccount = holeFt.GetSketchPointCount * ((Dir1 + Dir2) - lPtn.GetSkippedItemCount)
                                If (holeFt.ThruTapDrillDiameter * 1000) <= minholesize Then
                                    'do nothing
                                ElseIf (holeFt.ThruTapDrillDiameter * 1000) >= pulsesize Then
                                    flowDrillCount += fccount
                                    HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                    HoleCutTimeOxy += (fccount * HoleStuff)

                                    HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                    HoleCutTimeNit += (fccount * HoleStuff)
                                Else
                                    flowDrillCount += fccount
                                    VarPlace1 = HoleFormula.IndexOf("*")
                                    VarPlace2 = HoleFormula.IndexOf("^")
                                    FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                    FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                    FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                    HoleCutTimeOxy += (fccount * HoleStuff)
                                End If
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
                                holeSize = ((holeFt.ThruHoleDiameter * 1000) * Math.PI)
                                If holeFt.ThruHoleDiameter * 1000 <= minholesize Then
                                    'do nothing
                                ElseIf (holeFt.ThruHoleDiameter * 1000) >= pulsesize Then
                                    holecount += fccount
                                    HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                    HoleCutTimeOxy += (fccount * HoleStuff)

                                    HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                    HoleCutTimeNit += (fccount * HoleStuff)
                                Else
                                    holecount += fccount
                                    VarPlace1 = HoleFormula.IndexOf("*")
                                    VarPlace2 = HoleFormula.IndexOf("^")
                                    FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                    FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                    FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                    HoleCutTimeOxy += (fccount * HoleStuff)
                                End If
                            End If
                        End If
                        'Tapped Holes
                        If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("Tapped") AndAlso Not ptnFt.Name.ToString.Contains("Flow") Then
                            If TappedHole.Contains(holeType) Then
                                ptnHoleList.Add(ptnFt.Name)
                                fccount = (holeFt.GetSketchPointCount) * ((cPtn.TotalInstances + cPtn.TotalInstances2) - cPtn.GetSkippedItemCount)
                                If holeFt.ThruTapDrillDiameter <> 0 Then
                                    tpholesize = holeFt.ThruTapDrillDiameter
                                ElseIf holeFt.TapDrillDiameter <> 0 Then
                                    tpholesize = holeFt.TapDrillDiameter
                                End If
                                holeSize = ((tpholesize * 1000) * Math.PI)
                                If holeFt.ThruTapDrillDiameter * 1000 <= minholesize Then
                                    'do nothing
                                ElseIf holeSize >= pulsesize Then
                                    tapCount += fccount
                                    HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                    HoleCutTimeOxy += (fccount * HoleStuff)

                                    HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                    HoleCutTimeNit += (fccount * HoleStuff)
                                Else
                                    tapCount += fccount
                                    VarPlace1 = HoleFormula.IndexOf("*")
                                    VarPlace2 = HoleFormula.IndexOf("^")
                                    FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                    FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                    FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                    HoleCutTimeOxy += (fccount * HoleStuff)
                                End If
                            End If
                        End If
                        'Countersink Holes
                        If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("CSK") Then
                            If CounterHole.Contains(holeType) Then
                                ptnHoleList.Add(ptnFt.Name)
                                fccount = (holeFt.GetSketchPointCount) * ((cPtn.TotalInstances + cPtn.TotalInstances2) - cPtn.GetSkippedItemCount)
                                If holeFt.ThruHoleDiameter * 1000 <= minholesize Then
                                    'do nothing
                                ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= pulsesize Then
                                    countersnkcount += fccount
                                    HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                    HoleCutTimeOxy += (fccount * HoleStuff)

                                    HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                    HoleCutTimeNit += (fccount * HoleStuff)
                                Else
                                    countersnkcount += fccount
                                    VarPlace1 = HoleFormula.IndexOf("*")
                                    VarPlace2 = HoleFormula.IndexOf("^")
                                    FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                    FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                    FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                    HoleCutTimeOxy += (fccount * HoleStuff)
                                    CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                                End If
                            End If
                        End If
                        'Flowdrill Holes
                        If ptnFt.GetTypeName = "HoleWzd" And ptnFt.Name.ToString.Contains("Flow") Then
                            ptnFlowList.Add(ptnFt.Name)
                            fccount = (holeFt.GetSketchPointCount) * ((cPtn.TotalInstances + cPtn.TotalInstances2) - cPtn.GetSkippedItemCount)
                            If (holeFt.ThruTapDrillDiameter * 1000) <= minholesize Then
                                'do nothing
                            ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= pulsesize Then
                                flowDrillCount += fccount
                                HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                HoleCutTimeOxy += (fccount * HoleStuff)

                                HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                HoleCutTimeNit += (fccount * HoleStuff)
                            Else
                                flowDrillCount += fccount
                                VarPlace1 = HoleFormula.IndexOf("*")
                                VarPlace2 = HoleFormula.IndexOf("^")
                                FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                HoleCutTimeOxy += (fccount * HoleStuff)
                            End If
                        End If
                    Next
                End If
                If ft.GetTypeName.Contains("MirrorPattern") Then '---------------------------------------------------------MIRROR FEATURE---------------------------------------------------------------------
line1:              ft.GetDefinition()
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
                                        fccount = holeFt.GetSketchPointCount
                                        holeSize = ((holeFt.ThruHoleDiameter * 1000) * Math.PI)
                                        If (holeFt.ThruHoleDiameter * 1000) <= minholesize Then
                                            'do nothing
                                        ElseIf (holeFt.ThruHoleDiameter * 1000) >= pulsesize Then
                                            holecount += fccount
                                            HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                            HoleCutTimeOxy += (fccount * HoleStuff)

                                            HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                            HoleCutTimeNit += (fccount * HoleStuff)
                                        Else
                                            holecount += fccount
                                            VarPlace1 = HoleFormula.IndexOf("*")
                                            VarPlace2 = HoleFormula.IndexOf("^")
                                            FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                            FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                            FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                            HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                            HoleCutTimeOxy += (fccount * HoleStuff)
                                        End If
                                    End If
                                End If
                                'Tapped Holes
                                If TappedHole.Contains(holeType) And Not ptnHoleList.Contains(ft.Name) And ptnFt.Name.ToString.Contains("Tapped") AndAlso Not ptnFt.Name.ToString.Contains("Flow") Then
                                    fccount = holeFt.GetSketchPointCount

                                    If holeFt.ThruTapDrillDiameter <> 0 Then
                                        tpholesize = holeFt.ThruTapDrillDiameter
                                    ElseIf holeFt.TapDrillDiameter <> 0 Then
                                        tpholesize = holeFt.TapDrillDiameter
                                    End If
                                    holeSize = ((tpholesize * 1000) * Math.PI)
                                    If (holeFt.ThruTapDrillDiameter * 1000) <= minholesize Then
                                        'do nothing
                                    ElseIf holeSize >= pulsesize Then
                                        tapCount += fccount
                                        HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                        HoleCutTimeOxy += (fccount * HoleStuff)

                                        HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                        HoleCutTimeNit += (fccount * HoleStuff)
                                    Else
                                        tapCount += fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTimeOxy += (fccount * HoleStuff)
                                    End If
                                End If
                                'Countersink Holes
                                If CounterHole.Contains(holeType) And Not ptnHoleList.Contains(ft.Name) And ptnFt.Name.ToString.Contains("CSK") Then
                                    fccount = (holeFt.GetSketchPointCount)
                                    If ((holeFt.ThruHoleDiameter * 1000) * Math.PI) <= minholesize Then
                                        'do nothing
                                    ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= pulsesize Then
                                        countersnkcount += fccount
                                        HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                        HoleCutTimeOxy += (fccount * HoleStuff)

                                        HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                        HoleCutTimeNit += (fccount * HoleStuff)
                                    Else
                                        countersnkcount += fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTimeOxy += (fccount * HoleStuff)
                                        CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                                    End If
                                End If
                                'Flowdrill Holes
                                If ptnFt.Name.ToString.Contains("Flow") Then
                                    fccount = holeFt.GetSketchPointCount
                                    If (holeFt.ThruTapDrillDiameter * 1000) <= minholesize Then
                                        'do nothing
                                    ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= pulsesize Then
                                        flowDrillCount += fccount
                                        HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                        HoleCutTimeOxy += (fccount * HoleStuff)

                                        HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                        HoleCutTimeNit += (fccount * HoleStuff)
                                    Else
                                        flowDrillCount += fccount
                                        VarPlace1 = HoleFormula.IndexOf("*")
                                        VarPlace2 = HoleFormula.IndexOf("^")
                                        FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                        FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                        FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                        HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                        HoleCutTimeOxy += (fccount * HoleStuff)
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
                                            holeSize = ((holeFt.ThruHoleDiameter * 1000) * Math.PI)
                                            If holeFt.ThruHoleDiameter * 1000 <= minholesize Then
                                                'do nothing
                                            ElseIf (holeFt.ThruHoleDiameter * 1000) >= pulsesize Then
                                                holecount += fccount
                                                HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                                HoleCutTimeOxy += (fccount * HoleStuff)

                                                HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                                HoleCutTimeNit += (fccount * HoleStuff)
                                            Else
                                                holecount += fccount
                                                VarPlace1 = HoleFormula.IndexOf("*")
                                                VarPlace2 = HoleFormula.IndexOf("^")
                                                FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                                FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                HoleCutTimeOxy += (fccount * HoleStuff)
                                            End If
                                        End If
                                    End If
                                    'Tapped Holes
                                    If ptFt.GetTypeName = "HoleWzd" And ptFt.Name.ToString.Contains("Tapped") AndAlso Not ptFt.Name.ToString.Contains("Flow") Then
                                        If TappedHole.Contains(holeType) Then
                                            ptnHoleList.Add(ptnFt.Name)
                                            fccount = (holeFt.GetSketchPointCount) * ((cPtn.TotalInstances + cPtn.TotalInstances2) - cPtn.GetSkippedItemCount)
                                            If holeFt.ThruTapDrillDiameter <> 0 Then
                                                tpholesize = holeFt.ThruTapDrillDiameter
                                            ElseIf holeFt.TapDrillDiameter <> 0 Then
                                                tpholesize = holeFt.TapDrillDiameter
                                            End If
                                            holeSize = ((tpholesize * 1000) * Math.PI)
                                            If holeFt.ThruTapDrillDiameter * 1000 <= minholesize Then
                                                'do nothing
                                            ElseIf holeSize >= pulsesize Then
                                                tapCount += fccount
                                                HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                                HoleCutTimeOxy += (fccount * HoleStuff)

                                                HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                                HoleCutTimeNit += (fccount * HoleStuff)
                                            Else
                                                tapCount += fccount
                                                VarPlace1 = HoleFormula.IndexOf("*")
                                                VarPlace2 = HoleFormula.IndexOf("^")
                                                FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                                FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                HoleCutTimeOxy += (fccount * HoleStuff)
                                            End If
                                        End If
                                    End If
                                    'Countersink Holes
                                    If ptFt.GetTypeName = "HoleWzd" And ptFt.Name.ToString.Contains("CSK") Then
                                        If CounterHole.Contains(holeType) Then
                                            ptnHoleList.Add(ptnFt.Name)
                                            fccount = (holeFt.GetSketchPointCount) * ((cPtn.TotalInstances + cPtn.TotalInstances2) - cPtn.GetSkippedItemCount)
                                            If holeFt.ThruHoleDiameter * 1000 <= minholesize Then
                                                'do nothing
                                            ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= pulsesize Then
                                                countersnkcount += fccount
                                                HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                                HoleCutTimeOxy += (fccount * HoleStuff)

                                                HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                                HoleCutTimeNit += (fccount * HoleStuff)
                                            Else
                                                countersnkcount += fccount
                                                VarPlace1 = HoleFormula.IndexOf("*")
                                                VarPlace2 = HoleFormula.IndexOf("^")
                                                FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                                FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                HoleCutTimeOxy += (fccount * HoleStuff)
                                                CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                                            End If
                                        End If
                                    End If
                                    'Flowdrill Holes
                                    If ptFt.GetTypeName = "HoleWzd" And ptFt.Name.ToString.Contains("Flow") Then
                                        ptnFlowList.Add(ptnFt.Name)
                                        fccount = (holeFt.GetSketchPointCount) * ((cPtn.TotalInstances + cPtn.TotalInstances2) - cPtn.GetSkippedItemCount)
                                        If (holeFt.ThruTapDrillDiameter * 1000) <= minholesize Then
                                            'do nothing
                                        ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= pulsesize Then
                                            flowDrillCount += fccount
                                            HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                            HoleCutTimeOxy += (fccount * HoleStuff)

                                            HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                            HoleCutTimeNit += (fccount * HoleStuff)
                                        Else
                                            flowDrillCount += fccount
                                            VarPlace1 = HoleFormula.IndexOf("*")
                                            VarPlace2 = HoleFormula.IndexOf("^")
                                            FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                            FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                            FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                            HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                            HoleCutTimeOxy += (fccount * HoleStuff)
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
                                                If lPtn.D1TotalInstances = 1 Then
                                                    Dir1 = 1
                                                Else
                                                    Dir1 = lPtn.D1TotalInstances - 1
                                                End If
                                                If lPtn.D2TotalInstances = 1 Then
                                                    Dir2 = 1
                                                Else
                                                    Dir2 = lPtn.D2TotalInstances - 1
                                                End If
                                                fccount = ptnFt.GetFaceCount * ((Dir1 * Dir2) - lPtn.GetSkippedItemCount)
                                                holeSize = ((holeFt.ThruHoleDiameter * 1000) * Math.PI)
                                                If (holeFt.ThruHoleDiameter * 1000) <= minholesize Then
                                                    'do nothing
                                                ElseIf (holeFt.ThruHoleDiameter * 1000) >= pulsesize Then
                                                    holecount += fccount
                                                    HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                                    HoleCutTimeOxy += (fccount * HoleStuff)

                                                    HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                                    HoleCutTimeNit += (fccount * HoleStuff)
                                                Else
                                                    holecount += fccount
                                                    VarPlace1 = HoleFormula.IndexOf("*")
                                                    VarPlace2 = HoleFormula.IndexOf("^")
                                                    FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                    FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                                    FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                    HoleCutTimeOxy += (fccount * HoleStuff)
                                                End If
                                            End If
                                        End If
                                        'Tapped Holes
                                        If parent.GetTypeName = "HoleWzd" And parent.Name.ToString.Contains("Tapped") AndAlso Not parent.Name.ToString.Contains("Flow") Then
                                            If TappedHole.Contains(holeType) Then
                                                ptnHoleList.Add(ptnFt.Name)
                                                If lPtn.D1TotalInstances = 1 Then
                                                    Dir1 = 1
                                                Else
                                                    Dir1 = lPtn.D1TotalInstances - 1
                                                End If
                                                If lPtn.D2TotalInstances = 1 Then
                                                    Dir2 = 1
                                                Else
                                                    Dir2 = lPtn.D2TotalInstances - 1
                                                End If
                                                fccount = ptnFt.GetFaceCount * ((Dir1 * Dir2) - lPtn.GetSkippedItemCount)
                                                If holeFt.ThruTapDrillDiameter <> 0 Then
                                                    tpholesize = holeFt.ThruTapDrillDiameter
                                                ElseIf holeFt.TapDrillDiameter <> 0 Then
                                                    tpholesize = holeFt.TapDrillDiameter
                                                End If
                                                holeSize = ((tpholesize * 1000) * Math.PI)
                                                If holeFt.ThruTapDrillDiameter * 1000 <= minholesize Then
                                                    'do nothing
                                                ElseIf holeSize >= pulsesize Then
                                                    tapCount += fccount
                                                    HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                                    HoleCutTimeOxy += (fccount * HoleStuff)

                                                    HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                                    HoleCutTimeNit += (fccount * HoleStuff)
                                                Else
                                                    tapCount += fccount
                                                    VarPlace1 = HoleFormula.IndexOf("*")
                                                    VarPlace2 = HoleFormula.IndexOf("^")
                                                    FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                    FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                                    FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                    HoleCutTimeOxy += (fccount * HoleStuff)
                                                End If
                                            End If
                                        End If
                                        'Countersink Holes
                                        If parent.GetTypeName = "HoleWzd" And parent.Name.ToString.Contains("CSK") Then
                                            If CounterHole.Contains(holeType) Then
                                                ptnHoleList.Add(ptnFt.Name)
                                                If lPtn.D1TotalInstances = 1 Then
                                                    Dir1 = 1
                                                Else
                                                    Dir1 = lPtn.D1TotalInstances - 1
                                                End If
                                                If lPtn.D2TotalInstances = 1 Then
                                                    Dir2 = 1
                                                Else
                                                    Dir2 = lPtn.D2TotalInstances - 1
                                                End If
                                                fccount = ptnFt.GetFaceCount * ((Dir1 * Dir2) - lPtn.GetSkippedItemCount)
                                                If holeFt.ThruHoleDiameter * 1000 <= minholesize Then
                                                    'do nothing
                                                ElseIf ((holeFt.ThruHoleDiameter * 1000) * Math.PI) >= pulsesize Then
                                                    countersnkcount += fccount
                                                    HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                                    HoleCutTimeOxy += (fccount * HoleStuff)

                                                    HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                                    HoleCutTimeNit += (fccount * HoleStuff)
                                                Else
                                                    countersnkcount += fccount
                                                    VarPlace1 = HoleFormula.IndexOf("*")
                                                    VarPlace2 = HoleFormula.IndexOf("^")
                                                    FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                    FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                                    FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                    HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                    HoleCutTimeOxy += (fccount * HoleStuff)
                                                    CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                                                End If
                                            End If
                                        End If
                                        'Flowdrill Holes
                                        If parent.GetTypeName = "HoleWzd" And parent.Name.ToString.Contains("Flow") Then
                                            ptnFlowList.Add(ptnFt.Name)
                                            If lPtn.D1TotalInstances = 1 Then
                                                Dir1 = 1
                                            Else
                                                Dir1 = lPtn.D1TotalInstances - 1
                                            End If
                                            If lPtn.D2TotalInstances = 1 Then
                                                Dir2 = 1
                                            Else
                                                Dir2 = lPtn.D2TotalInstances - 1
                                            End If
                                            fccount = ptnFt.GetFaceCount * ((Dir1 * Dir2) - lPtn.GetSkippedItemCount)
                                            If (holeFt.ThruTapDrillDiameter * 1000) <= minholesize Then
                                                'do nothing
                                            ElseIf ((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) >= pulsesize Then
                                                flowDrillCount += fccount
                                                HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                                HoleCutTimeOxy += (fccount * HoleStuff)

                                                HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                                                HoleCutTimeNit += (fccount * HoleStuff)
                                            Else
                                                flowDrillCount += fccount
                                                VarPlace1 = HoleFormula.IndexOf("*")
                                                VarPlace2 = HoleFormula.IndexOf("^")
                                                FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                                                FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                                                FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                                                HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                                                HoleCutTimeOxy += (fccount * HoleStuff)
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
                        fccount = holeFt.GetSketchPointCount
                        holeSize = ((holeFt.ThruHoleDiameter * 1000) * Math.PI)
                        If (holeFt.ThruHoleDiameter * 1000) <= minholesize Then
                            'do nothing
                        ElseIf (holeFt.ThruHoleDiameter * 1000) >= pulsesize Then
                            holecount += fccount
                            HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                            HoleCutTimeOxy += (fccount * HoleStuff)

                            HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                            HoleCutTimeNit += (fccount * HoleStuff)
                        Else
                            holecount += fccount
                            VarPlace1 = PulseFormula.IndexOf("*")
                            VarPlace2 = PulseFormula.IndexOf("^")
                            FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                            FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                            FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                            HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                            HoleCutTimeOxy += (fccount * HoleStuff)
                        End If
                    End If
                    '--Tapped Hole--
                    If TappedHole.Contains(holetype) And Not ptnHoleList.Contains(ft.Name) And ft.Name.ToString.Contains("Tapped") AndAlso Not ft.Name.ToString.Contains("Flow") Then
                        fccount = holeFt.GetSketchPointCount
                        If holeFt.ThruTapDrillDiameter <> 0 Then
                            tpholesize = holeFt.ThruTapDrillDiameter
                        ElseIf holeFt.TapDrillDiameter <> 0 Then
                            tpholesize = holeFt.TapDrillDiameter
                        End If
                        holeSize = ((tpholesize * 1000) * Math.PI)
                        If (tpholesize * 1000) <= minholesize Then
                            'do nothing
                        ElseIf (tpholesize * 1000) >= pulsesize Then
                            tapCount += fccount
                            HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                            HoleCutTimeOxy += (fccount * HoleStuff)

                            HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                            HoleCutTimeNit += (fccount * HoleStuff)
                        Else
                            tapCount += fccount
                            VarPlace1 = HoleFormula.IndexOf("*")
                            VarPlace2 = HoleFormula.IndexOf("^")
                            FormPart1 = CDbl(Strings.Left(Replace(PulseFormula, "(", ""), VarPlace1 - 1))
                            FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                            FormPart3 = Replace(Replace(PulseFormula, Strings.Left(PulseFormula, VarPlace2 + 1), ""), ")", "")
                            HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                            HoleCutTimeOxy += (fccount * HoleStuff)
                        End If
                    End If
                    '--Countersink Hole--
                    If CounterHole.Contains(holetype) And Not ptnHoleList.Contains(ft.Name) And ft.Name.ToString.Contains("CSK") Then
                        fccount = holeFt.GetSketchPointCount
                        holeSize = ((holeFt.ThruHoleDiameter * 1000) * Math.PI)
                        If (holeFt.ThruHoleDiameter * 1000) <= minholesize Then
                            'do nothing
                        ElseIf (holeFt.ThruHoleDiameter * 1000) >= pulsesize Then
                            countersnkcount += fccount

                            HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                            HoleCutTimeOxy += (fccount * HoleStuff)

                            HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                            HoleCutTimeNit += (fccount * HoleStuff)
                        Else
                            countersnkcount += fccount
                            VarPlace1 = HoleFormula.IndexOf("*")
                            VarPlace2 = HoleFormula.IndexOf("^")
                            FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                            FormPart2 = (((holeFt.ThruHoleDiameter * 1000) * Math.PI) + smThickness)
                            FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                            HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                            HoleCutTimeOxy += (fccount * HoleStuff)
                            CounterSinkRemov += ((holeFt.CounterSinkDiameter * 1000) * Math.PI) * fccount
                        End If
                    End If
                    '--Flowdrill--
                    If ft.Name.ToString.Contains("Flow") Then
                        fccount = holeFt.GetSketchPointCount
                        If (holeFt.ThruTapDrillDiameter * 1000) <= minholesize Then
                            'do nothing
                        ElseIf (holeFt.ThruTapDrillDiameter * 1000) >= pulsesize Then
                            flowDrillCount += fccount
                            HoleStuff = (holeSize + smThickness) * CuttingTimeOxygen    'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                            HoleCutTimeOxy += (fccount * HoleStuff)

                            HoleStuff = (holeSize + smThickness) * CuttingTimeNitrogen  'Add sheetmetal thickness to hole size since a cut is made from the inside of the circle that equals the sheet thickness
                            HoleCutTimeNit += (fccount * HoleStuff)
                        Else
                            flowDrillCount += fccount
                            VarPlace1 = HoleFormula.IndexOf("*")
                            VarPlace2 = HoleFormula.IndexOf("^")
                            FormPart1 = CDbl(Strings.Left(Replace(HoleFormula, "(", ""), VarPlace1 - 1))
                            FormPart2 = (((holeFt.ThruTapDrillDiameter * 1000) * Math.PI) + smThickness)
                            FormPart3 = Replace(Replace(HoleFormula, Strings.Left(HoleFormula, VarPlace2 + 1), ""), ")", "")
                            HoleStuff = FormPart1 * (FormPart2 ^ FormPart3)
                            HoleCutTimeOxy += (fccount * HoleStuff)
                        End If
                    End If
                End If
            End If
NextFeature: Next

        'All holes added up
        totalholes = holecount + tapCount + countersnkcount + flowDrillCount

        'Actual length of non-hole features
        ActualCutLength = cuttingLengthOuter + (cuttingLengthInner - CounterSinkRemov)

        'Length of normal cuts X cut time based on thickness + number of pierces X pierce time based on sheet thickness
        PerimCutTime = ((Math.Round(ActualCutLength, 2) * cuttimecalc) + (pierceCount * piercetime))

        'Normal cut times plus hole cut times
        TotalCutTime = PerimCutTime + HoleCutTimeOxy

        'Travel Time added equals piercecount X average move time between features (in seconds)
        TotalCutTime += (pierceCount * 0.43)
        t2 = Date.Now
        Debug.Print("Hole Information collected" & " (" & DateDiff(DateInterval.Second, t1, t2) & ")")
        t1 = Date.Now
        Debug.Print("Create Routings")
        '-------ROUTING-------
        Dim dtNAV As New DataTable
        sqlstr = "SELECT [Routing No_], [Operation No_], No_, Description, [Run Time]
                    FROM [OMT-Veyhl USA Corporation$Routing Line]
                    WHERE [Routing No_] LIKE '" & PartNum & "'"
        dtNAV = dbNAV.FillTableFromSql(sqlstr)

        Dim dt As New DataTable
        dt.Columns.Add("OperationNo")
        dt.Columns.Add("WorkCenter")
        dt.Columns.Add("WorkCenterName")
        dt.Columns.Add("RunTime(Minutes)")
        dt.Columns.Add("MoveTime")
        dt.Columns.Add("RoutingLinkCode")
        dt.Columns.Add("Fixture")
        dt.Columns.Add("WorkInstructions")
        dt.Columns.Add("Comments")

        'Everything new going to 102
        'Round to the nearest second

        'Benders
        'anything less than 500 routed to 301
        'anything .25in and greater to 305








        Dim OPnum As Integer = 20
        'Programming
        dt.Rows.Add(10, "091 PRG-FB", "Programming - Flat Bed Laser", 0.41667, 0, "", "", "", "") 'Used to be 1 minute, now 25 seconds
        'Lasers
        'If (Description.ToUpper.Contains("STUD PLATE") Or Description.ToUpper.Contains("MOTOR BOX") Or Description.ToUpper.Contains("INSERT") Or Description.ToUpper.Contains("INSRT")) And smThickness < 4.76 Then
        '    dt.Rows.Add(OPnum, "114 FB LSR", "Flat Bed Laser (Mazak #1)", TotalCutTime / 60, 0, "CONS", "", "", "")
        '    OPnum += 10
        'ElseIf (Description.ToUpper.Contains("FOOT GUSSET") Or Description.ToUpper.Contains("FOOT TOP PLATE") Or Description.ToUpper.Contains("TOP SUPPORT")) And smThickness < 4.76 Then
        '    dt.Rows.Add(OPnum, "103 FB LSR", "Foot Cell Flat Bed Lsr-Mazak#2", TotalCutTime / 60, 0, "CONS", "", "", "")
        '    OPnum += 10
        'If smThickness < 4.2 Then
        '    dt.Rows.Add(OPnum, "101 FB LSR", "Oxyen Flat Bed Laser-2 Trumpf", TotalCutTime / 60, 0, "CONS", "", "", "")
        '    OPnum += 10
        'ElseIf smThickness >= 4.2 Then
        dt.Rows.Add(OPnum, "102 LSR NI", "Nitrogen Flat Bed Laser-Trumpf", TotalCutTime / 60, 0, "CONS", "", "", "")
        OPnum += 10
        'End If
        'Drill Presses
        If flowDrillCount <> 0 Then
            dt.Rows.Add(OPnum, "HURCO", "Flo Drill", (flowDrillCount * 15) / 60, 0, "", "", "", "")
            OPnum += 10
        End If
        If countersnkcount <> 0 Then
            dt.Rows.Add(OPnum, "202 C.SINK", "Countersink", (countersnkcount * 20) / 60, 0, "", "", "", "")
            OPnum += 10
        End If
        If tapCount <> 0 Then
            If Description.Contains("MOTOR BOX") = True And (Description.Contains("KIMBALL") = False Or Description.Contains("CRANK") = False) Then
                dt.Rows.Add(OPnum, "KB 209 ROBO/BEND", "KB Robo Drill/BEND OP", 22 / 60, 0, "", "", "", "")
                OPnum += 10
            ElseIf smThickness >= 6.35 Then
                dt.Rows.Add(OPnum, "207 TAP HYDRAULIC", "Tap - Hydraulic Flex Arm", (tapCount * 20) / 60, 0, "", "", "", "")
                OPnum += 10
            ElseIf smThickness >= 4.18 Then
                dt.Rows.Add(OPnum, "203 TAP", "Tap", (tapCount * 10) / 60, 0, "", "", "", "")
                OPnum += 10
            Else
                dt.Rows.Add(OPnum, "203 TAP", "Tap", (tapCount * 10) / 60, 0, "", "", "", "")
                OPnum += 10
            End If
        End If

        'Bending
        If bendCount > 0 Then
            If maxbendlength > 1000 Then
                If partWeight < 50 Then
                    If smThickness > 2.66 Then
                        If Revision <> "X1" Then
                            dt.Rows.Add(OPnum, "304 BRAKE", "Large Brake Press", (bendCount * 10) / 60, 0, "", "", "", "")
                            OPnum += 10
                        Else
                            dt.Rows.Add(OPnum, "305 BRAKE", "Large Misc Brake Press", (bendCount * 10) / 60, 0, "", "", "", "")
                            OPnum += 10
                        End If
                    Else
                        dt.Rows.Add(OPnum, "305 BRAKE", "Large Misc Brake Press", (bendCount * 10) / 60, 0, "", "", "", "")
                        OPnum += 10
                    End If
                Else
                    If smThickness > 2.66 Then
                        If Revision <> "X1" Then
                            dt.Rows.Add(OPnum, "304 BRAKE", "Large Brake Press", (bendCount * 20) / 60, 0, "", "", "", "")
                            OPnum += 10
                        Else
                            dt.Rows.Add(OPnum, "305 BRAKE", "Large Misc Brake Press", (bendCount * 20) / 60, 0, "", "", "", "")
                            OPnum += 10
                        End If
                    Else
                        dt.Rows.Add(OPnum, "305 BRAKE", "Large Misc Brake Press", (bendCount * 20) / 60, 0, "", "", "", "")
                        OPnum += 10
                    End If
                End If
            Else
                If partWeight < 50 Then
                    If (maxbendlength > 500 And smThickness >= 2.66) Or (maxbendlength > 300 And smThickness >= 6.35) Then
                        If Revision <> "X1" Then
                            dt.Rows.Add(OPnum, "304 BRAKE", "Large Brake Press", (bendCount * 10) / 60, 0, "", "", "", "")
                            OPnum += 10
                        Else
                            dt.Rows.Add(OPnum, "305 BRAKE", "Large Misc Brake Press", (bendCount * 10) / 60, 0, "", "", "", "")
                            OPnum += 10
                        End If
                    ElseIf Revision = "X1" Then
                        dt.Rows.Add(OPnum, "305 BRAKE", "Large Misc Brake Press", (bendCount * 10) / 60, 0, "", "", "", "")
                        OPnum += 10
                    ElseIf smThickness >= 2.66 Then
                        dt.Rows.Add(OPnum, "301 BRAKE", "Brake Press", (bendCount * 10) / 60, 0, "", "", "", "")
                        OPnum += 10
                    Else
                        dt.Rows.Add(OPnum, "301 BRAKE", "Brake Press", (bendCount * 10) / 60, 0, "", "", "", "")
                        OPnum += 10
                    End If
                Else
                    If (maxbendlength > 500 And smThickness >= 2.66) Or (maxbendlength > 300 And smThickness >= 6.35) Then
                        If Revision <> "X1" Then
                            dt.Rows.Add(OPnum, "304 BRAKE", "Large Brake Press", (bendCount * 20) / 60, 0, "", "", "", "")
                            OPnum += 10
                        Else
                            dt.Rows.Add(OPnum, "305 BRAKE", "Large Misc Brake Press", (bendCount * 20) / 60, 0, "", "", "", "")
                            OPnum += 10
                        End If
                    Else
                        dt.Rows.Add(OPnum, "305 BRAKE", "Large Misc Brake Press", (bendCount * 20) / 60, 0, "", "", "", "")
                        OPnum += 10
                    End If
                End If
            End If
        End If
        t2 = Date.Now
        Debug.Print("Routing Created" & " (" & DateDiff(DateInterval.Second, t1, t2) & ")")

        Dim rt As Integer = dt.Rows.Count
        dt.Rows(rt - 1).Item("MoveTime") = 630

        If Not ECNNumber.Contains("ECN") Then
            If ECNNumber.Length = 4 Then
                ECNNumber = "ECN-0" & ECNNumber
            Else
                ECNNumber = "ECN-" & ECNNumber
            End If
        End If

        If dt.Rows.Count <> dtNAV.Rows.Count Then
            'dt = dtNAV
            'For Each r As DataRow In dt.Rows
            '    If r.Item("WorkCenter").ToString.Contains("10") Or r.Item("WorkCenter").ToString.Contains("114") Then
            '        r.Item("RunTime(Minutes)") = TotalCutTime / 60
            '    End If
            'Next
        End If

#Region "Testing"
        'Form1.txtMatDesc.Text = MaterialDesc
        'Form1.txtMatPN.Text = MaterialPN
        'Form1.txtEcnNum.Text = ECNNumber
        'Form1.TxtOutsideCut.Text = cuttingLengthOuter
        'Form1.TxtInsideCut.Text = cuttingLengthInner
        'Form1.TxtTotalHoleCut.Text = totalHoleLength
        'Form1.TxtActualCutLength.Text = ActualCutLength
        'Form1.TxtThickness.Text = smThickness
        'Form1.TxtBends.Text = bendCount
        'Form1.txtWeight.Text = partWeight
        'Form1.TxtTotalHoles.Text = totalholes
        'Form1.TxtNormalHoles.Text = holecount
        'Form1.TxtTappedHoles.Text = tapCount
        'Form1.TxtCounterSink.Text = countersnkcount
        'Form1.TxtFlowDrill.Text = flowDrillCount
        'Form1.DgvRouting.DataSource = dt

        'For Each r As DataRow In dt.Rows
        '    Form1.dtRoutingResult.Rows.Add(PartNum, r.Item("OperationNo"), r.Item("WorkCenter"), r.Item("WorkCenterName"), r.Item("RunTime(Minutes)"), r.Item("MoveTime"))
        'Next
#End Region

        sqlstr = "DELETE FROM [dbo].[ECN ME Fabrication]
                                                   WHERE [ECN No_] LIKE '" & ECNNumber & "' AND [Part No_] LIKE '" & PartNum & "'"
        SqlConn.ExecuteNonQueryForSql(sqlstr)

        sqlstr = "DELETE FROM [dbo].[ECN ME Fabrication Calculated]
                                                   WHERE [ECN No_] LIKE '" & ECNNumber & "' AND [Part No_] LIKE '" & PartNum & "'"
        SqlConn.ExecuteNonQueryForSql(sqlstr)

        For Each r As DataRow In dt.Rows
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
    End Sub

    Private Sub CutTimeTable()
        Dim dt As DataTable
        sqlstr = "SELECT *
                      FROM dbo.FlatBedCutTime
                      WHERE SizeMM = '" & smThickness & "'"
            dt = SqlConn.FillTableFromSql(sqlstr)
        If dt.Rows.Count > 0 Then
            cuttimecalc = dt.Rows(0).Item(3)
            pulsesize = dt.Rows(0).Item("MaxPulseSize")
            minholesize = dt.Rows(0).Item("MinHoleSize")
            piercetime = dt.Rows(0).Item("PierceTime")
            HoleFormula = dt.Rows(0).Item("NormalHoleFormula")
            PulseFormula = dt.Rows(0).Item("PulseHoleFormula")
            Try
                CuttingTimeNitrogen = dt.Rows(0).Item("NitrogenNormSecPerMM")
            Catch ex As Exception
            End Try
            Try
                CuttingTimeOxygen = dt.Rows(0).Item("OxygenNormSecPerMM")
            Catch ex As Exception
            End Try
        End If
    End Sub
End Class
