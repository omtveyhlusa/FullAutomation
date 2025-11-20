Imports EPDM.Interop.epdm
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Public Class CreateDXF

    Public Shared Sub DXF(ByVal ItemCategory As String, ByVal Filename As String, ByVal Revision As String, ByVal swApp As SldWorks, ByVal swModel As ModelDoc2, ByVal swPart As PartDoc, ByVal pdmFile As IEdmFile12, ByVal foldID As Integer)

        'Get Configuration
        Dim cfgNames As Object
        Dim cfgName As String
        Dim i As Integer
        Dim mcc As Integer
        cfgNames = swModel.GetConfigurationNames()
        i = 0
        cfgName = cfgNames(i)
        mcc = swModel.GetConfigurationCount
        Do Until InStr(cfgName.ToLower, "sm-flat-pattern") <> 0 Or InStr(cfgName.ToLower, "flat pattern") <> 0 Or i > swModel.GetConfigurationCount
            Try
                cfgName = cfgNames(i)
            Catch ex As Exception
            End Try
            i += 1
        Loop

        'exit if not found
        If i > swModel.GetConfigurationCount Then
            Debug.Print("Flat Pattern configuration not found!")
            'retErr = "Flat-Pattern configuration not found."
            'swApp.ExitApp()
            'swApp.CloseAllDocuments(True)
            Exit Sub
        End If

        swModel.GetConfigurationByName(cfgName(i))
        swModel.ShowConfiguration2(cfgName)

        Dim varAlignment As Object
        Dim dataAlignment(11) As Double
        Dim varViews As Object
        Dim dataViews(1) As String
        Dim options As Integer

        dataAlignment(0) = 0.0#
        dataAlignment(1) = 0.0#
        dataAlignment(2) = 0.0#
        dataAlignment(3) = 1.0#
        dataAlignment(4) = 0.0#
        dataAlignment(5) = 0.0#
        dataAlignment(6) = 0.0#
        dataAlignment(7) = 1.0#
        dataAlignment(8) = 0.0#
        dataAlignment(9) = 0.0#
        dataAlignment(10) = 0.0#
        dataAlignment(11) = 1.0#
        varAlignment = dataAlignment
        dataViews(0) = "*Current"
        dataViews(1) = "*Front"
        varViews = dataViews
        options = 33
        'options = 1
        'Try
        '    My.Computer.FileSystem.DeleteFile("\\omt-file01\Shared\Engineering\OMT-Veyhl USA Print\00_Pending DXFS\" & Filename & "-" & Revision & ".dxf")
        'Catch ex As Exception
        'End Try
        'swPart.ExportToDWG2("\\omt-file05\OV Prints\00_Pending DXFS\" & Filename & "-" & Revision & " " & Date.Now.ToString("yyyy-MM-dd HH_mm_ss") & ".dxf", pdmFile.GetLocalPath(foldID), swExportToDWG_e.swExportToDWG_ExportSheetMetal, True, varAlignment, False, False, options, Nothing)
        'If Dir("\\omt-file05\OV Prints\00_Pending DXFS\" & Filename & "-" & Revision & " " & Date.Now.ToString("yyyy-MM-dd HH_mm_ss") & ".dxf") = "" Then
        '    PDM_Class.errmsg = "Failed to create DXF"
        'End If
        swPart.ExportToDWG2("\\omt-file05\OV Prints\00_Pending DXFS\" & Filename & "-" & Revision & ".dxf", pdmFile.GetLocalPath(foldID), swExportToDWG_e.swExportToDWG_ExportSheetMetal, True, varAlignment, False, False, options, Nothing)
        If Dir("\\omt-file05\OV Prints\00_Pending DXFS\" & Filename & "-" & Revision & ".dxf") = "" Then
            PDM_Class.errmsg = "Failed to create DXF"
        End If
    End Sub
End Class
