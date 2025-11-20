Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Public Class CreateSTP
    Public Shared Sub STP(ByVal ItemCategory As String, ByVal Filename As String, ByVal Revision As String, ByVal swApp As SldWorks, ByVal swModel As ModelDoc2)
        'Get Configuration
        Dim cfgNames As Object
        Dim cfgName As String
        Dim i As Integer
        Dim mcc As Integer
        cfgNames = swModel.GetConfigurationNames()
        i = 0
        cfgName = cfgNames(i)
        Try
            mcc = swModel.GetConfigurationCount
            Do Until InStr(cfgName.ToLower, "default") <> 0 Or InStr(cfgName.ToLower, "original") <> 0 Or i > swModel.GetConfigurationCount
                Try
                    cfgName = cfgNames(i)
                Catch ex As Exception
                End Try
                i += 1
            Loop

            swModel.GetConfigurationByName(cfgName(i))
            swModel.ShowConfiguration2(cfgName)
        Catch ex As Exception
        End Try

        If ItemCategory.Contains("96") Then
            'Dim fullSTPpath As String = "\\omt-file05\OV Prints\00_STEPS\_NAV SW STEPS\Sheet Metal Items\" & Filename & "-" & Revision & " " & Date.Now.ToString("yyyy-MM-dd HH_mm_ss") & ".STEP"
            Dim fullSTPpath As String = "\\omt-file05\OV Prints\00_STEPS\_NAV SW STEPS\Sheet Metal Items\" & Filename & "-" & Revision & ".STEP"
            swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swStepAP, 203)
            swModel.Extension.SaveAs(fullSTPpath, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, 0, 0)
            Debug.Print("File Created: " & My.Computer.FileSystem.FileExists(fullSTPpath))
        ElseIf {"86", "31"}.contains(ItemCategory) Then
            Try
                'Dim fullSTPpath As String = "\\omt-file05\OV Prints\00_Pending STEPS\" & Filename & "-" & Revision & " " & Date.Now.ToString("yyyy-MM-dd HH_mm_ss") & ".STEP"
                Dim fullSTPpath As String = "\\omt-file05\OV Prints\00_Pending STEPS\" & Filename & "-" & Revision & ".STEP"
                swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swStepAP, 203)
                swModel.Extension.SaveAs(fullSTPpath, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, 0, 0)
                Debug.Print("File Created: " & My.Computer.FileSystem.FileExists(fullSTPpath))
                If Dir(fullSTPpath) = "" Then
                    PDM_Class.errmsg = "Failed to create STEP"
                End If
            Catch ex As Exception
                PDM_Class.errmsg = "Failed to make STEP"
                Exit Sub
            End Try
        End If
    End Sub
End Class
