Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Public Class CreatePDF
    Public Shared Sub PDF(ByVal Filename As String, ByVal Revision As String, ByVal Model As ModelDoc2, ByVal Drawing As DrawingDoc, ByVal App As SldWorks)
        Dim swModelDocExt As ModelDocExtension

        'Fixes tables to have a line thickness of 0.18
        'Dim swView As View
        'Dim dwgTable As TableAnnotation
        'Dim swAnn As Annotation
        'swView = Drawing.GetFirstView
        'Do While Not swView Is Nothing
        '    dwgTable = swView.GetFirstTableAnnotation
        '    Do While Not dwgTable Is Nothing
        '        swAnn = dwgTable.GetAnnotation

        '        dwgTable.GridLineWeight = 1
        '        dwgTable.BorderLineWeight = 1

        '        dwgTable = dwgTable.GetNext
        '    Loop
        '    swView = swView.GetNextView
        'Loop

        swModelDocExt = Model.Extension
        Dim path As String = "\\omt-file05\OV Prints\00_Pending PDFS\"
        Dim fileinfo As String = Filename & "-" & Revision
        Dim ext As String = ".PDF"
        Dim pdffileloc As String = path & fileinfo & ext
        'Try
        '    My.Computer.FileSystem.DeleteFile(pdffileloc)
        'Catch ex As Exception
        'End Try
        Try
            swModelDocExt.SaveAs2(pdffileloc, 0, 1, Nothing, "", False, 0, 0)
            'Model.Extension.SaveAs(pdffileloc, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, "", "")
            If Dir(pdffileloc) = "" Then
                Debug.Print("Create PDF failed")
                PDM_Class.errmsg = "Failed to create PDF"
            End If
        Catch ex As Exception
            Debug.Print("Create PDF failed")
            PDM_Class.errmsg = "Failed to create PDF"
        End Try
    End Sub
End Class
