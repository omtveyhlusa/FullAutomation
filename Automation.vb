Imports EPDM.Interop.epdm
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic.FileIO
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Data.Common
Imports System.Deployment.Application
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Security
Imports DecryptionDLL.DecrypterDLL
Public Class Automation
    'Solidworks PDM
    Public vault1 As EdmVault5

    'Solidworks


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

    Public Retval As Integer
    Public err As String
    Public warn As String

    Dim dbECNActive As Boolean
    Dim dbPDMActive As Boolean
    Dim dbNAVActive As Boolean

    'sql
    Public _dbECN As SQLdb.SqlDatabase
    Public _dbPDM As SQLdb.SqlDatabase
    Public _dbNAV As SQLdb.SqlDatabase

    Dim sqlstr As String

    Sub New(ByVal fileID As Integer, ByVal foldID As Integer, ByRef retErr As String, ByRef ProgType As String)
        GetSQLAccess()

        'Get latest version of file
        Dim vault1 As IEdmVault18 = New EdmVault5
        If Not vault1.IsLoggedIn Then
            vault1.LoginAuto("OMT-Sandbox", 0)
        End If
        'Dim vault2 As IEdmVault18 = New EdmVault5
        'If Not vault2.IsLoggedIn Then
        '    vault2.LoginAuto("OMT-Sandbox", 0)
        'End If
        Dim pdmFile As IEdmFile12 = vault1.GetObject(EdmObjectType.EdmObject_File, fileID)
        Dim pos As IEdmPos5 = pdmFile.GetFirstFolderPosition
        Dim errors As Integer = Nothing
        Dim warnings As Integer = Nothing
        pdmFile.GetFileCopy(0)
        'restart
        Dim swApp As SldWorks
        Dim swModel As ModelDoc2
        Dim swPart As PartDoc
        Dim swAssembly As AssemblyDoc
        Dim swDrawing As DrawingDoc
        swApp = New SldWorks

        'Open file in Solidworks
        Debug.Print("Filename: " & pdmFile.Name)
        If pdmFile.Name.ToUpper.Contains("SLDPRT") Then
            swPart = swApp.OpenDoc6(pdmFile.GetLocalPath(foldID), swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", errors, warnings)
            swModel = swPart
            swApp.Visible = True
        ElseIf pdmFile.Name.ToUpper.Contains("SLDASM") Then
            'Commenting out until assemblies get figured out.
            'GetAllFiles(fileID)
            'swAssembly = swApp.OpenDoc6(pdmFile.GetLocalPath(foldID), swDocumentTypes_e.swDocASSEMBLY, swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", errors, warnings)
            'swModel = swAssembly
            'swApp.Visible = True
        ElseIf pdmFile.Name.ToUpper.Contains("SLDDRW") Then
            GetAllFiles(fileID)
            swDrawing = swApp.OpenDoc6(pdmFile.GetLocalPath(foldID), swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", errors, warnings)
            swModel = swDrawing
            swApp.Visible = True
        End If

        'give the sw file time to open
        Threading.Thread.Sleep(5000)

        Try

            'Collect part property information
            Dim swCustPropMgr As CustomPropertyManager
            PartNum = Path.GetFileNameWithoutExtension(pdmFile.GetLocalPath(foldID))
            If Not pdmFile.Name.ToUpper.Contains("DRW") Then
                swCustPropMgr = swModel.Extension.CustomPropertyManager("")
            Else
                swCustPropMgr = swDrawing.Extension.CustomPropertyManager("")
            End If
            ECNNumber = swCustPropMgr.Get("ECN Number")
            Description = swCustPropMgr.Get("Description")
            Revision = swCustPropMgr.Get("Revision")
            ItemCategory = swCustPropMgr.Get("ITEM GROUP")
            MaterialDesc = swCustPropMgr.Get("matDesc")
            MaterialPN = swCustPropMgr.Get("MATERIAL PN")

            Debug.Print("ECN Number: " & ECNNumber)
            Debug.Print("Part Number: " & PartNum)
            Debug.Print("Description: " & Description)
            Debug.Print("Revision: " & Revision)
            Debug.Print("Item Category: " & ItemCategory)
            Debug.Print("Material Description: " & MaterialDesc)
            Debug.Print("Material Part Number: " & MaterialPN)

            If ECNNumber.Contains("fromparent+") Then
                ECNNumber = Replace(ECNNumber, "fromparent+", "")
            End If

            If pdmFile.Name.ToUpper.Contains("SLDDRW") Then
                Debug.Print("Create PDF initiated")
                CreatePDF.PDF(Path.GetFileNameWithoutExtension(pdmFile.GetLocalPath(foldID)), Revision, swModel, swDrawing, swApp)
                Debug.Print("Create PDF finished")
                If ItemCategory.Contains("99") Then
                    Debug.Print("Assembly Automation initiated")
                    'Dim ASY As New Routing_Assembly(PartNum, Description, "99", ECNNumber)
                    Debug.Print("Assembly Automation finished")
                End If
            ElseIf pdmFile.Name.ToUpper.Contains("SLDASM") Then
                If ItemCategory.Contains("95") Then
                    Debug.Print("Weldment Routing Information Initiated")
                    'Dim WLD As New Routing_Weldment(PartNum, swModel, pdmFile, swApp)
                    Debug.Print("Weldment Routing Information Finished")
                    Debug.Print("Weldment Paint Information initiated")
                    'Dim PNT As New Routing_Paint(PartNum, Revision, "95", swModel, swPart, swAssembly, swApp)
                    Debug.Print("Weldment Paint Information finished")
                End If
            ElseIf pdmFile.Name.ToUpper.Contains("SLDPRT") Then
                Debug.Print("Create STP initiated")
                CreateSTP.STP(ItemCategory, Path.GetFileNameWithoutExtension(pdmFile.GetLocalPath(foldID)), Revision, swApp, swModel)
                Debug.Print("Create STP finished")
                If ItemCategory.Contains("96") Then
                    Debug.Print("Create DXF initiated")
                    CreateDXF.DXF(ItemCategory, Path.GetFileNameWithoutExtension(pdmFile.GetLocalPath(foldID)), Revision, swApp, swModel, swPart, pdmFile, foldID)
                    Debug.Print("Create DXF finished")
                    Debug.Print("Flat Sheet Information initiated")
                    Dim SM As New Routing_SheetMetal(PartNum, swModel, fileID, foldID)
                    Debug.Print("Flat Sheet Information finished")
                    Debug.Print("Flat Sheet Paint Information initiated")
                    'Dim PNT As New Routing_Paint(PartNum, Revision, "96", swModel, swPart, swAssembly, swApp)
                    Debug.Print("Flat Sheet Paint Information finished")
                ElseIf ItemCategory.Contains("86") Then
                    Dim TB As New Routing_TubeMetal(PartNum, swModel, swApp, fileID, foldID)
                    Debug.Print("Tube Paint Information initiated")
                    'Dim PNT As New Routing_Paint(PartNum, Revision, "86", swModel, swPart, swAssembly, swApp)
                    Debug.Print("Tube Paint Information finished")
                End If
            Else
                'do nothing
            End If
            swApp.CloseAllDocuments(True)
            'swApp.QuitDoc(swModel.GetTitle)
            swPart = Nothing
            swModel = Nothing
            swDrawing = Nothing

            swApp.ExitApp()
            'swApp.CloseAllDocuments(True)
            Try
                If My.Computer.FileSystem.FileExists("C:\temp\TempFile.SLDPRT") Then
                    My.Computer.FileSystem.DeleteFile("C:\temp\TempFile.SLDPRT")
                End If
            Catch ex As Exception
            End Try
        Catch ex As Exception
            '    If ex.Message.Contains("RPC server is unavailable") Then
            '        For Each MyProcess In Process.GetProcessesByName("SLDWORKS")
            '            MyProcess.Kill()
            '        Next

            '        For Each MyProcess In Process.GetProcessesByName("sldworks_fs")
            '            MyProcess.Kill()
            '        Next

            '        For Each MyProcess In Process.GetProcessesByName("sldExitApp")
            '            MyProcess.Kill()
            '        Next

            '        GoTo Restart
            '    End If
        End Try
    End Sub
    Private Sub GetAllFiles(ByVal fileID As Integer)
        Dim vault1 As IEdmVault18 = New EdmVault5
        If Not vault1.IsLoggedIn Then
            vault1.LoginAuto("OMT-Sandbox", 0)
        End If

        Dim dt As New DataTable
        Dim dtresults As New DataTable
        dtresults.Columns.Add("FileID")
        dtresults.Columns.Add("FoldID")
        dtresults.Columns.Add("Filename")
        dtresults.Columns.Add("Type")
        dtresults.Columns.Add("Checked")

        'Initial pull
        sqlstr = "SELECT Filename, D.DocumentID, refs.XRefProjectID, D.ExtensionID
                    FROM xrefs AS Refs INNER JOIN
                    Documents AS D ON Refs.XRefDocument = D.DocumentID
                    WHERE refs.DocumentID = " & fileID & " AND refs.RevNr = (SELECT MAX(CAST(RevNr AS INT)) FROM XRefs WHERE xrefs.DocumentID = refs.DocumentID)"
        dt = _dbPDM.FillTableFromSql(sqlstr)
        For Each r As DataRow In dt.Rows
            If r.Item("ExtensionID") = 5 Then
                dtresults.Rows.Add(r.Item("DocumentID"), r.Item("XRefProjectID"), r.Item("Filename"), r.Item("ExtensionID"), "")
            ElseIf r.Item("ExtensionID") = 4 Then
                dtresults.Rows.Add(r.Item("DocumentID"), r.Item("XRefProjectID"), r.Item("Filename"), r.Item("ExtensionID"), "0")
            End If
        Next

        Dim i As Integer = 0
        Do Until i = dtresults.Rows.Count
            If dtresults.Rows(i).Item("Type") = 4 And dtresults.Rows(i).Item("Checked") = "0" Then
                dtresults.Rows(i).Item("Checked") = "1"
                sqlstr = "SELECT Filename, D.DocumentID, refs.XRefProjectID, D.ExtensionID
                    FROM xrefs AS Refs INNER JOIN
                    Documents AS D ON Refs.XRefDocument = D.DocumentID
                    WHERE refs.DocumentID = " & dtresults.Rows(i).Item("FileID") & " AND refs.RevNr = (SELECT MAX(CAST(RevNr AS INT)) FROM XRefs WHERE xrefs.DocumentID = refs.DocumentID)"
                dt = _dbPDM.FillTableFromSql(sqlstr)
                For Each r As DataRow In dt.Rows
                    If r.Item("ExtensionID") = 5 Then
                        dtresults.Rows.Add(r.Item("DocumentID"), r.Item("XRefProjectID"), r.Item("Filename"), r.Item("ExtensionID"), "")
                    ElseIf r.Item("ExtensionID") = 4 Then
                        dtresults.Rows.Add(r.Item("DocumentID"), r.Item("XRefProjectID"), r.Item("Filename"), r.Item("ExtensionID"), "0")
                    End If
                Next
                i = 0
            End If
            i += 1
        Loop

        For Each r As DataRow In dtresults.Rows
            If r.Item("Type") = 5 Then
                Dim pdmFile As IEdmFile12 = vault1.GetObject(EdmObjectType.EdmObject_File, r.Item("FileID"))
                Dim pos As IEdmPos5 = pdmFile.GetFirstFolderPosition
                pdmFile.GetFileCopy(0)
            End If
        Next
        For Each r As DataRow In dtresults.Rows
            If r.Item("Type") = 4 Then
                Dim pdmFile As IEdmFile12 = vault1.GetObject(EdmObjectType.EdmObject_File, r.Item("FileID"))
                Dim pos As IEdmPos5 = pdmFile.GetFirstFolderPosition
                pdmFile.GetFileCopy(0)
            End If
        Next
    End Sub
    Private Sub GetSQLAccess()
        Dim Info() As String
        'ECN
        If dbECNActive = False Then
            If My.Computer.FileSystem.FileExists("\\omt-file01\Support\Public\CustomApps\Database Common Files\SQLECN.enc") Then
                Info = Decrypt(IO.File.ReadAllBytes("\\omt-file01\Support\Public\CustomApps\Database Common Files\SQLECN.enc"))
                _dbECN = New SQLdb.SqlDatabase(Info(0), Info(1), Info(2), Info(3))
                dbECNActive = True
                Debug.Print("ECN SQL found?: " & dbECNActive.ToString)
            End If
        End If

        'PDM
        If dbPDMActive = False Then
            If My.Computer.FileSystem.FileExists("\\omt-file01\Support\Public\CustomApps\Database Common Files\SQLPDM.enc") Then
                Info = Decrypt(IO.File.ReadAllBytes("\\omt-file01\Support\Public\CustomApps\Database Common Files\SQLPDM.enc"))

                _dbPDM = New SQLdb.SqlDatabase(Info(0), Info(1), Info(2), Info(3))
                dbPDMActive = True
                Debug.Print("PDM SQL found?: " & dbPDMActive.ToString)
            End If
        End If

        'NAV
        If dbNAVActive = False Then
            If My.Computer.FileSystem.FileExists("\\omt-file01\Support\Public\CustomApps\Database Common Files\SQLNAV.enc") Then
                Info = Decrypt(IO.File.ReadAllBytes("\\omt-file01\Support\Public\CustomApps\Database Common Files\SQLNAV.enc"))

                _dbNAV = New SQLdb.SqlDatabase(Info(0), Info(1), Info(2), Info(3))
                dbNAVActive = True
                Debug.Print("NAV SQL found?: " & dbNAVActive.ToString)
            End If
        End If

        If dbECNActive = False Or dbPDMActive = False Or dbNAVActive = False Then
            Application.Exit()
        End If
    End Sub
End Class

