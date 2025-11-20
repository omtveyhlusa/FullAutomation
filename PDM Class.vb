Imports EPDM.Interop.epdm
Imports EPDM.Interop.EPDMResultCode
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports System.Collections.Specialized
Imports System.IO
Imports System.Runtime.InteropServices

Public Class PDM_Class
    Implements IEdmAddIn5
    <ComVisible(True)>
    Public Sub GetAddInInfo(ByRef poInfo As EdmAddInInfo, poVault As IEdmVault5, poCmdMgr As IEdmCmdMgr5) Implements IEdmAddIn5.GetAddInInfo
        Try
            poInfo.mbsAddInName = "ECN Automation"
            poInfo.mbsCompany = "OMT-Veyhl USA"
            poInfo.mbsDescription = "ECN Automation Tool"
            'Year/Month/Day/version
            poInfo.mlAddInVersion = 2024022701

            'Required PDM version
            poInfo.mlRequiredVersionMajor = 31
            poInfo.mlRequiredVersionMinor = 5

            'poCmdMgr.AddCmd(1, "ECN Automation Test")
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskRun)
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetup)
        Catch ex As Runtime.InteropServices.COMException
            MsgBox("HRESULT = 0x" + ex.ErrorCode.ToString("X") + vbCrLf + ex.Message)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData() As EdmCmdData) Implements IEdmAddIn5.OnCmd
        Select Case poCmd.meCmdType
            Case EdmCmdType.EdmCmd_Menu
                Dim fileID As Integer = ppoData(0).mlObjectID1
                Dim foldID As Integer = ppoData(0).mlObjectID3
                Dim retErr As String = Nothing

                Dim CreateFiles As New Automation(fileID, foldID, retErr, "AUTO")
            Case EdmCmdType.EdmCmd_TaskSetup
                OnTaskSetup(poCmd, ppoData)
            Case EdmCmdType.EdmCmd_TaskSetupButton
                OnTaskSetupButton(poCmd, ppoData)
            Case EdmCmdType.EdmCmd_TaskRun
                OnTaskRun(poCmd, ppoData)
        End Select
    End Sub

    Dim currentSetupPage As setupPage

    Public Sub OnTaskSetup(ByRef poCmd As EdmCmd, ByRef ppoData() As EdmCmdData)
        Dim props As IEdmTaskProperties = poCmd.mpoExtra
        If Not props Is Nothing Then
            props.TaskFlags = EdmTaskFlag.EdmTask_SupportsChangeState + EdmTaskFlag.EdmTask_SupportsDetails + EdmTaskFlag.EdmTask_SupportsInitExec + EdmTaskFlag.EdmTask_SupportsInitForm

            'create the menu command for this task
            Dim cmds As EdmTaskMenuCmd() = New EdmTaskMenuCmd(0) {}
            cmds(0).mbsMenuString = "Tasks\\ECN Automation"
            cmds(0).mbsStatusBarHelp = "This command creates ECN files and automates ECN data."
            cmds(0).mlCmdID = 1
            cmds(0).mlEdmMenuFlags = EdmMenuFlags.EdmMenu_Nothing
            props.SetMenuCmds(cmds)

            'create the task setup page
            currentSetupPage = New setupPage()
            currentSetupPage.CreateControl()
            currentSetupPage.loadData(poCmd)
            Dim page(0) As EdmTaskSetupPage
            page(0).mbsPageName = "ECN Automation"
            page(0).mlPageHwnd = currentSetupPage.Handle.ToInt32()
            page(0).mpoPageImpl = currentSetupPage
            props.SetSetupPages(page)
        End If
    End Sub

    Public Sub OnTaskSetupButton(ByRef poCmd As EdmCmd, ByRef ppoData() As EdmCmdData)
        If poCmd.mbsComment = "OK" And Not IsNothing(currentSetupPage) Then
            currentSetupPage.storeData(poCmd)
        End If
        currentSetupPage = Nothing
    End Sub
    Public Shared errmsg As String
    Public Sub OnTaskRun(ByRef poCmd As EdmCmd, ByRef ppoData() As EdmCmdData)
        'Get the property interface used to access the framework
        errmsg = Nothing
        Dim inst As IEdmTaskInstance
        inst = poCmd.mpoExtra
        'Dim dxfPath As String = inst.GetValEx("fPath").ToString
        Dim retErr As String = ""
        On Error GoTo ErrHand
        'Inform the framework that the task has started
        inst.SetStatus(EdmTaskStatus.EdmTaskStat_Running)

        'Format a message to be displayed in the task list
        Dim msg As String
        msg = "Creating files..."


        'Dim idx As Integer
        'idx = 1
        'Dim maxPos As Integer
        'maxPos = 200
        'inst.SetProgressRange(maxPos, 0, msg + CStr(idx))
        'While idx < maxPos
        'Update progress bar that shows in the task list
        'inst.SetProgressPos(idx, msg + CStr(idx))
        '    idx = idx + 1

        'Handle the cancel button in the task list
        If inst.GetStatus() = EdmTaskStatus.EdmTaskStat_CancelPending Then
            inst.SetStatus(EdmTaskStatus.EdmTaskStat_DoneCancelled)
            Exit Sub
        End If

        'Call the class that creates the dxf
        For Each file In ppoData
            Dim fileID As Integer = file.mlObjectID1
            Dim foldID As Integer = file.mlObjectID2



            Dim CreateFiles As New Automation(fileID, foldID, retErr, "AUTO")

            'write the error log to a txt file
            If retErr <> "" Then
                GoTo ErrHand
            End If
        Next

        'Handle temporary suspension of the task
        If inst.GetStatus() = EdmTaskStatus.EdmTaskStat_SuspensionPending Then
            inst.SetStatus(EdmTaskStatus.EdmTaskStat_Suspended)
            While inst.GetStatus() = EdmTaskStatus.EdmTaskStat_Suspended
                System.Threading.Thread.Sleep(1000)
            End While
            If inst.GetStatus() = EdmTaskStatus.EdmTaskStat_ResumePending Then
                inst.SetStatus(EdmTaskStatus.EdmTaskStat_Running)
            End If
        End If
        'End While
        'Inform the framework that the task has successfully completed 
        inst.SetStatus(EdmTaskStatus.EdmTaskStat_DoneOK)
        Exit Sub

ErrHand:
        'Return errors to the framework by failing the task
        Dim v11 As IEdmVault11

        v11 = poCmd.mpoVault

        If Not errmsg = Nothing Then
            inst.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, "", "The task failed, " & errmsg)
        Else
            If Not retErr = Nothing Then
                inst.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, "", "The task failed, " & retErr)
            Else
                inst.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, Err.Number, "The task failed, " & v11.GetErrorMessage(Err.Number))
            End If
        End If





    End Sub
End Class
