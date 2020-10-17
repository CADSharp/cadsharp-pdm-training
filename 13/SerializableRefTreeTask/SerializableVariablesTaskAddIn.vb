Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports EPDM.Interop.epdm

<Guid("B652ED7E-5831-4013-911A-3469C3F83F56")>
<ComVisible(True)>
Public Class SerializableVariablesTaskAddIn
    Implements IEdmAddIn5


    Public Sub GetAddInInfo(ByRef poInfo As EdmAddInInfo, poVault As IEdmVault5, poCmdMgr As IEdmCmdMgr5) Implements IEdmAddIn5.GetAddInInfo
#Region "GetAddInInfo implementation"
        Try
            Console.WriteLine("Loading add-in...")

            poInfo.mbsAddInName = "Serializable Variables Task"
            poInfo.mbsCompany = "CADSharp LLC"
            poInfo.mbsDescription = "Serializable properties of a PDM file."

            poInfo.mlRequiredVersionMajor = 10
            poInfo.mlRequiredVersionMinor = 1
            poInfo.mlAddInVersion = 1



            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetup)

            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetupButton)

            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskLaunch)

            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskRun)

            Console.WriteLine("Add-in loaded.")

        Catch ex As Exception
            Console.WriteLine($"Message = {ex.Message} - StackTrace = {ex.StackTrace}")
        End Try
#End Region

    End Sub

    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData() As EdmCmdData) Implements IEdmAddIn5.OnCmd

        Select Case poCmd.meCmdType
            Case EdmCmdType.EdmCmd_TaskSetup
                OnTaskSetup(poCmd, ppoData)
                Exit Select
            Case EdmCmdType.EdmCmd_TaskSetupButton
                OnTaskSetupButton(poCmd, ppoData)
                Exit Select
            Case EdmCmdType.EdmCmd_TaskLaunch
                OnTaskLaunch(poCmd, ppoData)
                Exit Select
            Case EdmCmdType.EdmCmd_TaskRun
                OnTaskRun(poCmd, ppoData)
                Exit Select
        End Select


    End Sub



    Public settingsSetupUserControl As New Settings
    Sub OnTaskSetup(ByRef poCmd As EdmCmd, ByRef ppoData() As EdmCmdData)



        Dim taskProperties As IEdmTaskProperties = poCmd.mpoExtra

        taskProperties.TaskFlags = EdmTaskFlag.EdmTask_SupportsInitExec

        Dim settingsSetupSetup As New EdmTaskSetupPage
        settingsSetupUserControl.CreateControl()

        settingsSetupUserControl.Load(poCmd)

        settingsSetupSetup.mbsPageName = settingsSetupUserControl.Name
        settingsSetupSetup.mlPageHwnd = settingsSetupUserControl.Handle.ToInt32()
        settingsSetupSetup.mpoPageImpl = settingsSetupUserControl

        taskProperties.SetSetupPages(New EdmTaskSetupPage() {settingsSetupSetup})

        'add menu item
        Dim cmd As EdmTaskMenuCmd
        cmd.mbsMenuString = "Serialize all variables..."
        cmd.mbsStatusBarHelp = "Serialize all variables..."
        cmd.mlCmdID = 2051
        cmd.mlEdmMenuFlags = EdmMenuFlags.EdmMenu_ContextMenuItem

        Dim cmds As New List(Of EdmTaskMenuCmd)
        cmds.Add(cmd)
        taskProperties.SetMenuCmds(cmds.ToArray())

    End Sub
    Sub OnTaskSetupButton(ByRef poCmd As EdmCmd, ByRef ppoData() As EdmCmdData)

        settingsSetupUserControl.Store(poCmd)

    End Sub


    Sub OnTaskLaunch(ByRef poCmd As EdmCmd, ByRef ppoData() As EdmCmdData)

    End Sub
    Sub OnTaskRun(ByRef poCmd As EdmCmd, ByRef ppoData() As EdmCmdData)

        Try
            Dim taskInstance As IEdmTaskInstance = poCmd.mpoExtra
            Dim vault As IEdmVault5 = poCmd.mpoVault

            Dim serializationType As SerializationType
            Dim saveLocation As String
            saveLocation = taskInstance.GetValEx(Settings.saveLocationKey)
            Dim serializationRet As Boolean = [Enum].TryParse(Of SerializationType)(taskInstance.GetValEx(Settings.serializationTypeKey), True, serializationType)

            For Each ppoDatum As EdmCmdData In ppoData


                Dim affectedFile As IEdmFile5


                If (ppoDatum.mlObjectID1 <> 0) Then
                    affectedFile = vault.GetObject(EdmObjectType.EdmObject_File, ppoDatum.mlObjectID1)
                    Dim serializedVariables = SerializationHelper.SerializeProperties(affectedFile, serializationType)
                    If (String.IsNullOrWhiteSpace(serializedVariables) = False) Then
                        System.IO.File.WriteAllText($"{saveLocation}\{System.IO.Path.GetFileNameWithoutExtension(affectedFile.Name)}.{serializationType.ToString()}", serializedVariables)
                    End If
                End If
            Next
        Catch ex As Exception

        End Try


    End Sub




End Class
