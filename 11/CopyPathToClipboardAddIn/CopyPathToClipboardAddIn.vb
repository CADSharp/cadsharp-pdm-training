Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports EPDM.Interop.epdm

<ComVisible(True)>
<Guid("409B3FDC-058F-4767-940A-5568EAC190B5")>
Public Class CopyPathToClipboardAddIn
    Implements IEdmAddIn5

    Dim commandID As Integer = 12354
    Public Sub GetAddInInfo(ByRef poInfo As EdmAddInInfo, poVault As IEdmVault5, poCmdMgr As IEdmCmdMgr5) Implements IEdmAddIn5.GetAddInInfo

#Region "GetAddInInfo implementation "
        Try
            Console.WriteLine("Loading add-in...")

            poInfo.mbsAddInName = "Copy to clipboard addin"
            poInfo.mlRequiredVersionMajor = 10
            poInfo.mlRequiredVersionMinor = 1
            poInfo.mbsCompany = "CADSharp LLC"

            poCmdMgr.AddCmd(commandID, "Copy path", EdmMenuFlags.EdmMenu_OnlyInContextMenu, "Copy path command", 0, -1)
            Console.WriteLine("Loaded add-in.")
        Catch ex As Exception
            Console.WriteLine(ex.Message + " " + ex.StackTrace)
        End Try
#End Region

    End Sub

    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData() As EdmCmdData) Implements IEdmAddIn5.OnCmd


#Region "On Cmd"
        Dim vault As IEdmVault5
        vault = poCmd.mpoVault
        Dim affectFolder = vault.GetObject(EdmObjectType.EdmObject_File, ppoData.First.mlObjectID1)
        If (poCmd.meCmdType = EdmCmdType.EdmCmd_Menu) Then
            If poCmd.mlCmdID = commandID Then
                Console.WriteLine($"Command with ID {commandID} triggered a file called {affectFolder.Name} [{ppoData.First.mlObjectID1}] ")
            End If
        End If

#End Region
    End Sub
End Class
