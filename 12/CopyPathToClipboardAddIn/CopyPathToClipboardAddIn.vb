Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports EPDM.Interop.epdm

<ComVisible(True)>
<Guid("409B3FDC-058F-4767-940A-5568EAC190B5")>
Public Class CopyPathToClipboardAddIn
    Implements IEdmAddIn5

    Dim commandID As Integer = 12345
    Public Sub GetAddInInfo(ByRef poInfo As EdmAddInInfo, poVault As IEdmVault5, poCmdMgr As IEdmCmdMgr5) Implements IEdmAddIn5.GetAddInInfo
#Region "GetAddInInfo implementation"
        Try
            Console.WriteLine("Loading add-in...")

            poInfo.mbsAddInName = "Copy to clipboard add-in"
            poInfo.mbsCompany = "CADSharp LLC"
            poInfo.mlRequiredVersionMajor = 10
            poInfo.mlRequiredVersionMinor = 1
            poInfo.mlAddInVersion = 1

            poCmdMgr.AddCmd(commandID, "Copy path", EdmMenuFlags.EdmMenu_OnlyInContextMenu, "Copy path to clipboard command", "Copy path to clipboard command", -1, 0)

            Console.WriteLine("Add-in loaded.")
        Catch ex As Exception
            Console.WriteLine($"Message = {ex.Message} - StackTrace = {ex.StackTrace}")
        End Try
#End Region
    End Sub

    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData() As EdmCmdData) Implements IEdmAddIn5.OnCmd
#Region "On Cmd"

        Dim vault As IEdmVault5
        vault = poCmd.mpoVault
        Dim stringBuilder As New Text.StringBuilder

        If poCmd.meCmdType = EdmCmdType.EdmCmd_Menu Then
            If poCmd.mlCmdID = commandID Then

                For Each ppoDatum As EdmCmdData In ppoData

                    Dim affectedFile As IEdmFile5
                    Dim affectedFolder As IEdmFolder5
                    Dim fileParentFolder As IEdmFolder5

                    Dim path As String

                    If (ppoDatum.mlObjectID1 <> 0) Then
                        affectedFile = vault.GetObject(EdmObjectType.EdmObject_File, ppoDatum.mlObjectID1)
                        fileParentFolder = affectedFile.GetNextFolder(affectedFile.GetFirstFolderPosition)
                        path = $"{fileParentFolder.LocalPath}\{affectedFile.Name}"
                    Else
                        affectedFolder = vault.GetObject(EdmObjectType.EdmObject_Folder, ppoDatum.mlObjectID2)
                        path = affectedFolder.LocalPath
                    End If

                    Console.WriteLine($"copied to clipboard: {path}")

                    stringBuilder.AppendLine(path)

                Next

            End If
        End If

        Dim paths As String = String.Empty
        paths = stringBuilder.ToString()
        If (String.IsNullOrWhiteSpace(paths) = False) Then
            Clipboard.SetText(paths)
        End If

#End Region

    End Sub
End Class
