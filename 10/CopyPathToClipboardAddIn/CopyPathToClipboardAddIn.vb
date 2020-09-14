Imports System.Runtime.InteropServices
Imports EPDM.Interop.epdm

<ComVisible(True)>
<Guid("409B3FDC-058F-4767-940A-5568EAC190B5")>
Public Class CopyPathToClipboardAddIn
    Implements IEdmAddIn5

    Public Sub GetAddInInfo(ByRef poInfo As EdmAddInInfo, poVault As IEdmVault5, poCmdMgr As IEdmCmdMgr5) Implements IEdmAddIn5.GetAddInInfo
        Throw New NotImplementedException()
    End Sub

    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData() As EdmCmdData) Implements IEdmAddIn5.OnCmd
        Throw New NotImplementedException()
    End Sub
End Class
