Imports EPDM.Interop.epdm
Module Program

    Dim vault As IEdmVault5
    Dim vaultClass As EdmVault5
    Dim username As String
    Dim password As String
    Dim vaultName As String
    Sub init()

        username = "admin"
        password = "admin"
        vaultName = "PDM2020"

    End Sub

    Sub Main()

        init()

        vaultClass = New EdmVault5()
        vault = vaultClass

        'list all vault by name 
        Dim vaultNames = GetVaultNames(vault, False)

        Console.WriteLine($"Available vaults:")
        For Each vaultName In vaultNames
            Console.WriteLine(vaultName)
        Next

        'login to vault
        Console.WriteLine($"Attempting to login to {vaultName}...")
        Dim loginReturn As Boolean


        loginReturn = Login(vault, username, password, vaultName)


        If loginReturn Then
            Console.WriteLine("Login successful")
        Else
            Console.WriteLine("Login failed. Please check your credentials.")
        End If


        Dim folder As IEdmFolder5
        Dim filePath As String = "C:\PDM2020\Example\Axle_&.sldprt"
        Dim axlePart = vault.GetFileFromPath(filePath, folder)

        Dim handle As Int32
        handle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle.ToInt32()

        ' explanation:
        ' Locked = False -> file Is checked In 
        ' Locked = True  -> file Is checked out

        If (axlePart.IsLocked = False) Then

            'check file out 
            Console.WriteLine($"Checking out {axlePart.Name}...")
            axlePart.LockFile(folder.ID, handle, 0)
            Console.WriteLine($"Checked out {axlePart.Name} = {axlePart.IsLocked}")


            Console.WriteLine("Sleeping for several seconds...")
            System.Threading.Thread.Sleep(3000)

            Dim ret As String = String.Empty

            While (ret <> "y" And ret <> "n")

                Console.WriteLine("Would you like to undo check out? [y] | [n]")
                ret = Console.ReadLine()

            End While

            If (ret.Trim().ToLower() = "y") Then
                'undo check out 
                Console.WriteLine($"Undoing checkout of {axlePart.Name}...")
                axlePart.UndoLockFile(handle, True)
                Console.WriteLine($"Undo checkout of {axlePart.Name} = {Not axlePart.IsLocked}")
            Else
                'check check out 
                Console.WriteLine($"Checking in {axlePart.Name}...")
                axlePart.UnlockFile(handle, "Checking file back into the vault.", 0)
                Console.WriteLine($"Checked in {axlePart.Name} = {Not axlePart.IsLocked}")

            End If

        Else
                ' print who has the file checked out 
                Dim userId = axlePart.LockedByUserID
            Dim user = vault.GetObject(EdmObjectType.EdmObject_User, userId)
            Console.WriteLine($"file is checked by {user.Name}")
        End If



    End Sub


#Region "05"
    Public Function GetAllFiles(ByVal folder As IEdmFolder5) As IEdmFile5()

        Dim files As New List(Of IEdmFile5)

        Dim Position = folder.GetFirstFilePosition
        While Not Position.IsNull
            Dim iteratingFile = folder.GetNextFile(Position)
            files.Add(iteratingFile)
        End While

        Return files.ToArray()

    End Function

    Public Function GetAllSubFolders(ByVal folder As IEdmFolder5) As IEdmFolder5()

        Dim folders As New List(Of IEdmFolder5)
        Dim Position = folder.GetFirstSubFolderPosition
        While Not Position.IsNull
            Dim iteratingSubFolder = folder.GetNextSubFolder(Position)
            folders.AddRange(GetAllFiles(iteratingSubFolder))
        End While

        Return folders.ToArray()

    End Function

    Public Function GetAllFiles(ByVal folder As IEdmFolder5, ByVal recursive As Boolean) As IEdmFile5()

        Dim files As New List(Of IEdmFile5)
        If (recursive = False) Then
            Return GetAllFiles(folder)
        Else

            files.AddRange(GetAllFiles(folder))

            Dim Position = folder.GetFirstSubFolderPosition

            While Not Position.IsNull
                Dim iteratingSubFolder = folder.GetNextSubFolder(Position)

                files.AddRange(GetAllFiles(iteratingSubFolder))
                Dim iteratingSubFolderFiles = GetAllFiles(iteratingSubFolder, True)
                files.AddRange(iteratingSubFolderFiles)

            End While

        End If
        Return files.ToArray()
    End Function
    Public Function GetAllFiles(ByVal folder As IEdmFolder5, ByVal recursive As Boolean, ByRef CompletionResult As GetAllResult_e, Optional ByVal CancellationRequest As ICancellationRequest = Nothing) As IEdmFile5()
        Dim files As New List(Of IEdmFile5)
        If (recursive = False) Then
            Return GetAllFiles(folder)
        Else
            files.AddRange(GetAllFiles(folder))

            Try
                Dim Position = folder.GetFirstSubFolderPosition
                While Not Position.IsNull
                    Dim iteratingSubFolder = folder.GetNextSubFolder(Position)
                    files.AddRange(GetAllFiles(iteratingSubFolder))
                    'check if user has requested cancellation 
                    If (CancellationRequest IsNot Nothing) Then
                        If CancellationRequest.IsCancellationRequested Then
                            Throw New CancellationRequestThrownException()
                        End If
                    End If
                    Dim iteratingSubFolderFiles = GetAllFiles(iteratingSubFolder, True)
                    files.AddRange(iteratingSubFolderFiles)
                End While
            Catch ex As CancellationRequestThrownException
                CompletionResult = GetAllResult_e.CancelledByUser
                Return files.ToArray()
            End Try

        End If
        CompletionResult = GetAllResult_e.Completed
        Return files.ToArray()
    End Function
#End Region

#Region "03"
    Public Function GetVaultNames(ByVal vault As IEdmVault13, ByVal onlyLogged As Boolean) As String()

        Dim vaultViews As EdmViewInfo()

        vault.GetVaultViews(vaultViews, onlyLogged)

        Dim vaultNames As New List(Of String)
        If vaultViews IsNot Nothing Then

            For Each vaultView In vaultViews
                vaultNames.Add(vaultView.mbsVaultName)
            Next
        End If
        Return vaultNames.ToArray()

    End Function

    Public Function Login(ByVal vault As IEdmVault5, ByVal username As String, ByVal password As String, ByVal vaultName As String) As Boolean
        Try
            vault.Login(username, password, vaultName)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function LoginAuto(ByVal vault As IEdmVault5, ByVal vaultName As String) As Boolean
        Try
            Dim handle As Integer
            handle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle.ToInt32()
            vault.LoginAuto(vaultName, handle)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function LoginExternal(ByVal vault As IEdmVault5, ByVal username As String, ByVal password As String, ByVal vaultName As String) As Boolean
        Try
            Dim handle As Integer
            handle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle.ToInt32()
            vault.LoginEx(username, password, vaultName, 0)

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region

End Module

#Region "05- cancellation classes "

Public Enum GetAllResult_e
    Completed
    CancelledByUser
End Enum
Public Interface ICancellationRequest

    Sub Cancel()
    Property IsCancellationRequested() As Boolean

End Interface

Public Class CancellationRequestThrownException
    Inherits Exception


End Class
Public Class CancellationRequest
    Implements ICancellationRequest

    Dim cancelled As Boolean = False
    Public Property IsCancellationRequested As Boolean Implements ICancellationRequest.IsCancellationRequested
        Get
            Return cancelled
        End Get
        Private Set(value As Boolean)
            cancelled = value
        End Set
    End Property

    Public Sub Cancel() Implements ICancellationRequest.Cancel
        IsCancellationRequested = True
    End Sub
End Class

#End Region