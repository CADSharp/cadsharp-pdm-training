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


        'get root folder 
        Dim rootFolder As IEdmFolder5
        rootFolder = vault.RootFolder

        Dim exampleSubFolder = rootFolder.GetSubFolder("Example")

        Console.WriteLine("")
        Console.WriteLine("GetFiles non-recursive:")
        Console.WriteLine("")
        Dim files = GetAllFiles(exampleSubFolder)
        Console.WriteLine("")
        Console.WriteLine($"Found {files.Length} files.")
        Console.WriteLine("")

        For Each file As IEdmFile5 In files
            Console.WriteLine(file.Name)
        Next
        Console.WriteLine("")
        Console.WriteLine("GetFiles recursive:")
        Console.WriteLine("")

        files = GetAllFiles(exampleSubFolder, True)

        Console.WriteLine("")
        Console.WriteLine($"Found {files.Length} files.")
        Console.WriteLine("")

        For Each file As IEdmFile5 In files
            Console.WriteLine(file.Name)
        Next

        Console.WriteLine("")
        Console.WriteLine("GetFiles recursive with cancellation after 1 seconds (method is delayed by 2 seconds):")
        Console.WriteLine("")

        Dim cancellationRequest = New CancellationRequest()
        Dim thread = New System.Threading.Thread(New Threading.ThreadStart(Sub()
                                                                               System.Threading.Thread.Sleep(1000)
                                                                               cancellationRequest.Cancel()
                                                                           End Sub))

        thread.Start()

        Dim completionResult As GetAllResult_e
        files = GetAllFiles(exampleSubFolder, True, completionResult, cancellationRequest)

        Console.WriteLine("")
        Console.WriteLine($"Result: {completionResult.ToString()}")
        Console.WriteLine($"Found {files.Length} files.")
        Console.WriteLine("")

        For Each file As IEdmFile5 In files
            Console.WriteLine(file.Name)
        Next

#Region "sample code from 03"
        'Dim axleFirst = exampleSubFolder.GetFile("axle_&.sldprt")

        ''list file meta data
        'Console.WriteLine("File data:")
        'Console.WriteLine($"File name: {axleFirst.Name}")
        'Console.WriteLine($"File ID: {axleFirst.ID}")
        'Console.WriteLine($"File Version: {axleFirst.CurrentVersion}")
        'Console.WriteLine($"File Rev: {axleFirst.CurrentRevision}")
        'Console.WriteLine($"File State: {axleFirst.CurrentState.Name}")



        ''attempt to get file file from path
        'Dim folder As IEdmFolder5
        'Dim filePath As String = "C:\PDM2020\Example\Axle_&.sldprt"
        'axleFirst = vault.GetFileFromPath(filePath, folder)

        ''get file using id 
        'Dim id As Integer = axleFirst.ID
        'axleFirst = vault.GetObject(EdmObjectType.EdmObject_File, id)

        ''get file by doing a search
        'Dim search As IEdmSearch5
        'search = vault.CreateSearch()
        'search.FileName = "axle_&.sldprt"
        'search.FindFiles = True

        ''find file by search 
        'Console.WriteLine("Searching...")
        'Dim searchResult = search.GetFirstResult()

        'While (searchResult IsNot Nothing)
        '    Console.WriteLine($"Found: {searchResult.Name} [{searchResult.Path}]")
        '    searchResult = search.GetNextResult()
        'End While

#End Region

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
                    System.Threading.Thread.Sleep(2000)
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