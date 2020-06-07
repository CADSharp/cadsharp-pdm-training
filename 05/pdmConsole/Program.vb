Imports EPDM.Interop.epdm
Imports System.Collections.Generic
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
        Console.WriteLine("Getting files - Non recursive")
        Console.WriteLine("")

        Dim files = GetAllFiles(exampleSubFolder)

        Console.WriteLine($"Found {files.Length} files.")
        Console.WriteLine("")

        For Each file As IEdmFile5 In files
            Console.WriteLine(file.Name)
        Next

        Console.WriteLine("")
        Console.WriteLine("Getting files - recursive")
        Console.WriteLine("")

        files = GetAllFiles(exampleSubFolder, True)

        Console.WriteLine($"Found {files.Length} files.")
        Console.WriteLine("")

        For Each file As IEdmFile5 In files
            Console.WriteLine(file.Name)
        Next

        Console.WriteLine("")
        Console.WriteLine("Getting files - recursive with cancellation")
        Console.WriteLine("")

        Dim cancellationRequest As New CancellationRequest
        Dim completionState As CompletionState_e

        Dim cancellationThread As Threading.Thread
        cancellationThread = New Threading.Thread(New Threading.ThreadStart(Sub()
                                                                                System.Threading.Thread.Sleep(500)
                                                                                cancellationRequest.Cancel()
                                                                            End Sub))
        cancellationThread.Start()
        files = GetAllFiles(exampleSubFolder, True, completionState, cancellationRequest)




        Console.WriteLine($"Found {files.Length} files.")
        Console.WriteLine("")

        For Each file As IEdmFile5 In files
            Console.WriteLine(file.Name)
        Next


        Console.ReadLine()

#Region "sample code from 04"
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


    Public Function GetAllFiles(ByVal folder As IEdmFolder5) As IEdmFile5()

        Dim files As New List(Of IEdmFile5)

        Dim position = folder.GetFirstFilePosition

        While (Not position.IsNull)

            Dim iteratingFile As IEdmFile5
            iteratingFile = folder.GetNextFile(position)
            files.Add(iteratingFile)
        End While
        Return files.ToArray()
    End Function

    Public Function GetAllFiles(ByVal folder As IEdmFolder5, ByVal GetAllSubFolders As Boolean) As IEdmFile5()

        Dim files As New List(Of IEdmFile5)

        If GetAllSubFolders = False Then
            Return GetAllFiles(folder)
        Else

            files.AddRange(GetAllFiles(folder))

            Dim position = folder.GetFirstSubFolderPosition

            While (Not position.IsNull)
                Dim iteratingSubFolder As IEdmFolder5
                iteratingSubFolder = folder.GetNextSubFolder(position)
                files.AddRange(GetAllFiles(iteratingSubFolder, True))
            End While

            Return files.ToArray()

        End If
    End Function
    Public Function GetAllFiles(ByVal folder As IEdmFolder5, ByVal GetAllSubFolders As Boolean, ByRef completionState As CompletionState_e, Optional cancellationRequest As ICancellationRequest = Nothing) As IEdmFile5()

        Dim files As New List(Of IEdmFile5)

        If GetAllSubFolders = False Then
            Return GetAllFiles(folder)
        Else

            files.AddRange(GetAllFiles(folder))

            Try
                Dim position = folder.GetFirstSubFolderPosition
                While (Not position.IsNull)
                    Dim iteratingSubFolder As IEdmFolder5
                    iteratingSubFolder = folder.GetNextSubFolder(position)
                    System.Threading.Thread.Sleep(1000)
                    If cancellationRequest IsNot Nothing Then
                        If cancellationRequest.IsCancellationRequestThrown Then
                            Throw New CancellationRequestThrownException
                        End If
                    End If
                    files.AddRange(GetAllFiles(iteratingSubFolder, True, completionState, cancellationRequest))
                End While
            Catch ex As CancellationRequestThrownException
                completionState = CompletionState_e.CancelledByUser
                Return files.ToArray()
            End Try
            completionState = CompletionState_e.CompletedSuccessfully
            Return files.ToArray()

        End If
    End Function

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

    Public Enum CompletionState_e
        CancelledByUser
        CompletedSuccessfully
    End Enum

End Module


Public Class CancellationRequestThrownException
    Inherits Exception
End Class

Public Class CancellationRequest
    Implements ICancellationRequest

    Private Cancelled As Boolean = False
    Public Property IsCancellationRequestThrown As Boolean Implements ICancellationRequest.IsCancellationRequestThrown
        Get
            Return Cancelled
        End Get
        Set(value As Boolean)
            Cancelled = value
        End Set
    End Property

    Public Sub Cancel() Implements ICancellationRequest.Cancel
        IsCancellationRequestThrown = True
    End Sub
End Class

Public Interface ICancellationRequest
    Sub Cancel()
    Property IsCancellationRequestThrown() As Boolean
End Interface


