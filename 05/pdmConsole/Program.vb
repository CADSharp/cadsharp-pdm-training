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
        Console.WriteLine("Get All Files Non-recursive:")
        Console.WriteLine("")
        Dim files = GetAllFiles(exampleSubFolder)
        Console.WriteLine("")
        Console.WriteLine($"Found: {files.Count} files.")
        Console.WriteLine("")

        For Each file As IEdmFile5 In files
            Console.WriteLine(file.Name)
        Next

        Console.WriteLine("")
        Console.WriteLine("Get All Files recursive:")
        Console.WriteLine("")

        files = GetAllFiles(exampleSubFolder, True)
        Console.WriteLine("")
        Console.WriteLine($"Found: {files.Count} files.")
        Console.WriteLine("")

        For Each file As IEdmFile5 In files
            Console.WriteLine(file.Name)
        Next


        Console.WriteLine("")
        Console.WriteLine("Get All Files recursive (with cancellation):")
        Console.WriteLine("")
        Dim cancellationRequest As New CancellationRequest()
        Dim completion As CompletionResult_e
        Dim thread = New System.Threading.Thread(New Threading.ThreadStart(Sub()
                                                                               System.Threading.Thread.Sleep(1000)
                                                                               cancellationRequest.Cancel()

                                                                           End Sub))

        thread.Start()

        files = GetAllFiles(exampleSubFolder, True, completion, cancellationRequest)
        Console.WriteLine("")
        Console.WriteLine($"Completion: {completion.ToString()}")
        Console.WriteLine($"Found: {files.Count} files.")
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

        Dim position = folder.GetFirstFilePosition()

        While Not position.IsNull

            Dim iteratingFile As IEdmFile5
            iteratingFile = folder.GetNextFile(position)

            files.Add(iteratingFile)

        End While

        Return files.ToArray

    End Function

    Public Function GetAllSubFolders(ByVal folder As IEdmFolder5) As IEdmFolder5()
        Dim folders As New List(Of IEdmFolder5)
        Dim position = folder.GetFirstSubFolderPosition
        While (Not position.IsNull)
            Dim iteratingSubFolder As IEdmFolder5
            iteratingSubFolder = folder.GetNextSubFolder(position)
            folders.Add(iteratingSubFolder)
        End While
        Return folders.ToArray()
    End Function
    Public Function GetAllFiles(ByVal folder As IEdmFolder5, ByVal GetAllSubFolders As Boolean) As IEdmFile5()
        Dim files As New List(Of IEdmFile5)

        If (GetAllSubFolders = False) Then
            Return GetAllFiles(folder)
        Else
            files.AddRange(GetAllFiles(folder))
            Dim position As IEdmPos5 = folder.GetFirstSubFolderPosition()
            While (Not position.IsNull)

                Dim iteratingSubFolder As IEdmFolder5

                iteratingSubFolder = folder.GetNextSubFolder(position)
                files.AddRange(GetAllFiles(iteratingSubFolder))
                Dim iteratingSubSubFolderFiles = GetAllFiles(iteratingSubFolder, True)
                files.AddRange(iteratingSubSubFolderFiles)

            End While
        End If

        Return files.ToArray()
    End Function

    Public Function GetAllFiles(ByVal folder As IEdmFolder5, ByVal GetAllSubFolders As Boolean, ByRef CompletionResult As CompletionResult_e, ByVal CancellationRequest As ICancellationRequest)
        Dim files As New List(Of IEdmFile5)

        If (GetAllSubFolders = False) Then
            Return GetAllFiles(folder)
        Else
            Try

                files.AddRange(GetAllFiles(folder))
                Dim position As IEdmPos5 = folder.GetFirstSubFolderPosition()
                While (Not position.IsNull)

                    Dim iteratingSubFolder As IEdmFolder5

                    iteratingSubFolder = folder.GetNextSubFolder(position)


                    System.Threading.Thread.Sleep(2000)

                    If (CancellationRequest.IsCancellationRequestThrown) Then
                        Throw New CancellationRequestThrownException
                    End If

                    files.AddRange(GetAllFiles(iteratingSubFolder))

                    Dim iteratingSubSubFolderFiles = GetAllFiles(iteratingSubFolder, True)
                    files.AddRange(iteratingSubSubFolderFiles)

                End While
            Catch ex As CancellationRequestThrownException
                CompletionResult = CompletionResult_e.CancelledByUser
                Return files.ToArray()
            End Try
        End If
        CompletionResult = CompletionResult_e.Complete
        Return files.ToArray()
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

    Public Enum CompletionResult_e

        Unknown
        Complete
        CancelledByUser

    End Enum


End Module

Public Class CancellationRequestThrownException
    Inherits Exception
End Class
Public Interface ICancellationRequest
    Sub Cancel()
    Property IsCancellationRequestThrown() As Boolean
End Interface

Public Class CancellationRequest
    Implements ICancellationRequest

    Dim cancelled As Boolean = False
    Public Property IsCancellationRequestThrown As Boolean Implements ICancellationRequest.IsCancellationRequestThrown
        Get
            Return cancelled
        End Get
        Set(value As Boolean)
            cancelled = value
        End Set
    End Property

    Public Sub Cancel() Implements ICancellationRequest.Cancel
        IsCancellationRequestThrown = True
    End Sub
End Class
