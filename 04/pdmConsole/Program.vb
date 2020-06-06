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


        vaultClass = New EdmVault5()
        vault = vaultClass

        'list all vault by name 
        Dim vaultNames = GetVaultNames(vault, False)

        Console.WriteLine($"Available vaults:")
        For Each vaultName In vaultNames
            Console.WriteLine(vaultName)
        Next

        'ask for user input 
        Console.Write("Please enter a vault to log to:")
        vaultName = Console.ReadLine()

        Console.Write("Please enter a username:")
        username = Console.ReadLine()

        Console.Write("Please enter a password:")
        password = Console.ReadLine()


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

        'get example subfolder 
        Dim exampleSubFolder As IEdmFolder5
        exampleSubFolder = rootFolder.GetSubFolder("Example")

        'get axle part 
        Dim axlePart As IEdmFile5

        axlePart = exampleSubFolder.GetFile("axle_&.sldprt")

        'print part metadata
        Console.WriteLine($"Meta data:")
        Console.WriteLine($"")

        Console.WriteLine($"File name: {axlePart.Name}")
        Console.WriteLine($"File ID: {axlePart.ID}")
        Console.WriteLine($"File version: {axlePart.CurrentVersion}")
        Console.WriteLine($"File revision: {axlePart.CurrentRevision}")
        Console.WriteLine($"File workflow state: {axlePart.CurrentState.Name}")

        'get axle part using file path

        Dim folder As IEdmFolder5
        Dim filePath As String
        filePath = "C:\PDM2020\Example\axle_&.sldprt"

        axlePart = vault.GetFileFromPath(filePath, folder)

        'get file using GetObject 
        Dim id As Integer
        id = axlePart.ID
        axlePart = vault.GetObject(EdmObjectType.EdmObject_File, id)

        'get file using pdm search 
        Dim search As IEdmSearch5
        search = vault.CreateSearch()

        search.FindFiles = True
        search.FileName = "axle_&.sldprt"

        Console.WriteLine("Searching for file...")

        Dim searchResult = search.GetFirstResult()

        While searchResult IsNot Nothing

            Console.WriteLine($"Found: {searchResult.Name} [{searchResult.Path}]")
            searchResult = search.GetNextResult
        End While


        Console.ReadLine()

    End Sub

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
