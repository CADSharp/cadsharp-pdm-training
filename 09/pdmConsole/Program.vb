Imports System.IO
Imports System.Runtime.InteropServices
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
        For Each vaultName_ In vaultNames
            Console.WriteLine(vaultName_)
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


        Dim rootFolder As IEdmFolder5
        rootFolder = vault.RootFolder
        Dim exampleSubFolder As IEdmFolder5

        exampleSubFolder = rootFolder.GetSubFolder("Example")

        Dim filePath As String
        filePath = "C:\PDM2020\Example\Full_Grill_Assembly.sldasm"

        Dim folder As IEdmFolder5
        Dim GrillAssemblyFile As IEdmFile5

        GrillAssemblyFile = vault.GetFileFromPath(filePath, folder)

        Dim handle As Integer
        handle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle.ToInt32()

        Dim layoutName As String
        Dim bominCSV As String

        layoutName = GetFirstLayoutName(vault)

        bominCSV = GetBOMCSV(GrillAssemblyFile, layoutName, "@")

        Console.WriteLine(bominCSV)

        Console.ReadLine()

    End Sub
#Region "09"

    Public Function GetFirstLayoutName(vault2 As IEdmVault9) As String

        Dim bomMgr As IEdmBomMgr = vault2.CreateUtility(EdmUtility.EdmUtil_BomMgr)

        Dim ppoRetLayouts() As EdmBomLayout = Nothing

        bomMgr.GetBomLayouts(ppoRetLayouts)

        If ppoRetLayouts Is Nothing Or ppoRetLayouts.Length = 0 Then
            Throw New Exception("Cannot find any BOM layouts.")
        End If

        Return ppoRetLayouts.First.mbsLayoutName

    End Function
    Public Function GetBOMCSV(ByVal file As IEdmFile5, ByVal layoutName As String, ByVal configurationName As String) As String

        If file Is Nothing Then
            Throw New NullReferenceException("the file parameter is null at PDM::GetBOMCSV.")
        End If

        If String.IsNullOrWhiteSpace(configurationName) Then
            configurationName = "@"
        End If

        Try
            Dim file7 = TryCast(file, IEdmFile7)

            Dim computedBOM = file7.GetComputedBOM(layoutName, 0, configurationName, EdmBomFlag.EdmBf_ShowSelected)

            If computedBOM Is Nothing Then
                Throw New Exception($"Failed to get computed BOM.", Nothing)
            End If

            Dim computedBOM3 As IEdmBomView3 = computedBOM

            Dim tempFolder = System.IO.Path.GetTempPath()

            Dim tempFolderDir = New DirectoryInfo(tempFolder)

            Dim bomCSVFilePath = Path.Combine(tempFolderDir.FullName, $"{System.IO.Path.GetFileNameWithoutExtension(file.Name)}--{layoutName}.txt")

            computedBOM3.SaveToCSV(bomCSVFilePath, True)

            Dim txt = System.IO.File.ReadAllText(bomCSVFilePath)

            Return txt

        Catch e As COMException
            Throw New Exception($"PDM Exception: {e.Message}", e)

        Catch e As Exception

            Throw New Exception($".NET Exception: {e.Message}", e)

        End Try
    End Function

#End Region
#Region "08"
    Public Function BatchUpdateDataCardVariables(ByVal folder As IEdmFolder5, ByVal vault7 As IEdmVault7, ByVal Variables As IEdmVariable5(), ByVal configurationNames As String(), ByVal Values As Object()) As EdmBatchError2()


        Dim Update As IEdmBatchUpdate2
        Update = vault7.CreateUtility(EdmUtility.EdmUtil_BatchUpdate)
        Dim VariableMgr As IEdmVariableMgr5
        VariableMgr = vault7


        Dim Search As IEdmSearch5

        Search = vault7.CreateUtility(EdmUtility.EdmUtil_Search)
        Search.FindFiles = True
        Search.FindFolders = False
        Search.StartFolderID = folder.ID

        Dim Result As IEdmSearchResult5
        Result = Search.GetFirstResult

        While Not Result Is Nothing
            For Each variable As IEdmVariable5 In Variables
                Dim valueIndex As Integer = Array.IndexOf(Variables, variable)
                Dim configurationIndex = valueIndex
                Dim value As Object = Values(valueIndex)
                Dim configurationName As String = configurationNames(configurationIndex)
                Update.SetVar(Result.ID, variable.ID, value, configurationName, EdmBatchFlags.EdmBatch_AllConfigs)
            Next
            Result = Search.GetNextResult
        End While



        Dim Errors() As EdmBatchError2 = Nothing
        Dim errorSize As Integer = Update.CommitUpdate(Errors, Nothing)

        Return Errors
    End Function

    Public Function BatchCheckIn(ByVal Vault As IEdmVault8, ByVal files() As IEdmFile5, ByVal Handle As Integer, Optional locks As EdmUnlockBuildTreeFlags = EdmUnlockBuildTreeFlags.Eubtf_MayUnlock) As Boolean

        Try
            ' create batch unlocker object 
            Dim batchUnlocker As IEdmBatchUnlock2 = Vault.CreateUtility(EdmUtility.EdmUtil_BatchUnlock)

            ' create the selection list 
            Dim list = New List(Of EdmSelItem)

            Dim selectedFile

            For Each afile In files

                selectedFile = New EdmSelItem

                Dim aPos As IEdmPos5 = afile.GetFirstFolderPosition

                Dim aFolder As IEdmFolder5 = afile.GetNextFolder(aPos)

                selectedFile.mlDocID = afile.ID

                selectedFile.mlProjID = aFolder.ID

                list.Add(selectedFile)
            Next


            Dim ppoSelection() As EdmSelItem = Nothing

            ppoSelection = list.ToArray

            batchUnlocker.AddSelection(Vault, ppoSelection)

            'create tree 
            batchUnlocker.CreateTree(Handle, locks)

            ' unlock file 
            batchUnlocker.UnlockFiles(Handle, Nothing)


            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function


#End Region



#Region "07"

    Public Function SetDatacardVariableValue(ByVal file As IEdmFile5, ByVal VariableName As String, ByVal Value As Object, Optional ByVal configurationName As String = "@") As Tuple(Of Boolean, String)
        Dim enumeratorVariable As IEdmEnumeratorVariable5 = Nothing
        Try
            enumeratorVariable = file.GetEnumeratorVariable()
            enumeratorVariable.SetVar(VariableName, configurationName, Value)
            enumeratorVariable.CloseFile(True)
            Return New Tuple(Of Boolean, String)(True, String.Empty)
        Catch ex As COMException
            If (enumeratorVariable IsNot Nothing) Then
                enumeratorVariable.CloseFile(True)
            End If
            Return New Tuple(Of Boolean, String)(False, ex.Message)
        End Try

    End Function
    Public Function GetDatacardVariableValue(ByVal file As IEdmFile5, ByVal VariableName As String, Optional ByVal configurationName As String = "@") As Tuple(Of Boolean, String, Object)
        Dim enumeratorVariable As IEdmEnumeratorVariable5 = Nothing
        Try
            Dim value As Object = Nothing
            enumeratorVariable = file.GetEnumeratorVariable()
            enumeratorVariable.GetVar(VariableName, configurationName, value)
            enumeratorVariable.CloseFile(True)
            Return New Tuple(Of Boolean, String, Object)(True, String.Empty, value)
        Catch ex As COMException
            If (enumeratorVariable IsNot Nothing) Then
                enumeratorVariable.CloseFile(True)
            End If
            Return New Tuple(Of Boolean, String, Object)(False, ex.Message, Nothing)
        End Try

    End Function
    Public Function GetDatacardVariableValueByVersion(ByVal file As IEdmFile5, ByVal VariableName As String, ByVal Version As Integer, ByVal folderID As Integer, Optional ByVal configurationName As String = "@") As Tuple(Of Boolean, String, Object)

        Dim enumeratorVariable As IEdmEnumeratorVariable8 = Nothing
        Try
            Dim ppRetVariables As Object() = Nothing
            Dim getVarData As EdmGetVarData
            Dim value As Object = Nothing
            enumeratorVariable = file.GetEnumeratorVariable()
            enumeratorVariable.GetVersionVars(Version, folderID, ppRetVariables, New String() {configurationName}, getVarData)
            enumeratorVariable.CloseFile(True)

            For Each ppRetVariable As IEdmVariableValue6 In ppRetVariables
                If (ppRetVariable.VariableName = VariableName) Then
                    value = ppRetVariable.GetValue(configurationName)
                    Exit For
                End If
            Next

            Return New Tuple(Of Boolean, String, Object)(True, String.Empty, value)
        Catch ex As COMException
            If (enumeratorVariable IsNot Nothing) Then
                enumeratorVariable.CloseFile(True)
            End If
            Return New Tuple(Of Boolean, String, Object)(False, ex.Message, Nothing)
        End Try



    End Function
    Public Function GetVariableValueFromDb(ByVal file As IEdmFile5, ByVal VariableName As String, Optional ByVal configurationName As String = "@") As Tuple(Of Boolean, String, Object)
        Dim enumeratorVariable As IEdmEnumeratorVariable8 = Nothing
        Try
            Dim value As Object = Nothing
            enumeratorVariable = file.GetEnumeratorVariable()
            enumeratorVariable.GetVarFromDb(VariableName, configurationName, value)
            enumeratorVariable.CloseFile(True)
            Return New Tuple(Of Boolean, String, Object)(True, String.Empty, value)
        Catch ex As COMException
            If (enumeratorVariable IsNot Nothing) Then
                enumeratorVariable.CloseFile(True)
            End If
            Return New Tuple(Of Boolean, String, Object)(False, ex.Message, Nothing)
        End Try

    End Function
    Public Function GetAllVariables(ByVal variableMgr As IEdmVariableMgr5) As IEdmVariable5()

        Dim variables As New List(Of IEdmVariable5)
        Dim position = variableMgr.GetFirstVariablePosition()
        While Not position.IsNull
            Dim iteratingVariable = variableMgr.GetNextVariable(position)
            variables.Add(iteratingVariable)
        End While

        Return variables.ToArray()


    End Function

#End Region

#Region "06"

    Public Function IsCheckedOut(ByVal file As IEdmFile5) As Boolean
        If (file IsNot Nothing) Then
            Return file.IsLocked
        End If
        Return False
    End Function
    Public Function IsCheckedIn(ByVal file As IEdmFile5) As Boolean
        If (file IsNot Nothing) Then
            Return Not file.IsLocked
        End If
        Return False
    End Function
    Public Function CheckIn(ByVal file As IEdmFile5, Optional comment As String = "", Optional handle As Integer = 0) As Boolean
        If (file IsNot Nothing) Then
            file.UnlockFile(handle, comment, 0)
            Return IsCheckedIn(file)
        End If
        Return False
    End Function
    Public Function CheckOut(ByVal file As IEdmFile5, ByVal folder As IEdmFolder5, Optional handle As Integer = 0) As Boolean
        If (file IsNot Nothing) Then
            file.LockFile(folder.ID, handle)
            Return IsCheckedOut(file)
        End If
        Return False
    End Function




#End Region

#Region "05"


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
#End Region

#Region "04"
    'no helper sub or functions here
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
Public Enum CompletionState_e
    CancelledByUser
    CompletedSuccessfully
End Enum
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


#End Region