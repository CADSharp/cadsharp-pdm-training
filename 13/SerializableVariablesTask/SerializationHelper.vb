Imports System.IO
Imports System.Xml.Serialization
Imports EPDM.Interop.epdm
Imports Newtonsoft.Json

Public Class SerializationHelper

    Public Shared Function SerializeProperties(ByVal file As IEdmFile5, ByVal serializationType As SerializationType) As String

        Dim serializationOutput As String = String.Empty

        Dim configurationNames As New List(Of String)


        Dim folder As IEdmFolder5 = Nothing
        Dim version As Integer = 0
        Dim ppoRetVariables As Object() = Nothing
        Dim ppoConfigurations As Object() = Nothing
        Dim ppoRetData As EdmGetVarData = Nothing

        If file Is Nothing Then
            Throw New NullReferenceException("file")
        End If


        Dim enumerator As IEdmEnumeratorVariable10
        enumerator = file.GetEnumeratorVariable

        folder = file.GetNextFolder(file.GetFirstFolderPosition)
        version = file.CurrentVersion


        Dim strList = file.GetConfigurations(version)

        Dim pos = strList.GetHeadPosition()
        While (pos.IsNull = False)
            configurationNames.Add(strList.GetNext(pos))
        End While

        ppoConfigurations = configurationNames.ToArray

        enumerator.GetVersionVars(version, folder.ID, ppoRetVariables, ppoConfigurations, ppoRetData)

        Dim serializableFile As New File
        serializableFile.Name = file.Name
        serializableFile.ID = file.ID


        Dim variables As New List(Of Variable)

        For Each ppoRetVariable As IEdmVariableValue6 In ppoRetVariables
            Dim variable As New Variable
            variable.Name = ppoRetVariable.VariableName
            For Each configurationName As String In configurationNames
                Dim Value As Object
                Value = ppoRetVariable.GetValue(configurationName)
                variable.ConfigurationValuePairs.Add(New KeyValuePair(Of String, Object)(configurationName, Value))
                serializableFile.AddVariable(variable)
            Next


        Next

        If serializationType = SerializationType.Json Then
            Dim serializedText = JsonConvert.SerializeObject(serializableFile)
            Return serializedText
        Else
            Dim x As System.Xml.Serialization.XmlSerializer = New System.Xml.Serialization.XmlSerializer(serializableFile.GetType())


            Dim textWriter As New StringWriter()

            x.Serialize(textWriter, serializableFile)

            serializationOutput = textWriter.ToString()

            textWriter.Close()
            textWriter.Dispose()

        End If

        Return serializationOutput

    End Function

End Class

Public Class File

    Public Sub AddVariable(ByVal variable As Variable)
        Dim array As New ArrayList
        array.AddRange(Variables)
        array.Add(variable)
        Variables = array.ToArray().Cast(Of Variable).ToArray()
    End Sub


    Private _Variables As Variable() = New Variable() {}
    Public Property Variables() As Variable()
        Get
            Return _Variables
        End Get
        Set(ByVal value As Variable())
            _Variables = value
        End Set
    End Property

    Private _Name As String
    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
        End Set
    End Property

    Private _ID As Integer
    Public Property ID() As Integer
        Get
            Return _ID
        End Get
        Set(ByVal value As Integer)
            _ID = value
        End Set
    End Property


End Class



<XmlType(TypeName:="ConfigurationValuePair")>
<Serializable>
Public Structure KeyValuePair(Of K, V)

    Public Sub New(ByVal __k As K, ByVal __v As V)
        Key = __k
        Value = __v
    End Sub

    Private _Key As K
    Public Property Key() As K
        Get
            Return _Key
        End Get
        Set(ByVal value As K)
            _Key = value
        End Set
    End Property


    Private _Value As V
    Public Property Value() As V
        Get
            Return _Value
        End Get
        Set(ByVal value As V)
            _Value = value
        End Set
    End Property

End Structure

Public Class Variable
    Private _name As String
    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Private _ConfigurationValuePairs As List(Of KeyValuePair(Of String, Object)) = New List(Of KeyValuePair(Of String, Object))()
    Public Property ConfigurationValuePairs() As List(Of KeyValuePair(Of String, Object))
        Get
            Return _ConfigurationValuePairs
        End Get
        Set(ByVal value As List(Of KeyValuePair(Of String, Object)))
            _ConfigurationValuePairs = value
        End Set
    End Property
End Class