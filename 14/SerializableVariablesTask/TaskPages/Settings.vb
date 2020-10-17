Imports EPDM.Interop.epdm


Public Enum SerializationType

    Json
    XML

End Enum
Public Class Settings

    Public Const serializationTypeKey As String = "SerializationType"
    Public Const saveLocationKey As String = "SaveLocation"

    Public Sub Load(ByRef poCmd As EdmCmd)

        Dim taskProperties As IEdmTaskProperties = poCmd.mpoExtra


        'load data from taskproperties 
        Dim serializationType As SerializationType

        Dim serializationTypeValEx As String = taskProperties.GetValEx(serializationTypeKey)

        Dim serializtionRet As Boolean
        serializtionRet = [Enum].TryParse(Of SerializationType)(serializationTypeValEx, True, serializationType)

        Dim saveLocation As String
        saveLocation = taskProperties.GetValEx(saveLocationKey)

        Select Case serializationType
            Case SerializationType.Json
                jsonRadioButton.Checked = True
                xmlRadioButton.Checked = False
                Exit Select
            Case SerializationType.XML
                jsonRadioButton.Checked = False
                xmlRadioButton.Checked = True
                Exit Select
        End Select

        saveLocationTextBox.Text = saveLocation

    End Sub

    Public Sub Store(ByRef poCmd As EdmCmd)

        Dim taskProperties As IEdmTaskProperties = poCmd.mpoExtra

        Dim serializationType As SerializationType
        Dim saveLocation As String

        saveLocation = saveLocationTextBox.Text

        If jsonRadioButton.Checked = True Then
            serializationType = SerializationType.Json
        Else
            serializationType = SerializationType.XML
        End If

        taskProperties.SetValEx(serializationTypeKey, serializationType.ToString())
        taskProperties.SetValEx(saveLocationKey, saveLocation)

    End Sub
End Class
