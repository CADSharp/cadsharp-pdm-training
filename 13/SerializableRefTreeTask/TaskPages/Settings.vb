Imports EPDM.Interop.epdm

Public Class Settings

    Public Const serializationTypeKey As String = "SerializationType"
    Public Const saveLocationKey As String = "SaveLocation"

    Public Sub Load(ByRef poCmd As EdmCmd)

        Dim taskProperties As IEdmTaskProperties = poCmd.mpoExtra

        'load data from taskproperties 
        Dim serializationType As SerializationType
        Dim serializationTypeValEx As String = taskProperties.GetValEx(serializationTypeKey)
        Dim saveLocation As String


        saveLocation = taskProperties.GetValEx(saveLocationKey)
        Dim seriazationRet As Boolean = [Enum].TryParse(Of SerializationType)(serializationTypeValEx, True, serializationType)


        saveLocationTxtBox.Text = saveLocation

        Select Case serializationType
            Case SerializationType.XML
                jsonRadioButton.Checked = False
                xmlRadioButton.Checked = True
                Exit Select
            Case SerializationType.Json
                jsonRadioButton.Checked = True
                xmlRadioButton.Checked = false 
                Exit Select
        End Select

    End Sub

    Public Sub Store(ByRef poCmd As EdmCmd)
        'store data from task properties

        Dim taskProperties As IEdmTaskProperties = poCmd.mpoExtra


        Dim serializationType As SerializationType
        Dim saveLocation As String


        saveLocation = saveLocationTxtBox.Text

        If jsonRadioButton.Checked = True Then
            serializationType = SerializationType.Json
        Else
            serializationType = SerializationType.XML
        End If

        taskProperties.SetValEx(saveLocationKey, saveLocation)
        taskProperties.SetValEx(serializationTypeKey, serializationType.ToString())
    End Sub

End Class
