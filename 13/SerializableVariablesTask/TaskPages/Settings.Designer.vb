<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Settings
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.jsonRadioButton = New System.Windows.Forms.RadioButton()
        Me.xmlRadioButton = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.saveLocationTextBox = New System.Windows.Forms.TextBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.xmlRadioButton)
        Me.GroupBox1.Controls.Add(Me.jsonRadioButton)
        Me.GroupBox1.Location = New System.Drawing.Point(32, 22)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(447, 160)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Serialization type"
        '
        'jsonRadioButton
        '
        Me.jsonRadioButton.AutoSize = True
        Me.jsonRadioButton.Location = New System.Drawing.Point(47, 45)
        Me.jsonRadioButton.Name = "jsonRadioButton"
        Me.jsonRadioButton.Size = New System.Drawing.Size(47, 17)
        Me.jsonRadioButton.TabIndex = 0
        Me.jsonRadioButton.TabStop = True
        Me.jsonRadioButton.Text = "Json"
        Me.jsonRadioButton.UseVisualStyleBackColor = True
        '
        'xmlRadioButton
        '
        Me.xmlRadioButton.AutoSize = True
        Me.xmlRadioButton.Location = New System.Drawing.Point(47, 77)
        Me.xmlRadioButton.Name = "xmlRadioButton"
        Me.xmlRadioButton.Size = New System.Drawing.Size(47, 17)
        Me.xmlRadioButton.TabIndex = 1
        Me.xmlRadioButton.TabStop = True
        Me.xmlRadioButton.Text = "XML"
        Me.xmlRadioButton.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.saveLocationTextBox)
        Me.GroupBox2.Location = New System.Drawing.Point(32, 188)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(447, 124)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Save location:"
        '
        'saveLocationTextBox
        '
        Me.saveLocationTextBox.Location = New System.Drawing.Point(47, 44)
        Me.saveLocationTextBox.Name = "saveLocationTextBox"
        Me.saveLocationTextBox.Size = New System.Drawing.Size(356, 20)
        Me.saveLocationTextBox.TabIndex = 0
        '
        'Settings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Settings"
        Me.Size = New System.Drawing.Size(515, 340)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents xmlRadioButton As Windows.Forms.RadioButton
    Friend WithEvents jsonRadioButton As Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents saveLocationTextBox As Windows.Forms.TextBox
End Class
