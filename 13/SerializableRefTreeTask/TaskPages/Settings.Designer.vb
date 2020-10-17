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
        Me.saveLocationTxtBox = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.jsonRadioButton = New System.Windows.Forms.RadioButton()
        Me.xmlRadioButton = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'saveLocationTxtBox
        '
        Me.saveLocationTxtBox.Location = New System.Drawing.Point(23, 153)
        Me.saveLocationTxtBox.Name = "saveLocationTxtBox"
        Me.saveLocationTxtBox.Size = New System.Drawing.Size(425, 20)
        Me.saveLocationTxtBox.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 124)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Save location:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.xmlRadioButton)
        Me.GroupBox1.Controls.Add(Me.jsonRadioButton)
        Me.GroupBox1.Location = New System.Drawing.Point(23, 21)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(425, 91)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "GroupBox1"
        '
        'jsonRadioButton
        '
        Me.jsonRadioButton.AutoSize = True
        Me.jsonRadioButton.Location = New System.Drawing.Point(35, 35)
        Me.jsonRadioButton.Name = "jsonRadioButton"
        Me.jsonRadioButton.Size = New System.Drawing.Size(53, 17)
        Me.jsonRadioButton.TabIndex = 0
        Me.jsonRadioButton.TabStop = True
        Me.jsonRadioButton.Text = "JSON"
        Me.jsonRadioButton.UseVisualStyleBackColor = True
        '
        'xmlButton
        '
        Me.xmlRadioButton.AutoSize = True
        Me.xmlRadioButton.Location = New System.Drawing.Point(35, 58)
        Me.xmlRadioButton.Name = "xmlButton"
        Me.xmlRadioButton.Size = New System.Drawing.Size(47, 17)
        Me.xmlRadioButton.TabIndex = 1
        Me.xmlRadioButton.TabStop = True
        Me.xmlRadioButton.Text = "XML"
        Me.xmlRadioButton.UseVisualStyleBackColor = True
        '
        'Settings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.saveLocationTxtBox)
        Me.Name = "Settings"
        Me.Size = New System.Drawing.Size(491, 247)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents saveLocationTxtBox As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents xmlRadioButton As Windows.Forms.RadioButton
    Friend WithEvents jsonRadioButton As Windows.Forms.RadioButton
End Class
