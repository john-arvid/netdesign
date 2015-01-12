<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormNDAskForSwitch
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Me.ButtonCancel = New System.Windows.Forms.Button()
        Me.LabelName = New System.Windows.Forms.Label()
        Me.LabelModel = New System.Windows.Forms.Label()
        Me.LabelPort = New System.Windows.Forms.Label()
        Me.LabelRow = New System.Windows.Forms.Label()
        Me.GroupBoxPorts = New System.Windows.Forms.GroupBox()
        Me.ComboBoxMedia = New System.Windows.Forms.ComboBox()
        Me.ComboBoxPurpose = New System.Windows.Forms.ComboBox()
        Me.LabelPurpose = New System.Windows.Forms.Label()
        Me.LabelMedia = New System.Windows.Forms.Label()
        Me.TextBoxName = New System.Windows.Forms.TextBox()
        Me.TextBoxModel = New System.Windows.Forms.TextBox()
        Me.TextBoxPort = New System.Windows.Forms.TextBox()
        Me.ButtonOk = New System.Windows.Forms.Button()
        Me.CheckBoxVertically = New System.Windows.Forms.CheckBox()
        Me.ComboBoxRow = New System.Windows.Forms.ComboBox()
        Me.GroupBoxPorts.SuspendLayout()
        Me.SuspendLayout()
        '
        'ButtonCancel
        '
        Me.ButtonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButtonCancel.Location = New System.Drawing.Point(270, 227)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(75, 23)
        Me.ButtonCancel.TabIndex = 10
        Me.ButtonCancel.Text = "Cancel"
        Me.ButtonCancel.UseVisualStyleBackColor = True
        '
        'LabelName
        '
        Me.LabelName.AutoSize = True
        Me.LabelName.Location = New System.Drawing.Point(12, 13)
        Me.LabelName.Name = "LabelName"
        Me.LabelName.Size = New System.Drawing.Size(48, 13)
        Me.LabelName.TabIndex = 2
        Me.LabelName.Text = "?? name"
        '
        'LabelModel
        '
        Me.LabelModel.AutoSize = True
        Me.LabelModel.Location = New System.Drawing.Point(12, 41)
        Me.LabelModel.Name = "LabelModel"
        Me.LabelModel.Size = New System.Drawing.Size(42, 13)
        Me.LabelModel.TabIndex = 3
        Me.LabelModel.Text = "?? type"
        '
        'LabelPort
        '
        Me.LabelPort.AutoSize = True
        Me.LabelPort.Location = New System.Drawing.Point(12, 82)
        Me.LabelPort.Name = "LabelPort"
        Me.LabelPort.Size = New System.Drawing.Size(82, 13)
        Me.LabelPort.TabIndex = 4
        Me.LabelPort.Text = "Number of ports"
        '
        'LabelRow
        '
        Me.LabelRow.AutoSize = True
        Me.LabelRow.Location = New System.Drawing.Point(184, 82)
        Me.LabelRow.Name = "LabelRow"
        Me.LabelRow.Size = New System.Drawing.Size(81, 13)
        Me.LabelRow.TabIndex = 5
        Me.LabelRow.Text = "Number of rows"
        '
        'GroupBoxPorts
        '
        Me.GroupBoxPorts.Controls.Add(Me.ComboBoxMedia)
        Me.GroupBoxPorts.Controls.Add(Me.ComboBoxPurpose)
        Me.GroupBoxPorts.Controls.Add(Me.LabelPurpose)
        Me.GroupBoxPorts.Controls.Add(Me.LabelMedia)
        Me.GroupBoxPorts.Location = New System.Drawing.Point(15, 123)
        Me.GroupBoxPorts.Name = "GroupBoxPorts"
        Me.GroupBoxPorts.Size = New System.Drawing.Size(330, 64)
        Me.GroupBoxPorts.TabIndex = 6
        Me.GroupBoxPorts.TabStop = False
        Me.GroupBoxPorts.Text = "Ports"
        '
        'ComboBoxMedia
        '
        Me.ComboBoxMedia.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxMedia.Items.AddRange(New Object() {"Copper", "Fiber"})
        Me.ComboBoxMedia.Location = New System.Drawing.Point(45, 29)
        Me.ComboBoxMedia.Name = "ComboBoxMedia"
        Me.ComboBoxMedia.Size = New System.Drawing.Size(61, 21)
        Me.ComboBoxMedia.Sorted = True
        Me.ComboBoxMedia.TabIndex = 3
        '
        'ComboBoxPurpose
        '
        Me.ComboBoxPurpose.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxPurpose.FormattingEnabled = True
        Me.ComboBoxPurpose.Items.AddRange(New Object() {"Control network", "Data network", "Data/Control network", "Management"})
        Me.ComboBoxPurpose.Location = New System.Drawing.Point(170, 29)
        Me.ComboBoxPurpose.Name = "ComboBoxPurpose"
        Me.ComboBoxPurpose.Size = New System.Drawing.Size(121, 21)
        Me.ComboBoxPurpose.Sorted = True
        Me.ComboBoxPurpose.TabIndex = 99
        '
        'LabelPurpose
        '
        Me.LabelPurpose.AutoSize = True
        Me.LabelPurpose.Location = New System.Drawing.Point(120, 29)
        Me.LabelPurpose.Name = "LabelPurpose"
        Me.LabelPurpose.Size = New System.Drawing.Size(46, 13)
        Me.LabelPurpose.TabIndex = 1
        Me.LabelPurpose.Text = "Purpose"
        '
        'LabelMedia
        '
        Me.LabelMedia.AutoSize = True
        Me.LabelMedia.Location = New System.Drawing.Point(0, 29)
        Me.LabelMedia.Name = "LabelMedia"
        Me.LabelMedia.Size = New System.Drawing.Size(36, 13)
        Me.LabelMedia.TabIndex = 0
        Me.LabelMedia.Text = "Media"
        '
        'TextBoxName
        '
        Me.TextBoxName.Location = New System.Drawing.Point(167, 13)
        Me.TextBoxName.Name = "TextBoxName"
        Me.TextBoxName.Size = New System.Drawing.Size(178, 20)
        Me.TextBoxName.TabIndex = 0
        Me.TextBoxName.Text = "Test"
        '
        'TextBoxModel
        '
        Me.TextBoxModel.Location = New System.Drawing.Point(167, 39)
        Me.TextBoxModel.Name = "TextBoxModel"
        Me.TextBoxModel.Size = New System.Drawing.Size(178, 20)
        Me.TextBoxModel.TabIndex = 2
        Me.TextBoxModel.Text = "test"
        '
        'TextBoxPort
        '
        Me.TextBoxPort.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.TextBoxPort.Location = New System.Drawing.Point(100, 82)
        Me.TextBoxPort.Name = "TextBoxPort"
        Me.TextBoxPort.Size = New System.Drawing.Size(65, 20)
        Me.TextBoxPort.TabIndex = 3
        Me.TextBoxPort.Text = "1"
        '
        'ButtonOk
        '
        Me.ButtonOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ButtonOk.Location = New System.Drawing.Point(30, 226)
        Me.ButtonOk.Name = "ButtonOk"
        Me.ButtonOk.Size = New System.Drawing.Size(118, 23)
        Me.ButtonOk.TabIndex = 13
        Me.ButtonOk.Text = "Create ??"
        Me.ButtonOk.UseVisualStyleBackColor = True
        '
        'CheckBoxVertically
        '
        Me.CheckBoxVertically.AutoSize = True
        Me.CheckBoxVertically.Location = New System.Drawing.Point(167, 204)
        Me.CheckBoxVertically.Name = "CheckBoxVertically"
        Me.CheckBoxVertically.Size = New System.Drawing.Size(98, 17)
        Me.CheckBoxVertically.TabIndex = 12
        Me.CheckBoxVertically.Text = "Orient vertically"
        Me.CheckBoxVertically.UseVisualStyleBackColor = True
        '
        'ComboBoxRow
        '
        Me.ComboBoxRow.FormattingEnabled = True
        Me.ComboBoxRow.Items.AddRange(New Object() {"1", "2"})
        Me.ComboBoxRow.Location = New System.Drawing.Point(270, 81)
        Me.ComboBoxRow.Name = "ComboBoxRow"
        Me.ComboBoxRow.Size = New System.Drawing.Size(49, 21)
        Me.ComboBoxRow.TabIndex = 14
        '
        'FormNDAskForSwitch
        '
        Me.AcceptButton = Me.ButtonOk
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.ButtonCancel
        Me.ClientSize = New System.Drawing.Size(354, 259)
        Me.Controls.Add(Me.ComboBoxRow)
        Me.Controls.Add(Me.ButtonOk)
        Me.Controls.Add(Me.CheckBoxVertically)
        Me.Controls.Add(Me.TextBoxPort)
        Me.Controls.Add(Me.TextBoxModel)
        Me.Controls.Add(Me.TextBoxName)
        Me.Controls.Add(Me.GroupBoxPorts)
        Me.Controls.Add(Me.LabelRow)
        Me.Controls.Add(Me.LabelPort)
        Me.Controls.Add(Me.LabelModel)
        Me.Controls.Add(Me.LabelName)
        Me.Controls.Add(Me.ButtonCancel)
        Me.MaximizeBox = False
        Me.Name = "FormNDAskForSwitch"
        Me.Text = "?? parameters"
        Me.GroupBoxPorts.ResumeLayout(False)
        Me.GroupBoxPorts.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
    Friend WithEvents LabelName As System.Windows.Forms.Label
    Friend WithEvents LabelModel As System.Windows.Forms.Label
    Friend WithEvents LabelPort As System.Windows.Forms.Label
    Friend WithEvents LabelRow As System.Windows.Forms.Label
    Friend WithEvents GroupBoxPorts As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBoxMedia As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxPurpose As System.Windows.Forms.ComboBox
    Friend WithEvents LabelPurpose As System.Windows.Forms.Label
    Friend WithEvents LabelMedia As System.Windows.Forms.Label
    Friend WithEvents TextBoxName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxModel As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxPort As System.Windows.Forms.TextBox
    Friend WithEvents ButtonOk As System.Windows.Forms.Button
    Friend WithEvents CheckBoxVertically As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBoxRow As System.Windows.Forms.ComboBox
End Class
