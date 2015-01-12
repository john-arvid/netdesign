<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NDAskForChassis
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ButtonAccept = New System.Windows.Forms.Button()
        Me.ButtonCancel = New System.Windows.Forms.Button()
        Me.TextBoxName = New System.Windows.Forms.TextBox()
        Me.TextBoxModel = New System.Windows.Forms.TextBox()
        Me.TextBoxPages = New System.Windows.Forms.TextBox()
        Me.TextBoxBlades = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TextBoxBladeType = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBoxPort = New System.Windows.Forms.TextBox()
        Me.GroupBoxPorts = New System.Windows.Forms.GroupBox()
        Me.ComboBoxMedia = New System.Windows.Forms.ComboBox()
        Me.ComboBoxPurpose = New System.Windows.Forms.ComboBox()
        Me.LabelPurpose = New System.Windows.Forms.Label()
        Me.LabelMedia = New System.Windows.Forms.Label()
        Me.LabelRow = New System.Windows.Forms.Label()
        Me.LabelPort = New System.Windows.Forms.Label()
        Me.CheckBoxVertically = New System.Windows.Forms.CheckBox()
        Me.ComboBoxRow = New System.Windows.Forms.ComboBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBoxPorts.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "?? name"
        '
        'ButtonAccept
        '
        Me.ButtonAccept.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ButtonAccept.Location = New System.Drawing.Point(10, 380)
        Me.ButtonAccept.Name = "ButtonAccept"
        Me.ButtonAccept.Size = New System.Drawing.Size(189, 23)
        Me.ButtonAccept.TabIndex = 2
        Me.ButtonAccept.Text = "Create ??"
        Me.ButtonAccept.UseVisualStyleBackColor = True
        '
        'ButtonCancel
        '
        Me.ButtonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButtonCancel.Location = New System.Drawing.Point(289, 380)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(85, 23)
        Me.ButtonCancel.TabIndex = 3
        Me.ButtonCancel.Text = "Cancel"
        Me.ButtonCancel.UseVisualStyleBackColor = True
        '
        'TextBoxName
        '
        Me.TextBoxName.Location = New System.Drawing.Point(242, 10)
        Me.TextBoxName.Name = "TextBoxName"
        Me.TextBoxName.Size = New System.Drawing.Size(132, 20)
        Me.TextBoxName.TabIndex = 4
        Me.TextBoxName.Text = "Test"
        '
        'TextBoxModel
        '
        Me.TextBoxModel.Location = New System.Drawing.Point(242, 39)
        Me.TextBoxModel.Name = "TextBoxModel"
        Me.TextBoxModel.Size = New System.Drawing.Size(132, 20)
        Me.TextBoxModel.TabIndex = 5
        Me.TextBoxModel.Text = "Test"
        '
        'TextBoxPages
        '
        Me.TextBoxPages.Location = New System.Drawing.Point(321, 83)
        Me.TextBoxPages.Name = "TextBoxPages"
        Me.TextBoxPages.Size = New System.Drawing.Size(53, 20)
        Me.TextBoxPages.TabIndex = 6
        Me.TextBoxPages.Text = "1"
        '
        'TextBoxBlades
        '
        Me.TextBoxBlades.Location = New System.Drawing.Point(321, 112)
        Me.TextBoxBlades.Name = "TextBoxBlades"
        Me.TextBoxBlades.Size = New System.Drawing.Size(53, 20)
        Me.TextBoxBlades.TabIndex = 7
        Me.TextBoxBlades.Text = "1"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(50, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "?? model"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 86)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(198, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Number of pages associated with this ??"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 115)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(135, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Number of blades per page"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ComboBoxRow)
        Me.GroupBox1.Controls.Add(Me.TextBoxBladeType)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.TextBoxPort)
        Me.GroupBox1.Controls.Add(Me.GroupBoxPorts)
        Me.GroupBox1.Controls.Add(Me.LabelRow)
        Me.GroupBox1.Controls.Add(Me.LabelPort)
        Me.GroupBox1.Location = New System.Drawing.Point(15, 142)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(359, 175)
        Me.GroupBox1.TabIndex = 11
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Define blade"
        '
        'TextBoxBladeType
        '
        Me.TextBoxBladeType.Location = New System.Drawing.Point(141, 59)
        Me.TextBoxBladeType.Name = "TextBoxBladeType"
        Me.TextBoxBladeType.Size = New System.Drawing.Size(100, 20)
        Me.TextBoxBladeType.TabIndex = 17
        Me.TextBoxBladeType.Text = "Blade"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(56, 59)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(61, 13)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "Blade Type"
        '
        'TextBoxPort
        '
        Me.TextBoxPort.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.TextBoxPort.Location = New System.Drawing.Point(141, 21)
        Me.TextBoxPort.Name = "TextBoxPort"
        Me.TextBoxPort.Size = New System.Drawing.Size(32, 20)
        Me.TextBoxPort.TabIndex = 10
        Me.TextBoxPort.Text = "1"
        '
        'GroupBoxPorts
        '
        Me.GroupBoxPorts.Controls.Add(Me.ComboBoxMedia)
        Me.GroupBoxPorts.Controls.Add(Me.ComboBoxPurpose)
        Me.GroupBoxPorts.Controls.Add(Me.LabelPurpose)
        Me.GroupBoxPorts.Controls.Add(Me.LabelMedia)
        Me.GroupBoxPorts.Location = New System.Drawing.Point(5, 105)
        Me.GroupBoxPorts.Name = "GroupBoxPorts"
        Me.GroupBoxPorts.Size = New System.Drawing.Size(348, 64)
        Me.GroupBoxPorts.TabIndex = 15
        Me.GroupBoxPorts.TabStop = False
        Me.GroupBoxPorts.Text = "Ports"
        '
        'ComboBoxMedia
        '
        Me.ComboBoxMedia.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxMedia.Items.AddRange(New Object() {"Copper", "Fiber"})
        Me.ComboBoxMedia.Location = New System.Drawing.Point(54, 29)
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
        Me.ComboBoxPurpose.Location = New System.Drawing.Point(221, 29)
        Me.ComboBoxPurpose.Name = "ComboBoxPurpose"
        Me.ComboBoxPurpose.Size = New System.Drawing.Size(121, 21)
        Me.ComboBoxPurpose.Sorted = True
        Me.ComboBoxPurpose.TabIndex = 99
        '
        'LabelPurpose
        '
        Me.LabelPurpose.AutoSize = True
        Me.LabelPurpose.Location = New System.Drawing.Point(144, 29)
        Me.LabelPurpose.Name = "LabelPurpose"
        Me.LabelPurpose.Size = New System.Drawing.Size(46, 13)
        Me.LabelPurpose.TabIndex = 1
        Me.LabelPurpose.Text = "Purpose"
        '
        'LabelMedia
        '
        Me.LabelMedia.AutoSize = True
        Me.LabelMedia.Location = New System.Drawing.Point(12, 29)
        Me.LabelMedia.Name = "LabelMedia"
        Me.LabelMedia.Size = New System.Drawing.Size(36, 13)
        Me.LabelMedia.TabIndex = 0
        Me.LabelMedia.Text = "Media"
        '
        'LabelRow
        '
        Me.LabelRow.AutoSize = True
        Me.LabelRow.Location = New System.Drawing.Point(179, 21)
        Me.LabelRow.Name = "LabelRow"
        Me.LabelRow.Size = New System.Drawing.Size(81, 13)
        Me.LabelRow.TabIndex = 14
        Me.LabelRow.Text = "Number of rows"
        '
        'LabelPort
        '
        Me.LabelPort.AutoSize = True
        Me.LabelPort.Location = New System.Drawing.Point(53, 21)
        Me.LabelPort.Name = "LabelPort"
        Me.LabelPort.Size = New System.Drawing.Size(82, 13)
        Me.LabelPort.TabIndex = 13
        Me.LabelPort.Text = "Number of ports"
        '
        'CheckBoxVertically
        '
        Me.CheckBoxVertically.AutoSize = True
        Me.CheckBoxVertically.Location = New System.Drawing.Point(134, 346)
        Me.CheckBoxVertically.Name = "CheckBoxVertically"
        Me.CheckBoxVertically.Size = New System.Drawing.Size(99, 17)
        Me.CheckBoxVertically.TabIndex = 12
        Me.CheckBoxVertically.Text = "Orient Vertically"
        Me.CheckBoxVertically.UseVisualStyleBackColor = True
        '
        'ComboBoxRow
        '
        Me.ComboBoxRow.FormattingEnabled = True
        Me.ComboBoxRow.Items.AddRange(New Object() {"1", "2"})
        Me.ComboBoxRow.Location = New System.Drawing.Point(266, 20)
        Me.ComboBoxRow.Name = "ComboBoxRow"
        Me.ComboBoxRow.Size = New System.Drawing.Size(46, 21)
        Me.ComboBoxRow.TabIndex = 18
        '
        'NDAskForChassis
        '
        Me.AcceptButton = Me.ButtonAccept
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.ButtonCancel
        Me.ClientSize = New System.Drawing.Size(386, 415)
        Me.Controls.Add(Me.CheckBoxVertically)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBoxBlades)
        Me.Controls.Add(Me.TextBoxPages)
        Me.Controls.Add(Me.TextBoxModel)
        Me.Controls.Add(Me.TextBoxName)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.ButtonAccept)
        Me.Controls.Add(Me.Label1)
        Me.Name = "NDAskForChassis"
        Me.Text = "Chassis ?? parameters"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBoxPorts.ResumeLayout(False)
        Me.GroupBoxPorts.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ButtonAccept As System.Windows.Forms.Button
    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
    Friend WithEvents TextBoxName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxModel As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxPages As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxBlades As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBoxPort As System.Windows.Forms.TextBox
    Friend WithEvents GroupBoxPorts As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBoxMedia As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxPurpose As System.Windows.Forms.ComboBox
    Friend WithEvents LabelPurpose As System.Windows.Forms.Label
    Friend WithEvents LabelMedia As System.Windows.Forms.Label
    Friend WithEvents LabelRow As System.Windows.Forms.Label
    Friend WithEvents LabelPort As System.Windows.Forms.Label
    Friend WithEvents TextBoxBladeType As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents CheckBoxVertically As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBoxRow As System.Windows.Forms.ComboBox
End Class
