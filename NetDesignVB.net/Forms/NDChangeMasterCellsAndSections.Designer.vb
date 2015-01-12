<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NDChangeMasterCellsAndSections
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
        Me.LabelMasterName = New System.Windows.Forms.Label()
        Me.LabelCellName = New System.Windows.Forms.Label()
        Me.LabelCellValue = New System.Windows.Forms.Label()
        Me.LabelSection = New System.Windows.Forms.Label()
        Me.CheckBoxDelete = New System.Windows.Forms.CheckBox()
        Me.ComboBoxSection = New System.Windows.Forms.ComboBox()
        Me.TextBoxMasterName = New System.Windows.Forms.TextBox()
        Me.TextBoxCellName = New System.Windows.Forms.TextBox()
        Me.TextBoxCellValue = New System.Windows.Forms.TextBox()
        Me.ButtonOK = New System.Windows.Forms.Button()
        Me.ButtonCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'LabelMasterName
        '
        Me.LabelMasterName.AutoSize = True
        Me.LabelMasterName.Location = New System.Drawing.Point(13, 19)
        Me.LabelMasterName.Name = "LabelMasterName"
        Me.LabelMasterName.Size = New System.Drawing.Size(70, 13)
        Me.LabelMasterName.TabIndex = 0
        Me.LabelMasterName.Text = "Master Name"
        '
        'LabelCellName
        '
        Me.LabelCellName.AutoSize = True
        Me.LabelCellName.Location = New System.Drawing.Point(13, 42)
        Me.LabelCellName.Name = "LabelCellName"
        Me.LabelCellName.Size = New System.Drawing.Size(55, 13)
        Me.LabelCellName.TabIndex = 1
        Me.LabelCellName.Text = "Cell Name"
        '
        'LabelCellValue
        '
        Me.LabelCellValue.AutoSize = True
        Me.LabelCellValue.Location = New System.Drawing.Point(12, 69)
        Me.LabelCellValue.Name = "LabelCellValue"
        Me.LabelCellValue.Size = New System.Drawing.Size(54, 13)
        Me.LabelCellValue.TabIndex = 2
        Me.LabelCellValue.Text = "Cell Value"
        '
        'LabelSection
        '
        Me.LabelSection.AutoSize = True
        Me.LabelSection.Location = New System.Drawing.Point(12, 104)
        Me.LabelSection.Name = "LabelSection"
        Me.LabelSection.Size = New System.Drawing.Size(43, 13)
        Me.LabelSection.TabIndex = 3
        Me.LabelSection.Text = "Section"
        '
        'CheckBoxDelete
        '
        Me.CheckBoxDelete.AutoSize = True
        Me.CheckBoxDelete.Location = New System.Drawing.Point(19, 139)
        Me.CheckBoxDelete.Name = "CheckBoxDelete"
        Me.CheckBoxDelete.Size = New System.Drawing.Size(63, 17)
        Me.CheckBoxDelete.TabIndex = 5
        Me.CheckBoxDelete.Text = "Delete?"
        Me.CheckBoxDelete.UseVisualStyleBackColor = True
        '
        'ComboBoxSection
        '
        Me.ComboBoxSection.FormattingEnabled = True
        Me.ComboBoxSection.Items.AddRange(New Object() {"User-Defined Cells", "Shape Data", "Hyperlinks"})
        Me.ComboBoxSection.Location = New System.Drawing.Point(147, 104)
        Me.ComboBoxSection.Name = "ComboBoxSection"
        Me.ComboBoxSection.Size = New System.Drawing.Size(121, 21)
        Me.ComboBoxSection.TabIndex = 6
        '
        'TextBoxMasterName
        '
        Me.TextBoxMasterName.Location = New System.Drawing.Point(125, 19)
        Me.TextBoxMasterName.Name = "TextBoxMasterName"
        Me.TextBoxMasterName.Size = New System.Drawing.Size(143, 20)
        Me.TextBoxMasterName.TabIndex = 7
        Me.TextBoxMasterName.Text = "Wire Bundle.77"
        '
        'TextBoxCellName
        '
        Me.TextBoxCellName.Location = New System.Drawing.Point(125, 42)
        Me.TextBoxCellName.Name = "TextBoxCellName"
        Me.TextBoxCellName.Size = New System.Drawing.Size(143, 20)
        Me.TextBoxCellName.TabIndex = 8
        '
        'TextBoxCellValue
        '
        Me.TextBoxCellValue.Location = New System.Drawing.Point(125, 69)
        Me.TextBoxCellValue.Name = "TextBoxCellValue"
        Me.TextBoxCellValue.Size = New System.Drawing.Size(143, 20)
        Me.TextBoxCellValue.TabIndex = 9
        '
        'ButtonOK
        '
        Me.ButtonOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ButtonOK.Location = New System.Drawing.Point(19, 227)
        Me.ButtonOK.Name = "ButtonOK"
        Me.ButtonOK.Size = New System.Drawing.Size(75, 23)
        Me.ButtonOK.TabIndex = 10
        Me.ButtonOK.Text = "OK"
        Me.ButtonOK.UseVisualStyleBackColor = True
        '
        'ButtonCancel
        '
        Me.ButtonCancel.CausesValidation = False
        Me.ButtonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButtonCancel.Location = New System.Drawing.Point(171, 227)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(75, 23)
        Me.ButtonCancel.TabIndex = 11
        Me.ButtonCancel.Text = "Cancel"
        Me.ButtonCancel.UseVisualStyleBackColor = True
        '
        'NDChangeMasterCellsAndSections
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.ButtonOK)
        Me.Controls.Add(Me.TextBoxCellValue)
        Me.Controls.Add(Me.TextBoxCellName)
        Me.Controls.Add(Me.TextBoxMasterName)
        Me.Controls.Add(Me.ComboBoxSection)
        Me.Controls.Add(Me.CheckBoxDelete)
        Me.Controls.Add(Me.LabelSection)
        Me.Controls.Add(Me.LabelCellValue)
        Me.Controls.Add(Me.LabelCellName)
        Me.Controls.Add(Me.LabelMasterName)
        Me.Name = "NDChangeMasterCellsAndSections"
        Me.Text = "NDChangeMasterCellsAndSections"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LabelMasterName As System.Windows.Forms.Label
    Friend WithEvents LabelCellName As System.Windows.Forms.Label
    Friend WithEvents LabelCellValue As System.Windows.Forms.Label
    Friend WithEvents LabelSection As System.Windows.Forms.Label
    Friend WithEvents CheckBoxDelete As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBoxSection As System.Windows.Forms.ComboBox
    Friend WithEvents TextBoxMasterName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxCellName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxCellValue As System.Windows.Forms.TextBox
    Friend WithEvents ButtonOK As System.Windows.Forms.Button
    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
End Class
