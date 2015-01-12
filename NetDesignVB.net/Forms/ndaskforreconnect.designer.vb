<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NDAskForReconnect
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
        Me.RadioButtonNewPage = New System.Windows.Forms.RadioButton()
        Me.RadioButtonExistingPage = New System.Windows.Forms.RadioButton()
        Me.ComboBoxExistingPage = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBoxNewPageName = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.TextBoxFileName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CheckBoxODC = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'RadioButtonNewPage
        '
        Me.RadioButtonNewPage.AutoSize = True
        Me.RadioButtonNewPage.Location = New System.Drawing.Point(10, 78)
        Me.RadioButtonNewPage.Name = "RadioButtonNewPage"
        Me.RadioButtonNewPage.Size = New System.Drawing.Size(75, 17)
        Me.RadioButtonNewPage.TabIndex = 0
        Me.RadioButtonNewPage.TabStop = True
        Me.RadioButtonNewPage.Text = "New Page"
        Me.RadioButtonNewPage.UseVisualStyleBackColor = True
        '
        'RadioButtonExistingPage
        '
        Me.RadioButtonExistingPage.AutoSize = True
        Me.RadioButtonExistingPage.Location = New System.Drawing.Point(10, 137)
        Me.RadioButtonExistingPage.Name = "RadioButtonExistingPage"
        Me.RadioButtonExistingPage.Size = New System.Drawing.Size(89, 17)
        Me.RadioButtonExistingPage.TabIndex = 1
        Me.RadioButtonExistingPage.TabStop = True
        Me.RadioButtonExistingPage.Text = "Existing Page"
        Me.RadioButtonExistingPage.UseVisualStyleBackColor = True
        '
        'ComboBoxExistingPage
        '
        Me.ComboBoxExistingPage.FormattingEnabled = True
        Me.ComboBoxExistingPage.Location = New System.Drawing.Point(86, 158)
        Me.ComboBoxExistingPage.Name = "ComboBoxExistingPage"
        Me.ComboBoxExistingPage.Size = New System.Drawing.Size(121, 21)
        Me.ComboBoxExistingPage.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 98)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(91, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "New Page Name:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 161)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(68, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Select Page:"
        '
        'TextBoxNewPageName
        '
        Me.TextBoxNewPageName.Location = New System.Drawing.Point(107, 95)
        Me.TextBoxNewPageName.Name = "TextBoxNewPageName"
        Me.TextBoxNewPageName.Size = New System.Drawing.Size(100, 20)
        Me.TextBoxNewPageName.TabIndex = 5
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button1.Location = New System.Drawing.Point(13, 220)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Ok"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.Location = New System.Drawing.Point(135, 220)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 7
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(145, 31)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(63, 23)
        Me.Button3.TabIndex = 15
        Me.Button3.Text = "Open File"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'TextBoxFileName
        '
        Me.TextBoxFileName.Location = New System.Drawing.Point(39, 33)
        Me.TextBoxFileName.Name = "TextBoxFileName"
        Me.TextBoxFileName.Size = New System.Drawing.Size(100, 20)
        Me.TextBoxFileName.TabIndex = 14
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 36)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(23, 13)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "File"
        '
        'CheckBoxODC
        '
        Me.CheckBoxODC.AutoSize = True
        Me.CheckBoxODC.Location = New System.Drawing.Point(12, 10)
        Me.CheckBoxODC.Name = "CheckBoxODC"
        Me.CheckBoxODC.Size = New System.Drawing.Size(104, 17)
        Me.CheckBoxODC.TabIndex = 16
        Me.CheckBoxODC.Text = "Other Document"
        Me.CheckBoxODC.UseVisualStyleBackColor = True
        '
        'NDAskForReconnect
        '
        Me.AcceptButton = Me.Button1
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button2
        Me.ClientSize = New System.Drawing.Size(219, 256)
        Me.Controls.Add(Me.CheckBoxODC)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.TextBoxFileName)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBoxNewPageName)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBoxExistingPage)
        Me.Controls.Add(Me.RadioButtonExistingPage)
        Me.Controls.Add(Me.RadioButtonNewPage)
        Me.Name = "NDAskForReconnect"
        Me.Text = "OPC connect to"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents RadioButtonNewPage As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonExistingPage As System.Windows.Forms.RadioButton
    Friend WithEvents ComboBoxExistingPage As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBoxNewPageName As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents TextBoxFileName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CheckBoxODC As System.Windows.Forms.CheckBox
End Class
