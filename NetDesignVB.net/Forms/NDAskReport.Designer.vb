<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NDAskReport
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.CheckBoxWholeDocument = New System.Windows.Forms.CheckBox()
        Me.CheckedListBoxPages = New System.Windows.Forms.CheckedListBox()
        Me.CheckBoxConnectedWire = New System.Windows.Forms.CheckBox()
        Me.CheckBoxAllData = New System.Windows.Forms.CheckBox()
        Me.CheckedListBoxData = New System.Windows.Forms.CheckedListBox()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button1.Location = New System.Drawing.Point(30, 205)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(91, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Create Report"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.Location = New System.Drawing.Point(207, 205)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'CheckBoxWholeDocument
        '
        Me.CheckBoxWholeDocument.AutoSize = True
        Me.CheckBoxWholeDocument.Location = New System.Drawing.Point(12, 24)
        Me.CheckBoxWholeDocument.Name = "CheckBoxWholeDocument"
        Me.CheckBoxWholeDocument.Size = New System.Drawing.Size(109, 17)
        Me.CheckBoxWholeDocument.TabIndex = 2
        Me.CheckBoxWholeDocument.Text = "Whole Document"
        Me.CheckBoxWholeDocument.UseVisualStyleBackColor = True
        '
        'CheckedListBoxPages
        '
        Me.CheckedListBoxPages.FormattingEnabled = True
        Me.CheckedListBoxPages.Location = New System.Drawing.Point(12, 47)
        Me.CheckedListBoxPages.Name = "CheckedListBoxPages"
        Me.CheckedListBoxPages.Size = New System.Drawing.Size(239, 79)
        Me.CheckedListBoxPages.TabIndex = 3
        '
        'CheckBoxConnectedWire
        '
        Me.CheckBoxConnectedWire.AutoSize = True
        Me.CheckBoxConnectedWire.Location = New System.Drawing.Point(30, 132)
        Me.CheckBoxConnectedWire.Name = "CheckBoxConnectedWire"
        Me.CheckBoxConnectedWire.Size = New System.Drawing.Size(127, 17)
        Me.CheckBoxConnectedWire.TabIndex = 4
        Me.CheckBoxConnectedWire.Text = "Connected wires only"
        Me.CheckBoxConnectedWire.UseVisualStyleBackColor = True
        '
        'CheckBoxAllData
        '
        Me.CheckBoxAllData.AutoSize = True
        Me.CheckBoxAllData.Location = New System.Drawing.Point(291, 24)
        Me.CheckBoxAllData.Name = "CheckBoxAllData"
        Me.CheckBoxAllData.Size = New System.Drawing.Size(76, 17)
        Me.CheckBoxAllData.TabIndex = 5
        Me.CheckBoxAllData.Text = "Everything"
        Me.CheckBoxAllData.UseVisualStyleBackColor = True
        '
        'CheckedListBoxData
        '
        Me.CheckedListBoxData.FormattingEnabled = True
        Me.CheckedListBoxData.Items.AddRange(New Object() {"Switch Name", "Switch Type", "Switch Port Number", "Switch Port Type", "Switch Port Media", "Processor Name", "Processor Type", "Processor Port Number", "Processor Port Type", "Processor Port Media", "Wire ID", "Wire Length", "Wire Type", "Wire Media", "Port Type", "Port Media", "Rack Location", "U Position"})
        Me.CheckedListBoxData.Location = New System.Drawing.Point(291, 47)
        Me.CheckedListBoxData.Name = "CheckedListBoxData"
        Me.CheckedListBoxData.Size = New System.Drawing.Size(382, 79)
        Me.CheckedListBoxData.TabIndex = 6
        '
        'NDAskReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(685, 255)
        Me.Controls.Add(Me.CheckedListBoxData)
        Me.Controls.Add(Me.CheckBoxAllData)
        Me.Controls.Add(Me.CheckBoxConnectedWire)
        Me.Controls.Add(Me.CheckedListBoxPages)
        Me.Controls.Add(Me.CheckBoxWholeDocument)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "NDAskReport"
        Me.Text = "NDAskReport"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents CheckBoxWholeDocument As System.Windows.Forms.CheckBox
    Friend WithEvents CheckedListBoxPages As System.Windows.Forms.CheckedListBox
    Friend WithEvents CheckBoxConnectedWire As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxAllData As System.Windows.Forms.CheckBox
    Friend WithEvents CheckedListBoxData As System.Windows.Forms.CheckedListBox
End Class
