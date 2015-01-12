Module Processor

    Public Sub HandleProcessor(ByRef processorShape As Visio.Shape)

        If CountShapesOnPageByName("Rack") < 1 Then
            MsgBox("You have to drop a rack on the page first")
            processorShape.Delete()
            Exit Sub
        End If

        Dim ProcessorForm As New FormNDAskForSwitch

        changeNameInForm(processorShape.Master.Name, ProcessorForm)

        ProcessorForm.AutoValidate = Windows.Forms.AutoValidate.Disable


        ProcessorForm.ShowDialog()

        If ProcessorForm.DialogResult = Windows.Forms.DialogResult.Cancel Then
            processorShape.Delete()
        Else
            Call CreateSwitch(processorShape, ProcessorForm.TextBoxName.Text, ProcessorForm.TextBoxModel.Text, _
                              ProcessorForm.TextBoxPort.Text, ProcessorForm.ComboBoxRow.Text, _
                              ProcessorForm.ComboBoxMedia.Text, ProcessorForm.ComboBoxPurpose.Text, _
                              ProcessorForm.CheckBoxVertically.Checked)
        End If

        
        ' Closing the form
        ProcessorForm.Close()
        ' And empties it
        ProcessorForm = Nothing

    End Sub


End Module