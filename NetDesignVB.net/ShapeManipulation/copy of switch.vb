Module switchtest
    ''' <summary>
    ''' Opens a form that request information from the user about the dropped
    ''' switch. Formats the form and sets default values. Calls the createswitch
    ''' function.
    ''' </summary>
    ''' <param name="switchShape">The switch shape</param>
    ''' <remarks></remarks>
    Public Sub handleSwitchTest(ByRef switchShape As Visio.Shape)

        If CountShapesOnPageByName("Rack") < 1 Then
            MsgBox("You have to drop a rack on the page first")
            switchShape.Delete()
            Exit Sub
        End If

        Dim SwitchForm As New FormNDAskForSwitch
        ' Replaces ?? in the form
        changeNameInForm(switchShape.Master.Name, SwitchForm)

        ' Validates only when I says so
        'SwitchForm.AutoValidate = Windows.Forms.AutoValidate.Disable

        ' Opens the form to the user
        SwitchForm.ShowDialog()

        ' Changes the data of the switch to the users input
        Call CreateSwitch(switchShape, SwitchForm)


        ' If the user exits the form, delete the shape
        If SwitchForm.DialogResult = Windows.Forms.DialogResult.Cancel Then
            switchShape.Delete()
        End If

        ' Closing the form
        SwitchForm.Close()
        ' And empties it
        SwitchForm = Nothing

        If Not IsUniqueName(switchShape.Text) Then
            MsgBox("This shape has not a unique name!")
            switchShape.Delete()
        End If

    End Sub
    ''' <summary>
    ''' Updates the switch with the user input, calls the AddPortToSwitch 
    ''' function, and alters the switch size according to the number of ports.
    ''' Orients the switch horizontally if checked, groups all the shapes 
    ''' together if checked
    ''' </summary>
    ''' <param name="switchShape">The switch shape</param>
    ''' <param name="form">The form with user input</param>
    ''' <remarks></remarks>
    Private Sub CreateSwitch(ByRef switchShape As Visio.Shape, ByVal form As FormNDAskForSwitch)
        Dim SwitchName As String = form.TextBoxName.Text
        Dim SwitchType As String = form.TextBoxModel.Text

        'TODO why .value?
        ' Set the switch name and Model from the user input
        switchShape.Cells("Prop.Name.Value").Formula = """" + SwitchName + """"
        switchShape.Cells("Prop.Model.Value").Formula = """" + SwitchType + """"
        switchShape.Text = SwitchName + " - " + SwitchType
        switchShape.Cells("Prop.NumberOfPorts").Formula = form.TextBoxPort.Text
        switchShape.Cells("User.Type").Formula = """" + switchShape.Master.Name + """"
        switchShape.Cells("User.MediaType").Formula = """" + form.ComboBoxMedia.Text + """"
        switchShape.Cells("User.MediaPurpose").Formula = """" + form.ComboBoxPurpose.Text + """"

        'Call ChangeSwitchSize(switchShape, form.TextBoxPort.Text, form)

    End Sub

    Private Sub ChangeSwitchSize(ByRef switchShape As Visio.Shape, ByVal numberOfPorts As Integer, ByVal form As FormNDAskForSwitch)

        Dim PortSize As Integer
        Dim PositionY As Integer

        ' Just a check for numbers of rows to be less or qual to number of ports
        '  If numberOfRows > numberOfPorts Then numberOfRows = numberOfPorts

        ' Calculate how many port there should be on one row
        '  NumberOfPortsPerRow = Int(numberOfPorts / numberOfRows + 0.999)
        ' Calculate the size of the port
        PortSize = switchShape.Cells("Width").Result("") / (2 * numberOfPorts + 2)

        ' Find the position of the switch
        PositionY = switchShape.Cells("PinY").Result("") - switchShape.Cells("LocPinY").Result("")

        ' Alter the switch size to fit all the ports
        switchShape.Cells("Height").Formula = "=" + CStr(PortSize * (2 * form.ComboBoxRow.Text + 2))
        switchShape.Cells("PinY").Formula = "=" + CStr(PositionY + 0.5 * switchShape.Cells("Height").Result(""))


    End Sub

End Module