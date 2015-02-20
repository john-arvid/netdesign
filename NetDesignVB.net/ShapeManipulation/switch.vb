Module switch
    ''' <summary>
    ''' Opens a form that request information from the user about the dropped
    ''' switch. Formats the form and sets default values. Calls the createswitch
    ''' function.
    ''' </summary>
    ''' <param name="switchShape">The switch shape</param>
    ''' <remarks></remarks>
    Public Sub handleSwitch(ByRef switchShape As Visio.Shape)

        If CountShapesOnPageByName("Rack") < 1 Then
            MsgBox("You have to drop a rack on the page first")
            switchShape.Delete()
            Exit Sub
        End If

        Dim SwitchForm As New FormNDAskForSwitch
        ' Replaces ?? in the form
        If switchShape.CellExists(_ShapeCategories, 0) Then
            changeNameInForm(switchShape.Cells(_ShapeCategories).ResultStr(""), SwitchForm)
        End If

        ' Validates only when I says so
        SwitchForm.AutoValidate = Windows.Forms.AutoValidate.Disable


        ' Opens the form to the user
        SwitchForm.ShowDialog()

        ' If the user exits the form, delete the shape
        If SwitchForm.DialogResult = Windows.Forms.DialogResult.Cancel Then
            switchShape.Delete()
        Else
            ' Changes the data of the switch to the users input
            Call CreateSwitch(switchShape, SwitchForm.TextBoxName.Text, SwitchForm.TextBoxModel.Text, _
                              SwitchForm.TextBoxPort.Text, SwitchForm.ComboBoxRow.Text, _
                              SwitchForm.ComboBoxMedia.Text, SwitchForm.ComboBoxPurpose.Text, _
                              SwitchForm.CheckBoxVertically.Checked)



        End If

        ' Closing the form
        SwitchForm.Close()
        ' And empties it
        SwitchForm = Nothing

    End Sub
    
    ''' <summary>
    ''' Updates the switch with the user input, calls the AddPortToSwitch 
    ''' function, and alters the switch size according to the number of ports.
    ''' Orients the switch horizontally if checked, groups all the shapes 
    ''' together if checked
    ''' </summary>
    ''' <param name="switchShape">The shape that was dropped</param>
    ''' <param name="switchName">The user inputed name</param>
    ''' <param name="switchType">The selected switch type</param>
    ''' <param name="portNumber">How many ports</param>
    ''' <param name="rowNumber">How many rows</param>
    ''' <param name="mediaType">What kind of media type</param>
    ''' <param name="mediaPurpose">What purpose of media</param>
    ''' <param name="vertical">Vertical or not</param>
    ''' <param name="document">In what document</param>
    ''' <param name="page">On what page</param>
    ''' <remarks></remarks>
    Public Sub CreateSwitch(ByRef switchShape As Visio.Shape, ByVal switchName As String, _
                            ByVal switchType As String, ByVal portNumber As String, _
                            ByVal rowNumber As String, ByVal mediaType As String, _
                            ByVal mediaPurpose As String, ByVal vertical As Boolean, _
                            Optional ByRef document As Visio.Document = Nothing, Optional ByRef page As Visio.Page = Nothing)
        Dim PortLists As New List(Of Visio.Shape)
        Dim RackShape As Visio.Shape
        Dim SwitchShapeParent As Visio.Shape

        RackShape = GetRackShape()
        If Not RackShape Is Nothing Then
            switchShape.Cells(_RackLocation).Formula = "=" + RackShape.Name + "!" + _RackLocation
        End If

        If document Is Nothing Then
            document = Globals.ThisAddIn.Application.ActiveDocument
        End If

        If page Is Nothing Then
            page = Globals.ThisAddIn.Application.ActivePage
        End If

        'Globals.ThisAddIn.Application.Documents.OpenEx("\\cern.ch\dfs\Users\j\jkibsgaa\Documents\Drawing1.vsdx", Visio.VisOpenSaveArgs.visOpenRW)
        Globals.ThisAddIn.Application.Windows.ItemEx(document.Name).Activate()
        Globals.ThisAddIn.Application.ActiveWindow.Page = document.Pages.ItemU(page.NameU)

        ' Set the switch name and Model from the user input
        switchShape.Cells("Prop.Name").Formula = """" + switchName + """"
        switchShape.Cells("Prop.Model").Formula = """" + switchType + """"
        switchShape.Cells("LockTextEdit").Formula = 0
        switchShape.Text = switchName + " - " + switchType
        switchShape.Cells("LockTextEdit").Formula = 1
        switchShape.Cells("Prop.NumberOfPorts").Formula = "GUARD(""" + portNumber + """)"

        ' Adds ports to the switch
        Call AddPortToSwitch(switchShape, portNumber, rowNumber, mediaType, mediaPurpose, PortLists)

        ' Rotates and/or groups the switch and ports

        ' Ensure that none is selected
        Call Globals.ThisAddIn.Application.ActiveWindow.DeselectAll()
        ' Select the switch shape
        Call Globals.ThisAddIn.Application.ActiveWindow.Select(switchShape, 2)

        ' Select all the ports that has been added
        For Each p In PortLists
            Call Globals.ThisAddIn.Application.ActiveWindow.Select(p, 2)
        Next

        ' Rotates all selected shapes
        If vertical Then
            Call Globals.ThisAddIn.Application.DoCmd(Visio.VisUICmds.visCmdObjectRotate90)
            Dim sinA As Double, cosA As Double, M As Double, N As Double
            sinA = Math.Sin(Math.PI / 2)
            cosA = Math.Cos(Math.PI / 2)
            M = switchShape.Cells("PinX").Result("")
            N = switchShape.Cells("PinY").Result("")

            ' Alters the port size
            For Each p In PortLists
                Dim x As Double, y As Double, x_tag As Double, y_tag As Double
                x = p.Cells("PinX").Result("")
                y = p.Cells("PinY").Result("")
                x_tag = x * cosA - y * sinA - M * cosA + N * sinA + M
                y_tag = x * sinA + y * cosA - M * sinA - N * cosA + N
                p.Cells("PinX").Formula = "=" + CStr(x_tag)
                p.Cells("PinY").Formula = "=" + CStr(y_tag)
            Next
        End If



        ' Groups all the shapes selected

        Call Globals.ThisAddIn.Application.DoCmd(Visio.VisUICmds.visCmdObjectGroup)
        SwitchShapeParent = switchShape.Parent
        'SwitchShapeParent.CellsSRC(Visio.VisSectionIndices.visSectionCharacter, 0, Visio.VisCellIndices.visCharacterSize).Formula = "=28*MIN(Height,Width)/(2*72)"


        ' Copy shape sheet data from the switch to the grouped container
        Call CopyShapeSheetData(switchShape, switchShape.Parent, Visio.VisSectionIndices.visSectionProp)
        Call CopyShapeSheetData(switchShape, switchShape.Parent, Visio.VisSectionIndices.visSectionUser)

        'switchShape.DeleteSection(Visio.VisSectionIndices.visSectionUser)
        'switchShape.DeleteSection(Visio.VisSectionIndices.visSectionProp)

        If Not (SwitchShapeParent.RowExists(Visio.VisSectionIndices.visSectionObject, _
                Visio.VisRowIndices.visRowTextXForm, _
                Visio.VisExistsFlags.visExistsAnywhere)) Then

            Call SwitchShapeParent.AddRow(Visio.VisSectionIndices.visSectionObject, _
            Visio.VisRowIndices.visRowTextXForm, _
            Visio.VisRowTags.visTagDefault)
        End If

        '// Set  the text transform formulas:


        If vertical Then
            SwitchShapeParent.CellsU("TxtAngle").FormulaForceU = "90 deg"
            SwitchShapeParent.CellsU("TxtPinY").FormulaForceU = "Height*0.5"
            SwitchShapeParent.CellsU("TxtPinX").FormulaForceU = "Width*0.2"
            SwitchShapeParent.CellsU("TxtWidth").FormulaForceU = "Height"
            SwitchShapeParent.CellsU("TxtHeight").FormulaForceU = "Height*0.4"
        Else
            SwitchShapeParent.CellsU("TxtPinY").FormulaForceU = "Height*0.8"
            SwitchShapeParent.CellsU("TxtHeight").FormulaForceU = "Height*0.4"
        End If

        Call UpdatePortWithGroup(PortLists, SwitchShapeParent)

        ' Set the switchshape to be the grouping, e.g the parent
        switchShape = SwitchShapeParent

        'Remove the glue and the annoying quick connect
        switchShape.CellsU("GlueType").Formula = 8


        'Run the garbage collector to empty the memory for "shapes"
        GC.Collect()
    End Sub
    ''' <summary>
    ''' Adds the requested number of ports to the switch, adjusts the size of
    ''' the ports according to how many ports there is. Sets the date of the port.
    ''' Updates a list that then contains all the newly added ports. 
    ''' </summary>
    ''' <param name="switchShape">The shape that the ports are added to</param>
    ''' <param name="numberOfPorts">How many ports should be added</param>
    ''' <param name="numberOfRows">How many rows the ports will be devided on</param>
    ''' <param name="typeOfPort">What type of port is added</param>
    ''' <param name="purposeOfPort">What purpose the port has</param>
    ''' <param name="portList">A list of all the ports that has been added</param>
    ''' <remarks></remarks>
    Private Sub AddPortToSwitch(ByRef switchShape As Visio.Shape, _
                                ByVal numberOfPorts As Integer, _
                                ByVal numberOfRows As Integer, _
                                ByVal typeOfPort As String, _
                                ByVal purposeOfPort As String, _
                                ByRef portList As List(Of Visio.Shape))

        Dim PortMaster As Visio.Master
        Dim newport As Visio.Shape
        Dim NumberOfPortsPerRow As Integer
        Dim PortSize As Double
        Dim PositionX As Double
        Dim PositionY As Double
        Dim PointX As Double
        Dim PointY As Double
        Dim Count As Integer = 1

        ' Get the stencil master
        PortMaster = Globals.ThisAddIn.Application.Documents.Item("NetdesignHidden.vssx").Masters.Item("Port")

        ' Just a check for numbers of rows to be less or qual to number of ports
        If numberOfRows > numberOfPorts Then numberOfRows = numberOfPorts

        ' Calculate how many port there should be on one row
        NumberOfPortsPerRow = Int(numberOfPorts / numberOfRows + 0.999)
        ' Calculate the size of the port
        PortSize = switchShape.Cells("Width").Result("") / (2 * NumberOfPortsPerRow + 2)

        ' Find the position of the switch
        PositionX = switchShape.Cells("PinX").Result("") - switchShape.Cells("LocPinX").Result("")
        PositionY = switchShape.Cells("PinY").Result("") - switchShape.Cells("LocPinY").Result("")

        ' Go through all the rows
        For i As Integer = 0 To numberOfRows - 1
            ' Find the position for the first port
            PointX = PositionX + (1.5 + (i Mod 2)) * PortSize
            PointY = PositionY + (2 * (numberOfRows - i) - 0.5) * PortSize

            ' Go throug all the ports for that row
            For j As Integer = 1 To Math.Min(NumberOfPortsPerRow, numberOfPorts - Count + 1)
                ' Drop the port on the page, needs to be assigned or it won't show
                newport = Globals.ThisAddIn.Application.ActivePage.Drop(PortMaster, 0, 0)
                ' Set the position of the port
                newport.Cells("PinX").Formula = PointX
                newport.Cells("PinY").Formula = PointY

                ' Set the size of the port
                newport.Cells("Width").Formula = CStr(PortSize)
                newport.Cells("Height").Formula = "=Width"

                ' Set the data of the port
                newport.Cells(_MediaType).Formula = """" + typeOfPort + """"
                newport.Cells(_Purpose).Formula = """" + purposeOfPort + """"
                newport.Cells(_PortNumber).Formula = """" + CStr(Count) + """"

                ' Count to the next port
                Count = Count + 1

                ' Change the position for the next port
                PointX = PointX + 2 * PortSize

                ' Add the port to the list
                portList.Add(newport)
            Next
        Next

        ' Alter the switch size to fit all the ports
        switchShape.Cells("Height").Formula = "=" + CStr(PortSize * (2 * numberOfRows + 2))
        switchShape.Cells("PinY").Formula = "=" + CStr(PositionY + 0.5 * switchShape.Cells("Height").Result(""))

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="portList"></param>
    ''' <param name="switchParent"></param>
    ''' <remarks></remarks>
    Private Sub UpdatePortWithGroup(ByRef portList As List(Of Visio.Shape), ByRef switchParent As Visio.Shape)
        Dim Port As Visio.Shape

        ' Set a reference to the switchparent cell's, this is done with the !
        For Each Port In portList
            Port.Cells(_RackLocation).Formula = """=" + switchParent.Name + "!" + _RackLocation + """"
            Port.Cells(_SwitchName).Formula = """=" + switchParent.Name + "!Prop.Name" + """"
            Port.Cells("User.UPosition").Formula = """=" + switchParent.Name + "!Prop.UPosition" + """"
            Port.Cells("User.SwitchType").Formula = """=" + switchParent.Name + "!" + _ShapeCategories + """"
        Next

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="switchShape"></param>
    ''' <remarks></remarks>
    Public Sub DeletePointingOPC(ByRef switchShape As Visio.Shape)

        Dim Page As Visio.Page = Globals.ThisAddIn.Application.ActivePage
        Dim Shape As Visio.Shape
        Dim Counter As Integer = Page.Shapes.Count()

        For i As Integer = Counter To 1 Step -1
            Shape = Page.Shapes.Item(i)
            If Shape.CellExists(_ShapeCategories, 0) Then
                If Shape.Cells(_ShapeCategories).ResultStr("") = "OPC" Then
                    If Shape.Cells(_SwitchName).ResultStr("") = switchShape.Cells(_ShapeName).ResultStr("") Then
                        Shape.Delete()
                    End If
                End If
            End If
        Next

    End Sub

End Module