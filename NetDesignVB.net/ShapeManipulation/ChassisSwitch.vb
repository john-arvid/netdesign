Module ChassisSwitch
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="chassisSwitchShape"></param>
    ''' <remarks></remarks>
    Public Sub HandleChassisSwitch(ByRef chassisSwitchShape As Visio.Shape)

        If CountShapesOnPageByName("Rack") < 1 Then
            MsgBox("You have to drop a rack on the page first")
            chassisSwitchShape.Delete()
            Exit Sub
        End If

        Dim ChassisSwitchForm As New NDAskForChassis

        ' Replaces ?? in the form
        If chassisSwitchShape.CellExists(_ShapeCategories, 0) Then
            changeNameInForm(chassisSwitchShape.Cells(_ShapeCategories).ResultStr(Visio.VisUnitCodes.visUnitsString), ChassisSwitchForm)
        End If

        ChassisSwitchForm.ShowDialog()

        ' If the user exits the form, delete the shape
        If ChassisSwitchForm.DialogResult = Windows.Forms.DialogResult.Cancel Then
            chassisSwitchShape.Delete()
        Else
            Call CreateChassisSwitch(chassisSwitchShape, ChassisSwitchForm)
        End If


        ChassisSwitchForm.Close()

        ChassisSwitchForm = Nothing

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="chassisSwitchShape"></param>
    ''' <param name="form"></param>
    ''' <remarks></remarks>
    Private Sub CreateChassisSwitch(ByRef chassisSwitchShape As Visio.Shape, form As NDAskForChassis)
        Dim Page As Visio.Page
        Dim CurrentPage As Visio.Page
        Dim HomeMaster As Visio.Master
        Dim NextPageMaster As Visio.Master
        Dim PreviousMaster As Visio.Master
        Dim HomeShape As Visio.Shape
        Dim NextPageShape As Visio.Shape
        Dim PreviousPageShape As Visio.Shape
        Dim BladeMaster As Visio.Master
        Dim BladeShape As Visio.Shape
        Dim ChassisPageMaster As Visio.Master
        Dim ChassisPage As Visio.Shape
        Dim LastChassisPage As Visio.Shape = Nothing
        Dim ChassisSwitchBase As Visio.Shape = Nothing
        Dim W As Double, H As Double, pinX As Double, pinY As Double, zeroX As Double, zeroY As Double, dy As Double

        CurrentPage = Globals.ThisAddIn.Application.ActivePage
        HomeMaster = Globals.ThisAddIn.Application.Documents.Item("NetdesignHidden.vssx").Masters.Item("Drill Up connector")
        NextPageMaster = Globals.ThisAddIn.Application.Documents.Item("NetdesignHidden.vssx").Masters.Item("Next Page")
        PreviousMaster = Globals.ThisAddIn.Application.Documents.Item("NetdesignHidden.vssx").Masters.Item("Prev Page")
        BladeMaster = Globals.ThisAddIn.Application.Documents.Item("Netdesign.vssx").Masters.Item("Blade")
        ChassisPageMaster = Globals.ThisAddIn.Application.Documents.Item("NetdesignHidden.vssx").Masters.Item("Chassis switch page")

        chassisSwitchShape.Cells(_ShapeName).Formula = """" + form.TextBoxName.Text + """"
        chassisSwitchShape.Cells(_ShapeModel).Formula = """" + form.TextBoxModel.Text + """"
        chassisSwitchShape.Cells("LockTextEdit").Formula = 0
        chassisSwitchShape.Text = form.TextBoxName.Text + " - " + form.TextBoxModel.Text
        chassisSwitchShape.Cells("LockTextEdit").Formula = 1

        W = 0.8 * chassisSwitchShape.Cells("Width").Result("")
        H = 0.15 * chassisSwitchShape.Cells("Height").Result("")
        pinX = chassisSwitchShape.Cells("PinX").Result("")
        pinY = chassisSwitchShape.Cells("PinY").Result("") - chassisSwitchShape.Cells("LocPinY").Result("") + 0.8 * chassisSwitchShape.Cells("Height").Result("")

        'Put the chassis page link in the chassis switch before adding the pages
        For i As Integer = 1 To form.TextBoxPages.Text
            'Find the last chassis page link
            For Each shape As Visio.Shape In chassisSwitchShape.Shapes
                If shape.Cells(_ShapeCategories).ResultStr(Visio.VisUnitCodes.visUnitsString) = "Chassis Switch Page" Then
                    LastChassisPage = shape
                Else
                    ChassisSwitchBase = shape
                End If
            Next

            'Set the position and size corresponding to the last chassis page
            If Not LastChassisPage Is Nothing Then
                zeroX = chassisSwitchShape.Cells("PinX").Result("") - chassisSwitchShape.Cells("LocPinX").Result("")
                zeroY = chassisSwitchShape.Cells("PinY").Result("") - chassisSwitchShape.Cells("LocPinY").Result("")

                W = LastChassisPage.Cells("Width").Result("")
                H = LastChassisPage.Cells("Height").Result("")
                pinX = zeroX + LastChassisPage.Cells("PinX").Result("")
                pinY = zeroY + LastChassisPage.Cells("PinY").Result("") - 1.2 * H

                If Not ChassisSwitchBase Is Nothing And pinY - H < zeroY Then
                    dy = zeroY - (pinY - H)
                    ChassisSwitchBase.Cells("Height").Formula = "=" & CStr(ChassisSwitchBase.Cells("Height").Result("") + dy)
                    ChassisSwitchBase.Cells("PinY").Formula = ChassisSwitchBase.Cells("PinY").Result("") - dy * 0.5
                End If

            End If

            ChassisPage = CurrentPage.Drop(ChassisPageMaster, pinX, pinY)

            ChassisPage.Cells("Width").Formula = "=" & CStr(W)
            ChassisPage.Cells("Height").Formula = "=" & CStr(H)
            ChassisPage.Cells("PinX").Formula = "=" & CStr(pinX)
            ChassisPage.Cells("PinY").Formula = "=" & CStr(pinY)
            ChassisPage.CellsSRC(Visio.VisSectionIndices.visSectionCharacter, 0, Visio.VisCellIndices.visCharacterSize).Formula = "=" & _
                CStr(ChassisSwitchBase.CellsSRC(Visio.VisSectionIndices.visSectionCharacter, 0, Visio.VisCellIndices.visCharacterSize).Result("")) & "*Height/" & CStr(H)

            ChassisPage.Text = "Page-" + i.ToString()
            ChassisPage.Cells("LockTextEdit").Formula = 1

            Call Globals.ThisAddIn.Application.ActiveWindow.Select(chassisSwitchShape, Visio.VisSelectArgs.visSelect)
            Call Globals.ThisAddIn.Application.ActiveWindow.Select(ChassisPage, Visio.VisSelectArgs.visSelect)
            Call Globals.ThisAddIn.Application.ActiveWindow.Selection.AddToGroup()

            'Remove the glue and the annoying quick connect
            ChassisPage.CellsU("GlueType").Formula = 8
        Next

        'Remove the glue and the annoying quick connect
        chassisSwitchShape.CellsU("GlueType").Formula = 8



        For i As Integer = 1 To form.TextBoxPages.Text
            Page = Globals.ThisAddIn.Application.ActiveDocument.Pages.Add()
            Page.Name = form.TextBoxName.Text + ":Page " + i.ToString()

            chassisSwitchShape.Shapes.Item(i + 1).Hyperlinks("OffPageConnector").SubAddress = Page.Name

            If Not Page.Shapes("ThePage").CellExists("User.IsPartOfChassisSwitch", False) Then
                Call Page.Shapes("ThePage").AddNamedRow(Visio.VisSectionIndices.visSectionUser, "IsPartOfChassisSwitch", 0)
            End If
            Page.Shapes("ThePage").Cells("User.IsPartOfChassisSwitch").FormulaForce = "=GUARD(1)"

            'Drop the navigation shapes on the current page
            HomeShape = Page.Drop(HomeMaster, 0, 255)
            NextPageShape = Page.Drop(NextPageMaster, 254, 0)
            PreviousPageShape = Page.Drop(PreviousMaster, 1, 0)



            HomeShape.Hyperlinks("OffPageConnector").SubAddress = CurrentPage.Name

            'Remove the shape text lock, force since it's guarded
            HomeShape.CellsU("LockTextEdit").FormulaForce = "GUARD(0)"
            HomeShape.Text = CurrentPage.Name
            HomeShape.CellsU("LockTextEdit").FormulaForce = "GUARD(1)"


            If form.CheckBoxVertically.Checked Then
                pinX = 0.05 * Page.Shapes("ThePage").Cells("PageWidth").Result("")
                pinY = 0.5 * Page.Shapes("ThePage").Cells("PageHeight").Result("")
            Else
                pinX = 0.5 * Page.Shapes("ThePage").Cells("PageWidth").Result("")
                pinY = 0.85 * Page.Shapes("ThePage").Cells("PageHeight").Result("")
            End If

            For j As Integer = 1 To form.TextBoxBlades.Text
                BladeShape = Page.Drop(BladeMaster, 5, 5)
                Call CreateSwitch(BladeShape, form.TextBoxName.Text, form.TextBoxBladeType.Text, form.TextBoxPort.Text, form.ComboBoxRow.Text, form.ComboBoxMedia.SelectedItem.ToString(), form.ComboBoxPurpose.SelectedItem.ToString(), form.CheckBoxVertically.Checked)
                If form.CheckBoxVertically.Checked Then
                    pinX = pinX + 1.2 * BladeShape.Cells("Width").Result("")
                Else
                    pinY = pinY - 1.2 * BladeShape.Cells("Height").Result("")
                End If

                BladeShape.Cells("PinX").Formula = "=" & CStr(pinX)
                BladeShape.Cells("PinY").Formula = "=" & CStr(pinY)

            Next
            If i <> form.TextBoxPages.Text Then
                'Remove the shape text lock, force since it's guarded
                NextPageShape.CellsU("LockTextEdit").FormulaForce = "GUARD(0)"
                NextPageShape.Hyperlinks("OffPageConnector").SubAddress = form.TextBoxName.Text + ":Page " + (i + 1).ToString()
                NextPageShape.Cells("User.TextTitle").Formula = """" + NextPageShape.Hyperlinks("OffPageConnector").SubAddress + """"
                NextPageShape.Text = NextPageShape.Cells("User.TextTitle").ResultStr(Visio.VisUnitCodes.visUnitsString)
                NextPageShape.CellsU("LockTextEdit").FormulaForce = "GUARD(1)"
            End If

            If i <> 1 Then
                'Remove the shape text lock, force since it's guarded
                PreviousPageShape.CellsU("LockTextEdit").FormulaForce = "GUARD(0)"
                PreviousPageShape.Hyperlinks("OffPageConnector").SubAddress = form.TextBoxName.Text + ":Page " + (i - 1).ToString()
                PreviousPageShape.Cells("User.TextTitle").Formula = """" + PreviousPageShape.Hyperlinks("OffPageConnector").SubAddress + """"
                PreviousPageShape.Text = PreviousPageShape.Cells("User.TextTitle").ResultStr(Visio.VisUnitCodes.visUnitsString)
                PreviousPageShape.CellsU("LockTextEdit").FormulaForce = "GUARD(1)"
            End If


        Next

    End Sub



End Module