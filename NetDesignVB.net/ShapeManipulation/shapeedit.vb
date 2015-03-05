Module ShapeEdit
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="switchShape"></param>
    ''' <param name="cell"></param>
    ''' <remarks></remarks>
    Public Sub UpdateSwitch(ByRef switchShape As Visio.Shape, ByRef cell As Visio.Cell)

        Dim Shape As Visio.Shape
        Dim Page As Visio.Page = Globals.ThisAddIn.Application.ActivePage

        ' Update the text of the switch
        Call UpdateShapeName(switchShape, cell)

        ' Update all the wires connected to the switch
        For Each Shape In switchShape.Shapes
            If Shape.CellExists(_ShapeCategories, 0) Then
                If Shape.Cells(_ShapeCategories).ResultStr(Visio.VisUnitCodes.visUnitsString) = "Port" Then
                    If Shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll1D, "Wire").Length > 0 Then
                        Call UpdateWire(Page.Shapes.ItemFromID(Shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll1D, "Wire")(0)))
                    End If
                End If
            End If
        Next

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="shape"></param>
    ''' <param name="cell"></param>
    ''' <remarks></remarks>
    Public Sub UpdatePort(ByRef shape As Visio.Shape, ByRef cell As Visio.Cell)

        Dim WireShape As Visio.Shape
        Dim WireShapeId As Integer

        ' Update the port text
        shape.Cells("LockTextEdit").Formula = 0
        shape.Text = shape.Cells("User.TextTitle").ResultStr(Visio.VisUnitCodes.visUnitsString)
        shape.Cells("LockTextEdit").Formula = 1

        Call ValidatePort(shape)

        ' If port is connected to a wire, update the wire
        If shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll1D, "Wire").Length() > 0 Then
            WireShapeId = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll1D, "Wire")(0)
            WireShape = Globals.ThisAddIn.Application.ActivePage.Shapes.ItemFromID(WireShapeId)
            Call UpdateWire(WireShape)
        End If

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="wireShape"></param>
    ''' <param name="cell"></param>
    ''' <remarks></remarks>
    Public Sub UpdateWire(ByRef wireShape As Visio.Shape, Optional ByRef cell As Visio.Cell = Nothing)

        Dim OtherShape As Visio.Shape
        Dim OtherShapeId As Integer
        Dim Page As Visio.Page = Globals.ThisAddIn.Application.ActivePage

        ' Make sure the wire is in front of everything else
        wireShape.BringToFront()

        ' If the wire is connected to a port, synch the wire to get the port information
        If wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "Port").Length > 0 Then
            OtherShapeId = wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "Port")(0)
            OtherShape = Page.Shapes.ItemFromID(OtherShapeId)
            Call SynchWire(OtherShape, wireShape)
        ElseIf wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing2D, "Port").Length > 0 Then
            OtherShapeId = wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing2D, "Port")(0)
            OtherShape = Page.Shapes.ItemFromID(OtherShapeId)
            Call SynchWire(OtherShape, wireShape)
        End If

        ' If the wire is connected to an OPC, update the OPC so the OPC get's information from the wire
        If wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "OPC").Length > 0 Then
            OtherShapeId = wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "OPC")(0)
            OtherShape = Page.Shapes.ItemFromID(OtherShapeId)
            Call UpdateOPC(OtherShape, wireShape)
        ElseIf wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing2D, "OPC").Length > 0 Then
            OtherShapeId = wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing2D, "OPC")(0)
            OtherShape = Page.Shapes.ItemFromID(OtherShapeId)
            Call UpdateOPC(OtherShape, wireShape)
        End If

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="rackShape"></param>
    ''' <param name="cell"></param>
    ''' <remarks></remarks>
    Public Sub UpdateRack(ByRef rackShape As Visio.Shape, ByRef cell As Visio.Cell)


        ' If the rack location cell has been changed, update every wire on the page
        If cell Is rackShape.Cells("Prop.RackLocation") Then
            For Each shape As Visio.Shape In Globals.ThisAddIn.Application.ActivePage.Shapes
                If shape.CellExists(_ShapeCategories, 0) Then
                    If shape.Cells(_ShapeCategories).ResultStr(Visio.VisUnitCodes.visUnitsString) = "Wire" Then
                        Call UpdateWire(shape)
                    End If
                End If
            Next
        End If

        Call UpdateRackText(rackShape)

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="shape"></param>
    ''' <param name="cell"></param>
    ''' <remarks></remarks>
    Public Sub UpdateShapeName(ByRef shape As Visio.Shape, ByRef cell As Visio.Cell)

        Dim OldText As String
        Dim NewText As String


        OldText = shape.Text

        If cell Is shape.Cells(_ShapeName) Then
            NewText = cell.ResultStr(Visio.VisUnitCodes.visUnitsString) + " - " + shape.Cells(_ShapeModel).ResultStr(Visio.VisUnitCodes.visUnitsString)
        ElseIf cell Is shape.Cells(_ShapeModel) Then
            NewText = shape.Cells(_ShapeName).ResultStr(Visio.VisUnitCodes.visUnitsString) + " - " + cell.ResultStr(Visio.VisUnitCodes.visUnitsString)
        Else
            Exit Sub
        End If
        shape.Cells("LockTextEdit").Formula = 0
        If Not OldText = NewText Then
            shape.Text = NewText
        End If
        shape.Cells("LockTextEdit").Formula = 1

    End Sub


    ''' <summary>
    ''' Update the chassis data if a relevant cell has been updated
    ''' </summary>
    ''' <param name="shape">The chassis processor or switch shape</param>
    ''' <param name="cell">The cell that has been changed</param>
    ''' <remarks>Currently only for name and UPosition</remarks>
    Public Sub UpdateChassis(ByRef shape As Visio.Shape, ByRef cell As Visio.Cell)
        Dim ChassisPageShape As Visio.Shape
        Dim Page As Visio.Page
        Dim BladeShape As Visio.Shape

        'Update the text of the chassis
        Call UpdateShapeName(shape, cell)

        'If it is a different cell than the UPosition, then exit
        If cell IsNot shape.Cells(_UPosition) Then
            Exit Sub
        End If

        'Go through every page in the chassis
        For Each ChassisPageShape In shape.Shapes

            'If the chassispage has a hyperlink
            If ChassisPageShape.SectionExists(Visio.VisSectionIndices.visSectionHyperlink, 0) Then
                'Set the page to the one that the hyperlink points to
                Page = Globals.ThisAddIn.Application.ActiveDocument.Pages(ChassisPageShape.Hyperlinks("OffpageConnector").SubAddress)

                'Go through every blade on the current page in the chassis
                For Each BladeShape In Page.Shapes
                    If BladeShape.CellExists(_ShapeCategories, 0) AndAlso BladeShape.Cells(_ShapeCategories).ResultStr(Visio.VisUnitCodes.visUnitsString) = "Blade" Then

                        'Copy the newly updated UPosition from the chassis to the blade
                        BladeShape.Cells(_UPosition).Formula = shape.Cells(_UPosition).Formula
                    End If
                Next
            End If
        Next

    End Sub
End Module