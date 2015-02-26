Module Rack


    Public Sub HandleRack(ByRef rack As Visio.Shape)

        'Check if valid drop of rack
        If Not (ValidateRack(rack)) Then
            Exit Sub
        End If

        Call UpdateRackText(rack)
        
    End Sub

    Public Sub UpdateRackText(ByRef rack As Visio.Shape)

        Dim ChildShape As Visio.Shape

        ChildShape = rack.Shapes(1)

        'Remove the shape text lock, force since it's guarded
        ChildShape.CellsU("LockTextEdit").FormulaForce = "GUARD(0)"

        ChildShape.Text = rack.Cells("Prop.RackName").ResultStr(Visio.VisUnitCodes.visUnitsString) + " - " + rack.Cells("Prop.RackLocation").ResultStr(Visio.VisUnitCodes.visUnitsString)



        'Add the shape text lock, force since it's guarded
        ChildShape.CellsU("LockTextEdit").FormulaForce = "GUARD(1)"


    End Sub
End Module