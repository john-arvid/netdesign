Module Blade

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="bladeShape"></param>
    ''' <remarks></remarks>
    Public Sub HandleBlade(ByRef bladeShape As Visio.Shape)
        Dim Page As Visio.Page = bladeShape.ContainingPage
        Dim Shape As Visio.Shape

        'This is true for a chassis blade, but not for a dropped blade
        If bladeShape.Parent IsNot Page Then
            Exit Sub
        End If

        'If the page that the shape has been dropped on is not a chassis page, delete the shape and exit
        If Not Page.Shapes("ThePage").CellExists("User.IsPartOfChassisSwitch", 0) OrElse Not Page.Shapes("ThePage").Cells("User.IsPartOfChassisSwitch").ResultInt("", 1) = "1" Then
            MsgBox("This shape can only be dropped on a chassis page!")
            bladeShape.Delete()
            Exit Sub
        End If

        'Get a user dialog same as the switch, telling the function that this is not a switch, therefor not checking if a rack is present on the page
        Call handleSwitch(bladeShape, False)

        For Each Shape In Page.Shapes
            If Shape.CellExists(_UPosition, 0) Then
                bladeShape.Cells(_UPosition).Formula = Shape.Cells(_UPosition).Formula
                Exit For
            End If
        Next

    End Sub


End Module