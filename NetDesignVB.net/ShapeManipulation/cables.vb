Module cables

    Public Sub handleWire(ByRef wireShape As Visio.Shape)



        '' This is just to show how to change a shapesheet cell that is not user
        ''created. This is not necessary since this is fixed in the stencil
        'shape.CellsU("ConFixedCode").Formula = 2



    End Sub

    Public Sub handleWireBundle(ByRef wireBundleShape As Visio.Shape)
        Static NoDialog As Integer = 0

        ' To avoid the dialog to open when the corresponding OPC is being dropped
        If NoDialog = 1 Then
            NoDialog = 0
            Exit Sub
        End If



    End Sub




End Module