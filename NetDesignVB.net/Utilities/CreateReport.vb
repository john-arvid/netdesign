

Module Report


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="page"></param>
    ''' <param name="allWires"></param>
    ''' <remarks></remarks>
    Public Sub CreateReport(ByRef page As Visio.Page, ByVal allWires As Boolean, Optional ByRef checkedItems As Windows.Forms.CheckedListBox.CheckedItemCollection = Nothing)

        Dim MyDocumentsPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim TextFile As String = MyDocumentsPath + "\NetDesignReport.txt"
        Dim ShapeInformation As StringBuilder = New StringBuilder()
        Dim Seperator As String = ","
        Dim Shape As Visio.Shape
        Dim FromShape As Visio.Shape
        Dim FromShapeId As VariantType

        If Not System.IO.File.Exists(TextFile) Then
            System.IO.File.Create(TextFile).Dispose()
        End If

        Call InitiateReport(ShapeInformation, Seperator)

        For Each Shape In page.Shapes
            If Shape.CellExists(_ShapeCategories, 0) Then
                If Shape.Cells(_ShapeCategories).ResultStr(Visio.VisUnitCodes.visUnitsString) = "Wire" Then
                    If Shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll2D, "").Length > 1 Then
                        FromShapeId = Shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "")(0)
                        FromShape = page.Shapes.ItemFromID(FromShapeId)
                        If Not FromShape.Cells(_ShapeCategories).ResultStr(Visio.VisUnitCodes.visUnitsString) = "OPC" Then
                            Call GetInformation(Shape, ShapeInformation, Seperator, page, checkedItems)
                        End If
                        'ElseIf check to include loose wires
                    End If
                End If
            End If
        Next

        Call WriteReportToFile(ShapeInformation, TextFile)



    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="shape"></param>
    ''' <param name="shapeInformation"></param>
    ''' <param name="seperator"></param>
    ''' <param name="page"></param>
    ''' <remarks></remarks>
    Private Sub GetInformation(ByRef shape As Visio.Shape, ByRef shapeInformation As StringBuilder, ByVal seperator As String, ByRef page As Visio.Page, ByRef checkedItems As Windows.Forms.CheckedListBox.CheckedItemCollection)
        Dim ToShape As Visio.Shape
        Dim ToShapeId As Integer
        Dim FromShape As Visio.Shape
        Dim FromShapeId As Integer

        ' Get the shapes that are connected to the wire
        ToShapeId = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing2D, "")(0)
        ToShape = page.Shapes.ItemFromID(ToShapeId)
        FromShapeId = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "")(0)
        FromShape = page.Shapes.ItemFromID(FromShapeId)

        shapeInformation.AppendLine()

        'Add information about the first shape
        For Each item In checkedItems
            Select Case item.ToString()
                Case "Rack Location"
                    shapeInformation.Append(ToShape.Cells(_RackLocation).ResultStr(Visio.VisUnitCodes.visUnitsString))
                Case "Switch Name"
                    shapeInformation.Append(ToShape.Cells(_SwitchName).ResultStr(Visio.VisUnitCodes.visUnitsString))
                Case "Switch Type"

                Case "Switch Port Number"

                Case "Switch Port Type"

                Case "Switch Port Media"

                Case "Processor Name"

                Case "Processor Type"

                Case "Processor Port Number"

                Case "Processor Port Type2"

                Case "Processor Port Media"

                Case "Wire ID"
                    shapeInformation.Append(shape.Cells(_WireID).ResultStr(Visio.VisUnitCodes.visUnitsString))
                Case "Wire Length"

                Case "Wire Type"

                Case "Wire Media"
                    shapeInformation.Append(shape.Cells(_MediaType).ResultStr(Visio.VisUnitCodes.visUnitsString))
                Case "Port Type"

                Case "Port Media"

                Case "U Position"
                    shapeInformation.Append(ToShape.Cells("User.UPosition").ResultStr(Visio.VisUnitCodes.visUnitsString))
            End Select
            shapeInformation.Append(seperator)
        Next

        'Add information about the second shape
        For Each item In checkedItems
            Select Case item.ToString()
                Case "Rack Location"
                    shapeInformation.Append(FromShape.Cells(_RackLocation).ResultStr(Visio.VisUnitCodes.visUnitsString))
                Case "Switch Name"
                    shapeInformation.Append(FromShape.Cells(_SwitchName).ResultStr(Visio.VisUnitCodes.visUnitsString))
                Case "Switch Type"

                Case "Switch Port Number"

                Case "Switch Port Type"

                Case "Switch Port Media"

                Case "Processor Name"

                Case "Processor Type"

                Case "Processor Port Number"

                Case "Processor Port Type2"

                Case "Processor Port Media"

                Case "Wire ID"
                    shapeInformation.Append(shape.Cells(_WireID).ResultStr(Visio.VisUnitCodes.visUnitsString))
                Case "Wire Length"

                Case "Wire Type"

                Case "Wire Media"
                    shapeInformation.Append(shape.Cells(_MediaType).ResultStr(Visio.VisUnitCodes.visUnitsString))
                Case "Port Type"

                Case "Port Media"

                Case "U Position"
                    shapeInformation.Append(FromShape.Cells("User.UPosition").ResultStr(Visio.VisUnitCodes.visUnitsString))
            End Select
            shapeInformation.Append(seperator)
        Next

        'Remove the last seperator
        shapeInformation.Remove(shapeInformation.Length - 1, 1)

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="textString"></param>
    ''' <param name="seperator"></param>
    ''' <remarks></remarks>
    Private Sub InitiateReport(ByRef textString As StringBuilder, ByVal seperator As String)

        textString.Append("Rack Name Destination")
        textString.Append(seperator)
        textString.Append("Switch Name Destination")
        textString.Append(seperator)
        textString.Append("U Position Destination")
        textString.Append(seperator)
        textString.Append("Rack Name Source")
        textString.Append(seperator)
        textString.Append("Switch Name Source")
        textString.Append(seperator)
        textString.Append("U Position Source")
        textString.Append(seperator)
        textString.Append("Wire Type")
        textString.Append(seperator)
        textString.Append("Wire ID")


        textString.AppendLine()

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="shapeInformation"></param>
    ''' <param name="textFile"></param>
    ''' <remarks></remarks>
    Private Sub WriteReportToFile(ByRef shapeInformation As StringBuilder, ByVal textFile As String)

        Using Outfile As System.IO.StreamWriter = New System.IO.StreamWriter(textFile, False)
            Outfile.Write(shapeInformation.ToString())
        End Using

    End Sub


End Module