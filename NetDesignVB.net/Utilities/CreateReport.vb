

Module Report


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="page"></param>
    ''' <param name="allWires"></param>
    ''' <remarks></remarks>
    Public Sub CreateReport(ByRef page As Visio.Page, ByVal allWires As Boolean)

        Dim MyDocumentsPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim TextFile As String = MyDocumentsPath + "\NetDesignReport.txt"
        Dim ShapeInformation As StringBuilder = New StringBuilder()
        Dim Seperator As String = ","
        Dim Shape As Visio.Shape
        Dim FromShape As Visio.Shape
        Dim FromShapeId As VariantType

        If Not System.IO.File.Exists(TextFile) Then
            System.IO.File.Create(TextFile).Dispose()
            Call InitiateReport(ShapeInformation, Seperator)
        End If

        'For Each Shape In page.Shapes
        '    If Shape.CellExists("User.msvShapeCategories", 0) Then
        '        If Shape.Cells("User.msvShapeCategories").ResultStr("") = "Wire" AndAlso Shape.Cells("User.NotReport").ResultInt("", 1) = 0 Then
        '            Call GetInformation(Shape, ShapeInformation, Seperator, page)
        '        ElseIf Shape.Cells("User.msvShapeCategories").ResultStr("") = "OPC" AndAlso Shape.Cells("User.NotReport").ResultInt("", 1) = 0 Then
        '            If Shape.Cells("User.OPCType").ResultStr("") = "Wire Bundle" Then
        '                For i As Integer = 1 To 12
        '                    Call GetInformation(Shape.Shapes.Item(i), ShapeInformation, Seperator, page)
        '                Next
        '            End If
        '        End If
        '    End If
        'Next


        For Each Shape In page.Shapes
            If Shape.CellExists("User.msvShapeCategories", 0) Then
                If Shape.Cells("User.msvShapeCategories").ResultStr("") = "Wire" Then
                    If Shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll2D, "").Length > 1 Then
                        FromShapeId = Shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "")(0)
                        FromShape = page.Shapes.ItemFromID(FromShapeId)
                        If Not FromShape.Cells("User.msvShapeCategories").ResultStr("") = "OPC" Then
                            Call GetInformation(Shape, ShapeInformation, Seperator, page)
                        End If
                        'ElseIf check to include loose wires
                    End If
                End If
            End If
        Next

        Call WriteReportToFile(ShapeInformation, TextFile)

        MsgBox("Report has been created and saved: " + TextFile, MsgBoxStyle.OkOnly)

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="shape"></param>
    ''' <param name="shapeInformation"></param>
    ''' <param name="seperator"></param>
    ''' <param name="page"></param>
    ''' <remarks></remarks>
    Private Sub GetInformation(ByRef shape As Visio.Shape, ByRef shapeInformation As StringBuilder, ByVal seperator As String, ByRef page As Visio.Page)
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
        shapeInformation.Append(ToShape.Cells("User.RackLocation").ResultStr(""))
        shapeInformation.Append(seperator)
        shapeInformation.Append(ToShape.Cells("User.SwitchName").ResultStr(""))
        shapeInformation.Append(seperator)
        shapeInformation.Append(ToShape.Cells("User.UPosition").ResultStr(""))
        shapeInformation.Append(seperator)
        shapeInformation.Append(FromShape.Cells("User.RackLocation").ResultStr(""))
        shapeInformation.Append(seperator)
        shapeInformation.Append(FromShape.Cells("User.SwitchName").ResultStr(""))
        shapeInformation.Append(seperator)
        shapeInformation.Append(FromShape.Cells("User.UPosition").ResultStr(""))
        shapeInformation.Append(seperator)
        shapeInformation.Append(shape.Cells("Prop.Media").ResultStr(""))
        shapeInformation.Append(seperator)
        shapeInformation.Append(shape.Cells("Prop.WireID").ResultStr(""))
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