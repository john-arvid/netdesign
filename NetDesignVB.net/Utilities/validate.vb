Module Validate

    ''' <summary>
    ''' Goes through every shape in the document, checks rules and gives the
    ''' user feedback
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub validateRules(ByVal ruleSet As Visio.ValidationRuleSet)

        Dim Pages As Visio.Pages
        Dim Page As Visio.Page
        Dim Shape As Visio.Shape
        Pages = Globals.ThisAddIn.Application.ActiveDocument.Pages


        For Each Page In Pages
            For Each Shape In Page.Shapes
                If Not Shape.Master Is Nothing Then
                    Select Case Shape.Master.Name
                        Case "Wire"
                            Call validateWire(Shape, ruleSet, Page)

                        Case "Switch"

                        Case "Processor"

                        Case "Rack"

                    End Select
                End If
            Next
        Next


    End Sub

    ''' <summary>
    ''' Validates the a wire. Responds to the user with a message when error
    ''' </summary>
    ''' <param name="wireShape">The current wire shape</param>
    ''' <remarks></remarks>
    Private Sub validateWire(ByRef wireShape As Visio.Shape, ByVal ruleSet As Visio.ValidationRuleSet, ByVal page As Visio.Page)

        ' Exit the sub if the wire does not have two connections
        If Not wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll2D, "").Length = 2 Then
            Exit Sub
        End If

        Dim ToShape As Visio.Shape
        Dim FromShape As Visio.Shape
        Dim FromShapeId As VariantType
        Dim ToShapeId As VariantType
        Dim ToShapeParent As Visio.Shape = Nothing
        Dim FromShapeParent As Visio.Shape = Nothing


        ' Set the shape variables, uses the id of the glued/connected shape
        FromShapeId = wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "")(0)
        FromShape = page.Shapes.ItemFromID(FromShapeId)
        ToShapeId = wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing2D, "")(0)
        ToShape = page.Shapes.ItemFromID(ToShapeId)

        Try
            ToShapeParent = ToShape.Parent
            FromShapeParent = FromShape.Parent
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
            System.Diagnostics.Debug.WriteLine("The shape had not a parent")
        End Try

        ' Using the cell user.type instead of master since 
        'If FromShape.Cells("User.Type").ResultStr("") = "Processor Test" And ToShape.Cells("User.Type").ResultStr("") = "Switch Test" Then
        '    'issue!
        '    Call AddIssue("Pro-Switch", ruleSet, page, wireShape)
        'End If

        If FromShape.Master.Name = "Processor Test" And ToShape.Master.Name = "Switch Test" Then
            'issue!
            Call AddIssue("Pro-Switch", ruleSet, page, wireShape)
        End If

        If FromShape.Master.Name = "Processor Test" And FromShape.Master.Name = "Processor Test" Then
            'issue!
            Call AddIssue("Pro-Pro", ruleSet, page, wireShape)
        End If

        If Not (FromShapeParent Is Nothing AndAlso ToShapeParent Is Nothing) Then
            If FromShapeParent Is ToShapeParent Then
                Call AddIssue("Wireloop", ruleSet, page, wireShape)
            End If
        End If
        

        If FromShape.Cells(_MediaType).ResultStr("") <> wireShape.Cells(_MediaType).ResultStr("") Then
            Call AddIssue("MediaType", ruleSet, page, wireShape)
        End If

        If ToShape.Cells(_MediaType).ResultStr("") <> wireShape.Cells(_MediaType).ResultStr("") Then
            Call AddIssue("MediaType", ruleSet, page, wireShape)
        End If

        If FromShape.Cells(_TransmissionSpeed).ResultStr("") <> wireShape.Cells(_TransmissionSpeed).ResultStr("") Then
            Call AddIssue("MediaSpeed", ruleSet, page, wireShape)
        End If

        If ToShape.Cells(_TransmissionSpeed).ResultStr("") <> wireShape.Cells(_TransmissionSpeed).ResultStr("") Then
            Call AddIssue("MediaSpeed", ruleSet, page, wireShape)
        End If

        If FromShape.Cells("Prop.Purpose").ResultStr("") <> wireShape.Cells("Prop.Purpose").ResultStr("") Then
            Call AddIssue("MediaPurpose", ruleSet, page, wireShape)
        End If

        If ToShape.Cells("Prop.Purpose").ResultStr("") <> wireShape.Cells("Prop.Purpose").ResultStr("") Then
            Call AddIssue("MediaPurpose", ruleSet, page, wireShape)
        End If

        If ToShape.Master.Name = "OPC" And FromShape.Master.Name = "OPC" Then
            Call AddIssue("OPCToOPC", ruleSet, page, wireShape)
        End If



    End Sub


    Public Sub ValidateWireConnection(ByRef connections As Visio.Connects)


        Dim ToShape As Visio.Shape = Nothing
        Dim wireShape As Visio.Shape = Nothing
        Dim OtherShape As Visio.Shape = Nothing
        Dim FromShape As Visio.Shape = Nothing
        Dim FromShapeId As VariantType
        Dim OtherShapeId As VariantType
        Dim Page As Visio.Page = Globals.ThisAddIn.Application.ActivePage

        'Set the shapes from the connection
        wireShape = connections.FromSheet
        ToShape = connections.ToSheet

        'If connected to an OPC
        If ToShape.Cells(_ShapeCategories).ResultStr("") = "OPC" Then
            Exit Sub
        End If

        ' Set the shape variables, uses the id of the glued/connected shape
        If wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "").Length > 0 Then
            FromShapeId = wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "")(0)
            FromShape = Page.Shapes.ItemFromID(FromShapeId)
        End If

        If wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing2D, "").Length > 0 Then
            OtherShapeId = wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing2D, "")(0)
            OtherShape = Page.Shapes.ItemFromID(OtherShapeId)
        End If



        'If the wire is connected in both ends
        If wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll2D, "").Length > 1 Then



            If FromShape.Cells("User.SwitchType").ResultStr("") = "Processor" Then
                If OtherShape.Cells("User.SwitchType").ResultStr("") = "Processor" Then
                    MsgBox("Processor connected to Processor is deprecated, think about it.")
                ElseIf OtherShape.Cells("User.SwitchType").ResultStr("") = "Switch" Then
                    MsgBox("Hierarchy problem, Processor can not be source when connected to switch")
                    wireShape.Disconnect(Visio.VisConnectorEnds.visConnectorBeginpoint, 2, 2, Visio.VisUnitCodes.visCentimeters)
                ElseIf OtherShape.Cells("User.SwitchType").ResultStr("") = "Blade" Then
                    MsgBox("Hierarchy problem, Processor can not be source when connected to router")
                    wireShape.Disconnect(Visio.VisConnectorEnds.visConnectorBeginpoint, 2, 2, Visio.VisUnitCodes.visCentimeters)
                End If
            ElseIf FromShape.Cells("User.SwitchType").ResultStr("") = "Switch" AndAlso OtherShape.Cells("User.SwitchType").ResultStr("") = "Blade" Then
                MsgBox("Hierarchy problem, switch can not be source when connected to router")
                wireShape.Disconnect(Visio.VisConnectorEnds.visConnectorBeginpoint, 2, 2, Visio.VisUnitCodes.visCentimeters)
            End If

        End If









        If wireShape.Cells(_MediaType).ResultStr("") <> ToShape.Cells(_MediaType).ResultStr("") Then
            MsgBox("Media type not same, change and reconnect")
            Try
                wireShape.Disconnect(Visio.VisConnectorEnds.visConnectorBeginpoint, 2, 2, Visio.VisUnitCodes.visCentimeters)
            Catch ex As Exception
                ' MsgBox(ex.Message)
            End Try

            Try
                wireShape.Disconnect(Visio.VisConnectorEnds.visConnectorEndPoint, 2, 2, Visio.VisUnitCodes.visCentimeters)
            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try

        ElseIf wireShape.Cells(_TransmissionSpeed).ResultStr("") <> ToShape.Cells(_TransmissionSpeed).ResultStr("") Then
            MsgBox("Transmission speed is not the same, change and reconnect")
            Try
                wireShape.Disconnect(Visio.VisConnectorEnds.visConnectorBeginpoint, 2, 2, Visio.VisUnitCodes.visCentimeters)
            Catch ex As Exception
                ' MsgBox(ex.Message)
            End Try

            Try
                wireShape.Disconnect(Visio.VisConnectorEnds.visConnectorEndPoint, 2, 2, Visio.VisUnitCodes.visCentimeters)
            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try

        ElseIf wireShape.Cells("Prop.Purpose").ResultStr("") <> ToShape.Cells("Prop.Purpose").ResultStr("") Then
            MsgBox("Media purpose is not the same, change and reconnect")
            Try
                wireShape.Disconnect(Visio.VisConnectorEnds.visConnectorBeginpoint, 2, 2, Visio.VisUnitCodes.visCentimeters)
            Catch ex As Exception
                ' MsgBox(ex.Message)
            End Try

            Try
                wireShape.Disconnect(Visio.VisConnectorEnds.visConnectorEndPoint, 2, 2, Visio.VisUnitCodes.visCentimeters)
            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try
        End If






        'If wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "").Length > 0 Then
        '    wireShape = wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "")(0)
        '    MsgBox("Blah")
        '    wireShape.Disconnect(Visio.VisConnectorEnds.visConnectorBothEnds, 2, 2, Visio.VisUnitCodes.visCentimeters)
        'End If

        'If wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing2D, "").Length > 0 Then
        '    ToShape = wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing2D, "")(0)

        '    MsgBox("aha")
        '    ToShape.Disconnect(Visio.VisConnectorEnds.visConnectorBothEnds, 2, 2, Visio.VisUnitCodes.visCentimeters)
        'End If

        'If Not wireShape Is Nothing Then
        '    If wireShape.Cells(MediaType).Formula <> wireShape.Cells(MediaType).Formula Then
        '        wireShape.Disconnect(Visio.VisConnectorEnds.visConnectorBothEnds, 2, 2, Visio.VisUnitCodes.visCentimeters)
        '    End If
        'End If

        'If Not ToShape Is Nothing Then

        'End If

    End Sub
    ''' <summary>
    ''' Check if there is any racks on the page
    ''' </summary>
    ''' <param name="rackShape">The dropped rack shape</param>
    ''' <remarks></remarks>
    Public Sub ValidateRack(ByRef rackShape As Visio.Shape)

        If CountShapesOnPageByName("Rack") > 1 Then
            MsgBox("Can only be one rack on the page!")
            rackShape.Delete()
        ElseIf Globals.ThisAddIn.Application.ActivePage.Shapes("ThePage").CellExists("User.IsPartOfChassisSwitch", False) Then
            If Globals.ThisAddIn.Application.ActivePage.Shapes("ThePage").Cells("User.IsPartOfChassisSwitch").Result("") = 1 Then
                MsgBox("You can't put a rack in a blade!")
                rackShape.Delete()
            End If
            End If



    End Sub

    ''' <summary>
    ''' Check if there is any other shape with the same name in the document
    ''' </summary>
    ''' <param name="uniqueShape">The shape that is checked</param>
    ''' <returns>True if there is no equal name, false if there are any similar names</returns>
    ''' <remarks></remarks>
    Public Function IsUniqueName(ByVal name As String, Optional ByRef uniqueShape As Visio.Shape = Nothing)

        Dim Page As Visio.Page
        Dim Shape As Visio.Shape
        Dim Document As Visio.Document

        If Not uniqueShape Is Nothing Then
            Document = uniqueShape.Document

            For Each Page In Document.Pages
                For Each Shape In Page.Shapes
                    If uniqueShape.ID <> Shape.ID AndAlso String.Compare(uniqueShape.Text, Shape.Text, False) = 0 Then
                        Return False
                    End If
                Next
            Next
        Else
            Document = Globals.ThisAddIn.Application.ActiveDocument
            For Each Page In Document.Pages
                For Each Shape In Page.Shapes
                    If Shape.CellExists(_ShapeName, False) Then
                        If String.Compare(name, Shape.Cells(_ShapeName).ResultStr(""), False) = 0 Then
                            Return False
                        End If
                    End If
                Next
            Next
        End If



        Return True
    End Function

End Module