Module Examples

    ' This is a example on how to check every key input from the user.
    'I will not do it this way since I want the user to type in whatever he 
    'wants, and then check it afterwards. Then hopefully he will learn.
    ' Whenever a user types a character in the textbox3, this will trigger and
    'check what kind of character it is. If it is not a backspace, and it is
    'a letter then the character is allowed. If it is a letter it is revoked
    'and the background becomes red.

    ' Remove the comment on handles to make this active again,
    'this function needs to be in a form, e.g NDAskForSwitch.vb

    ''' <summary>
    ''' EXAMPLE. Allows only numeric input, triggers on keypress event. 
    ''' Changes background color on textbox.
    ''' </summary>
    ''' <param name="sender">The object</param>
    ''' <param name="e">The event</param>
    ''' <remarks>This is deprecated since the user should be allowed
    ''' to write input without interuption</remarks>


    Private Sub Is_Numeric_Input(ByVal sender As Object, _
                                 ByVal e As System.Windows.Forms.KeyPressEventArgs) 'Handles TextBox3.KeyPress

        ' As long as not trying to erase a character
        If e.KeyChar <> ControlChars.Back Then

            ' If it is a number, background as white and allow the input.
            If IsNumeric(e.KeyChar) Then
                'TextBox3.BackColor = Drawing.Color.White
                e.Handled = False
            Else
                ' Background as red and the input will be discarded
                'TextBox3.BackColor = Drawing.Color.Red
                e.Handled = True
            End If
        End If
    End Sub

    ' This is an example to show how to get an added stencil and the show all
    'the masternames that the stencil haves. This stencil needs to be opened 
    'in the document from before, but when opening the stencil one needed
    'to have the whole path, in this case one would only need the name

    ''' <summary>
    ''' An example on how to get all the masternames in a stencil. 
    ''' Intetion to show how the get the stencil and iterate through.
    ''' </summary>
    ''' <remarks>Example</remarks>
    Public Sub ShowAllMasterNames()

        Dim PortMasterName() As String = Nothing
        Dim LowerBound As Integer
        Dim UpperBound As Integer

        Try
            ' Get the stencil that has been added to the document.
            Globals.ThisAddIn.Application.Documents.Item("Netdesign.vssx").Masters.GetNames(PortMasterName)
            LowerBound = LBound(PortMasterName)
            UpperBound = UBound(PortMasterName)

            While LowerBound <= UpperBound
                MsgBox(PortMasterName(LowerBound))
                LowerBound = LowerBound + 1
            End While

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ruleSet"></param>
    ''' <remarks></remarks>
    Private Sub RuleWireLoop(ByVal ruleSet As Visio.ValidationRuleSet)
        Dim Page As Visio.Page
        Dim Rule As Visio.ValidationRule
        Dim Shape As Visio.Shape
        Dim ToShape As Visio.Shape
        Dim FromShape As Visio.Shape
        Dim ShapeId As VariantType

        Rule = GetRule(ruleSet, "WireLoop")
        If Not Rule Is Nothing Then
            For Each Page In ruleSet.Document.Pages
                For Each Shape In Page.Shapes
                    If Shape.Master.Name = "Wire" And Shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll2D, "").Length > 1 Then
                        ShapeId = Shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "")(0)
                        ToShape = Page.Shapes.ItemFromID(ShapeId)
                        ShapeId = Shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing2D, "")(0)
                        FromShape = Page.Shapes.ItemFromID(ShapeId)
                        ToShape.Delete()

                    End If
                Next
            Next
        End If
    End Sub


   


End Module