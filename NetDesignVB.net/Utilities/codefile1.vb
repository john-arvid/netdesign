
Public Class test

    Dim activeshape As Visio.Shape
    Dim pages As Visio.Pages
    Dim activeDocument As Visio.Document
    Dim activeApplication As Visio.Application

    Public Sub New()
        activeApplication = Globals.ThisAddIn.Application
        activeDocument = activeApplication.ActiveDocument
        pages = activeDocument.Pages
    End Sub

    Public Sub unique()
        

        If activeApplication.ActiveWindow.Selection.Count = 0 Then
            MsgBox("No shape selected")
            Exit Sub
        End If
        activeshape = activeApplication.ActiveWindow.Selection.Item(1)

        If activeshape.CellExists(ShapeName, 0) Then
            For Each page As Visio.Page In pages
                checkName(activeshape, page)
            Next

        End If

    End Sub

    Private Sub checkName(ByVal activeShape As Visio.Shape, ByVal page As Visio.Page)
        Dim i As Integer = 0

        For Each shape As Visio.Shape In page.Shapes
            If shape.CellExists(ShapeName, 0) Then
                If activeShape.Cells(ShapeName).ResultStr("") = shape.Cells(ShapeName).ResultStr("") Then
                    i += 1
                End If
            End If
        Next

        If i > 1 Then
            MsgBox("This shape has a name that is equal to another shape")
        End If

    End Sub

End Class