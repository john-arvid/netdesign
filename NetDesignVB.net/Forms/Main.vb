Module Main

    Public Class Main

        Dim ShapeList As List(Of Shape)

        Public Sub CacheShapes()

            Dim TempShape As New Shape
            Dim TempGUID As Integer
            Dim TempName As String
            Dim TempPortCount As Integer
            Dim TempUPosition As Integer
            Dim TempMediaType As String
            Dim TempMediaPurpose As String
            Dim TempMediaSpeed As String
            Dim TempPortNumber As Integer
            

            Try
                If Globals.ThisAddIn.Application.ActiveDocument.Pages.Count > 0 Then
                    For Each page As Visio.Page In Globals.ThisAddIn.Application.ActiveDocument.Pages
                        For Each shape As Visio.Shape In page.Shapes
                            If shape.Master.Name = "Switch" Then
                                TempShape.Name = shape.Text
                                TempShape.GUID = shape.UniqueID(Visio.VisUniqueIDArgs.visGetOrMakeGUID)
                            End If


                        Next
                    Next
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(ex.Message + "1")
            End Try

            ShapeList.Add(TempShape)

            Globals.ThisAddIn.Application.ActivePage.sh()
        End Sub


    End Class

    Public Structure Shape

        Dim GUID As Integer
        Dim Name As String
        Dim PortCount As Integer
        Dim UPosition As Integer
        Dim Ports() As Port


    End Structure

    Public Structure Port

        Dim MediaType As String
        Dim MediaPurpose As String
        Dim MediaSpeed As String
        Dim PortNumber As Integer

    End Structure

End Module