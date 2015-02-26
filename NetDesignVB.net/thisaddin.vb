

Public Class ThisAddIn



    Private Sub ThisAddIn_Startup() Handles Me.Startup
        'These two foreach are to cache the shapes data, this will 
        'increase the document loading time, but it will decrease the first
        ' check of rules.

        'My.Application.Info.Version.ToString()
        'If System.Deployment.Application.ApplicationDeployment.CurrentDeployment.IsFirstRun Then
        '    MsgBox("First time running!")
        'End If

        Dim DialogResult As Integer = MsgBox("Do you want to enable NetDesign? test", MsgBoxStyle.YesNo)

        If DialogResult = MsgBoxResult.No Then
            Exit Sub
        End If

        Try
            If Globals.ThisAddIn.Application.ActiveDocument.Pages.Count > 0 Then
                For Each page As Visio.Page In Globals.ThisAddIn.Application.ActiveDocument.Pages
                    For Each shape As Visio.Shape In page.Shapes
                    Next
                Next
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message + "1")
        End Try

        'Dim path As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\Visual Studio 2012\Projects\NetDesignVB.net\NetDesignVB.net\bin\Debug\"
        'openDocument(Me.Application, path + "clean.vsdx")
        'openStencil(Me.Application, path + "Netdesign.vssx")

        Dim test As New AddAdvise
        test.DemoAddAdvise(Globals.ThisAddIn.Application)

        ' Create a new drawing.
        If Globals.ThisAddIn.Application.Documents.Count = 0 Then
            Globals.ThisAddIn.Application.Documents.Add("")

            ' Prepare the first page the new document
            Call PreparePage(Globals.ThisAddIn.Application.ActivePage)
        End If


        Call AddStencils()

        ' Add ruleset
        Call AddOrUpdateRuleSet()

        

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
