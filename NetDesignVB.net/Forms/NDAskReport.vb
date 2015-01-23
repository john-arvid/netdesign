Public Class NDAskReport


    Private Sub NDAskReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Page As Visio.Page
        Dim Pages As Visio.Pages

        Pages = Globals.ThisAddIn.Application.ActiveDocument.Pages

        For Each Page In Pages
            CheckedListBoxPages.Items.Add(Page.Name)
        Next

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Page As Visio.Page
        Dim Pages As Visio.Pages
        Dim AllWires As Boolean = True
        Dim Item As Object

        If CheckBoxConnectedWire.Checked Then
            AllWires = False
        End If

        If CheckBoxWholeDocument.Checked Then
            Pages = Globals.ThisAddIn.Application.ActiveDocument.Pages

            For Each Page In Pages
                Call CreateReport(Page, AllWires)
                GC.Collect()
            Next
        ElseIf (CheckedListBoxPages.CheckedItems.Count >= 1) Then
            For Each Item In CheckedListBoxPages.CheckedItems
                Page = Globals.ThisAddIn.Application.ActiveDocument.Pages.Item(Item.ToString())
                Call CreateReport(Page, AllWires)
                GC.Collect()
            Next
        Else
            MsgBox("Nothing was selected, I do nothing!", MsgBoxStyle.OkOnly)

        End If

    End Sub

    Private Sub CheckBoxWholeDocument_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxWholeDocument.CheckedChanged
        If CheckBoxWholeDocument.Checked Then
            CheckedListBoxPages.Enabled = False
        Else
            CheckedListBoxPages.Enabled = True
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub CheckedListBoxData_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBoxData.SelectedIndexChanged

    End Sub
End Class