Public Class NDAskReport


    Private Sub NDAskReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Page As Visio.Page
        Dim Pages As Visio.Pages

        Pages = Globals.ThisAddIn.Application.ActiveDocument.Pages

        For Each Page In Pages
            CheckedListBox1.Items.Add(Page.Name)
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
        ElseIf (CheckedListBox1.CheckedItems.Count >= 1) Then
            For Each Item In CheckedListBox1.CheckedItems
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
            CheckedListBox1.Enabled = False
        Else
            CheckedListBox1.Enabled = True
        End If
    End Sub
End Class