Public Class NDAskReport

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub NDAskReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Page As Visio.Page
        Dim Pages As Visio.Pages

        Pages = Globals.ThisAddIn.Application.ActiveDocument.Pages

        For Each Page In Pages
            CheckedListBoxPages.Items.Add(Page.Name)
        Next

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
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
                Call CreateReport(Page, AllWires, CheckedListBoxData.CheckedItems)
                GC.Collect()
            Next
        ElseIf (CheckedListBoxPages.CheckedItems.Count >= 1) Then
            For Each Item In CheckedListBoxPages.CheckedItems
                Page = Globals.ThisAddIn.Application.ActiveDocument.Pages.Item(Item.ToString())
                Call CreateReport(Page, AllWires, CheckedListBoxData.CheckedItems)
                GC.Collect()
            Next
        Else
            MsgBox("Nothing was selected, I do nothing!", MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        MsgBox("Report has been created and saved: Documents\NetDesignReport.txt", MsgBoxStyle.OkOnly)

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub CheckBoxWholeDocument_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxWholeDocument.CheckedChanged
        If CheckBoxWholeDocument.Checked Then
            CheckedListBoxPages.Enabled = False
        Else
            CheckedListBoxPages.Enabled = True
        End If
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub CheckedListBoxData_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBoxData.SelectedIndexChanged

    End Sub

    Private Sub CheckBoxAllData_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxAllData.CheckedChanged
        If CheckBoxAllData.Checked Then
            For i As Integer = 0 To CheckedListBoxData.Items.Count - 1
                CheckedListBoxData.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
            CheckedListBoxData.Enabled = False
        Else
            For i As Integer = 0 To CheckedListBoxData.Items.Count - 1
                CheckedListBoxData.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
            CheckedListBoxData.Enabled = True
        End If
    End Sub
End Class