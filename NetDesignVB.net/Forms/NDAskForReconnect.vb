Public Class NDAskForReconnect

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonNewPage.CheckedChanged
        TextBoxNewPageName.Enabled = True
        ComboBoxExistingPage.Enabled = False
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonExistingPage.CheckedChanged
        TextBoxNewPageName.Enabled = False
        ComboBoxExistingPage.Enabled = True
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxODC.CheckedChanged
        If CheckBoxODC.Checked Then
            TextBoxFileName.Enabled = True
            Button3.Enabled = True
        Else
            TextBoxFileName.Enabled = False
            Button3.Enabled = False
        End If
        ComboBoxExistingPage.Items.Clear()
    End Sub

    Private Sub NDAskForReconnect_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TextBoxFileName.Enabled = False
        Button3.Enabled = False

        If Globals.ThisAddIn.Application.ActiveDocument.Pages.Count > 1 Then
            RadioButtonExistingPage.Checked = True
            Call SetPageNames(Globals.ThisAddIn.Application.ActiveDocument)
        Else
            RadioButtonNewPage.Checked = True
            RadioButtonExistingPage.Enabled = False
        End If

    End Sub


    Private Sub SetPageNames(ByRef document As Visio.Document)
        'TODO: does not work properly, when a page has been added when doing a page copy, the list does not get populated with the added page, needs to do this manually here in the code.
        Dim PageNames() As String = Nothing

        ComboBoxExistingPage.Items.Clear()

        document.Pages.GetNames(PageNames)

        For Each PageName As String In PageNames
            If Not (PageName = Globals.ThisAddIn.Application.ActivePage.Name) OrElse (CheckBoxODC.Checked) Then
                ComboBoxExistingPage.Items.Add(PageName)
                RadioButtonExistingPage.Enabled = True
            End If
        Next

        ComboBoxExistingPage.SelectedIndex = 0
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk


        TextBoxFileName.Text = OpenFileDialog1.FileName.ToString()
        Call SetPageNames(Globals.ThisAddIn.Application.Documents.OpenEx(TextBoxFileName.Text, Visio.VisOpenSaveArgs.visOpenMinimized))
    End Sub


    
    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        If Globals.ThisAddIn.Application.ActiveDocument.Name = "Drawing1" Then
            MsgBox("This document needs to be saved first")
            My.Computer.Keyboard.SendKeys("^s", True)
            CancelButton.PerformClick()
        Else
            OpenFileDialog1.Title = "Please select a file"

            OpenFileDialog1.ShowDialog()


        End If
    End Sub

End Class