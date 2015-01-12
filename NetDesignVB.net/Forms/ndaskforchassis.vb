Public Class NDAskForChassis

    '' I don't need this constructor since I do this when I create the object
    '    Public Sub New(ByVal newLabel As String)
    '' Needs to have this in the constructor
    '    InitializeComponent()
    '' Changing the names of labels and button
    '    Call changeName(newLabel, Me)

    '    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, _
                              ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.FormClosing

        ' If the ok button was pressed
        If DialogResult = Windows.Forms.DialogResult.OK Then
            If Not ValidateChildren() Then
                e.Cancel = True
            Else
                ' Does not actually need this since e.cancel is default false
                e.Cancel = False
            End If
        End If

    End Sub


    Private Sub Textbox1_Validating(ByVal sender As Object, _
                                    ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBoxName.Validating
        If e.Cancel = True Then
            Exit Sub
        End If

        If TextBoxName.Text.Length = 0 Then
            MsgBox("You need to enter a name!", MsgBoxStyle.Critical)
            e.Cancel = True
        ElseIf Not IsUniqueName(TextBoxName.Text) Then
            MsgBox("You need to enter a unique name!", MsgBoxStyle.Critical)
            e.Cancel = True
        Else
            e.Cancel = False
        End If

    End Sub

    Private Sub NDAskForChassis_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBoxMedia.SelectedIndex = 0
        ComboBoxPurpose.SelectedIndex = 1
        ComboBoxRow.SelectedIndex = 0
    End Sub
End Class