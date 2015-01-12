Public Class FormNDAskForSwitch


#Region "Event Handeling"


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
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


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
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
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Textbox2_Validating(ByVal sender As Object, _
                                    ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBoxModel.Validating
        If e.Cancel Then
            Exit Sub
        End If

        If TextBoxModel.Text.Length = 0 Then
            MsgBox("You need to enter a type!", MsgBoxStyle.Critical)
            e.Cancel = True
        Else
            ' Does not actually need this since e.cancel is default false
            e.Cancel = False
        End If
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>

    Private Sub Textbox3_Validating(ByVal sender As Object, _
                                    ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBoxPort.Validating
        If e.Cancel Then
            Exit Sub
        End If

        If TextBoxPort.Text.Length = 0 Then
            MsgBox("You need to enter a number of ports!", MsgBoxStyle.Critical)
            e.Cancel = True
        ElseIf IsNumeric(TextBoxPort.Text) <> True Then
            MsgBox("Can only be numbers!", MsgBoxStyle.Critical)
            e.Cancel = True
        ElseIf TextBoxPort.Text > 99 Then
            MsgBox("This is insane, to many ports!", MsgBoxStyle.Critical)
            e.Cancel = True
        ElseIf TextBoxPort.Text < 0 Then
            MsgBox("This is not possible, you can't have negativ amount of ports!", MsgBoxStyle.Critical)
            e.Cancel = True
        Else
            ' Does not actually need this since e.cancel is default false
            e.Cancel = False
        End If
    End Sub

    

#End Region


    Private Sub FormNDAskForSwitch_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Set the default value
        ComboBoxPurpose.SelectedIndex = 1
        ComboBoxMedia.SelectedIndex = 0
        ComboBoxRow.SelectedIndex = 0


    End Sub

End Class