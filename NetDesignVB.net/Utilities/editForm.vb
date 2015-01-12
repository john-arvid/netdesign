Module editForm

    Public Sub changeNameInForm(ByVal newLabel As String, _
                          ByRef formControl As Windows.Forms.Form)

        ' Using a stringbuilder to replace the given string, slower but works.
        'if a problem, then replace to manual string concatenating
        Dim stringBuilder As New StringBuilder()

        Dim test As Windows.Forms.Control

        ' Change the name of the form
        stringBuilder.Append(formControl.Text)
        stringBuilder.Replace(replaceInForm, newLabel)
        formControl.Text = stringBuilder.ToString
        stringBuilder.Clear()

        ' Goes through every thing in the form and checks for the ??
        For Each test In formControl.Controls
            If test.Text.Contains(replaceInForm) Then
                stringBuilder.Append(test.Text)
                stringBuilder.Replace(replaceInForm, newLabel)
                test.Text = stringBuilder.ToString
                stringBuilder.Clear()
            End If
        Next

    End Sub
 
End Module

