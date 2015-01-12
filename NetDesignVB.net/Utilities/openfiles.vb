﻿Module openFiles

    Public Sub openStencil(ByVal applicationObj As Visio.Application, ByVal stencilName As String)

        Try
            If System.IO.File.Exists(stencilName) Then
                applicationObj.Documents.OpenEx(stencilName, (CShort(Visio.VisOpenSaveArgs.visOpenDocked) + CShort(Visio.VisOpenSaveArgs.visOpenRO)))
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

    End Sub

    Public Sub openDocument(ByVal applicationObj As Visio.Application, ByVal documentName As String, Optional ByVal openSaveArguments As Integer = Visio.VisOpenSaveArgs.visOpenDocked)

        Try
            If System.IO.File.Exists(documentName) Then
                applicationObj.Documents.OpenEx(documentName, (CShort(openSaveArguments)))
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

    End Sub

End Module