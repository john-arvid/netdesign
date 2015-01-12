Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

        'For i As Integer = 1 To Globals.ThisAddIn.Application.EventList.Count
        '    MsgBox(Globals.ThisAddIn.Application.EventList.Item(i).TargetArgs)
        'Next

        'Call ChangeMasterCellsAndSections()

        'Dim shape As Visio.Shape = Globals.ThisAddIn.Application.ActivePage.Shapes(1)
        'For Each item As Visio.Shape In shape.Shapes
        '    MsgBox(item.NameU)
        'Next

        Dim shape As Visio.Shape = Globals.ThisAddIn.Application.ActivePage.Shapes(1)
        MsgBox(shape.Shapes.Item(1).Cells("User.WireID").ResultStr(Visio.VisUnitCodes.visUnitsString))

        'Dim shape As Visio.Shape = Globals.ThisAddIn.Application.ActivePage.Shapes(1)
        'For Each item As Visio.Shape In shape.Shapes
        '    MsgBox(item.Cells("Prop.WireNumber").ResultStr(""))
        '    MsgBox(item.Cells("User.PortName").ResultStr(Visio.VisUnitCodes.visUnitsString))
        'Next

        'Dim mastershape As Visio.Master = Globals.ThisAddIn.Application.Documents("Netdesign.vssx").Masters.Item("Wire Bundle")
        'Dim shape As Visio.Shape = mastershape.Shapes(1)
        'shape.CellsU("GlueType").Formula = 8
        'For Each item As Visio.Shape In shape.Shapes
        '    item.CellsU("GlueType").Formula = 8
        '    If item.Shapes.Count > 0 Then
        '        For Each item2 As Visio.Shape In item.Shapes
        '            item2.CellsU("GlueType").Formula = 8
        '            If item2.Shapes.Count > 0 Then
        '                For Each item3 As Visio.Shape In item.Shapes
        '                    item3.CellsU("GlueType").Formula = 8

        '                Next
        '            End If
        '        Next
        '    End If
        'Next

    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles ReportButton.Click
        Dim Report As New NDAskReport

        Report.ShowDialog()

        Report.Close()

        My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)

    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        'Item 6 is the shape added event
        Globals.ThisAddIn.Application.EventList.Item(1).Enabled = False
        'openDocument(Globals.ThisAddIn.Application, "G:\Projects\ATLASTDAQNetworking\TDAQ Connectivity\0_NetDesign Visio\New Tool\Common.vsdx", Visio.VisOpenSaveArgs.visOpenMinimized)
        openDocument(Globals.ThisAddIn.Application, "\\cern.ch\dfs\Users\j\jkibsgaa\Documents\Common.vsdx", Visio.VisOpenSaveArgs.visOpenMinimized)
        Call Magic()

        My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
        Globals.ThisAddIn.Application.EventList.Item(1).Enabled = True
    End Sub
End Class
