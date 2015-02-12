﻿Imports System.Threading.Tasks
Imports Microsoft.Office.Interop.Visio.VisSectionIndices
Imports System.Text.RegularExpressions


Module Utilities

    Dim _list As New List(Of Windows.Forms.Control)

    ''' <summary>
    ''' Iterates through every control in the form, if a control has a child it
    ''' recursively check the child control items. E.g a groupbox.
    ''' </summary>
    ''' <param name="container">The form that has the controls</param>
    ''' <returns>A list of controls</returns>
    ''' <remarks>Does not check if the form is empty</remarks>
    Public Function GetControlChilds(ByRef container As Windows.Forms.Control)
        For Each child As Windows.Forms.Control In container.Controls
            _list.Add(child)
            If child.HasChildren Then
                GetControlChilds(child)
            End If
        Next

        Return _list
    End Function

    ''' <summary>
    ''' Copy shapesheet information from a shape to another, only information
    ''' from a specific section
    ''' </summary>
    ''' <param name="fromShape"></param>
    ''' <param name="toShape"></param>
    ''' <param name="sectionName">What section should be copied</param>
    ''' <remarks></remarks>
    Public Sub CopyShapeSheetData(ByVal fromShape As Visio.Shape, ByRef toShape As Visio.Shape, ByVal sectionName As Visio.VisSectionIndices)

        ' This will hold the row information
        Dim ToRow As Visio.Row
        Dim FromRow As Visio.Row

        ' If the section does not exist in the original shape, exit
        If Not fromShape.SectionExists(sectionName, False) Then
            'TODO: debug.write to immidiate
            Exit Sub
        End If

        ' Add the section in the destination shape if it does not exist
        If Not toShape.SectionExists(sectionName, False) Then
            toShape.AddSection(sectionName)
        End If

        ' Add all the row names to be sure that another cell won't get an
        'invalid reference
        For i As Integer = 0 To fromShape.Section(sectionName).Count - 1
            toShape.AddRow(sectionName, i, 0)
            FromRow = fromShape.Section(sectionName).Row(i)
            ToRow = toShape.Section(sectionName).Row(i)

            ToRow.Name = FromRow.Name
        Next

        ' Iterate trough every row and every cell, and copies the data(formula)
        For i As Integer = 0 To fromShape.Section(sectionName).Count - 1
            FromRow = fromShape.Section(sectionName).Row(i)
            ToRow = toShape.Section(sectionName).Row(i)

            For j As Integer = 0 To FromRow.Count - 1
                ToRow.Cell(j).Formula = FromRow.Cell(j).Formula
            Next
        Next

    End Sub


    ''' <summary>
    ''' Counts the number of given shapes on the current page
    ''' </summary>
    ''' <param name="shape">The shape that is being counted</param>
    ''' <returns>Number of given shapes on same page</returns>
    ''' <remarks>Gives the total, including the passed shape</remarks>
    Public Function CountShapesOnPage(ByVal shape As Visio.Shape) As Integer
        Dim Count As Integer = 0
        Dim Page As Visio.Page
        Dim Thing As Visio.Shape

        Page = shape.ContainingPage

        For Each Thing In Page.Shapes
            If Thing.Master.Name = shape.Master.Name Then
                Count += 1
            End If
        Next

        CountShapesOnPage = Count
    End Function

    ''' <summary>
    ''' Counts all the shapes with a specific name on the active page
    ''' </summary>
    ''' <param name="shapeName">Look for this name</param>
    ''' <returns>How many shapes with the specified name</returns>
    ''' <remarks></remarks>
    Public Function CountShapesOnPageByName(ByVal shapeName As String) As Integer
        Dim Count As Integer = 0
        Dim Page As Visio.Page
        Dim Shape As Visio.Shape

        Page = Globals.ThisAddIn.Application.ActivePage

        For Each Shape In Page.Shapes
            If Not Shape.Master Is Nothing Then
                If Shape.Master.Name = shapeName Then
                    Count += 1
                End If
            End If
        Next

        Return Count
    End Function

    ''' <summary>
    ''' Goes thorugh all the pages in the document and looks for an GUID
    ''' </summary>
    ''' <param name="document">Look through this document</param>
    ''' <param name="guid">Find this GUID</param>
    ''' <returns>The page with the specified GUID</returns>
    ''' <remarks></remarks>
    Public Function GetPageByGUID(ByRef document As Visio.Document, ByVal guid As String) As Visio.Page

        Dim Page As Visio.Page
        Dim GUIDPage As String

        For Each Page In document.Pages
            GUIDPage = Page.PageSheet.UniqueID(Visio.VisUniqueIDArgs.visGetGUID)
            If guid = GUIDPage Then
                Return Page
            End If
        Next

        Return Nothing

    End Function

    ''' <summary>
    ''' Called when a formula is changed, calls a specific function to handle 
    ''' the shape
    ''' </summary>
    ''' <param name="shape">The shape that was changed</param>
    ''' <param name="cell">The cell that was changed</param>
    ''' <remarks>Need a try catch to avoid exception when a shape without 
    ''' a master is changed</remarks>
    Public Sub FormulaChanged(ByRef shape As Visio.Shape, ByRef cell As Visio.Cell)

        ' Trigger the validation of the document when a formula is changed
        'Globals.ThisAddIn.Application.ActiveDocument.Validation.Validate()

        If shape.CellExists("User.msvShapeCategories", 0) AndAlso (cell.Section = visSectionProp) Then

            If shape.Cells("User.msvShapeCategories").ResultStr("") = "Switch" Then
                Call UpdateSwitch(shape, cell)
            ElseIf shape.Cells("User.msvShapeCategories").ResultStr("") = "Processor" Then
                Call UpdateSwitch(shape, cell)
            ElseIf shape.Cells("User.msvShapeCategories").ResultStr("") = "Blade" Then
                Call UpdateSwitch(shape, cell)
            ElseIf shape.Cells("User.msvShapeCategories").ResultStr("") = "Rack" Then
                Call UpdateRack(shape, cell)
            ElseIf shape.Cells("User.msvShapeCategories").ResultStr("") = "Wire" Then
                Call UpdateWire(shape, cell)
            ElseIf shape.Cells("User.msvShapeCategories").ResultStr("") = "Port" Then
                Call UpdatePort(shape, cell)
            ElseIf shape.Cells("User.msvShapeCategories").ResultStr("") = "Chassis Switch" Then
                Call UpdateSwitch(shape, cell)
            ElseIf shape.Cells("User.msvShapeCategories").ResultStr("") = "Chassis Processor" Then
                Call UpdateSwitch(shape, cell)
            End If
        End If


    End Sub


    ''' <summary>
    ''' Move shapesheet information from a shape to another
    ''' </summary>
    ''' <param name="toShape"></param>
    ''' <param name="fromShape"></param>
    ''' <remarks></remarks>
    Public Sub MoveInformation(ByVal toShape As Visio.Shape, ByRef fromShape As Visio.Shape)

        toShape.Cells("User.MediaType").Formula = """" + fromShape.Cells("Prop.Media").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
        toShape.Cells("User.MediaPurpose").Formula = """" + fromShape.Cells("Prop.Purpose").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
        toShape.Cells("User.MediaSpeed").Formula = """" + fromShape.Cells("Prop.TransmissionSpeed").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
        toShape.Cells("User.PortName").Formula = """" + fromShape.Cells("User.PortName").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
        toShape.Cells("User.SwitchName").Formula = """" + fromShape.Cells("User.SwitchName").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="connections"></param>
    ''' <param name="connecting"></param>
    ''' <remarks></remarks>
    Public Sub ConnectionChanged(ByRef connections As Visio.Connects, ByVal connecting As Boolean)

        'TODO: Add when disconnecting that the wire looses information.

        ' Fromsheet is always a wire, so test only on tosheet
        If connections.ToSheet.Cells("User.msvShapeCategories").ResultStr("") = "OPC" Then

            Call UpdateOPC(connections.ToSheet, connections.FromSheet)
        ElseIf connections.ToSheet.Cells("User.msvShapeCategories").ResultStr("") = "Port" Then
            If connecting Then
                connections.ToSheet.CellsU("FillForegndTrans").FormulaForceU = "0%"
            Else
                connections.ToSheet.CellsU("FillForegndTrans").FormulaForceU = "50%"
            End If
        End If
        Call ValidateWireConnection(connections)
        'TODO: this is happening at every connection. should it, or doesnt it matter?
        Call SynchWire(connections.ToSheet, connections.FromSheet)
        Call UpdateLabel(connections.ToSheet, connections.FromSheet)

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="connectedShape"></param>
    ''' <param name="wireShape"></param>
    ''' <remarks></remarks>
    Public Sub SynchWire(ByRef connectedShape As Visio.Shape, ByRef wireShape As Visio.Shape)

        If connectedShape.Cells("User.msvShapeCategories").ResultStr("") = "Port" Then
            wireShape.Cells("User.SwitchName").Formula = """" + connectedShape.Cells("User.SwitchName").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
            wireShape.Cells("User.PortName").Formula = """" + connectedShape.Cells("User.TextTitle").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
        ElseIf connectedShape.Cells("User.msvShapeCategories").ResultStr("") = "OPC" Then
            wireShape.Cells("User.SwitchName").Formula = """" + connectedShape.Cells("User.SwitchName").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
            wireShape.Cells("User.PortName").Formula = """" + connectedShape.Cells("User.PortName").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
        End If

        wireShape.BringToFront()

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="connectedShape"></param>
    ''' <param name="wireShape"></param>
    ''' <remarks></remarks>
    Public Sub UpdateLabel(ByRef connectedShape As Visio.Shape, ByRef wireShape As Visio.Shape)
        Dim IncomingNode As Visio.Shape
        Dim OutgoingNode As Visio.Shape


        'Remove the shape text lock, force since it's guarded
        wireShape.CellsU("LockTextEdit").FormulaForce = "GUARD(0)"

        If wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll2D, "").Length = 2 Then
            IncomingNode = wireShape.ContainingPage.Shapes.ItemFromID(wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "")(0))
            OutgoingNode = wireShape.ContainingPage.Shapes.ItemFromID(wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing2D, "")(0))

            If IncomingNode.Cells("User.msvShapeCategories").ResultStr("") = "OPC" Then
                wireShape.Text = IncomingNode.Text + "/" + OutgoingNode.Cells("User.SwitchName").ResultStr(Visio.VisUnitCodes.visUnitsString) + ":" + OutgoingNode.Cells("User.UPosition").ResultStr(Visio.VisUnitCodes.visUnitsString) + ":" + OutgoingNode.Text
            ElseIf OutgoingNode.Cells("User.msvShapeCategories").ResultStr("") = "OPC" Then
                wireShape.Text = OutgoingNode.Text + "/" + IncomingNode.Cells("User.SwitchName").ResultStr(Visio.VisUnitCodes.visUnitsString) + ":" + IncomingNode.Cells("User.UPosition").ResultStr(Visio.VisUnitCodes.visUnitsString) + ":" + IncomingNode.Text
            End If
        Else
            wireShape.Text = "Not complete yet!"
        End If


        'Add the shape text lock, force since it's guarded
        wireShape.CellsU("LockTextEdit").FormulaForce = "GUARD(1)"

    End Sub

    Public Function GetRackShape()
        Dim RackShape As Visio.Shape = Nothing
        Dim Shape As Visio.Shape
        Dim Shapes As Visio.Shapes

        Shapes = Globals.ThisAddIn.Application.ActivePage.Shapes

        For Each Shape In Shapes
            If Not Shape.Master Is Nothing Then
                If Shape.Master.Name = "Rack" Then
                    RackShape = Shape
                    Return RackShape
                End If
            End If
        Next

        Return RackShape
    End Function

    Public Sub Magic()

        Dim Document As Visio.Document = Globals.ThisAddIn.Application.Documents.Item(1)
        Dim OtherDocument As Visio.Document
        Dim Page As Visio.Page
        Dim NextPage As Visio.Page
        Dim RackMaster As Visio.Master
        Dim SwitchMaster As Visio.Master
        Dim WireMaster As Visio.Master
        Dim OPCMaster As Visio.Master
        Dim MasterBundle As Visio.Master
        Dim RackShape As Visio.Shape
        Dim RackShapeCopy As Visio.Shape
        Dim PreviousSwitch As Visio.Shape
        Dim NextSwitch As Visio.Shape
        Dim MainSwitchSmall As Visio.Shape
        Dim WireShape As Visio.Shape
        Dim WireShapeCopy As Visio.Shape
        Dim OPCShape As Visio.Shape
        Dim OPCCopy As Visio.Shape
        Dim OPCMain As Visio.Shape
        Dim OPCMainCopy As Visio.Shape
        Dim WireBundle As Visio.Shape
        Dim WireBundleCopy As Visio.Shape
        Dim WireID As Integer = 100000
        Dim ProgressBar As New ProgressBar()

        Dim NumberOfPorts As String = "24"
        Dim NumberOfPages As Integer = 4




        OtherDocument = Globals.ThisAddIn.Application.Documents.Item("Common.vsdx")

        RackMaster = Globals.ThisAddIn.Application.Documents.Item("Netdesign.vssx").Masters.Item("Rack")
        SwitchMaster = Globals.ThisAddIn.Application.Documents.Item("Netdesign.vssx").Masters.Item("Switch")
        WireMaster = Globals.ThisAddIn.Application.Documents.Item("Netdesign.vssx").Masters.Item("Wire")
        OPCMaster = Globals.ThisAddIn.Application.Documents.Item("Netdesign.vssx").Masters.Item("OPC")
        MasterBundle = Globals.ThisAddIn.Application.Documents.Item("Netdesign.vssx").Masters.Item("Wire Bundle")

        ProgressBar.Show()

        Dim PositionY3 = 10.8
        Dim Count As Integer = Document.Pages.Count
        For i As Integer = Count To Count + NumberOfPages

            Dim PositionX = 0.5 ' For the lower switch
            Dim PositionY = 2.0 ' For the lower switch
            Dim PositionX2 = 0.5 ' For the upper switch
            Dim PositionY2 = 9.0 ' For the upper switch

            Page = Document.Pages.Item(i)

            'Drop rack on page
            RackShape = Document.Pages.Item(i).Drop(RackMaster, 0, 0)
            RackShape.Cells("Prop.RackLocation").Formula = """" + Page.Name + """"

            'Create nextswitch
            NextSwitch = Document.Pages.Item(i).Drop(SwitchMaster, 4, 11)
            Call CreateSwitch(NextSwitch, "Test2", "TestModel2", NumberOfPorts, "2", "Copper", "Data network", False, Document, Page)
            NextSwitch.Cells("Width").Formula = """" + "181.3453mm" + """"
            NextSwitch.Cells("Height").Formula = """" + "21.7614mm" + """"

            'Create small mainswitch
            MainSwitchSmall = Document.Pages.Item(i).Drop(SwitchMaster, 4, 5.3)
            Call CreateSwitch(MainSwitchSmall, "Test3", "TestModel3", "1", "1", "Copper", "Data network", False, Document, Page)
            MainSwitchSmall.Cells("Width").Formula = """" + "20mm" + """"
            MainSwitchSmall.Cells("Height").Formula = """" + "20mm" + """"

            'Create the next page
            Document.Pages.Add()
            NextPage = Document.Pages.Item(i + 1)

            'Drop rack on next page
            RackShapeCopy = Document.Pages.Item(i + 1).Drop(RackMaster, 0, 0)
            RackShapeCopy.Cells("Prop.RackLocation").Formula = """" + NextPage.Name + """"

            'Create previous switch
            PreviousSwitch = Document.Pages.Item(i + 1).Drop(SwitchMaster, 4, 1.9)
            Call CreateSwitch(PreviousSwitch, "Test1", "TestModel1", NumberOfPorts, "2", "Copper", "Data network", False, Document, NextPage)
            PreviousSwitch.Cells("Width").Formula = """" + "181.3453mm" + """"
            PreviousSwitch.Cells("Height").Formula = """" + "21.7614mm" + """"


            'Create all the OPC for the switch goind to the next page
            For j As Integer = 1 To NumberOfPorts

                If j Mod 9 = 1 Then
                    PositionY = 2.0
                    PositionY2 = 9.0
                End If

                'Create OPC on the first page
                OPCShape = Document.Pages.Item(i).Drop(OPCMaster, PositionX2, PositionY2)

                'Create OPC on the second page
                OPCCopy = Document.Pages.Item(i + 1).Drop(OPCMaster, PositionX, PositionY)

                'Transfer information between OPC
                Call TransferOPCInfo(OPCShape, OPCCopy, False)

                'Create wire on the first page
                WireShape = Document.Pages.Item(i).Drop(WireMaster, PositionX2 + 0.2, PositionY2)

                'Create wire on the second page
                WireShapeCopy = Document.Pages.Item(i + 1).Drop(WireMaster, PositionX + 0.2, PositionY)

                'Connect wire on first page
                WireShape.Cells("BeginX").GlueTo(NextSwitch.Shapes(j + 1).Cells("AlignBottom"))
                WireShape.Cells("EndX").GlueTo(OPCShape.Cells("AlignRight"))

                'Connect wire on second page
                WireShapeCopy.Cells("BeginX").GlueTo(PreviousSwitch.Shapes(j + 1).Cells("AlignTop"))
                WireShapeCopy.Cells("EndX").GlueTo(OPCCopy.Cells("AlignRight"))

                'Add the cable id
                WireShape.Cells("Prop.WireID").Formula = """" + WireID.ToString() + """"
                WireShapeCopy.Cells("Prop.WireID").Formula = """" + WireID.ToString() + """"

                'Block reporting on the other end of wire
                WireShapeCopy.Cells("User.NotReport").Formula = """" + "1" + """"


                'Change the positions for the next iteration
                PositionX += 0.148
                PositionY += 0.2
                PositionX2 += 0.148
                PositionY2 -= 0.2

                WireID += 1
            Next

            'Create wire bundles and connect
            'Dim Multiplier As Integer = 1
            'For j As Integer = 1 To 4
            '    WireBundle = Document.Pages.Item(i).Drop(MasterBundle, 2 + (j / 3), 9)
            '    WireBundleCopy = Document.Pages.Item(i + 1).Drop(MasterBundle, 2 + (j / 3), 2)

            '    Call TransferOPCInfo(WireBundle.Shapes.Item(13), WireBundleCopy.Shapes.Item(13), False)
            '    WireBundle.Shapes.Item(13).Text = RackShapeCopy.Cells("Prop.RackLocation").ResultStr("")
            '    WireBundleCopy.Shapes.Item(13).Text = RackShape.Cells("Prop.RackLocation").ResultStr("")

            '    For k As Integer = 1 To 12
            '        WireBundle.Shapes.Item(k).Cells("EndX").GlueTo(NextSwitch.Shapes.Item((k + Multiplier)).Cells("AlignBottom"))
            '        WireBundleCopy.Shapes.Item(k).Cells("EndX").GlueTo(PreviousSwitch.Shapes.Item((k + Multiplier)).Cells("AlignTop"))


            '    Next

            '    Multiplier += 12
            'Next



            'Create ODC for mainswitch in first document
            OPCMain = Document.Pages.Item(i).Drop(OPCMaster, 3, 5.3)

            'Create ODC for mainswitch in second document
            OPCMainCopy = OtherDocument.Pages.Item(1).Drop(OPCMaster, 2, PositionY3)

            'Transfer information between the ODC
            Call TransferOPCInfo(OPCMain, OPCMainCopy, True, OtherDocument.Path + OtherDocument.Name, Document.Path + Document.Name)

            'Create wire on page in first document
            WireShape = Document.Pages.Item(i).Drop(WireMaster, 3.02, 5)

            'Create wire on page in second document
            WireShapeCopy = OtherDocument.Pages.Item(1).Drop(WireMaster, 2.02, PositionY3)

            'Add the cable id
            WireShape.Cells("Prop.WireID").Formula = """" + WireShape.NameID + """"
            WireShapeCopy.Cells("Prop.WireID").Formula = """" + WireShapeCopy.NameID + """"

            'Connect wire in first document
            WireShape.Cells("BeginX").GlueTo(MainSwitchSmall.Shapes(2).Cells("AlignLeft"))
            WireShape.Cells("EndX").GlueTo(OPCMain.Cells("AlignRight"))

            'Connect wire in second document
            WireShapeCopy.Cells("BeginX").GlueTo(OtherDocument.Pages.Item(1).Shapes(2).Shapes(i + 1).Cells("AlignLeft"))
            WireShapeCopy.Cells("EndX").GlueTo(OPCMainCopy.Cells("AlignRight"))
            PositionY3 -= 0.2

            ProgressBar.Text = i.ToString() + " / " + NumberOfPages.ToString()

            ProgressBar.ProgressBar1.Increment(100 / NumberOfPages)


            WireID += 1000

        Next


        ProgressBar.Close()


    End Sub


    Private Sub TransferOPCInfo(ByRef OPCShape As Visio.Shape, ByRef OPCCopy As Visio.Shape, _
                                ByVal ODC As Boolean, Optional ByVal otherDocumentPath As String = "", _
                                Optional ByVal firstDocumentPath As String = "")

        ' Set the necessary information in the OPC, this is needed because of
        ' the OPC command that is called in the shapesheet cell event doubleclick
        OPCShape.Hyperlinks("OffPageConnector").SubAddress = OPCCopy.ContainingPage.Name
        OPCCopy.Hyperlinks("OffPageConnector").SubAddress = OPCShape.ContainingPage.Name
        OPCShape.Cells("User.OPCShapeID").Formula = """" + OPCShape.UniqueID(1).ToString + """"
        OPCCopy.Cells("User.OPCShapeID").Formula = """" + OPCCopy.UniqueID(1).ToString + """"
        OPCShape.Cells("User.OPCDShapeID").Formula = """" + OPCCopy.UniqueID(1).ToString + """"
        OPCCopy.Cells("User.OPCDShapeID").Formula = """" + OPCShape.UniqueID(1).ToString + """"
        OPCShape.Cells("User.OPCDPageID").Formula = """" + OPCCopy.ContainingPage.PageSheet.UniqueID(Visio.VisUniqueIDArgs.visGetOrMakeGUID) + """"
        OPCCopy.Cells("User.OPCDPageID").Formula = """" + OPCShape.ContainingPage.PageSheet.UniqueID(Visio.VisUniqueIDArgs.visGetOrMakeGUID) + """"

        If ODC Then
            OPCShape.Hyperlinks("OffPageConnector").Address = otherDocumentPath
            OPCShape.Cells("User.OPCDDocID").Formula = """" + otherDocumentPath + """"
            OPCCopy.Hyperlinks("OffPageConnector").Address = firstDocumentPath
            OPCCopy.Cells("User.OPCDDocID").Formula = """" + firstDocumentPath + """"
        End If

    End Sub



    Public Sub ChangeMasterCellsAndSections()

        Dim Stencil As Visio.Document = Globals.ThisAddIn.Application.Documents.Item("Netdesign.vssx")
        Dim MasterShape As Visio.Master
        Dim GroupShape As Visio.Shape
        Dim SubShape As Visio.Shape
        Dim Counter As Integer = 1
        Dim Form As New NDChangeMasterCellsAndSections

        Form.ShowDialog()

        If Form.DialogResult = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If

        MasterShape = Stencil.Masters.Item(Form.TextBoxMasterName.Text)
        GroupShape = MasterShape.Shapes.Item(1)

        Call CheckAndAddSection(GroupShape, visSectionUser)
        Call CheckAndAddSection(GroupShape, visSectionProp)

        For Each SubShape In GroupShape.Shapes
            Call CheckAndAddSection(SubShape, visSectionUser)
            Call CheckAndAddSection(SubShape, visSectionProp)

            Call CheckAndAddCellRow(SubShape, Form.TextBoxCellName.Text, visSectionUser, Form.TextBoxCellValue.Text)
            'Call CheckAndAddCellRow(SubShape, "OPCShapeID", visSectionUser)

            Counter += 1
        Next


    End Sub

    Private Sub CheckAndAddSection(ByRef shape As Visio.Shape, ByVal section As Visio.VisSectionIndices)

        If Not shape.SectionExists(section, 0) Then
            shape.AddSection(section)
        End If
    End Sub

    Private Sub CheckAndAddCellRow(ByRef shape As Visio.Shape, ByVal cellName As String, ByVal section As Visio.VisSectionIndices, Optional ByVal cellValue As String = "")

        Dim SectionName As String = ""

        If section = visSectionUser Then
            SectionName = "User."
        ElseIf section = visSectionProp Then
            SectionName = "Prop."
        ElseIf section = visSectionHyperlink Then
            SectionName = "Hyperlinks."
        End If

        If Not shape.CellExists((SectionName + cellName), 0) Then
            shape.AddNamedRow(section, cellName, Visio.VisRowTags.visTagDefault)
        End If
        If cellValue.Contains("!") Then
            shape.Cells(SectionName + cellName).Formula = cellValue
        Else
            shape.Cells(SectionName + cellName).Formula = """" + cellValue + """"
        End If


    End Sub

    Private Sub DeleteCellRow(ByRef shape As Visio.Shape, ByVal cellName As String, ByVal section As Visio.VisSectionIndices)

        Dim SectionName As String = ""

        If section = visSectionUser Then
            SectionName = "User."
        ElseIf section = visSectionProp Then
            SectionName = "Prop."
        End If

        If shape.CellExists(SectionName + cellName, 0) Then
            shape.DeleteRow(section, 0)
        End If

    End Sub


    Public Sub MarkerHandler(ByRef subject As Visio.Application)

        MsgBox("This is not implemented yet")
        Exit Sub

        Dim MarkerInfo() As String = Split(subject.EventInfo(0), "/")
        Dim Document As Visio.Document
        Dim Page As Visio.Page
        Dim shape As Visio.Shape
        Dim Arguments As String = MarkerInfo(5)
        Dim Application As String = MarkerInfo(6)

        For i As Integer = MarkerInfo.GetLowerBound(0) To MarkerInfo.GetUpperBound(0)
            MarkerInfo(i) = Regex.Replace(MarkerInfo(i), ".*=", "")
        Next

        Document = Globals.ThisAddIn.Application.Documents.Item(Convert.ToInt16(MarkerInfo(1)))
        Page = Document.Pages.Item(Convert.ToInt16(MarkerInfo(2)))

        For Each item As Visio.Shape In Page.Shapes
            If item.NameU = MarkerInfo(4) Then
                shape = item
            End If
        Next



    End Sub


    Public Sub DeleteChassisPages(ByRef shape As Visio.Shape)

        Dim Page As Visio.Page

        For Each ChildShape As Visio.Shape In shape.Shapes
            If ChildShape.Cells("User.msvShapeCategories").ResultStr("") = "Chassis Switch Page" Then
                Page = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ChildShape.Hyperlinks("OffPageConnector").SubAddress)
                Page.Delete(1)
            End If
        Next


    End Sub


    Public Sub PreparePage(ByRef page As Visio.Page)

        page.PageSheet.Cells("DrawingSizeType").Formula = "=3"
        page.PageSheet.Cells("DrawingResizeType").Formula = "=2"

    End Sub

End Module