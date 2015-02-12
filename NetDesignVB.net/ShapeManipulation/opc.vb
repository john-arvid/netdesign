Module OPC



    Public Sub HandleOffPageConnector(OPCShape As Visio.Shape)
        Static NoDialog As Integer = 0

        ' To avoid the dialog to open when the corresponding OPC is being dropped
        If NoDialog = 1 Then
            NoDialog = 0
            Exit Sub
        End If

        'If OPCShape.Cells("User.OPCType").ResultStr("") = "Wire Bundle" Then
        '    OPCShape = OPCShape.Shapes.Item(13)
        'End If


        Dim OPCForm As New NDAskForReconnect

        ' Show the user dialog form
        OPCForm.ShowDialog()

        ' If the user exits the form, delete the shape
        If OPCForm.DialogResult = Windows.Forms.DialogResult.Cancel Then
            OPCShape.Delete()
        Else
            NoDialog = 1

            Dim OtherDocumentName As String = ""
            Dim PageName As String = ""
            Dim NewPageName As String = ""
            If OPCForm.RadioButtonNewPage.Checked Then
                NewPageName = OPCForm.TextBoxNewPageName.Text
            Else
                PageName = OPCForm.ComboBoxExistingPage.SelectedItem.ToString()
            End If

            If OPCForm.CheckBoxODC.Checked Then
                OtherDocumentName = OPCForm.TextBoxFileName.Text
            End If

            Call CreateOffPageConnector(OPCShape, OtherDocumentName, PageName, OPCForm.CheckBoxODC.Checked, NewPageName, OPCShape.Cells("User.OPCType").ResultStr(""))
        End If

        OPCForm.Close()
        OPCForm = Nothing

    End Sub

    Private Sub CreateOffPageConnector(ByRef OPCShape As Visio.Shape, ByVal otherDocumentName As String, ByVal pageName As String, ByVal ODC As Boolean, ByVal newPageName As String, ByVal OPCType As String)

        Dim OPCMaster As Visio.Master
        Dim OPCCopy As Visio.Shape
        Dim FirstDocument As Visio.Document
        Dim OtherDocument As Visio.Document
        Dim PageNames As String() = Nothing
        OPCMaster = Globals.ThisAddIn.Application.Documents.Item("Netdesign.vssx").Masters.Item(OPCType)
        FirstDocument = Globals.ThisAddIn.Application.ActiveDocument

        If ODC Then
            OtherDocument = Globals.ThisAddIn.Application.Documents.Item(otherDocumentName)
        Else
            OtherDocument = FirstDocument
        End If

        ' Drop on existing page or new page or in other document
        If newPageName = "" Then
            OPCCopy = OtherDocument.Pages.Item(pageName).Drop(OPCMaster, 0, 0)
        Else
            OtherDocument.Pages.Add.Name = newPageName
            OPCCopy = OtherDocument.Pages.Item(newPageName).Drop(OPCMaster, 0, 0)
        End If

        ''Set the OPC peer not to be reported, this will propegate through the connected wire.
        'OPCCopy.Cells("User.NotReport").Formula = 1

        'If OPCCopy.Cells("User.OPCType").ResultStr("") = "Wire Bundle" Then
        '    OPCCopy = OPCCopy.Shapes.Item(13)
        'End If

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
            OPCShape.Hyperlinks("OffPageConnector").Address = otherDocumentName
            OPCShape.Cells("User.OPCDDocID").Formula = """" + otherDocumentName + """"
            OPCCopy.Hyperlinks("OffPageConnector").Address = FirstDocument.Path + FirstDocument.Name
            OPCCopy.Cells("User.OPCDDocID").Formula = """" + FirstDocument.Path + FirstDocument.Name + """"
            If OPCShape.Cells("User.OPCType").ResultStr("") = "Patch Panel" Then
                OPCShape.Text = OtherDocument.Name + " : " + OPCCopy.ContainingPage.Name
                OPCCopy.Text = FirstDocument.Name + " : " + OPCShape.ContainingPage.Name
            End If
            OPCShape.CellsU("EventDblClick").Formula = "RUNADDONWARGS(""OPC"",""/CMD=5"")"
            OPCCopy.CellsU("EventDblClick").Formula = "RUNADDONWARGS(""OPC"",""/CMD=5"")"
        Else
            OPCShape.Hyperlinks("OffPageConnector").Address = ""
            OPCCopy.Hyperlinks("OffPageConnector").Address = ""
            If OPCShape.Cells("User.OPCType").ResultStr("") = "Patch Panel" Then
                OPCShape.Text = OPCCopy.ContainingPage.Name
                OPCCopy.Text = OPCShape.ContainingPage.Name
            End If
            OPCShape.CellsU("EventDblClick").Formula = "RUNADDONWARGS(""OPC"",""/CMD=2"")"
            OPCCopy.CellsU("EventDblClick").Formula = "RUNADDONWARGS(""OPC"",""/CMD=2"")"
        End If

    End Sub

    ''' <summary>
    ''' Update the OPC when it is connected to a wire
    ''' </summary>
    ''' <param name="connections">The shapes that is connected</param>
    ''' <remarks></remarks>
    Public Sub UpdateOPC(ByRef OPC As Visio.Shape, ByRef wireShape As Visio.Shape)

        If OPC.Cells("User.OPCType").ResultStr("") = "OPC" Then
            'Call MoveInformation(OPC, wireShape)
            Call SynchOPC(OPC, wireShape)
            Call UpdateText(OPC, wireShape)
        ElseIf OPC.Cells("User.OPCType").ResultStr("") = "Patch Panel Port" Then
            'Call MoveInformation(OPC, wireShape)
            Call SynchOPC(OPC, wireShape)
        ElseIf OPC.Cells("User.OPCType").ResultStr("") = "Wire Bundle" Then
            'Call MoveInformation(OPC, wireShape)
            Call SynchOPC(OPC, wireShape)
            wireShape.Cells("Prop.WireID").Formula = """" + OPC.Cells("User.WireID").ResultStr("") + """"
            'Call UpdateText(OPC, wireShape)
        End If

    End Sub

    

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="OPC"></param>
    ''' <remarks></remarks>
    Private Sub SynchOPC(ByRef OPC As Visio.Shape, ByVal wireShape As Visio.Shape)

        Dim OtherOPC As Visio.Shape
        Dim PortShape As Visio.Shape

        OtherOPC = GetOtherOPC(OPC)

        OPC.Cells("User.MediaType").Formula = """" + wireShape.Cells("Prop.Media").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
        OPC.Cells("User.MediaPurpose").Formula = """" + wireShape.Cells("Prop.Purpose").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
        OPC.Cells("User.MediaSpeed").Formula = """" + wireShape.Cells("Prop.TransmissionSpeed").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"

        OtherOPC.Cells("User.OtherMediaType").Formula = """" + wireShape.Cells("Prop.Media").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
        OtherOPC.Cells("User.OtherMediaPurpose").Formula = """" + wireShape.Cells("Prop.Purpose").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
        OtherOPC.Cells("User.OtherMediaSpeed").Formula = """" + wireShape.Cells("Prop.TransmissionSpeed").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"

        If wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll2D, "Port").Length = 1 Then

            PortShape = wireShape.ContainingPage.Shapes.ItemFromID(wireShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll2D, "Port")(0))

            OPC.Cells("User.RackLocation").Formula = """" + PortShape.Cells("User.RackLocation").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
            OPC.Cells("User.UPosition").Formula = """" + PortShape.Cells("User.UPosition").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
            OPC.Cells("User.SwitchType").Formula = """" + PortShape.Cells("User.SwitchType").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
            OPC.Cells("User.PortName").Formula = """" + PortShape.Cells("User.TextTitle").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
            OPC.Cells("User.SwitchName").Formula = """" + PortShape.Cells("User.SwitchName").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"

            OtherOPC.Cells("User.OtherUPosition").Formula = """" + PortShape.Cells("User.UPosition").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
            OtherOPC.Cells("User.OtherSwitchType").Formula = """" + PortShape.Cells("User.SwitchType").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
            OtherOPC.Cells("User.OtherPortName").Formula = """" + PortShape.Cells("User.TextTitle").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"
            OtherOPC.Cells("User.OtherSwitchName").Formula = """" + PortShape.Cells("User.SwitchName").ResultStr(Visio.VisUnitCodes.visUnitsString) + """"

        End If

        

    End Sub
    ''' <summary>
    ''' Update the text of the OPC's when a wire is being connected
    ''' </summary>
    ''' <param name="OPC">The OPC that the wire is being connected to</param>
    ''' <param name="wireShape">The wire that is being connected to the OPC</param>
    ''' <remarks></remarks>
    Private Sub UpdateText(ByRef OPC As Visio.Shape, ByRef wireShape As Visio.Shape)

        Dim OtherOPC As Visio.Shape
        Dim OtherSwitchName As String
        Dim OtherPortName As String
        Dim SwitchName As String
        Dim PortName As String

        'Get the other OPC
        OtherOPC = GetOtherOPC(OPC)

        'Get the names that will be used, Other means the opposite OPC's. 
        SwitchName = OPC.Cells("User.SwitchName").ResultStr(Visio.VisUnitCodes.visUnitsString)
        PortName = OPC.Cells("User.PortName").ResultStr(Visio.VisUnitCodes.visUnitsString)
        OtherSwitchName = OtherOPC.Cells("User.SwitchName").ResultStr(Visio.VisUnitCodes.visUnitsString)
        OtherPortName = OtherOPC.Cells("User.PortName").ResultStr(Visio.VisUnitCodes.visUnitsString)

        'Change the text of both OPC's
        OPC.Text = OtherOPC.ContainingPage.Name + ":" + OtherSwitchName + ":" + OtherPortName
        OtherOPC.Text = OPC.ContainingPage.Name + ":" + SwitchName + ":" + PortName


        'Check if the otherOPC has a wire connected, update that wire's label if so
        If OtherOPC.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll1D, "").Length = 1 Then
            Dim OtherWireShape As Visio.Shape

            OtherWireShape = OtherOPC.ContainingPage.Shapes.ItemFromID(OtherOPC.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll1D, "")(0))

            Call UpdateLabel(OtherOPC, OtherWireShape)

        End If

    End Sub

    Private Function GetOtherOPC(ByVal OPC As Visio.Shape) As Visio.Shape
        Dim Document As Visio.Document
        Dim Page As Visio.Page
        Dim OtherOPC As Visio.Shape = Nothing
        Dim OPCParent As Visio.Shape
        Dim OtherOPCNumber As Integer

        If OPC.Cells("User.OPCType").ResultStr("") = "Wire Bundle" Then
            OPCParent = OPC.Parent
            OtherOPCNumber = OPC.Cells("Prop.WireNumber").ResultInt("", 1)
        Else
            OPCParent = OPC
        End If

        If OPCParent.Hyperlinks("OffPageConnector").Address.Contains("!") OrElse OPCParent.Hyperlinks("OffPageConnector").Address = "" Then
            Document = Globals.ThisAddIn.Application.ActiveDocument
        Else
            openDocument(Globals.ThisAddIn.Application, OPCParent.Hyperlinks("OffPageConnector").Address, Visio.VisOpenSaveArgs.visOpenMinimized)
            Document = Globals.ThisAddIn.Application.Documents.Item(OPCParent.Hyperlinks("OffPageConnector").Address)
        End If

        Page = GetPageByGUID(Document, OPCParent.Cells("User.OPCDPageID").ResultStr(Visio.VisUnitCodes.visUnitsString))
        Try
            OtherOPC = Page.Shapes.ItemFromUniqueID(OPCParent.Cells("User.OPCDShapeID").ResultStr(Visio.VisUnitCodes.visUnitsString))
        Catch ex As Exception
            System.Console.Write(ex.Message)
        End Try

        If OPCParent.Cells("User.OPCType").ResultStr("") = "Patch Panel Port" Then
            OtherOPC = OtherOPC.Shapes(OPC.Cells("Prop.PortNumber").ResultInt("", 1))
        ElseIf OPCParent.Cells("User.OPCType").ResultStr("") = "Wire Bundle" Then
            OtherOPC = OtherOPC.Shapes(OtherOPCNumber)
        End If

        Return OtherOPC

    End Function

    Public Sub DeleteOtherOPC(ByRef OPC As Visio.Shape)

        Dim OtherOPC As Visio.Shape = Nothing

        OtherOPC = GetOtherOPC(OPC)


        ' To avoid the event to trigger again and create a loop
        If Not OtherOPC Is Nothing Then
            OtherOPC.Cells("User.msvShapeCategories").Formula = """" + "Dead" + """"
            OPC.Cells("User.msvShapeCategories").Formula = """" + "Dead" + """"
            OtherOPC.Delete()
        End If
    End Sub

    Public Sub DeleteAllOPCOnPage(ByRef page As Visio.Page)

        Dim Shape As Visio.Shape

        Dim Counter As Integer = page.Shapes.Count()

        For i As Integer = Counter To 1 Step -1
            Shape = page.Shapes.Item(i)
            If Shape.CellExists("User.msvShapeCategories", 0) Then
                If Shape.Cells("User.msvShapeCategories").ResultStr("") = "OPC" Then
                        Shape.Delete()
                End If
            End If
        Next

    End Sub


End Module