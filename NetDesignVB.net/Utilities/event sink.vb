'// EventSink.vb
'// <copyright>Copyright (c) Microsoft Corporation.  All rights reserved.
'// </copyright>
'// <summary>This file contains the implementation of EventSink class.</summary>

Imports System

'// <summary>This class is an event sink for Visio events. It handles event
'// notification by implementing the IVisEventProc interface, which is
'// defined in the Visio type library. In order to be notified of events,
'// an instance of this class must be passed as the eventSink argument in
'// calls to the AddAdvise method.
'//
'// This class demonstrates how to handle event notification for events
'// raised by at least one of Application, Document, and/or Page objects.
'// This class also demonstrates how to respond to QueryCancel events.
'</summary>

' Needed to add <System.Runtime.InteropServices.ComVisible(True)> to avoid an
' exception, I don't know why and internet does not know either. (It is 
' supposed to be true by default)
' I guess it's just that darn microsoft!
<System.Runtime.InteropServices.ComVisible(True)> Public Class EventSink
    Implements Microsoft.Office.Interop.Visio.IVisEventProc

    '// <summary>visEvtAdd is declared as a 2-byte value to avoid a run-time
    '// overflow error.</summary>
    Private Const visEvtAdd As Short = -32768

    Private Const eventSinkCaption As String = "Event Sink"
    Private Const tab As String = "\t"
    Private eventDescriptions As  _
    System.Collections.Specialized.StringDictionary

    '// <summary>The constructor initializes the event descriptions 
    '// dictionary.</summary>
    Public Sub New()

        initializeStrings()
    End Sub

    '// <summary>This method is called by Visio when an event in the
    '// EventList collection has been triggered. This method is an
    '// implementation of IVisEventProc.VisEventProc method.</summary>
    '// <param name="eventCode">Event code of the event that fired</param>
    '// <param name="source">Reference to source of the event</param>
    '// <param name="eventId">Unique identifier of the event object that 
    '// fired</param>
    '// <param name="eventSequenceNumber">Relative position of the event in 
    '// the event list</param>
    '// <param name="subject">Reference to the subject of the event</param>
    '// <param name="moreInformation">Additional information for the event
    '// </param>
    '// <returns>False to allow a QueryCancel operation or True to cancel 
    '// a QueryCancel operation. The return value is ignored by Visio unless 
    '// the event is a QueryCancel event.</returns>
    '// <seealso cref="Microsoft.Office.Interop.Visio.IVisEventProc"></seealso>
    Public Function VisEventProc( _
        ByVal eventCode As Short, _
        ByVal source As Object, _
        ByVal eventId As Integer, _
        ByVal eventSequenceNumber As Integer, _
        ByVal subject As Object, _
        ByVal moreInformation As Object) As Object _
        Implements Microsoft.Office.Interop.Visio.IVisEventProc.VisEventProc

        Dim message As String = ""
        Dim name As String = ""
        Dim eventInformation As String = ""
        Dim returnValue As Object = True

        Dim subjectApplication As Microsoft.Office.Interop.Visio.Application = Nothing
        'Dim subjectDocument As Microsoft.Office.Interop.Visio.Document
        Dim subjectPage As Microsoft.Office.Interop.Visio.Page
        'Dim subjectMaster As Microsoft.Office.Interop.Visio.Master
        'Dim subjectSelection As Microsoft.Office.Interop.Visio.Selection
        Dim subjectShape As Microsoft.Office.Interop.Visio.Shape
        Dim subjectCell As Microsoft.Office.Interop.Visio.Cell
        Dim subjectConnects As Microsoft.Office.Interop.Visio.Connects
        'Dim subjectStyle As Microsoft.Office.Interop.Visio.Style
        'Dim subjectWindow As Microsoft.Office.Interop.Visio.Window
        'Dim subjectMouseEvent As Microsoft.Office.Interop.Visio.MouseEvent
        'Dim subjectKeyboardEvent As Microsoft.Office.Interop.Visio.KeyboardEvent
        'Dim subjectDataRecordset As Microsoft.Office.Interop.Visio.DataRecordset
        'Dim subjectDataRecordsetChangedEvent As Microsoft.Office.Interop.Visio.DataRecordsetChangedEvent
        'Dim subjectRelatedPairEvent As Microsoft.Office.Interop.Visio.RelatedShapePairEvent
        'Dim subjectMovedSelectionEvent As Microsoft.Office.Interop.Visio.MovedSelectionEvent
        Dim subjectValidationRuleSet As Microsoft.Office.Interop.Visio.ValidationRuleSet

        Try

            Select Case (eventCode)

                ' Shape added event
                Case CShort(Visio.VisEventCodes.visEvtShape) + visEvtAdd
                    subjectShape = CType(subject, Microsoft.Office. _
                       Interop.Visio.Shape)
                    Call checkType(subjectShape)


                    ' Cell event codes
                Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    visEvtFormula) + _
                    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    visEvtMod)

                    ' Subject object is a Cell
                    subjectCell = CType(subject, Microsoft.Office.Interop. _
                        Visio.Cell)
                    subjectShape = subjectCell.Shape
                    subjectApplication = subjectCell.Application
                    'name = subjectShape.Name + "!" + subjectCell.Name
                    Call FormulaChanged(subjectShape, subjectCell)


                    ' connection made between shapes
                Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    visEvtConnect) + visEvtAdd


                    ' Subject object is a Connects collection
                    subjectConnects = CType(subject, Microsoft.Office. _
                        Interop.Visio.Connects)
                    subjectApplication = subjectConnects.Application
                    Call ConnectionChanged(subjectConnects, True)

                    ' connection removed between shapes
                Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                            visEvtConnect) + _
                            CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                            visEvtDel)

                    ' Subject object is a Connects collection
                    subjectConnects = CType(subject, Microsoft.Office. _
                        Interop.Visio.Connects)
                    Call ConnectionChanged(subjectConnects, False)


                    'Validation Event code
                Case CShort(Global.Microsoft.Office.Interop.Visio.VisEventCodes. _
                    visEvtCodeRuleSetValidated)

                    'Subject is ValidationRuleSet object
                    subjectValidationRuleSet = CType(subject, Global.Microsoft.Office.Interop. _
                        Visio.ValidationRuleSet)
                    Call validateRules(subjectValidationRuleSet)

                    ' Shape has been deleted
                Case CShort(Microsoft.Office.Interop.Visio. _
                    VisEventCodes.visEvtDel + _
                    Microsoft.Office.Interop.Visio.VisEventCodes.visEvtShape)

                    subjectShape = CType(subject, Microsoft.Office. _
                           Interop.Visio.Shape)

                    If subjectShape.CellExists(_ShapeCategories, 0) Then
                        If subjectShape.Cells(_ShapeCategories).ResultStr("") = "OPC" Then
                            Call DeleteOtherOPC(subjectShape)
                        ElseIf subjectShape.Cells(_ShapeCategories).ResultStr("") = "Switch" Then
                            Call DeletePointingOPC(subjectShape)
                        ElseIf subjectShape.Cells(_ShapeCategories).ResultStr("") = "Chassis Switch" Then
                            Call DeleteChassisPages(subjectShape)
                        ElseIf subjectShape.Cells(_ShapeCategories).ResultStr("") = "Chassis Processor" Then
                            Call DeleteChassisPages(subjectShape)
                        End If
                    End If


                    ' Before a page is deleted
                Case CShort(Visio.VisEventCodes.visEvtDel + Visio.VisEventCodes.visEvtPage)
                    ' Subject object is a Page
                    subjectPage = CType(subject, Microsoft.Office.Interop. _
                        Visio.Page)

                    ' Delete all the OPC on the page, so all the OPC peers are being deleted
                    Call DeleteAllOPCOnPage(subjectPage)


                Case CShort(Visio.VisEventCodes.visEvtMarker + Visio.VisEventCodes.visEvtApp)

                    ' Subject object is an Application
                    ' EventInfo is empty for most of these events.  However for
                    ' the Marker event, the EnterScope event and the ExitScope 
                    ' event eventinfo contains the context string. 
                    subjectApplication = CType(subject, Microsoft.Office. _
                        Interop.Visio.Application)

                    Call MarkerHandler(subjectApplication)

                    ' Page added event
                Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes.visEvtPage) + visEvtAdd

                    ' Subject object is a Page
                    subjectPage = CType(subject, Microsoft.Office.Interop.Visio.Page)
                    subjectApplication = subjectPage.Application
                    name = subjectPage.Name

                    Call PreparePage(subjectPage)


                    '    ' Document event codes
                    'Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtDoc) + _
                    'CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtDel), _
                    'CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtCodeBefDocSave), _
                    'CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtCodeBefDocSaveAs), _
                    'CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtCodeDocDesign), _
                    'CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtDoc) + visEvtAdd, _
                    'CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtDoc) + _
                    'CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtMod), _
                    'CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtCodeCancelDocClose), _
                    'CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtCodeDocCreate), _
                    'CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtCodeDocOpen), _
                    'CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtCodeDocSave), _
                    'CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtCodeDocSaveAs), _
                    'CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtCodeDocRunning), _
                    'CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtCodeQueryCancelDocClose), _
                    'CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    'visEvtRemoveHiddenInformation)

                    '    ' Subject object is a Document
                    '    '   Eventinfo may be non empty. 
                    '    '  (1) For DocumentChanged Event it may indicate what 
                    '    '   changed, e.g.  /pagereordered.                      
                    '    '  (2) For the Save and SaveAs related events, the eventInfo may contain
                    '    '   version information with the format /version=X where X is the file 
                    '    '   version number.  For SaveAs events, the eventInfo also contains the 
                    '    '   the full path for the save as action, in the format /saveasfile=
                    '    '   where the full path directly follows the equal sign.  If the save  
                    '    '   action is the result AutoSave, then the string /saveasfile= will be 
                    '    '   replaced by /autosavefile=.
                    '    '  (3) For RemoveHiddenInformation, the eventInfo
                    '    '   indicates the data that were removed. The various types 
                    '    '   are represented by the following strings: 
                    '    '   /visRHIPersonalInfo, /visRHIMasters, /visRHIStyles,
                    '    '   /visRHIDataRecordsets, /visRHIValidationRules. The /visRHIStyles
                    '    '   string appears when themes, datagraphics or styles were removed.
                    '    subjectDocument = CType(subject,  _
                    '        Microsoft.Office.Interop.Visio.Document)
                    '    subjectApplication = subjectDocument.Application
                    '    name = subjectDocument.Name

                    '    ' Page event codes
                    'Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtPage) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtDel), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtPage) + visEvtAdd, _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtPage) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtMod), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeCancelPageDel), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeQueryCancelPageDel)


                    '    ' Subject object is a Page
                    '    subjectPage = CType(subject, Microsoft.Office.Interop. _
                    '        Visio.Page)
                    '    subjectApplication = subjectPage.Application
                    '    name = subjectPage.Name

                    '    ' Master event codes
                    'Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtMaster) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtDel), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtMaster) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtMod), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeCancelMasterDel), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtMaster) + visEvtAdd, _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeQueryCancelMasterDel)

                    '    ' Subject object is a Master
                    '    subjectMaster = CType(subject, Microsoft.Office. _
                    '        Interop.Visio.Master)
                    '    subjectApplication = subjectMaster.Application
                    '    name = subjectMaster.Name

                    '    ' Selection event codes
                    'Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeBefSelDel), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeSelAdded), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeCancelSelDel), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeCancelConvertToGroup), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeQueryCancelUngroup), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeQueryCancelConvertToGroup), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeQueryCancelSelDel), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeCancelUngroup), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '     visEvtCodeQueryCancelSelGroup), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeCancelSelGroup)

                    '    ' Subject object is a Selection
                    '    subjectSelection = CType(subject, Microsoft.Office. _
                    '        Interop.Visio.Selection)
                    '    subjectApplication = subjectSelection.Application



                    '    ' Shape event codes
                    'Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtShape) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtDel), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeShapeBeforeTextEdit), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtShape) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtMod), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeShapeExitTextEdit), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeShapeParentChange), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtText) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtMod), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtShapeDataGraphicChanged)

                    '    ' Subject object is a Shape 
                    '    ' EventInfo is normally empty but for ShapeChanged events
                    '    ' it indicates what changed. The possible EventInfo strings
                    '    ' are: /name, /data1, /data2, /data3, /uniqueid, /ink, /listorder. 
                    '    ' The /ink string will only appear when ink strokes are added/deleted
                    '    ' from an ink shape.
                    '    ' The /listorder string will only appear for a list Shape when its list 
                    '    ' members are re-ordered. 
                    '    subjectShape = CType(subject, Microsoft.Office. _
                    '       Interop.Visio.Shape)
                    '    subjectApplication = subjectShape.Application
                    '    name = subjectShape.Name

                    'Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtShapeLinkAdded), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtShapeLinkDeleted)

                    '    ' Subject object is a Shape
                    '    '  EventInfo provides the recordset ID and rowID 
                    '    '  participating in the link as  
                    '    ' /DataRecordsetID=<ID> and /DataRowID=<ID2>
                    '    subjectShape = CType(subject, Microsoft.Office. _
                    '    Interop.Visio.Shape)
                    '    subjectApplication = subjectShape.Application
                    '    name = subjectShape.Name




                    '    ' Style event codes
                    'Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtStyle) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtDel), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtStyle) + visEvtAdd, _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtStyle) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtMod), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeCancelStyleDel), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeQueryCancelStyleDel)

                    '    ' Subject object is a Style
                    '    subjectStyle = CType(subject, Microsoft.Office. _
                    '        Interop.Visio.Style)
                    '    subjectApplication = subjectStyle.Application
                    '    name = subjectStyle.Name

                    '    ' Window event codes
                    'Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtWindow) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtDel), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeBefWinPageTurn), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtWindow) + visEvtAdd, _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtWindow) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtMod), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeWinPageTurn), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeBefWinSelDel), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeCancelWinClose), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtApp) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtWinActivate), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeWinSelChange), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeViewChanged), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeQueryCancelWinClose)

                    '    ' Subject object is a Window
                    '    subjectWindow = CType(subject, Microsoft.Office. _
                    '        Interop.Visio.Window)
                    '    subjectApplication = subjectWindow.Application
                    '    name = subjectWindow.Caption

                    '    ' Application event codes
                    'Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtApp) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtAfterModal), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeAfterResume), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtApp) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtAppActivate), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtApp) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtAppDeactivate), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtApp) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtObjActivate), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtApp) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtObjDeactivate), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtApp) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtBeforeModal), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtApp) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtBeforeQuit), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeBeforeSuspend), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeEnterScope), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeExitScope), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtApp) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtMarker), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeBefForcedFlush), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeAfterForcedFlush), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtApp) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtNonePending), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeWinOnAddonKeyMSG), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeQueryCancelQuit), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeQueryCancelSuspend), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeCancelQuit), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeCancelSuspend), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtApp) + _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtIdle)

                    '    ' Subject object is an Application
                    '    ' EventInfo is empty for most of these events.  However for
                    '    ' the Marker event, the EnterScope event and the ExitScope 
                    '    ' event eventinfo contains the context string. 
                    '    subjectApplication = CType(subject, Microsoft.Office. _
                    '        Interop.Visio.Application)

                    '    ' Mouse Event 
                    'Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeMouseDown), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeMouseMove), _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeMouseUp)

                    '    ' Subject is MouseEvent object
                    '    ' Note Mouse events can also be canceled. 
                    '    ' EventInfo may be non-empty for MouseMove events, and 
                    '    ' contains information about the DragState which is also
                    '    ' exposed as a property on the MouseEvent object.                        
                    '    subjectMouseEvent = CType(subject, Microsoft.Office. _
                    '    Interop.Visio.MouseEvent)

                    '    ' Keyboard Event		
                    'Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '     visEvtCodeKeyDown), _
                    '     CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '     visEvtCodeKeyPress), _
                    '     CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '     visEvtCodeKeyUp)

                    '    ' Subject is KeyboardEvent object
                    '    ' Note KeyboardEvents can also be canceled. 
                    '    subjectKeyboardEvent = CType(subject, Microsoft.Office. _
                    '    Interop.Visio.KeyboardEvent)

                    '    ' Data Recordset Event codes

                    '    ' Data Recordset event with DataRecordset object
                    'Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtDataRecordset) + visEvtAdd, _
                    '    CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtDataRecordset) + CShort(Microsoft.Office.Interop. _
                    '    Visio.VisEventCodes.visEvtDel)

                    '    ' Subject object is a DataRecordset
                    '    subjectDataRecordset = CType(subject, Microsoft.Office. _
                    '     Interop.Visio.DataRecordset)

                    '    ' Data Recordset events with DataRecordsetChangedEvent Object
                    'Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtMod) + CShort(Microsoft.Office.Interop.Visio. _
                    '    VisEventCodes.visEvtDataRecordset)

                    '    ' Subject is DataRecordsetChangedEvent Object
                    '    subjectDataRecordsetChangedEvent = CType(subject, Microsoft.Office. _
                    '    Interop.Visio.DataRecordsetChangedEvent)

                    '    'Relationship Event codes -- containers & callouts
                    'Case CShort(Global.Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeContainerRelationshipAdded), _
                    '    CShort(Global.Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeContainerRelationshipDeleted), _
                    '    CShort(Global.Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeCalloutRelationshipAdded), _
                    '    CShort(Global.Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeCalloutRelationshipDeleted)

                    '    ' Subject is RelatedShapePairEvent Object
                    '    ' For Container Events, the FromShapeID is the ID of the container shape 
                    '    ' and the ToShapeID is the ID of the member Shape.
                    '    ' For Callout Events, the FromShapeID is the ID of callout shape and 
                    '    ' the ToShapeID is the ID of the target shape.
                    '    subjectRelatedPairEvent = CType(subject, Global.Microsoft.Office.Interop. _
                    '        Visio.RelatedShapePairEvent)

                    '    'Subprocess event code
                    'Case CShort(Global.Microsoft.Office.Interop.Visio.VisEventCodes. _
                    '    visEvtCodeSelectionMovedToSubprocess)

                    '    'Subject is MovedSelectionEvent Object
                    '    subjectMovedSelectionEvent = CType(subject, Global.Microsoft.Office.Interop. _
                    '        Visio.MovedSelectionEvent)



                Case Else
                    name = "Unknown"
                    subjectApplication = Nothing
            End Select

            ' get a description for this event code
            message = getEventDescription(eventCode)

            ' append the name of the subject object
            If (name.Length > 0) Then
                message += ": " + name
            End If

            ' append event info when it is available
            If (Not subjectApplication Is Nothing) Then

                eventInformation = subjectApplication.EventInfo( _
                    Microsoft.Office.Interop.Visio.VisEventCodes. _
                    visEvtIdMostRecent)

                If (Not eventInformation Is Nothing) Then
                    message += tab + eventInformation
                End If
            End If

            ' append moreInformation when it is available
            If (Not moreInformation Is Nothing) Then
                message += tab + moreInformation.ToString()
            End If

            ' get the targetArgs string from the event object. targetArgs
            ' are added to the event object in the AddAdvise method
            Dim events As Microsoft.Office.Interop.Visio.EventList
            Dim thisEvent As Microsoft.Office.Interop.Visio.Event
            Dim sourceType As String
            Dim targetArgs As String

            sourceType = source.GetType().FullName
            If (sourceType = _
                "Microsoft.Office.Interop.Visio.ApplicationClass") Then

                events = CType(source, Microsoft.Office.Interop.Visio. _
                Application).EventList
            ElseIf (sourceType = _
                "Microsoft.Office.Interop.Visio.DocumentClass") Then

                events = CType(source, Microsoft.Office.Interop.Visio. _
                Document).EventList
            ElseIf (sourceType = _
                "Microsoft.Office.Interop.Visio.PageClass") Then

                events = CType(source, Microsoft.Office.Interop.Visio. _
                Page).EventList
            Else
                events = Nothing
            End If

            If (Not events Is Nothing) Then

                thisEvent = events.ItemFromID(eventId)
                targetArgs = thisEvent.TargetArgs

                ' append targetArgs when it is available
                If (targetArgs.Length > 0) Then
                    message += " " + targetArgs
                End If
            End If

            ' Write the event info to the output window
            System.Diagnostics.Debug.WriteLine(message)

            ' if this is a QueryCancel event then prompt the user
            returnValue = getQueryCancelResponse(eventCode, subject)

        Catch err As System.Runtime.InteropServices.COMException
            System.Diagnostics.Debug.WriteLine(err.Message)
        End Try


        'GC.Collect()

        Return returnValue
    End Function

    '// <summary>This method prompts the user to continue or cancel. If the
    '// alertResponse value is set in this Visio instance then its value 
    '// will be used and the dialog will be suppressed.</summary>
    '// <param name="eventCode">Event code of the event that fired</param>
    '// <param name="subject">Reference to subject of the event</param>
    '// <returns>False to allow the QueryCancel operation or True to cancel 
    '// the QueryCancel operation.</returns>
    Private Shared Function getQueryCancelResponse( _
        ByVal eventCode As Short, _
        ByVal subject As Object) As Object

        Const docCloseCancelPrompt As String = _
            "Are you sure you want to close the document?"
        Const pageDeleteCancelPrompt As String = _
            "Are you sure you want to delete the page?"
        Const masterDeleteCancelPrompt As String = _
            "Are you sure you want to delete the master?"
        Const ungroupCancelPrompt As String = _
            "Are you sure you want to ungroup the selected shapes?"
        Const convertToGroupCancelPrompt As String = _
            "Are you sure you want to convert the selected shapes to a group?"
        Const selectionDeleteCancelPrompt As String = _
            "Are you sure you want to delete the selected shapes?"
        Const styleDeleteCancelPrompt As String = _
            "Are you sure you want to delete the style?"
        Const windowCloseCancelPrompt As String = _
            "Are you sure you want to close the window?"
        Const quitCancelPrompt As String = _
            "Are you sure you want to quit Visio?"
        Const suspendCancelPrompt As String = _
            "Are you sure you want to suspend Visio?"
        Const cancelGroupPrompt As String = _
            "Are you sure you want to group the selected shapes?"

        Dim subjectApplication As Microsoft.Office.Interop.Visio.Application
        Dim subjectDocument As Microsoft.Office.Interop.Visio.Document
        Dim subjectPage As Microsoft.Office.Interop.Visio.Page
        Dim subjectMaster As Microsoft.Office.Interop.Visio.Master
        Dim subjectSelection As Microsoft.Office.Interop.Visio.Selection
        Dim subjectStyle As Microsoft.Office.Interop.Visio.Style
        Dim subjectWindow As Microsoft.Office.Interop.Visio.Window

        Dim prompt As String
        Dim subjectName As String
        Dim alertResponse As Short
        Dim isQueryCancelEvent As Boolean = True
        Dim returnValue As Object = False

        Select Case (eventCode)

            ' Query Document Close
            Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                visEvtCodeQueryCancelDocClose)

                subjectDocument = CType(subject, Microsoft.Office.Interop. _
                    Visio.Document)
                subjectName = subjectDocument.Name
                subjectApplication = subjectDocument.Application
                prompt = docCloseCancelPrompt + System.Environment. _
                NewLine + subjectName

                ' Query Cancel Page Delete
            Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                visEvtCodeQueryCancelPageDel)

                subjectPage = CType(subject, Microsoft.Office.Interop. _
                    Visio.Page)
                subjectName = subjectPage.NameU
                subjectApplication = subjectPage.Application
                prompt = pageDeleteCancelPrompt + System.Environment. _
                    NewLine + subjectName

                ' Query Cancel Master Delete
            Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                visEvtCodeQueryCancelMasterDel)

                subjectMaster = CType(subject, Microsoft.Office.Interop. _
                    Visio.Master)
                subjectName = subjectMaster.NameU
                subjectApplication = subjectMaster.Application
                prompt = masterDeleteCancelPrompt + System.Environment. _
                    NewLine + subjectName

                ' Query Cancel Ungroup
            Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                visEvtCodeQueryCancelUngroup)

                subjectSelection = CType(subject, Microsoft.Office. _
                    Interop.Visio.Selection)
                subjectApplication = subjectSelection.Application
                prompt = ungroupCancelPrompt

                ' Query Cancel Convert To Group
            Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                visEvtCodeQueryCancelConvertToGroup)

                subjectSelection = CType(subject, Microsoft.Office. _
                    Interop.Visio.Selection)
                subjectApplication = subjectSelection.Application
                prompt = convertToGroupCancelPrompt

                ' Query Cancel Selection Delete
            Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                visEvtCodeQueryCancelSelDel)

                subjectSelection = CType(subject, Microsoft.Office. _
                    Interop.Visio.Selection)
                subjectApplication = subjectSelection.Application
                prompt = selectionDeleteCancelPrompt

                ' Query Cancel Style Delete
            Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                visEvtCodeQueryCancelStyleDel)

                subjectStyle = CType(subject, Microsoft.Office.Interop. _
                    Visio.Style)
                subjectName = subjectStyle.NameU
                subjectApplication = subjectStyle.Application
                prompt = styleDeleteCancelPrompt + System.Environment. _
                    NewLine + subjectName

                ' Query Cancel Window Close
            Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                visEvtCodeQueryCancelWinClose)

                subjectWindow = CType(subject, Microsoft.Office.Interop. _
                    Visio.Window)
                subjectName = subjectWindow.Caption()
                subjectApplication = subjectWindow.Application
                prompt = windowCloseCancelPrompt + System.Environment. _
                    NewLine + subjectName

                ' Query Cancel Quit
            Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                visEvtCodeQueryCancelQuit)

                subjectApplication = CType(subject, Microsoft.Office. _
                    Interop.Visio.Application)
                prompt = quitCancelPrompt

                ' Query Cancel Suspend
            Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                visEvtCodeQueryCancelSuspend)

                subjectApplication = CType(subject, Microsoft.Office. _
                    Interop.Visio.Application)
                prompt = suspendCancelPrompt

                ' Query Cancel Group
            Case CShort(Microsoft.Office.Interop.Visio.VisEventCodes. _
                visEvtCodeQueryCancelSelGroup)

                subjectSelection = CType(subject, Microsoft.Office. _
                    Interop.Visio.Selection)
                subjectApplication = subjectSelection.Application
                prompt = cancelGroupPrompt

            Case Else
                ' This event is not cancelable.
                isQueryCancelEvent = False
                subjectApplication = Nothing
                prompt = ""
        End Select

        If (isQueryCancelEvent = True) Then

            ' check for an alertResponse setting in Visio
            If (Not subjectApplication Is Nothing) Then
                alertResponse = subjectApplication.AlertResponse
            End If

            If (alertResponse <> 0) Then

                ' if alertResponse is No or Cancel then cancel this event
                ' by returning true
                If ((alertResponse = System.Windows.Forms.DialogResult.No) _
                    Or (alertResponse = System.Windows.Forms.DialogResult. _
                    Cancel)) Then

                    returnValue = True
                End If
            Else

                ' alertResponse is not set so prompt the user
                Dim result As System.Windows.Forms.DialogResult

                result = System.Windows.Forms.MessageBox.Show(prompt, _
                    eventSinkCaption, _
                    System.Windows.Forms.MessageBoxButtons.YesNo, _
                    System.Windows.Forms.MessageBoxIcon.Question)

                If (result = System.Windows.Forms.DialogResult.No) Then
                    returnValue = True
                End If
            End If
        End If

        getQueryCancelResponse = returnValue
    End Function

    '// <summary>This method adds an event description to the
    '// eventDescriptions dictionary.</summary>
    '// <param name="eventCode">Event code of the event</param>
    '// <param name="description">Short description of the event</param>
    Private Sub addEventDescription( _
        ByVal eventCode As Short, _
        ByVal description As String)

        Dim key As String
        key = Convert.ToString(eventCode, _
            System.Globalization.CultureInfo.InvariantCulture)
        eventDescriptions.Add(key, description)
    End Sub

    '// <summary>This method returns a short description for the given
    '// eventCode.</summary>
    '// <param name="eventCode">Event code</param>
    '// <returns>Short description of the eventCode</returns>
    Private Function getEventDescription(ByVal eventCode As Short) As String
        Dim description As String
        Dim key As String

        key = Convert.ToString(eventCode, _
            System.Globalization.CultureInfo.InvariantCulture)
        description = eventDescriptions(key)

        If (description Is Nothing) Then
            description = "NoEventDescription"
        End If
        getEventDescription = description
    End Function

    '// <summary>This method populates the eventDescriptions dictionary
    '// with a short description of each Visio event code.</summary>
    Private Sub initializeStrings()

        eventDescriptions = _
            New System.Collections.Specialized.StringDictionary

        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtApp) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtAfterModal), "AfterModal")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeAfterResume), "AfterResume")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtApp) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtAppActivate), "AppActivated")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtApp) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtAppDeactivate), "AppDeactivated")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtApp) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtObjActivate), "AppObjActivated")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtApp) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtObjDeactivate), "AppObjDeactivated")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtDoc) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtDel), "BeforeDocumentClose")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeBefDocSave), "BeforeDocumentSave")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeBefDocSaveAs), "BeforeDocumentSaveAs")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtMaster) + CShort(Microsoft.Office.Interop. _
         Visio.VisEventCodes.visEvtDel), "BeforeMasterDelete")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtApp) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtBeforeModal), "BeforeModal")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtPage) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtDel), "BeforePageDelete")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtApp) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtBeforeQuit), "BeforeQuit")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeBefSelDel), "BeforeSelectionDelete")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtShape) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtDel), "BeforeShapeDelete")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeShapeBeforeTextEdit), _
            "BeforeShapeTextEdit")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtStyle) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtDel), "BeforeStyleDelete")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeBeforeSuspend), "BeforeSuspend")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtWindow) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtDel), "BeforeWindowClose")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeBefWinPageTurn), "BeforeWindowPageTurn")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeBefWinSelDel), "BeforeWindowSelDelete")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCell) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtMod), "CellChanged")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtConnect) + visEvtAdd, "ConnectionsAdded")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtConnect) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtDel), "ConnectionsDeleted")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeCancelConvertToGroup), _
            "ConvertToGroupCanceled")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeDocDesign), "DesignModeEntered")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtDoc) + visEvtAdd, "DocumentAdded")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtDoc) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtMod), "DocumentChanged")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeCancelDocClose), "DocumentCloseCanceled")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeDocCreate), "DocumentCreated")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeDocOpen), "DocumentOpened")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeDocSave), "DocumentSaved")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeDocSaveAs), "DocumentSavedAs")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeEnterScope), "EnterScope")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeExitScope), "ExitScope")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtFormula) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtMod), "FormulaChanged")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeKeyDown), "KeyDown")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeKeyPress), "KeyPress")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeKeyUp), "KeyUp")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtMaster) + visEvtAdd, "MasterAdded")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtApp) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtMarker), "MarkerEvent")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtMaster) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtMod), "MasterChanged")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeCancelMasterDel), "MasterDeleteCanceled")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeMouseDown), "MouseDown")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeMouseMove), "MouseMove")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeMouseUp), "MouseUp")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeBefForcedFlush), _
            "MustFlushScopeBeginning")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeAfterForcedFlush), "MustFlushScopeEnded")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtApp) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtNonePending), "NoEventsPending")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeWinOnAddonKeyMSG), _
            "OnKeystrokeMessageForAddon")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtPage) + visEvtAdd, "PageAdded")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtPage) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtMod), "PageChanged")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeCancelPageDel), "PageDeleteCanceled")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeQueryCancelConvertToGroup), _
            "QueryCancelConvertToGroup")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeQueryCancelDocClose), _
            "QueryCancelDocumentClose")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeQueryCancelMasterDel), _
            "QueryCancelMasterDelete")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeQueryCancelPageDel), _
            "QueryCancelPageDelete")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeQueryCancelQuit), "QuerCancelQuit")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeQueryCancelSelDel), _
            "QueryCancelSelectionDelete")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeQueryCancelStyleDel), _
            "QueryCancelStyleDelete")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeQueryCancelSuspend), _
            "QueryCancelSuspend")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeQueryCancelUngroup), _
            "QueryCancelUngroup")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeQueryCancelWinClose), _
            "QueryCancelWindowClose")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeCancelQuit), "QuitCanceled")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeDocRunning), "RunModeEntered")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeSelAdded), "SelectionAdded")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeWinSelChange), "SelectionChanged")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeCancelSelDel), "SelectionDeleteCanceled")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtShape + visEvtAdd), "ShapeAdded")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtShape) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtMod), "ShapeChanged")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeShapeExitTextEdit), "ShapeExitedTextEdit")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeShapeParentChange), "ShapeParentChanged")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeShapeDelete), "ShapesDeleted")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtStyle) + visEvtAdd, "StyleAdded")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtStyle) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtMod), "StyleChanged")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeCancelStyleDel), "StyleDeleteCanceled")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeCancelSuspend), "SuspendCanceled")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtText) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtMod), "TextChanged")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeCancelUngroup), "UngroupCanceled")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeViewChanged), "ViewChanged")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtApp) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtIdle), "VisioIsIdle")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtApp) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtWinActivate), "WindowActivated")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeCancelWinClose), "WindowCloseCanceled")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtWindow) + visEvtAdd, "WindowOpened")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtWindow) + CShort(Microsoft.Office.Interop. _
            Visio.VisEventCodes.visEvtMod), "WindowChanged")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeWinPageTurn), "WindowTurnedToPage")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtShapeDataGraphicChanged), "ShapeDataGraphicChanged")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtShapeLinkAdded), "ShapeLinkAdded")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtShapeLinkDeleted), "ShapeLinkDeleted")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtRemoveHiddenInformation), "RemoveHiddenInformation")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeCancelSelGroup), "GroupCanceled")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeQueryCancelSelGroup), "QueryCancelGroup")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtDataRecordset) + visEvtAdd, "DataRecordsetAdded")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtDataRecordset) + CShort(Microsoft.Office. _
            Interop.Visio.VisEventCodes.visEvtDel), "BeforeDataRecordsetDelete")
        addEventDescription(CShort(Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtDataRecordset) + CShort(Microsoft.Office. _
            Interop.Visio.VisEventCodes.visEvtMod), "DataRecordsetChanged")
        addEventDescription(CShort(Global.Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeCalloutRelationshipAdded), "CalloutRelationshipAdded")
        addEventDescription(CShort(Global.Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeCalloutRelationshipDeleted), "CalloutRelationshipDeleted")
        addEventDescription(CShort(Global.Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeContainerRelationshipAdded), "ContainerRelationshipAdded")
        addEventDescription(CShort(Global.Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeContainerRelationshipDeleted), "ContainerRelationshipDeleted")
        addEventDescription(CShort(Global.Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeSelectionMovedToSubprocess), "SelectionMovedToSubProcess")
        addEventDescription(CShort(Global.Microsoft.Office.Interop.Visio. _
            VisEventCodes.visEvtCodeRuleSetValidated), "RuleSetValidated")
    End Sub

End Class

