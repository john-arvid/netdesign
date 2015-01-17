'// DemoAddAdvise.vb
'// <copyright>Copyright (c) Microsoft Corporation.  All rights reserved.
'// </copyright>
'// <summary>This module demonstrates how to use AddAdvise to tell Visio what
'// events this is interested in listening to.</summary>

Imports System

Module DemoAddAdviseSample

    '// <summary>The EventSink class will sink the events added in this
    '// class. It will write event information to the debug output window.
    '// </summary>
    Public eventHandler As EventSink

    '// <summary>This procedure uses the AddAdvise method on an event list
    '// to tell Visio which events to monitor and which object will handle
    '// those events. This procedure also shows how to use the AddAdvise
    '// method at both the application and document levels, each of which
    '// has its own EventList collection.</summary>
    '// <param name="theApplication">Reference to the Visio application
    '// object </param>
    Public Class AddAdvise

        Public Sub DemoAddAdvise( _
        ByVal theApplication As Microsoft.Office.Interop.Visio.Application)

            ' Declare visEvtAdd as a 2-byte value to avoid a run-time overflow
            ' error.
            Const visEvtAdd As Short = -32768

            Dim eventsDocument As Microsoft.Office.Interop.Visio.EventList
            Dim eventsApplication As Microsoft.Office.Interop.Visio.EventList
            Dim addedDocument As Microsoft.Office.Interop.Visio.Document

            Try
                eventHandler = New EventSink

                'openDocument(theApplication, "\\cern.ch\dfs\Users\j\jkibsgaa\Documents\Drawing1.vsdx")

                


                'Call openDocument(theApplication, "G:\Projects\ATLASTDAQNetworking\TDAQ Connectivity\0_NetDesign Visio\New Tool\Stencils\Netdesign.vssx", Visio.VisOpenSaveArgs.visAddDocked + Visio.VisOpenSaveArgs.visOpenRO)
                'Call openDocument(theApplication, "G:\Projects\ATLASTDAQNetworking\TDAQ Connectivity\0_NetDesign Visio\New Tool\Stencils\NetdesignHidden.vssx", Visio.VisOpenSaveArgs.visOpenHidden + Visio.VisOpenSaveArgs.visOpenRO)
                

                


                ' Get the EventList collection of the application.


                eventsApplication = theApplication.EventList


                '' Get the EventList collection of this document.
                'eventsDocument = Globals.ThisAddIn.Application.ActiveDocument.EventList

                Globals.ThisAddIn.Application.EventsEnabled = True
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Add events for which Visio will send notification to
                ' the EventSink class.

                ' Add the ShapeAdded event.
                eventsApplication.AddAdvise(CShort(Microsoft.Office.Interop.Visio. _
                   VisEventCodes.visEvtShape + CShort(visEvtAdd)), _
                  eventHandler, "", "ShapeAdded")

                '' Add the cells formula changed event.
                'eventsApplication.AddAdvise(CShort(Microsoft.Office.Interop.Visio.VisEventCodes.visEvtFormula + Visio.VisEventCodes.visEvtMod), eventHandler, "", "CellChanged")

                ' Add the connection made event.
                eventsApplication.AddAdvise(CShort(visEvtAdd + Visio.VisEventCodes.visEvtConnect), eventHandler, "", "ConnectionMade")

                ' Add the connection removed event.
                eventsApplication.AddAdvise(CShort(Visio.VisEventCodes.visEvtDel + Visio.VisEventCodes.visEvtConnect), eventHandler, "", "ConnectionRemoved")

                ' Add the BeforeShapeDelete event.
                eventsApplication.AddAdvise(CShort(Microsoft.Office.Interop.Visio.VisEventCodes.visEvtDel + Microsoft.Office.Interop.Visio.VisEventCodes.visEvtShape), eventHandler, "", "BeforeShapeDelete")

                ' Add the before page deleted event.
                eventsApplication.AddAdvise(CShort(Visio.VisEventCodes.visEvtDel + Visio.VisEventCodes.visEvtPage), eventHandler, "", "BeforePageDelete")

                ' Add the markerevent. This is used when a shapesheet call is made.
                eventsApplication.AddAdvise(CShort(Visio.VisEventCodes.visEvtMarker + Visio.VisEventCodes.visEvtApp), eventHandler, "", "Marker")

                ' Add the PageAdded event.
                eventsApplication.AddAdvise(CShort(Microsoft.Office.Interop.Visio.VisEventCodes.visEvtPage + visEvtAdd), eventHandler, "", "PageAdded")

                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                '' Add the QueryCancelSelectionDelete event.
                'eventsDocument.AddAdvise(CShort(Microsoft.Office.Interop.Visio. _
                '    VisEventCodes.visEvtCodeQueryCancelSelDel), _
                '    eventHandler, "", "")



                '' Add the BeforeDocumentClose event.
                'eventsDocument.AddAdvise(CShort(Microsoft.Office.Interop.Visio. _
                '    VisEventCodes.visEvtDel + _
                '    Microsoft.Office.Interop.Visio.VisEventCodes.visEvtDoc), _
                '    eventHandler, "", "")

                '' Add the ApplicationQuit event.
                'eventsApplication.AddAdvise(CShort(Microsoft.Office.Interop.Visio. _
                '    VisEventCodes.visEvtApp + _
                '    Microsoft.Office.Interop.Visio.VisEventCodes. _
                '    visEvtBeforeQuit), _
                '    eventHandler, "", "")

                '' Add the WindowTurnToPage event.
                'eventsApplication.AddAdvise(CShort(Microsoft.Office.Interop.Visio. _
                '    VisEventCodes.visEvtCodeWinPageTurn), _
                '    eventHandler, "", "")

                '' Add the Validate ruleset event
                'eventsApplication.AddAdvise(CShort(Microsoft.Office.Interop.Visio. _
                '    VisEventCodes.visEvtCodeRuleSetValidated), _
                '    eventHandler, "", "")


            Catch err As System.Runtime.InteropServices.COMException
                System.Diagnostics.Debug.WriteLine(err.Message)
            End Try

        End Sub

        
    End Class

    Public Sub DeleteEvent()

    End Sub

    Public Sub AddEventAgain()
        'Globals.ThisAddIn.Application.EventList.AddAdvise(CShort(Microsoft.Office.Interop.Visio.VisEventCodes.visEvtFormula + Visio.VisEventCodes.visEvtMod), eventHandler, "", "")
        Globals.ThisAddIn.Application.EventsEnabled = True
        ' Add the ShapeAdded event.
        Globals.ThisAddIn.Application.EventList.AddAdvise(CShort(Microsoft.Office.Interop.Visio.VisEventCodes.visEvtShape + CShort(-32768)), eventHandler, "", "")
    End Sub
End Module
