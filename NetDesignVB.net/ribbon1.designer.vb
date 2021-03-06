﻿Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.ReportButton = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.UpdateShapesButton = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.ReportButton)
        Me.Group1.Items.Add(Me.UpdateShapesButton)
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'Button1
        '
        Me.Button1.Label = "Unique"
        Me.Button1.Name = "Button1"
        '
        'ReportButton
        '
        Me.ReportButton.Label = "Create Report"
        Me.ReportButton.Name = "ReportButton"
        '
        'Button3
        '
        Me.Button3.Label = "Magic"
        Me.Button3.Name = "Button3"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button1)
        Me.Group2.Items.Add(Me.Button3)
        Me.Group2.Label = "Group2"
        Me.Group2.Name = "Group2"
        Me.Group2.Visible = False
        '
        'UpdateShapesButton
        '
        Me.UpdateShapesButton.Label = "Update the shapes"
        Me.UpdateShapesButton.Name = "UpdateShapesButton"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Visio.Drawing"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ReportButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents UpdateShapesButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
