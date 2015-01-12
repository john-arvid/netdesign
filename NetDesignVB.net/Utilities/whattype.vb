Module whatType

    ' Check what kind of shape has been added to the page and the call the
    'correct handle function. This uses the constant names in constants.vb

    ' This can be done in two ways, either check the masters name (stencil)
    'or check the Shapesheet prompt value of ObjType. I choose now to
    'use the prompt value.
    ' After some testing I found out that I needed to check another thing to
    'differ in a good way. Some shapes has a chassis/chassy and the shape also
    'have another cell that need to be checked.
    'So first I check the ObjType, and the the SmartGroupType if the ObjType
    'is of a chassis/chassy.
    'That is why I have a case inside a case, later if possible I could do an
    'architectural design change that would make this easier.
    'TODO maybe use mastername?


    Public Sub checkType(ByVal shape As Visio.Shape)


        '' If this cell exist then the SmartGroupType will to, need to change 
        ''this whole thing if I do any architectural design changes.
        'If (shape.CellExists("User.ATLAS_TDAQ_ObjType", 0)) Then

        '    ' Do a check for every known type
        '    Select Case (shape.Cells("User.ATLAS_TDAQ_ObjType.Prompt").ResultStr(""))

        '        ' If it is a chassis/chassy then it have to be one of these
        '        Case shpSmartGroupChassis, shpSmartGroupChassy

        '            Select Case (shape.Cells("User.SmartGroupType.Prompt").ResultStr(""))
        '                Case shpSwitch
        '                    If Not shape.Master Is Nothing Then
        '                        If shape.Master.NameU = "Switch" Then
        '                            Call handleSwitch(shape)
        '                        End If
        '                    End If
        '                Case shpSwitchBlade
        '                    MsgBox(shape.Cells("User.SmartGroupType.Prompt").ResultStr(""))

        '                Case shpProcessor
        '                    'MsgBox(shape.Cells("User.SmartGroupType.Prompt").ResultStr(""))
        '                    If Not shape.Master Is Nothing Then
        '                        If shape.Master.NameU = "Processor" Then
        '                            Call HandleProcessor(shape)
        '                        End If
        '                    End If

        '                Case shpChassisSwitch
        '                    MsgBox(shape.Cells("User.SmartGroupType.Prompt").ResultStr(""))
        '            End Select

        '        Case shpChassisSwitchPageLink
        '            MsgBox(shape.Cells("User.ATLAS_TDAQ_ObjType.Prompt").ResultStr(""))

        '        Case shpDrillUpConnector
        '            MsgBox(shape.Cells("User.ATLAS_TDAQ_ObjType.Prompt").ResultStr(""))

        '        Case shpNextPageConnector
        '            MsgBox(shape.Cells("User.ATLAS_TDAQ_ObjType.Prompt").ResultStr(""))

        '        Case shpOPC
        '            MsgBox(shape.Cells("User.ATLAS_TDAQ_ObjType.Prompt").ResultStr(""))

        '        Case shpPort
        '            'MsgBox(shape.NameU)

        '        Case shpPrevPageConnector
        '            MsgBox(shape.Cells("User.ATLAS_TDAQ_ObjType.Prompt").ResultStr(""))

        '        Case shpRackAsPage
        '            MsgBox(shape.Cells("User.ATLAS_TDAQ_ObjType.Prompt").ResultStr(""))

        '        Case shpThickLine
        '            MsgBox(shape.Cells("User.ATLAS_TDAQ_ObjType.Prompt").ResultStr(""))

        '        Case shpUndefined
        '            MsgBox(shape.Cells("User.ATLAS_TDAQ_ObjType.Prompt").ResultStr(""))

        '        Case shpThickLine
        '            MsgBox(shape.Cells("User.ATLAS_TDAQ_ObjType.Prompt").ResultStr(""))

        '        Case shpWire

        '        Case shpWirePortLabel
        '            MsgBox(shape.Cells("User.ATLAS_TDAQ_ObjType.Prompt").ResultStr(""))

        '        Case shpWireSignalLabel
        '            MsgBox(shape.Cells("User.ATLAS_TDAQ_ObjType.Prompt").ResultStr(""))

        '            'Case Else
        '            'MsgBox("Not found: " + shape.Cells("User.ATLAS_TDAQ_ObjType.Prompt").ResultStr(""))

        '    End Select

        'End If

        

        If Not shape.Master Is Nothing Then

            Select Case shape.Master.Name

                Case "Switch"
                    Call handleSwitch(shape)

                Case "Wire"
                    Call handleWire(shape)

                Case "Rack"
                    Call ValidateRack(shape)

                Case "Chassis Switch"
                    Call HandleChassisSwitch(shape)

                Case "Chassis Processor"
                    Call HandleChassisSwitch(shape)

                Case "Processor"
                    Call HandleProcessor(shape)

            End Select
        End If

        If shape.CellExists("User.msvShapeCategories", 0) Then
            If shape.Cells("User.msvShapeCategories").ResultStr("") = "OPC" Then
                Call HandleOffPageConnector(shape)
            End If
        End If

    End Sub

End Module