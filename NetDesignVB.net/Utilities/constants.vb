Module constants

    ' The string that is searched for when replacing in a form, to help differ
    'between switch and processor e.g
    Public Const replaceInForm As String = "??"

#Region "ShapeTypeNames"

    ' The different type of shapes
    Public Const shpProcessor As String = "ATLAS_TDAQ_OBJ_Processor"
    Public Const shpUndefined As String = "ATLAS_TDAQ_OBJ_Undefined"
    Public Const shpWire As String = "ATLAS_TDAQ_Wire"
    Public Const shpPort As String = "ATLAS_TDAQ_Port"
    Public Const shpOPC As String = "ATLAS_TDAQ_OBJ_OffPageConnector"
    Public Const shpWireSignalLabel As String = "ATLAS_TDAQ_OBJ_WireSignalLabel"
    Public Const shpWirePortLabel As String = "ATLAS_TDAQ_OBJ_WireLabel"
    Public Const shpDrillUpConnector As String = "ATLAS_TDAQ_OBJ_DrillUpConnector"
    Public Const shpNextPageConnector As String = "ATLAS_TDAQ_OBJ_NextPageConnector"
    Public Const shpPrevPageConnector As String = "ATLAS_TDAQ_OBJ_PrevPageConnector"
    Public Const shpThickLine As String = "ATLAS_TDAQ_OBJ_ThickLine"
    Public Const shpRackAsPage As String = "ATLAS_TDAQ_OBJ_RackAsPage"
    Public Const shpChassisSwitch As String = "ATLAS_TDAQ_OBJ_ChassisSwitch"
    Public Const shpChassisSwitchPageLink As String = "ATLAS_TDAQ_OBJ_ChassisSwitchPageLink"
    Public Const shpSwitchBlade As String = "ATLAS_TDAQ_OBJ_SwitchBlade"
    Public Const shpSwitch As String = "ATLAS_TDAQ_OBJ_Switch"
    Public Const shpSmartGroupChassis As String = "ATLAS_TDAQ_OBJ_SmartGroupChassis"
    Public Const shpSmartGroupChassy As String = "ATLAS_TDAQ_OBJ_SmartGroupChassy"

#End Region

#Region "Shapesheet cells"

    'Constants for the shapesheet cells
    Public Const UPosition As String = "Prop.UPosition"
    Public Const ShapeName As String = "Prop.Name"
    Public Const ShapeModel As String = "Prop.Model"
    Public Const TransmissionSpeed As String = "Prop.TransmissionSpeed"
    Public Const MediaType As String = "Prop.Media"
    Public Const WireID As String = "Prop.WireID"



#End Region

End Module