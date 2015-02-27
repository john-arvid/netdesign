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

    'Constants for the shapesheet cells, these will help when doing changes to the shapesheets
    Public Const _UPosition As String = "Prop.UPosition"
    Public Const _ShapeName As String = "Prop.Name"
    Public Const _ShapeModel As String = "Prop.Model"
    Public Const _TransmissionSpeed As String = "Prop.TransmissionSpeed"
    Public Const _MediaType As String = "Prop.Media"
    Public Const _WireID As String = "Prop.WireID"
    Public Const _ShapeCategories As String = "User.msvShapeCategories"
    Public Const _SwitchName As String = "User.SwitchName"
    ' OPC and Port shape 
    Public Const _PortName As String = "User.PortName"
    ' Port and Patch-Panel Port shape
    Public Const _PortNumber As String = "Prop.PortNumber"
    Public Const _Purpose As String = "Prop.Purpose"
    Public Const _RackLocation As String = "User.RackLocation"
    Public Const _Version As String = "User.Version"




#End Region


#Region "Document names"

    Public Const _Stencils As String = "Netdesign.vssx"
    Public Const _HiddenStencils As String = "NetdesignHidden.vssx"




#End Region



End Module