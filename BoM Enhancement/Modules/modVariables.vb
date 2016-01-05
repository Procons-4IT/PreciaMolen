Public Module modVariables

    Public oApplication As clsListener
    Public strSQL As String
    Public cfl_Text As String
    Public cfl_Btn As String
    Public objSourceForm As SAPbouiCOM.Form
    Public frmSourceMatrix As SAPbouiCOM.Matrix
    Public frmSourceFormUD As String
    Public frmSourceForm As SAPbouiCOM.Form
    Public frmSourcePMForm As SAPbouiCOM.Form
    Public frmSourceQCOR As SAPbouiCOM.Form
    Public CompanyDecimalSeprator As String
    Public CompanyThousandSeprator As String
    Public intSelectedMatrixrow As Integer = 0
    Public Enum ValidationResult As Integer
        CANCEL = 0
        OK = 1
    End Enum

    Public Enum DocumentType As Integer
        RENTAL_QUOTATION = 1
        RENTAL_ORDER
        RENTAL_RETURN
    End Enum
    Public Const frm_BoM As String = "672"
    Public Const frm_ProductionOrder As String = "65211"

    Public Const frm_Phase As String = "frm_Phase"
    Public Const mnu_Phase As String = "PR0002"
    Public Const xml_Phase As String = "frm_Phase.xml"

    Public Const frm_Seq As String = "frm_Seq"
    Public Const mnu_Seq As String = "PR0003"
    Public Const xml_Seq As String = "frm_Seq.xml"


    Public Const frm_AirportMapping As String = "frm_Airport"
    Public Const mnu_Aiport As String = "AV0001"
    Public Const xml_Aiport As String = "xml_AirportVendormapping.xml"

    Public Const frm_BatchSelect As String = "42"
    Public Const frm_BatchSetup As String = "41"

    Public Const frm_Invoice As String = "133"
    Public Const mnu_Grossprofit As String = "5891"
    Public Const frm_Grossprofit As String = "241"
    Public Const frm_Delivery As String = "140"
    Public Const frm_Order As String = "139"

    Public Const frm_QCItemMaster As String = "frm_QCItemMaster"
    Public Const xml_QCItemMaster As String = "frm_QCItemMaster.xml"

    Public Const frm_WAREHOUSES As Integer = 62
    Public Const frm_SalesOrder As Integer = 139
    Public Const frm_ServiceCall As String = "60110"

    Public Const Suppliercode As String = "S000001"
    Public Const frm_BatchOrders As String = "frm_BatchOrders"
    Public Const mnu_FIND As String = "1281"
    Public Const mnu_ADD As String = "1282"
    Public Const mnu_CLOSE As String = "1286"
    Public Const mnu_NEXT As String = "1288"
    Public Const mnu_PREVIOUS As String = "1289"
    Public Const mnu_FIRST As String = "1290"
    Public Const mnu_LAST As String = "1291"
    Public Const mnu_ADD_ROW As String = "1292"
    Public Const mnu_DELETE_ROW As String = "1293"
    Public Const mnu_TAX_GROUP_SETUP As String = "8458"
    Public Const mnu_DEFINE_ALTERNATIVE_ITEMS As String = "11531"
    Public Const mnu_BatchOrders As String = "DABT_411"

    Public Const xml_MENU As String = "Menu.xml"
    Public Const xml_MENU_REMOVE As String = "RemoveMenus.xml"
    Public Const xml_BatchOrders As String = "BatchOrders.xml"


End Module
