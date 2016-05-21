Public Module modVariables
    Public oApplication As clsListener
    Public strSQL As String
    Public cfl_Text As String
    Public cfl_Btn As String
    Public CompanyDecimalSeprator As String
    Public CompanyThousandSeprator As String
    Public frmSourceMatrix As SAPbouiCOM.Matrix
    Public frm_SourceBoM As SAPbouiCOM.Form
    Public frm_SourceBoM1 As SAPbouiCOM.Form
    Public frm_SourceQuotation As SAPbouiCOM.Form
    Public sPath, strSelectedFilepath, strSelectedFolderPath As String
    Public frm_SourceProjectPhase As SAPbouiCOM.Form
    Public frm_SourceProjectPhase1 As SAPbouiCOM.Form
    Public frm_ProjectPhaseRow As Integer = 0
    Public frm_ProjectPhaseRow1 As Integer = 0
    Public blnIsHana As Boolean = False
  

    Public Enum ValidationResult As Integer
        CANCEL = 0
        OK = 1
    End Enum
    Public Const frm_BoM As String = "672"
    Public Const frm_WAREHOUSES As Integer = 62
    Public Const frm_ITEM_MASTER As Integer = 150
    Public Const frm_INVOICES As Integer = 133
    Public Const frm_GRPO As Integer = 143
    Public Const frm_ORDR As Integer = 139
    Public Const frm_GR_INVENTORY As Integer = 721
    Public Const frm_Project As Integer = 711
    Public Const frm_ProdReceipt As Integer = 65214
    Public Const frm_GoodsIssue As String = "720"
    Public Const frm_Delivery As Integer = 140
    Public Const frm_SaleReturn As Integer = 180
    Public Const frm_ARCreditMemo As Integer = 179
    Public Const frm_Customer As Integer = 134
    Public Const frm_SalesQuation As String = "149"
    Public Const frm_PurchaseRequest As String = "1470000200"

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

    Public Const xml_MENU As String = "Menu.xml"
    Public Const xml_MENU_REMOVE As String = "RemoveMenus.xml"

    Public Const mnu_BarCode As String = "Menu_B01"
    Public Const xml_BarCode As String = "frm_BarCode.xml"
    Public Const frm_BarCode As String = "frm_BarCode"

    Public Const mnu_OSCL As String = "Menu_02"
    Public Const xml_OSCL As String = "frm_OSCL.xml"
    Public Const frm_OSCL As String = "frm_OSCL"

    Public Const mnu_OPRT As String = "Menu_03"
    Public Const xml_OPRT As String = "frm_OPRT.xml"
    Public Const frm_OPRT As String = "frm_OPRT"

    Public Const mnu_BoM_Template As String = "Menu_P012"
    Public Const frm_BoM_Template As String = "frm_S01"
    Public Const xml_BoM_Template As String = "frm_BomApproval.xml"


    Public Const mnu_SubProject As String = "Menu_P02"
    Public Const frm_SubProject As String = "Z_OSUP"
    Public Const xml_SubProject As String = "Z_OSUP.xml"

    Public Const mnu_ProjectPhase As String = "Menu_P03"
    Public Const frm_ProjectPhase As String = "Z_OPRPH"
    Public Const xml_ProjectPhase As String = "Z_OPRPH.xml"

    Public Const frm_BOMRef As String = "Z_PRPH2"
    Public Const xml_BOMRef As String = "Z_PRPH2.xml"

    Public Const frm_BOMRef1 As String = "Z_PRPH3"
    Public Const xml_BOMRef1 As String = "Z_PRPH3.xml"

    Public Const mnu_ProjectEstimation As String = "Menu_P04"
    Public Const frm_Estimation As String = "Z_OQUT"
    Public Const xml_Estimation As String = "Z_OQUT.xml"


    Public Const mnu_BoM_Estimation As String = "Menu_P034"
    Public Const frm_BoM_Estimation As String = "frm_P02"
    Public Const xml_BoM_Estimation As String = "xml_BoMEstimation.xml"

    Public Const mnu_BoM_Approval As String = "Menu_P05"
    Public Const frm_BoM_Approval As String = "frm_P03"
    Public Const xml_BoM_Approval As String = "xml_BoMApproval.xml"


    Public Const frm_BoM_Wizard As String = "frm_P014"
    Public Const xml_BoM_Wizard As String = "xml_BoMWizard.xml"

    Public Const frm_BoM_Wizard_PO As String = "frm_P114"
    Public Const xml_BoM_Wizard_PO As String = "xml_BoMWizard_PO.xml"

    Public Const frm_BoM_Summary As String = "frm_S05"
    Public Const xml_BoM_Summary As String = "xml_BoMSummary.xml"

    Public Const frm_AppHisDetails As String = "frm_AppHisDetails"
    Public Const xml_AppHisDetails As String = "xml_AppHisDetails.xml"

    Public Const mnu_POWizard As String = "Menu_P07"
    Public Const frm_PO_Wizard As String = "frm_P015"
    Public Const xml_PO_Wizard As String = "xml_POWizard.xml"

    Public Const mnu_PBWizard As String = "Menu_P06"
    Public Const frm_PBWizar As String = "frm_P016"
    Public Const xml_PBWizar As String = "xml_PBWizard.xml"

End Module
