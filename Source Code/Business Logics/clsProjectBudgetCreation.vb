Public Class clsProjectbudgetCreation
    Inherits clsBase
    Private oMatrix As SAPbouiCOM.Matrix
    Dim oStatic As SAPbouiCOM.StaticText
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oGrid As SAPbouiCOM.Grid
    Private oCheck As SAPbouiCOM.CheckBoxColumn
    Private oBankGrid As SAPbouiCOM.Grid
    Private oCreditGrid_U As SAPbouiCOM.Grid
    Private oCreditGrid_P As SAPbouiCOM.Grid
    Private ocombo As SAPbouiCOM.ComboBoxColumn

    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim blnFormLoaded As Boolean = False

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_PBWizar, frm_PBWizar)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            ' oApplication.Utilities.setEdittextvalue(oForm, "12", aCode)
            ' oApplication.Utilities.setEdittextvalue(oForm, "13", aSlpCode)
            oForm.PaneLevel = 1
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub addcontrols(ByVal aforma As SAPbouiCOM.Form)
        Try
            oApplication.Utilities.AddControls(oForm, "_301", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", 0, 0, "2", "Generate Barcode", 120)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Function GetSelectedDocuments(aform As SAPbouiCOM.Form) As String
        Dim strDocNum As String = "10000"
        oGrid = aform.Items.Item("7").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oCheck = oGrid.Columns.Item("Select")
            If oCheck.IsChecked(intRow) Then
                strDocNum = strDocNum & "," & oGrid.DataTable.GetValue("DocEntry", intRow)
            End If

        Next
        Return strDocNum
    End Function
    Private Sub PopulateEstimationsDetails(aform As SAPbouiCOM.Form, aChoice As String)
        Dim oRec As SAPbobsCOM.Recordset
        Dim strItem As String
        aform.Freeze(True)
        If aChoice = "Header" Then
            oGrid = aform.Items.Item("7").Specific
            'oGrid.DataTable.ExecuteQuery("SELECT T0.""DocEntry"", T0.""DocNum"", T0.""CreateDate"",T0.""U_Z_PrjCode"",T0.""U_Z_PrjName"",""U_Z_SupPrjCode"",""U_Z_SupPrjName"",""U_Z_GLAcc"",""U_Z_FreeText"",T0.""U_Z_CardCode"",T0.""U_Z_SlpCode"",T0.""U_Z_TotalCost"" ""Total Cost"", T0.""U_Z_Remarks"", ' ' As ""Select"" FROM ""@Z_OQUT""  T0 where T0.""U_Z_DocStatus""='A' and T0.""U_Z_AppStatus""='A' and T0.""U_Z_SlpCode""='" & oApplication.Utilities.getEditTextvalue(aform, "13") & "' and  T0.""U_Z_CardCode""='" & oApplication.Utilities.getEditTextvalue(aform, "12") & "' order by ""DocEntry"" Desc")
            Dim sCode As String = oApplication.Utilities.getEditTextvalue(aform, "13")
            sCode = "SELECT T0.""DocEntry"", T0.""DocNum"", T0.""CreateDate"",T0.""U_Z_PrjCode"",T0.""U_Z_PrjName"",""U_Z_SupPrjCode"",""U_Z_SupPrjName"",""U_Z_GLAcc"",""U_Z_FreeText"",T0.""U_Z_CardCode"",T0.""U_Z_SlpCode"",T0.""U_Z_TotalCost"" ""Total Cost"", T0.""U_Z_Remarks"", 'Y' As ""Select"" FROM ""@Z_OQUT""  T0 where T0.""DocEntry"" in (Select T1.""DocEntry"" from ""@Z_QUT1"" T1 where isnull(Cast(T1.""U_Z_PONO"" as Varchar),'')='') and T0.""U_Z_DocStatus""='A' and T0.""U_Z_AppStatus""='A' and T0.""U_Z_CardCode""='" & oApplication.Utilities.getEditTextvalue(aform, "13") & "' order by ""DocEntry"" Desc"
            ' oGrid.DataTable.ExecuteQuery("SELECT T0.""DocEntry"", T0.""DocNum"", T0.""CreateDate"",T0.""U_Z_PrjCode"",T0.""U_Z_PrjName"",""U_Z_SupPrjCode"",""U_Z_SupPrjName"",""U_Z_GLAcc"",""U_Z_FreeText"",T0.""U_Z_CardCode"",T0.""U_Z_SlpCode"",T0.""U_Z_TotalCost"" ""Total Cost"", T0.""U_Z_Remarks"", 'Y' As ""Select"" FROM ""@Z_OQUT""   T0 where T0.""DocEntry"" in (Select T1.""DocEntry"" from ""@Z_QUT1"" T1 where isnull(Convert(Varchar,T1.""U_Z_PONO""),'')='') and T0.""U_Z_DocStatus""='A' and T0.""U_Z_AppStatus""='A' and T0.""U_Z_PrjCode""='" & oApplication.Utilities.getEditTextvalue(aform, "13") & "' order by ""DocEntry"" Desc")
            If blnIsHana = True Then
                oGrid.DataTable.ExecuteQuery("SELECT T0.""DocEntry"", T0.""DocNum"", T0.""CreateDate"",T0.""U_Z_PrjCode"",T0.""U_Z_PrjName"",""U_Z_SupPrjCode"",""U_Z_SupPrjName"",""U_Z_GLAcc"",""U_Z_FreeText"",T0.""U_Z_CardCode"",T0.""U_Z_SlpCode"",T0.""U_Z_TotalCost"" ""Total Cost"", T0.""U_Z_Remarks"", 'Y' As ""Select"" FROM ""@Z_OQUT""   T0 where   T0.""U_Z_DocStatus""='A' and T0.""U_Z_AppStatus""='A' and T0.""U_Z_PrjCode""='" & oApplication.Utilities.getEditTextvalue(aform, "13") & "' order by ""DocEntry"" Desc")
            Else
                oGrid.DataTable.ExecuteQuery("SELECT T0.""DocEntry"", T0.""DocNum"", T0.""CreateDate"",T0.""U_Z_PrjCode"",T0.""U_Z_PrjName"",""U_Z_SupPrjCode"",""U_Z_SupPrjName"",""U_Z_GLAcc"",""U_Z_FreeText"",T0.""U_Z_CardCode"",T0.""U_Z_SlpCode"",T0.""U_Z_TotalCost"" ""Total Cost"", T0.""U_Z_Remarks"", 'Y' As ""Select"" FROM ""@Z_OQUT""   T0 where   T0.""U_Z_DocStatus""='A' and T0.""U_Z_AppStatus""='A' and T0.""U_Z_PrjCode""='" & oApplication.Utilities.getEditTextvalue(aform, "13") & "' order by ""DocEntry"" Desc")
            End If
            oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Estimation No"
            oGrid.Columns.Item("DocEntry").Visible = False
            oGrid.Columns.Item("DocNum").TitleObject.Caption = "Estimation Number"
            oGrid.Columns.Item("DocNum").Editable = False
            oEditTextColumn = oGrid.Columns.Item("DocNum")
            oEditTextColumn.LinkedObjectType = "2"
            oGrid.Columns.Item("CreateDate").TitleObject.Caption = "Create Date"
            oGrid.Columns.Item("CreateDate").Editable = False
            oGrid.Columns.Item("U_Z_CardCode").TitleObject.Caption = "Customer Code"
            oGrid.Columns.Item("U_Z_CardCode").Editable = False
            oEditTextColumn = oGrid.Columns.Item("U_Z_CardCode")
            oEditTextColumn.LinkedObjectType = "2"
            oGrid.Columns.Item("U_Z_SlpCode").TitleObject.Caption = "Sales Person"
            oGrid.Columns.Item("U_Z_SlpCode").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oComboColumn = oGrid.Columns.Item("U_Z_SlpCode")
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("select ""SlpCode"",""SlpName"" from OSLP order by ""SlpCode""")
            For introw As Integer = 0 To oRecordSet.RecordCount - 1
                oComboColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
                oRecordSet.MoveNext()
            Next
            oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            oGrid.Columns.Item("U_Z_SlpCode").Editable = False
            'oGrid.Columns.Item("U_Z_Desc").TitleObject.Caption = "Description"
            'oGrid.Columns.Item("U_Z_Desc").Editable = False
            oGrid.Columns.Item("U_Z_Remarks").TitleObject.Caption = "Remarks"
            oGrid.Columns.Item("U_Z_Remarks").Editable = False
            oGrid.Columns.Item("U_Z_PrjCode").TitleObject.Caption = "Project Code"
            oGrid.Columns.Item("U_Z_PrjCode").Visible = False
            oGrid.Columns.Item("U_Z_PrjName").TitleObject.Caption = "Project Name"
            oGrid.Columns.Item("U_Z_PrjName").Editable = False
            oGrid.Columns.Item("U_Z_SupPrjCode").TitleObject.Caption = "Sub Project Code"
            oGrid.Columns.Item("U_Z_SupPrjCode").Editable = False
            oGrid.Columns.Item("U_Z_SupPrjName").TitleObject.Caption = "Sub Project Name"
            oGrid.Columns.Item("U_Z_SupPrjName").Editable = False
            oGrid.Columns.Item("U_Z_GLAcc").Visible = False
            oGrid.Columns.Item("U_Z_FreeText").TitleObject.Caption = "Text"
            oGrid.Columns.Item("U_Z_FreeText").Editable = False
            oGrid.Columns.Item("Total Cost").TitleObject.Caption = "Total Cost"
            oGrid.Columns.Item("Total Cost").Editable = False
            oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("Select").TitleObject.Caption = "Select"
            oGrid.Columns.Item("Select").Editable = False
            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
        Else
            Dim strDocNumbers As String = GetSelectedDocuments(aform)
            oGrid = aform.Items.Item("9").Specific
            If blnIsHana = True Then
                strItem = "SELECT T0.""DocEntry"",T1.""DocNum"", T0.""LineId"",T1.""U_Z_PrjCode"" ""ProjectCode"",T1.""U_Z_PrjName"" ""ProjectName"",T1.""U_Z_SupPrjCode"",T1.""U_Z_SupPrjName"" ""Phase"",T0.""U_Z_ItemCode"", T0.""U_Z_ItemDesc"", T0.""U_Z_Price"", T0.""U_Z_Qty"", T0.""U_Z_Total"", 'Y' As ""Select"" FROM ""@Z_QUT1""  T0"
                strItem = strItem & " inner Join ""@Z_OQUT"" T1 on T1.""DocEntry""=T0.""DocEntry"" where  T0.""U_Z_ItemCode""<>'' AND T0.""DocEntry"" in (" & strDocNumbers & ")"
            Else
                strItem = "SELECT T0.""DocEntry"",T1.""DocNum"", T0.""LineId"",T1.""U_Z_PrjCode"" ""ProjectCode"",T1.""U_Z_PrjName"" ""ProjectName"",T1.""U_Z_SupPrjCode"",T1.""U_Z_SupPrjName"" ""Phase"",T0.""U_Z_ItemCode"", T0.""U_Z_ItemDesc"", T0.""U_Z_Price"", T0.""U_Z_Qty"", T0.""U_Z_Total"", 'Y' As ""Select"" FROM ""@Z_QUT1""  T0"
                strItem = strItem & " inner Join ""@Z_OQUT"" T1 on T1.""DocEntry""=T0.""DocEntry"" where   T0.""U_Z_ItemCode""<>'' AND T0.""DocEntry"" in (" & strDocNumbers & ")"
            End If
            oGrid.DataTable.ExecuteQuery(strItem)
            oGrid.Columns.Item("ProjectCode").Editable = False
            oGrid.Columns.Item("ProjectName").Editable = False
            oGrid.Columns.Item("U_Z_SupPrjCode").Editable = False
            oGrid.Columns.Item("Phase").Editable = False

            oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Estimation No"
            oGrid.Columns.Item("DocEntry").Editable = False
            oGrid.Columns.Item("DocNum").Visible = False
            oGrid.Columns.Item("LineId").Visible = False

            oGrid.Columns.Item("U_Z_SupPrjCode").Visible = False
            oGrid.Columns.Item("U_Z_ItemCode").TitleObject.Caption = "Activity"
            oGrid.Columns.Item("U_Z_ItemCode").Editable = False
            oEditTextColumn = oGrid.Columns.Item("U_Z_ItemCode")
            oEditTextColumn.LinkedObjectType = "4"
            oGrid.Columns.Item("U_Z_ItemDesc").TitleObject.Caption = "Activity Description"
            oGrid.Columns.Item("U_Z_ItemDesc").Editable = False

            oGrid.Columns.Item("U_Z_Price").TitleObject.Caption = "Unit Price"
            oGrid.Columns.Item("U_Z_Price").Editable = False

            oGrid.Columns.Item("U_Z_Qty").TitleObject.Caption = "Quantity"
            oGrid.Columns.Item("U_Z_Qty").Editable = False
            oGrid.Columns.Item("U_Z_Total").TitleObject.Caption = "Sales Price"
            oGrid.Columns.Item("U_Z_Total").Editable = False
            oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("Select").TitleObject.Caption = "Select"
            oGrid.Columns.Item("Select").Editable = False
            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
        End If
        aform.Freeze(False)
    End Sub
    Private Sub SelectAll(aForm As SAPbouiCOM.Form, aFlag As Boolean)
        aForm.Freeze(True)
        If aForm.PaneLevel = 2 Then
            oGrid = aForm.Items.Item("7").Specific
        Else
            oGrid = aForm.Items.Item("9").Specific
        End If
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oCheck = oGrid.Columns.Item("Select")
            oCheck.Check(intRow, aFlag)
        Next
        aForm.Freeze(False)

    End Sub
    Private Function AddUDT(ByVal aform As SAPbouiCOM.Form, ByVal aDocEntry As Integer) As Boolean
        Try
            aform.Freeze(True)
            Dim strDocEntry, strLineId, firstName, LastName, strBPCode, strBPName As String
            Dim oRec As SAPbobsCOM.Recordset
            Dim oChild As SAPbobsCOM.GeneralData
            Dim oChildren, ochildern1 As SAPbobsCOM.GeneralDataCollection
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
            Dim oCompanyService As SAPbobsCOM.CompanyService
            oCompanyService = oApplication.Company.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("Z_PRJ")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRec1 As SAPbobsCOM.Recordset
            Dim strPhase, strSubProject, strEstimation, strBoM, strBoMRef, strQuery, strProject, strActivity, strBomLineQuery, strLineNo, strPONo, strDate As String
            Dim dtPostingdate As Date
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oGrid = aform.Items.Item("9").Specific
            strDate = oApplication.Utilities.getEditTextvalue(aform, "17")
            dtPostingdate = oApplication.Utilities.GetDateTimeValue(strDate)
            'Dim sCode As String
            'sCode = "SELECT T0.""DocEntry"", T0.""DocNum"", T0.""CreateDate"",T0.""U_Z_PrjCode"",T0.""U_Z_PrjName"",""U_Z_SupPrjCode"",""U_Z_SupPrjName"",""U_Z_GLAcc"",""U_Z_FreeText"",T0.""U_Z_CardCode"",T0.""U_Z_SlpCode"",T0.""U_Z_TotalCost"" ""Total Cost"", T0.""U_Z_Remarks"", 'Y' As ""Select"" FROM ""@Z_OQUT""  T0 where T0.""DocEntry"" in (Select T1.""DocEntry"" from ""@Z_QUT1"" T1 where isnull(Convert(Varchar,T1.""U_Z_PONO""),'')='') and T0.""U_Z_DocStatus""='A' and T0.""U_Z_AppStatus""='A' and T0.""U_Z_CardCode""='" & oApplication.Utilities.getEditTextvalue(aform, "13") & "' order by ""DocEntry"" Desc"
            'If blnIsHana = True Then
            '    sCode = "SELECT T0.""DocEntry"", T0.""DocNum"", T0.""CreateDate"",T0.""U_Z_PrjCode"",T0.""U_Z_PrjName"",""U_Z_SupPrjCode"",""U_Z_SupPrjName"",""U_Z_GLAcc"",""U_Z_FreeText"",T0.""U_Z_CardCode"",T0.""U_Z_SlpCode"",T0.""U_Z_TotalCost"" ""Total Cost"", T0.""U_Z_Remarks"", 'Y' As ""Select"" FROM ""@Z_OQUT""  T0 where T0.""DocEntry"" in (Select T1.""DocEntry"" from ""@Z_QUT1"" T1 where ifnull(Cast(T1.""U_Z_PONO"" as Varchar),'')='') and T0.""U_Z_DocStatus""='A' and T0.""U_Z_AppStatus""='A' and T0.""U_Z_CardCode""='" & oApplication.Utilities.getEditTextvalue(aform, "13") & "' order by ""DocEntry"" Desc"
            'End If
            'Dim oTemp As SAPbobsCOM.Recordset
            'oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oTemp.DoQuery(sCode)
            Dim strCode As String = oApplication.Utilities.getEditTextvalue(aform, "13") ' oCombobox.Selected.Value
            strProject = strCode
            Dim blnExits As Boolean = False
            oRec.DoQuery("Select * from ""@Z_HPRJ"" where ""U_Z_PRJCODE""='" & strProject & "'")
            If oRec.RecordCount > 0 Then
                aDocEntry = oRec.Fields.Item("DocEntry").Value
                oApplication.Utilities.Message("Project Budget already exists for the selected project.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aform.Freeze(False)
                Return False
                blnExits = True
            Else
                oRec.DoQuery("select * from ONNM where ""ObjectCode""='Z_PRJ'")
                aDocEntry = oRec.Fields.Item("AutoKey").Value
            End If
            Dim strPrjName, strstatus, strBudget, strExpe, strFromdate, strTodate, strApproval As String
            strApproval = "N"
            strstatus = "E"
            Dim strCardCode1, strCardName, strEmpID1, strInternal, strEMPName As String
            strBudget = "0" ' oApplication.Utilities.getEditTextvalue(aform, "8")
            strExpe = "0" ' oApplication.Utilities.getEditTextvalue(aform, "23")
            strFromdate = "" ' oApplication.Utilities.getEditTextvalue(aform, "19")
            strTodate = "" ' oApplication.Utilities.getEditTextvalue(aform, "21")
            strCardCode1 = "" ' oApplication.Utilities.getEditTextvalue(aform, "35")
            strCardName = "" ' oApplication.Utilities.getEditTextvalue(aform, "37")
            ' oCheckBox = aform.Items.Item("48").Specific
            Dim dtFrom, dtTo As Date
            dtFrom = oApplication.Utilities.GetDateTimeValue(strFromdate)
            dtTo = oApplication.Utilities.GetDateTimeValue(strTodate)
            Dim dblBudget, dblExp As Double
            dblBudget = 0 ' oApplication.Utilities.getDocumentQuantity(strBudget)
            dblExp = 0 'oApplication.Utilities.getDocumentQuantity(strExpe)
            oRec.DoQuery("Select * from OPRJ where ""PrjCode""='" & strCode & "'")
            strPrjName = oRec.Fields.Item("PrjName").Value
            dtFrom = oRec.Fields.Item("ValidFrom").Value
            dtTo = oRec.Fields.Item("ValidTo").Value
            If oRec.Fields.Item("U_Z_INTERNAL").Value = "Y" Then
                strInternal = "Y"
            Else
                strInternal = "N"
            End If
            strCardCode1 = oRec.Fields.Item("U_Z_CARDCODE").Value
            strCardName = oRec.Fields.Item("U_Z_CARDNAME").Value
            strEmpID1 = oRec.Fields.Item("U_Z_EMPID").Value ' oApplication.Utilities.getEditTextvalue(aform, "edEmp")
            strEMPName = oRec.Fields.Item("U_Z_EMPNAME").Value ' oApplication.Utilities.getEditTextvalue(aform, "49")
            Dim dblTotalHour As Double = 0 'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEditTextvalue(aform, "edTotHours"))
            Dim dblTotalCost As Double = 0 'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEditTextvalue(aform, "edTotCost"))
            If blnExits = False Then
                oGeneralData.SetProperty("U_Z_CARDCODE", strCardCode1)
                oGeneralData.SetProperty("U_Z_CARDNAME", strCardName)
                oGeneralData.SetProperty("U_Z_EMPID", strEmpID1)
                oGeneralData.SetProperty("U_Z_INTERNAL", strInternal)
                oGeneralData.SetProperty("U_Z_EMPNAME", strEMPName)
                oGeneralData.SetProperty("U_Z_PRJCODE", strCode)
                oGeneralData.SetProperty("U_Z_PRJNAME", strPrjName)
                oGeneralData.SetProperty("U_Z_BUDGET", dblBudget)
                oGeneralData.SetProperty("U_Z_TOTALEXPENSE", dblExp)
                'oGeneralData.SetProperty("U_Z_FROMDATE", dtFrom)
                ' oGeneralData.SetProperty("U_Z_TODATE", dtTo)
                oGeneralData.SetProperty("U_Z_STATUS", strstatus)
                oGeneralData.SetProperty("U_Z_APPROVAL", strApproval)
                oGeneralData.SetProperty("U_Z_TOTHOURS", dblTotalHour)
                oGeneralData.SetProperty("U_Z_TOTCOST", dblTotalCost)
                '  oGeneralData.SetProperty("U_Z_SLPCODE", oApplication.Utilities.getEditTextvalue(aform, "63"))
                ' oGeneralData.SetProperty("U_Z_SLPNAME", oApplication.Utilities.getEditTextvalue(aform, "64"))
                ' oGeneralData.SetProperty("U_Z_CUSTCNTID", oApplication.Utilities.getEditTextvalue(aform, "140"))
                oChildren = oGeneralData.Child("Z_PRJ1")
                For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    strEstimation = oGrid.DataTable.GetValue("DocEntry", intRow)
                    strProject = oGrid.DataTable.GetValue("ProjectCode", intRow)
                    strPhase = oGrid.DataTable.GetValue("Phase", intRow)
                    strActivity = oGrid.DataTable.GetValue("U_Z_ItemCode", intRow)
                    strLineNo = oGrid.DataTable.GetValue("LineId", intRow)
                    oChild = oChildren.Add()
                    oChild.SetProperty("U_Z_MODNAME", strPhase)
                    oChild.SetProperty("U_Z_ACTNAME", strActivity)
                    oChild.SetProperty("U_Z_TYPE", "R")
                    oChild.SetProperty("U_Z_QUANTITY", oGrid.DataTable.GetValue("U_Z_Qty", intRow))
                    oChild.SetProperty("U_Z_ORDER", "N")
                    oChild.SetProperty("U_Z_STATUS", "I")
                Next
                oChildren = oGeneralData.Child("Z_PRJ1")
                For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    strEstimation = oGrid.DataTable.GetValue("DocEntry", intRow)
                    strProject = oGrid.DataTable.GetValue("ProjectCode", intRow)
                    strPhase = oGrid.DataTable.GetValue("Phase", intRow)
                    strActivity = oGrid.DataTable.GetValue("U_Z_ItemCode", intRow)
                    strLineNo = oGrid.DataTable.GetValue("LineId", intRow)
                    oChild = oChildren.Add()
                    oChild.SetProperty("U_Z_MODNAME", strPhase)
                    oChild.SetProperty("U_Z_ACTNAME", strActivity)
                    oChild.SetProperty("U_Z_TYPE", "I")
                    oChild.SetProperty("U_Z_QUANTITY", oGrid.DataTable.GetValue("U_Z_Qty", intRow))
                    oChild.SetProperty("U_Z_ORDER", "N")
                    oChild.SetProperty("U_Z_STATUS", "I")
                    Dim stCode As String
                    stCode = oApplication.Utilities.getMaxCode("@Z_PRJ2", "U_Z_BOQREF")
                    oChild.SetProperty("U_Z_BOQ", stCode)
                    oApplication.Utilities.AddToUDT_Table(strProject, strPrjName, strPhase, strActivity, stCode)
                Next
                oGeneralService.Add(oGeneralData)
            End If
            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            aform.Freeze(False)
            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Return False
        End Try
    End Function

    Private Function AddtoDocument(aform As SAPbouiCOM.Form) As Boolean

        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oProduction As SAPbobsCOM.ProductionOrders
        Dim oRec, oRec1 As SAPbobsCOM.Recordset
        Dim strPhase, strSubProject, strEstimation, strBoM, strBoMRef, strQuery, strProject, strActivity, strBomLineQuery, strLineNo, strPONo, strDate As String
        Dim dtPostingdate As Date
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If AddUDT(oForm, 0) = True Then
            Return True
        Else
            Return False
        End If
        oGrid = aform.Items.Item("9").Specific
        strDate = oApplication.Utilities.getEditTextvalue(aform, "17")
        dtPostingdate = oApplication.Utilities.GetDateTimeValue(strDate)
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            strEstimation = oGrid.DataTable.GetValue("DocEntry", intRow)
            strProject = oGrid.DataTable.GetValue("ProjectCode", intRow)
            strPhase = oGrid.DataTable.GetValue("Phase", intRow)
            strActivity = oGrid.DataTable.GetValue("U_Z_ItemCode", intRow)
            strLineNo = oGrid.DataTable.GetValue("LineId", intRow)
            If blnIsHana = True Then
                strQuery = "select T0.""U_Z_Code"" ,T1.""U_Z_ItemCode"" ,T1.""U_Z_ItemName"" ,T1.""U_Z_BaseQty"",ifnull(T1.""U_Z_BoMRef"",'') ""BoMRef"" from ""@Z_OPRPH"" T0 Inner Join ""@Z_PRPH1"" T1 on T1.""DocEntry""=T0.""DocEntry"""
            Else
                strQuery = "select T0.""U_Z_Code"" ,T1.""U_Z_ItemCode"" ,T1.""U_Z_ItemName"" ,T1.""U_Z_BaseQty"",isnull(T1.""U_Z_BoMRef"",'') ""BoMRef"" from ""@Z_OPRPH"" T0 Inner Join ""@Z_PRPH1"" T1 on T1.""DocEntry""=T0.""DocEntry"""
            End If
            strQuery = strQuery & " where T0.""U_Z_Code""='" & strActivity & "'"
            oRec.DoQuery(strQuery)
            strPONo = ""
            For intMain As Integer = 0 To oRec.RecordCount - 1

                Dim blnLineExits As Boolean = False
                oProduction = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                oProduction.UserFields.Fields.Item("U_Z_PrjCode").Value = strProject
                oProduction.UserFields.Fields.Item("U_Z_SubPrj").Value = strPhase
                oProduction.UserFields.Fields.Item("U_Z_Phase").Value = strActivity
                oProduction.UserFields.Fields.Item("U_Z_EstNo").Value = strEstimation
                oProduction.Project = strProject
                oProduction.PlannedQuantity = oGrid.DataTable.GetValue("U_Z_Qty", intRow)
                oProduction.ItemNo = oRec.Fields.Item("U_Z_ItemCode").Value
                oProduction.PostingDate = dtPostingdate
                oProduction.DueDate = dtPostingdate.AddMonths(1)
                oProduction.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
                If oRec.Fields.Item("BoMRef").Value <> "" Then
                    strBomLineQuery = "Select ""U_Z_Type"",""U_Z_ItemCode"",""U_Z_BaseQty"",""U_Z_Cost"",""U_Z_WhsCode"",""U_Z_UoM"",""U_Z_PlnList"" from ""@Z_PRPH2"" where ""U_Z_PHRef""='" & oRec.Fields.Item("BoMRef").Value & "'"
                Else
                    strBomLineQuery = "Select * from ITT1 where ""Father""='" & oRec.Fields.Item("U_Z_ItemCode").Value & "'"
                    strBomLineQuery = "select ""Type"",""Code"",""Quantity"",""OrigPrice"",""Warehouse"",""Uom"",""PriceList""  from ITT1  where ""Father""='" & oRec.Fields.Item("U_Z_ItemCode").Value & "'"
                End If
                oRec1.DoQuery(strBomLineQuery)
                For intloop As Integer = 0 To oRec1.RecordCount - 1
                    If intloop > 0 Then
                        oProduction.Lines.Add()
                    End If
                    oProduction.Lines.SetCurrentLine(intloop)
                    If oRec1.Fields.Item(0).Value = "4" Then
                        oProduction.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                    ElseIf oRec1.Fields.Item(0).Value = "290" Then
                        oProduction.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                    End If
                    oProduction.Lines.ItemNo = oRec1.Fields.Item(1).Value
                    oProduction.Lines.BaseQuantity = oRec1.Fields.Item(2).Value
                    oProduction.Lines.PlannedQuantity = oGrid.DataTable.GetValue("U_Z_Qty", intRow)
                    oProduction.Lines.Warehouse = oRec1.Fields.Item(4).Value
                    blnLineExits = True
                    oRec1.MoveNext()
                Next
                If blnLineExits = True Then
                    If oProduction.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    Else
                        Dim strDocno As String
                        oApplication.Company.GetNewObjectCode(strDocno)
                        If strPONo = "" Then
                            strPONo = strDocno
                        Else
                            strPONo = strPONo & " ," & strDocno
                        End If
                    End If
                End If
                oRec1.DoQuery("Update ""@Z_QUT1"" set ""U_Z_PONO""='" & strPONo & "' where ""DocEntry""=" & strEstimation & " and ""LineId""=" & strLineNo)

                oRec.MoveNext()
            Next
        Next
        Return True

    End Function
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_PBWizar Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "7" And pVal.ColUID = "DocNum" Then
                                    Dim oobj As New clsProjectEstimation
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    oComboColumn = oGrid.Columns.Item("U_Z_SlpCode")
                                    oobj.LoadForm_View(oGrid.DataTable.GetValue("DocNum", pVal.Row), oComboColumn.GetSelectedValue(pVal.Row).Value)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                If pVal.ItemUID = "9" And pVal.ColUID = "U_Z_ItemCode" Then
                                    oGrid = oForm.Items.Item("9").Specific
                                    Dim oobj As New clsProjectPhase
                                    oobj.LoadForm_View(oGrid.DataTable.GetValue("U_Z_ItemCode", pVal.Row))
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "4"
                                        If oApplication.Utilities.getEditTextvalue(oForm, "15") = "" Then
                                            If oForm.PaneLevel = "1" Then
                                                oApplication.Utilities.Message("Project code missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                        If oApplication.Utilities.getEditTextvalue(oForm, "17") = "" Then
                                            'If oForm.PaneLevel = "1" Then
                                            '    oApplication.Utilities.Message("Posting Date  missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            '    BubbleEvent = False
                                            '    Exit Sub
                                            'End If
                                            '  oApplication.Utilities.setEdittextvalue(oForm,"17",Now.Date)

                                        End If
                                End Select
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                blnFormLoaded = True
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "3"
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                    Case "4"
                                        If oApplication.Utilities.getEditTextvalue(oForm, "15") = "" Then

                                        Else

                                            oForm.PaneLevel = oForm.PaneLevel + 1
                                            If oForm.PaneLevel = 2 Then
                                                PopulateEstimationsDetails(oForm, "Header")
                                            End If
                                            If oForm.PaneLevel = 3 Then
                                                PopulateEstimationsDetails(oForm, "Trans")
                                            End If
                                        End If
                                    Case "10"
                                        SelectAll(oForm, True)
                                    Case "11"
                                        SelectAll(oForm, False)

                                    Case "5"
                                        If oApplication.SBO_Application.MessageBox("Do you want to create the Project Budget Details?", , "Continue", "Cancel") = 2 Then
                                            Exit Sub
                                        End If
                                        If AddtoDocument(oForm) = True Then
                                            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            'oForm.PaneLevel = 2
                                            'PopulateEstimationsDetails(oForm, "Header")
                                            oForm.Close()
                                        End If
                                End Select

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim strValue As String
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects
                                    'If 1 = 1 Then 'pVal.ItemUID = "4" Or pVal.ItemUID = "6" Then
                                    '    strValue = oDataTable.GetValue(CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).ChooseFromListAlias, 0)
                                    '    Try
                                    '        oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                    '    Catch ex As Exception
                                    '        oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                    '    End Try
                                       
                                    '    End If


                                   
                                If pVal.ItemUID = "15" Then
                                    strValue = oDataTable.GetValue("PrjCode", 0)
                                    If pVal.ItemUID = "15" Then
                                        oApplication.Utilities.setEdittextvalue(oForm, "13", strValue)
                                        oApplication.Utilities.setEdittextvalue(oForm, "15", strValue)
                                    End If
                                    End If
                                Catch ex As Exception

                                End Try


                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_PBWizard
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If

            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Data Event"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Validations"
    Private Function Validation(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim strFBank, strTBank As String

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

#End Region

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub















#End Region
End Class
