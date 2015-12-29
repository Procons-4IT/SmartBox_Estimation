Public Class clsBoMApproval
    Inherits clsBase

    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBaseDocNo, strQuery As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Dim oRec As SAPbobsCOM.Recordset

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub


    Public Sub LoadForm()
        Try

            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_BoM_Approval) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oForm = oApplication.Utilities.LoadForm(xml_BoM_Approval, frm_BoM_Approval)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            HeaderGridBind(oForm)
            HeaderSumGridBind(oForm)
            oForm.PaneLevel = 1
            oForm.Items.Item("5").TextStyle = 7
            oForm.Items.Item("1000001").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oGrid = oForm.Items.Item("4").Specific
            oGrid.Columns.Item("RowsHeader").Click(0, False, False)
            oGrid = oForm.Items.Item("9").Specific
            oGrid.Columns.Item("RowsHeader").Click(0, False, False)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub HeaderGridBind(ByVal aForm As SAPbouiCOM.Form)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oGrid = aForm.Items.Item("4").Specific
        Try
            strQuery = "Select DocEntry,DocNum,CreateDate,U_Z_CardCode,U_Z_SlpCode,U_Z_Desc,U_Z_DocStatus,U_Z_AppStatus,U_Z_Remarks,U_Z_CurApprover,U_Z_NxtApprover,Creator from [@Z_OQUT] where U_Z_AppStatus='P' and U_Z_DocStatus='C' and ( U_Z_CurApprover='" & oApplication.Company.UserName & "' OR U_Z_NxtApprover='" & oApplication.Company.UserName & "')"
            oGrid.DataTable.ExecuteQuery(strQuery)
            FormatHeadGrid(aForm, "Header", "4")
            assignMatrixLineno(oGrid, aForm)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub HeaderSumGridBind(ByVal aForm As SAPbouiCOM.Form)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oGrid = aForm.Items.Item("9").Specific
        Try
            strQuery = "Select DocEntry,DocNum,CreateDate,U_Z_CardCode,U_Z_SlpCode,U_Z_Desc,U_Z_DocStatus,U_Z_AppStatus,U_Z_Remarks,U_Z_CurApprover,U_Z_NxtApprover,Creator from [@Z_OQUT] where U_Z_DocStatus='C' and ( U_Z_CurApprover='" & oApplication.Company.UserName & "' OR U_Z_NxtApprover='" & oApplication.Company.UserName & "')"
            oGrid.DataTable.ExecuteQuery(strQuery)
            FormatHeadGrid(aForm, "Header", "9")
            assignMatrixLineno(oGrid, aForm)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub FormatHeadGrid(ByVal aForm As SAPbouiCOM.Form, ByVal aChoice As String, ByVal gridId As String)
        Dim oGridCombo As SAPbouiCOM.ComboBoxColumn
        Try
            If aChoice = "Header" Then
                oGrid = aForm.Items.Item(gridId).Specific
                oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Document Number"
                oGrid.Columns.Item("DocEntry").Editable = False
                oGrid.Columns.Item("DocNum").TitleObject.Caption = "Estimation Number"
                oGrid.Columns.Item("DocNum").Editable = False
                oEditTextColumn = oGrid.Columns.Item("DocNum")
                oEditTextColumn.LinkedObjectType = "2"
                oGrid.Columns.Item("U_Z_CardCode").TitleObject.Caption = "Customer Code"
                oGrid.Columns.Item("U_Z_CardCode").Editable = False
                oEditTextColumn = oGrid.Columns.Item("U_Z_CardCode")
                oEditTextColumn.LinkedObjectType = "2"
                Dim oComboColumn As SAPbouiCOM.ComboBoxColumn
                Dim oRecordset As SAPbobsCOM.Recordset

                oGrid.Columns.Item("U_Z_SlpCode").TitleObject.Caption = "Sales Person"
                oGrid.Columns.Item("U_Z_SlpCode").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                oComboColumn = oGrid.Columns.Item("U_Z_SlpCode")
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery("select SlpCode,SlpName  from OSLP order by SlpCode")
                For introw As Integer = 0 To oRecordSet.RecordCount - 1
                    oComboColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
                    oRecordSet.MoveNext()
                Next
                oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description

                oGrid.Columns.Item("U_Z_SlpCode").Editable = False

                oGrid.Columns.Item("CreateDate").TitleObject.Caption = "Document Date"
                oGrid.Columns.Item("CreateDate").Editable = False
                oGrid.Columns.Item("U_Z_Desc").TitleObject.Caption = "BoM Description"
                oGrid.Columns.Item("U_Z_Desc").Editable = False
                oGrid.Columns.Item("U_Z_DocStatus").TitleObject.Caption = "Document Status"
                oGrid.Columns.Item("U_Z_DocStatus").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                oGridCombo = oGrid.Columns.Item("U_Z_DocStatus")
                oGridCombo.ValidValues.Add("D", "Draft")
                oGridCombo.ValidValues.Add("C", "Confirmed")
                oGridCombo.ValidValues.Add("L", "Cancelled")
                oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                oGrid.Columns.Item("U_Z_DocStatus").Editable = False
                oGrid.Columns.Item("U_Z_AppStatus").TitleObject.Caption = "Approval Status"
                oGrid.Columns.Item("U_Z_AppStatus").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                oGridCombo = oGrid.Columns.Item("U_Z_AppStatus")
                oGridCombo.ValidValues.Add("P", "Pending")
                oGridCombo.ValidValues.Add("A", "Approved")
                oGridCombo.ValidValues.Add("R", "Rejected")
                oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                oGrid.Columns.Item("U_Z_Remarks").TitleObject.Caption = "Remarks"
                oGrid.Columns.Item("U_Z_Remarks").Editable = True
                oGrid.Columns.Item("U_Z_CurApprover").TitleObject.Caption = "Current Approver"
                oGrid.Columns.Item("U_Z_CurApprover").Editable = False
                oGrid.Columns.Item("U_Z_NxtApprover").TitleObject.Caption = "Next Approver"
                oGrid.Columns.Item("U_Z_NxtApprover").Editable = False
                oGrid.Columns.Item("Creator").Visible = False
                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                oGrid.AutoResizeColumns()
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub ViewWorkSheet(ByVal sForm As SAPbouiCOM.Form, ByVal RefCode As String, ByVal gridId As String)
        oGrid = sForm.Items.Item(gridId).Specific
        Try
            strSQL = "SELECT U_Z_ItemCode,U_Z_ItemDesc,U_Z_Size,U_Z_Qty,U_Z_Price,U_Z_Total,U_Z_Spec FROM [@Z_QUT1]  T0 where T0.DocEntry='" & RefCode & "'"
            oGrid.DataTable.ExecuteQuery(strSQL)
            Formatgrid(oGrid)
            assignMatrixLineno(oGrid, sForm)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#Region "FormatGrid"
    Private Sub Formatgrid(ByVal agrid As SAPbouiCOM.Grid)
        agrid.Columns.Item("U_Z_ItemCode").TitleObject.Caption = "Item Code"
        oEditTextColumn = agrid.Columns.Item("U_Z_ItemCode")
        oEditTextColumn.LinkedObjectType = "4"
        agrid.Columns.Item("U_Z_ItemDesc").TitleObject.Caption = "Item Description"
        agrid.Columns.Item("U_Z_Spec").TitleObject.Caption = "Item Specification"
        agrid.Columns.Item("U_Z_Size").TitleObject.Caption = "Size"
        agrid.Columns.Item("U_Z_Qty").TitleObject.Caption = "Quantity"
        agrid.Columns.Item("U_Z_Price").TitleObject.Caption = "Unit Price"
        agrid.Columns.Item("U_Z_Total").TitleObject.Caption = "Total"
        oEditTextColumn = agrid.Columns.Item("U_Z_Total")
        oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
        agrid.AutoResizeColumns()
        agrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
    Public Sub assignMatrixLineno(ByVal aGrid As SAPbouiCOM.Grid, ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        For intNo As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            aGrid.RowHeaders.SetText(intNo, intNo + 1)
        Next
        aGrid.Columns.Item("RowsHeader").TitleObject.Caption = "#"
        aform.Freeze(False)
    End Sub
#End Region
#Region "Approval Functions"
    Public Sub addUpdateDocument(ByVal aForm As SAPbouiCOM.Form)
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        oCompanyService = oApplication.Company.GetCompanyService()
        Dim otestRs As SAPbobsCOM.Recordset
        Dim oChild As SAPbobsCOM.GeneralData
        Dim strCode, strQuery As String
        Dim strEmpName As String = ""
        Dim blnRecordExists As Boolean = False
        Dim HeadDocEntry, UserLineId As Integer
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oComboBox1, oCombobox2 As SAPbouiCOM.ComboBox
        Try
            If oApplication.SBO_Application.MessageBox("Documents once approved can not be changed. Do you want Continue?", , "Contine", "Cancel") = 2 Then
                Exit Sub
            End If
            oGeneralService = oCompanyService.GetGeneralService("P_APHIS")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otestRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oGrid = aForm.Items.Item("4").Specific
            Dim strDocEntry As String = ""
            Dim strDocType1, HeaderCode As String
            Dim strEmpID As String = ""
            Dim strLeaveType As String = ""
            If oGrid.DataTable.Rows.Count > 0 Then
                For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    strDocEntry = oGrid.DataTable.GetValue("DocEntry", index)
                    strEmpID = oGrid.DataTable.GetValue("Creator", index)
                    strQuery = "select T0.DocEntry,T1.LineId from [@P_OAPPT] T0 JOIN [@P_APPT2] T1 on T0.DocEntry=T1.DocEntry"
                    strQuery += " JOIN [@P_APPT1] T2 on T1.DocEntry=T2.DocEntry"
                    strQuery += " where T0.U_Z_DocType='B' AND (T2.U_Z_OUser='" & strEmpID & "' OR T1.U_Z_AUser='" & oApplication.Company.UserName & "')"
                    otestRs.DoQuery(strQuery)
                    If otestRs.RecordCount > 0 Then
                        HeadDocEntry = otestRs.Fields.Item(0).Value
                        UserLineId = otestRs.Fields.Item(1).Value
                    End If

                    strQuery = "Select * from [@P_APHIS] where U_Z_DocEntry='" & strDocEntry & "' and U_Z_DocType='B' and U_Z_ApproveBy='" & oApplication.Company.UserName & "'"
                    oRecordSet.DoQuery(strQuery)
                    If oRecordSet.RecordCount > 0 Then
                        oGeneralParams.SetProperty("DocEntry", oRecordSet.Fields.Item("DocEntry").Value)
                        oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                        oGeneralData.SetProperty("U_Z_AppStatus", oGrid.DataTable.GetValue("U_Z_AppStatus", index))
                        oGeneralData.SetProperty("U_Z_Remarks", oGrid.DataTable.GetValue("U_Z_Remarks", index))
                        oGeneralData.SetProperty("U_Z_ADocEntry", HeadDocEntry)
                        oGeneralData.SetProperty("U_Z_ALineId", UserLineId)
                        Dim oTemp As SAPbobsCOM.Recordset
                        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oTemp.DoQuery("Select * ,isnull(""firstName"",'') +  ' ' + isnull(""middleName"",'') +  ' ' + isnull(""lastName"",'') 'EmpName' from OHEM where ""userid""=" & oApplication.Company.UserSignature)
                        If oTemp.RecordCount > 0 Then
                            oGeneralData.SetProperty("U_Z_EmpId", oTemp.Fields.Item("empID").Value.ToString())
                            oGeneralData.SetProperty("U_Z_EmpName", oTemp.Fields.Item("EmpName").Value)
                            strEmpName = oTemp.Fields.Item("EmpName").Value
                        Else
                            oGeneralData.SetProperty("U_Z_EmpId", "")
                            oGeneralData.SetProperty("U_Z_EmpName", "")
                        End If
                        oGeneralService.Update(oGeneralData)
                    ElseIf (strDocEntry <> "" And strDocEntry <> "0") Then
                        Dim oTemp As SAPbobsCOM.Recordset
                        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oTemp.DoQuery("Select * ,isnull(""firstName"",'') + ' ' + isnull(""middleName"",'') +  ' ' + isnull(""lastName"",'') 'EmpName' from OHEM where ""userid""=" & oApplication.Company.UserSignature)
                        If oTemp.RecordCount > 0 Then
                            oGeneralData.SetProperty("U_Z_EmpId", oTemp.Fields.Item("empID").Value.ToString())
                            oGeneralData.SetProperty("U_Z_EmpName", oTemp.Fields.Item("EmpName").Value)
                            strEmpName = oTemp.Fields.Item("EmpName").Value
                        Else
                            oGeneralData.SetProperty("U_Z_EmpId", "")
                            oGeneralData.SetProperty("U_Z_EmpName", "")
                        End If
                        oGeneralData.SetProperty("U_Z_DocEntry", strDocEntry.ToString())
                        oGeneralData.SetProperty("U_Z_DocType", "B")
                        oGeneralData.SetProperty("U_Z_AppStatus", oGrid.DataTable.GetValue("U_Z_AppStatus", index))
                        oGeneralData.SetProperty("U_Z_Remarks", oGrid.DataTable.GetValue("U_Z_Remarks", index))
                        oGeneralData.SetProperty("U_Z_ApproveBy", oApplication.Company.UserName)
                        oGeneralData.SetProperty("U_Z_Approvedt", System.DateTime.Now)
                        oGeneralData.SetProperty("U_Z_ADocEntry", HeadDocEntry)
                        oGeneralData.SetProperty("U_Z_ALineId", UserLineId)
                        oGeneralService.Add(oGeneralData)
                    End If
                    updateFinalStatus(aForm, HeadDocEntry, strDocEntry, strEmpID, oGrid.DataTable.GetValue("U_Z_AppStatus", index), oGrid.DataTable.GetValue("U_Z_Remarks", index))
                    If oGrid.DataTable.GetValue("U_Z_AppStatus", index) = "A" Then
                        SendMessage(strDocType1, strDocEntry, oGrid.DataTable.GetValue("U_Z_AppStatus", index), HeadDocEntry, strEmpName, oApplication.Company.UserName, "B")
                    End If

                Next
            End If
            HeaderGridBind(oForm)
            HeaderSumGridBind(oForm)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub updateFinalStatus(ByVal aForm As SAPbouiCOM.Form, ByVal strTemplateNo As String, ByVal strDocEntry As String, ByVal aEmpID As String, ByVal strStatus As String, ByVal Remarks As String)
        Try
            Dim intLineID As Integer
            Dim strMessageUser, StrMailMessage, sQuery As String
            Dim oRecordSet, oTemp As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If strStatus = "A" Then
                sQuery = " Select T2.DocEntry "
                sQuery += " From [@P_APPT2] T2 "
                sQuery += " JOIN [@P_OAPPT] T3 ON T2.DocEntry = T3.DocEntry  "
                sQuery += " JOIN [@P_APPT1] T4 ON T4.DocEntry = T3.DocEntry  "
                sQuery += " Where  U_Z_AFinal = 'Y'"
                sQuery += " And T2.U_Z_AUser = '" + oApplication.Company.UserName + "' And T3.U_Z_DocType = 'B'"
                oRecordSet.DoQuery(sQuery)
                If Not oRecordSet.EoF Then
                    strQuery = "Update [@Z_OQUT] set U_Z_AppStatus='A' where DocEntry='" & strDocEntry & "'"
                    oRecordSet.DoQuery(strQuery)

                    StrMailMessage = "BoM Estimation has been Approved for the Document number :" & CInt(strDocEntry)
                    UserMessage(StrMailMessage, strDocEntry, aEmpID)
                End If
            ElseIf strStatus = "R" Then
                sQuery = " Select T2.DocEntry "
                sQuery += " From [@P_APPT2] T2 "
                sQuery += " JOIN [@P_OAPPT] T3 ON T2.DocEntry = T3.DocEntry  "
                sQuery += " JOIN [@P_APPT1] T4 ON T4.DocEntry = T3.DocEntry  "
                sQuery += " Where T2.U_Z_AUser = '" + oApplication.Company.UserName + "' And T3.U_Z_DocType = 'B'"
                oRecordSet.DoQuery(sQuery)
                If Not oRecordSet.EoF Then
                    strQuery = "Update [@Z_OQUT] set U_Z_AppStatus='R',U_Z_Remarks='" & Remarks & "' where DocEntry='" & strDocEntry & "'"
                    oRecordSet.DoQuery(strQuery)
                    StrMailMessage = "BoM Estimation has been Rejected for the Document number :" & CInt(strDocEntry)
                    UserMessage(StrMailMessage, strDocEntry, aEmpID)
                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub UserMessage(ByVal strMessage As String, ByVal strDocEntry As String, ByVal SAPUser As String)
        Dim oCmpSrv As SAPbobsCOM.CompanyService
        Dim oMessageService As SAPbobsCOM.MessagesService
        Dim oMessage As SAPbobsCOM.Message
        Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns
        Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn
        Dim oLines As SAPbobsCOM.MessageDataLines
        Dim oLine As SAPbobsCOM.MessageDataLine
        Dim oRecipientCollection As SAPbobsCOM.RecipientCollection
        oCmpSrv = oApplication.Company.GetCompanyService()
        oMessageService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService)
        oMessage = oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)
        oMessage.Subject = " BoM Estimation Approval Notification "
        oMessage.Text = strMessage
        oRecipientCollection = oMessage.RecipientCollection
        oRecipientCollection.Add()
        oRecipientCollection.Item(0).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
        oRecipientCollection.Item(0).UserCode = SAPUser
        pMessageDataColumns = oMessage.MessageDataColumns
        pMessageDataColumn = pMessageDataColumns.Add()
        pMessageDataColumn.ColumnName = "Document Number"
        oLines = pMessageDataColumn.MessageDataLines()
        oLine = oLines.Add()
        oLine.Value = strDocEntry
        oMessageService.SendMessage(oMessage)
    End Sub

    Public Sub SendMessage(ByVal strReqType As String, ByVal strReqNo As String, ByVal strAppStatus As String _
        , ByVal strTemplateNo As String, ByVal strOrginator As String, ByVal strAuthorizer As String, ByVal enDocType As String)
        Try
            Dim strQuery As String
            Dim strMessageUser As String
            Dim intLineID As Integer
            Dim oRecordSet, oTemp As SAPbobsCOM.Recordset
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim oMessageService As SAPbobsCOM.MessagesService
            Dim oMessage As SAPbobsCOM.Message
            Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns
            Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn
            Dim oLines As SAPbobsCOM.MessageDataLines
            Dim oLine As SAPbobsCOM.MessageDataLine
            Dim oRecipientCollection As SAPbobsCOM.RecipientCollection
            oCmpSrv = oApplication.Company.GetCompanyService()
            oMessageService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService)
            oMessage = oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select LineId From [@P_APPT2] Where DocEntry = '" & strTemplateNo & "' And U_Z_AUser = '" & strAuthorizer & "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                intLineID = CInt(oRecordSet.Fields.Item(0).Value)
                strQuery = "Select Top 1 U_Z_AUser From [@P_APPT2] Where  DocEntry = '" & strTemplateNo & "' And LineId > '" & intLineID.ToString() & "' and isnull(U_Z_AMan,'')='Y'  Order By LineId Asc "
                oRecordSet.DoQuery(strQuery)

                If Not oRecordSet.EoF Then
                    strMessageUser = oRecordSet.Fields.Item(0).Value
                    oMessage.Subject = " BoM Estimation Need Your Approval "
                    Dim strMessage As String = ""
                    strMessage = " Requested by  :" & oApplication.Company.UserName & ": Document Number : " & strReqNo

                    strQuery = "Update [@Z_OQUT] set U_Z_CurApprover='" & oApplication.Company.UserName & "',U_Z_NxtApprover='" & strMessageUser & "' where DocEntry='" & strReqNo & "'"
                    oTemp.DoQuery(strQuery)

                    oMessage.Text = "BoM Estimation " & " " & strMessage & " Needs Your Approval "
                    oRecipientCollection = oMessage.RecipientCollection
                    oRecipientCollection.Add()
                    oRecipientCollection.Item(0).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
                    oRecipientCollection.Item(0).UserCode = strMessageUser
                    pMessageDataColumns = oMessage.MessageDataColumns
                    pMessageDataColumn = pMessageDataColumns.Add()
                    pMessageDataColumn.ColumnName = "Document Number"
                    oLines = pMessageDataColumn.MessageDataLines()
                    oLine = oLines.Add()
                    oLine.Value = strReqNo
                    oMessageService.SendMessage(oMessage)
                End If
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BoM_Approval Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "4" Or pVal.ItemUID = "9") And pVal.ColUID = "DocNum" Then
                                    Dim oobj As New clsProjectEstimation
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    Dim ocomboColumn As SAPbouiCOM.ComboBoxColumn
                                    ocomboColumn = oGrid.Columns.Item("U_Z_SlpCode")
                                    oobj.LoadForm_View(oGrid.DataTable.GetValue("DocNum", pVal.Row), ocomboColumn.GetSelectedValue(pVal.Row).Value)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "4" Or pVal.ItemUID = "9" Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    Dim strDocEntry As String = oGrid.DataTable.GetValue("DocEntry", pVal.Row)
                                    Dim oOBj As New clsAppHisDetails
                                    oOBj.LoadForm(oForm, strDocEntry)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "1000001"
                                        oForm.PaneLevel = 1
                                        oGrid = oForm.Items.Item("4").Specific
                                        oGrid.Columns.Item("RowsHeader").Click(0)
                                    Case "7"
                                        oForm.PaneLevel = 2
                                        oGrid = oForm.Items.Item("9").Specific
                                        oGrid.Columns.Item("RowsHeader").Click(0)
                                    Case "3"
                                        Dim intRet As Integer = oApplication.SBO_Application.MessageBox("Are you sure want to submit the document?", 2, "Yes", "No", "")
                                        If intRet = 1 Then
                                            addUpdateDocument(oForm)
                                        End If
                                End Select
                                If (pVal.ItemUID = "9" Or pVal.ItemUID = "4") And pVal.ColUID = "RowsHeader" And pVal.Row > -1 Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    Dim strDocEntry As String = oGrid.DataTable.GetValue("DocEntry", pVal.Row)
                                    oForm.Freeze(True)
                                    If pVal.ItemUID = "4" Then
                                        ViewWorkSheet(oForm, strDocEntry, "6")
                                    Else
                                        ViewWorkSheet(oForm, strDocEntry, "10")
                                    End If
                                    oForm.Freeze(False)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim strValue As String
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If 1 = 1 Then 'pVal.ItemUID = "4" Or pVal.ItemUID = "6" Then
                                        strValue = oDataTable.GetValue(CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).ChooseFromListAlias, 0)
                                        Try
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                        Catch ex As Exception
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                        End Try
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
                Case mnu_BoM_Approval
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
