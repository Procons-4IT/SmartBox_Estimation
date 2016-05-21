Public Class clsProductionWizard
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
            oForm = oApplication.Utilities.LoadForm(xml_PO_Wizard, frm_PO_Wizard)
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
            sCode = "SELECT T0.""DocEntry"", T0.""DocNum"", T0.""CreateDate"",T0.""U_Z_PrjCode"",T0.""U_Z_PrjName"",""U_Z_SupPrjCode"",""U_Z_SupPrjName"",""U_Z_GLAcc"",""U_Z_FreeText"",T0.""U_Z_CardCode"",T0.""U_Z_SlpCode"",T0.""U_Z_TotalCost"" ""Total Cost"", T0.""U_Z_Remarks"", ' ' As ""Select"" FROM ""@Z_OQUT""  T0 where T0.""DocEntry"" in (Select T1.""DocEntry"" from ""@Z_QUT1"" T1 where isnull(Cast(T1.""U_Z_PONO"" as Varchar),'')='') and T0.""U_Z_DocStatus""='A' and T0.""U_Z_AppStatus""='A' and T0.""U_Z_CardCode""='" & oApplication.Utilities.getEditTextvalue(aform, "13") & "' order by ""DocEntry"" Desc"
            sCode = "SELECT T0.""DocEntry"", T0.""DocNum"", T0.""CreateDate"",T0.""U_Z_PrjCode"",T0.""U_Z_PrjName"",""U_Z_SupPrjCode"",""U_Z_SupPrjName"",""U_Z_GLAcc"",""U_Z_FreeText"",T0.""U_Z_CardCode"",T0.""U_Z_SlpCode"",T0.""U_Z_TotalCost"" ""Total Cost"", T0.""U_Z_Remarks"", 'N' As ""Select"" FROM ""@Z_OQUT""   T0 where T0.""DocEntry"" in (Select T1.""DocEntry"" from ""@Z_QUT1"" T1 where ifnull(CAST(T1.""U_Z_PONO"",Varchar),'')='') and T0.""U_Z_DocStatus""='A' and T0.""U_Z_AppStatus""='A' and T0.""U_Z_PrjCode""='" & oApplication.Utilities.getEditTextvalue(aform, "13") & "' order by ""DocEntry"" Desc"
            If blnIsHana = True Then
                oGrid.DataTable.ExecuteQuery("SELECT T0.""DocEntry"", T0.""DocNum"", T0.""CreateDate"",T0.""U_Z_PrjCode"",T0.""U_Z_PrjName"",""U_Z_SupPrjCode"",""U_Z_SupPrjName"",""U_Z_GLAcc"",""U_Z_FreeText"",T0.""U_Z_CardCode"",T0.""U_Z_SlpCode"",T0.""U_Z_TotalCost"" ""Total Cost"", T0.""U_Z_Remarks"", 'N' As ""Select"" FROM ""@Z_OQUT""   T0 where T0.""DocEntry"" in (Select T1.""DocEntry"" from ""@Z_QUT1"" T1 where ifnull(CAST(T1.""U_Z_PONO"" as Varchar),'')='') and T0.""U_Z_DocStatus""='A' and T0.""U_Z_AppStatus""='A' and T0.""U_Z_PrjCode""='" & oApplication.Utilities.getEditTextvalue(aform, "13") & "' order by ""DocEntry"" Desc")
            Else
                oGrid.DataTable.ExecuteQuery("SELECT T0.""DocEntry"", T0.""DocNum"", T0.""CreateDate"",T0.""U_Z_PrjCode"",T0.""U_Z_PrjName"",""U_Z_SupPrjCode"",""U_Z_SupPrjName"",""U_Z_GLAcc"",""U_Z_FreeText"",T0.""U_Z_CardCode"",T0.""U_Z_SlpCode"",T0.""U_Z_TotalCost"" ""Total Cost"", T0.""U_Z_Remarks"", 'N' As ""Select"" FROM ""@Z_OQUT""   T0 where T0.""DocEntry"" in (Select T1.""DocEntry"" from ""@Z_QUT1"" T1 where isnull(Convert(Varchar,T1.""U_Z_PONO""),'')='') and T0.""U_Z_DocStatus""='A' and T0.""U_Z_AppStatus""='A' and T0.""U_Z_PrjCode""='" & oApplication.Utilities.getEditTextvalue(aform, "13") & "' order by ""DocEntry"" Desc")
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
            oGrid.Columns.Item("Select").Editable = True
            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
        Else
            Dim strDocNumbers As String = GetSelectedDocuments(aform)
            oGrid = aform.Items.Item("9").Specific
            If blnIsHana = True Then
                strItem = "SELECT T0.""DocEntry"",T1.""DocNum"", T0.""LineId"",T1.""U_Z_PrjCode"" ""ProjectCode"",T1.""U_Z_PrjName"" ""ProjectName"",T1.""U_Z_SupPrjCode"",T1.""U_Z_SupPrjName"" ""Phase"",T0.""U_Z_ItemCode"", T0.""U_Z_ItemDesc"", T0.""U_Z_Price"", T0.""U_Z_Qty"", T0.""U_Z_Total"", 'Y' As ""Select"" FROM ""@Z_QUT1""  T0"
                strItem = strItem & " inner Join ""@Z_OQUT"" T1 on T1.""DocEntry""=T0.""DocEntry"" where ifnull(Cast(T0.""U_Z_PONO"" as Varchar),'')='' and  T0.""U_Z_ItemCode""<>'' AND T0.""DocEntry"" in (" & strDocNumbers & ")"
            Else
                strItem = "SELECT T0.""DocEntry"",T1.""DocNum"", T0.""LineId"",T1.""U_Z_PrjCode"" ""ProjectCode"",T1.""U_Z_PrjName"" ""ProjectName"",T1.""U_Z_SupPrjCode"",T1.""U_Z_SupPrjName"" ""Phase"",T0.""U_Z_ItemCode"", T0.""U_Z_ItemDesc"", T0.""U_Z_Price"", T0.""U_Z_Qty"", T0.""U_Z_Total"", 'Y' As ""Select"" FROM ""@Z_QUT1""  T0"
                strItem = strItem & " inner Join ""@Z_OQUT"" T1 on T1.""DocEntry""=T0.""DocEntry"" where isnull(Convert(Varchar,T0.""U_Z_PONO""),'')='' and  T0.""U_Z_ItemCode""<>'' AND T0.""DocEntry"" in (" & strDocNumbers & ")"
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
    Private Function ValidateBoMItem(ByVal aItem As String) As Boolean
        Dim businessObject As SAPbobsCOM.Recordset = DirectCast(modVariables.oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        businessObject.DoQuery(("Select * from OITM where ""ItemCode""='" & aItem & "'"))
        If (businessObject.Fields.Item("TreeType").Value = "N") Then
            Return False
        Else
            Return True
        End If

        'Return Operators.ConditionalCompareObjectEqual(businessObject.Fields.Item("TreeType").Value, "N", False)
    End Function
    Private Function ValidateNonBOM(ByVal aItem As String, ByVal aType As String) As Boolean
        Dim businessObject As SAPbobsCOM.Recordset = DirectCast(modVariables.oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        If (aType = "290") Then
            Return True
        End If
        businessObject.DoQuery(("Select * from OITM where ""ItemCode""='" & aItem & "'"))
        If (businessObject.Fields.Item("TreeType").Value <> "N") Then
            Return False
        End If
        Return True
        'Return Operators.ConditionalCompareObjectEqual(businessObject.Fields.Item("TreeType").Value, "N", False)
    End Function

    Private Function AddtoDocument(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim businessObject As SAPbobsCOM.Recordset = DirectCast(modVariables.oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim recordset2 As SAPbobsCOM.Recordset = DirectCast(modVariables.oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Me.oGrid = DirectCast(aform.Items.Item("9").Specific, SAPbouiCOM.Grid)
        Dim dateString As String = modVariables.oApplication.Utilities.getEditTextvalue(aform, "17")
        Dim dateTimeValue As DateTime = modVariables.oApplication.Utilities.GetDateTimeValue(dateString)
        modVariables.oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Dim num6 As Integer = (Me.oGrid.DataTable.Rows.Count - 1)
        Dim i As Integer = 0
        Do While (i <= num6)
            Dim str11 As String
            Dim str6 As String = Convert.ToString(Me.oGrid.DataTable.GetValue("DocEntry", i))
            Dim str10 As String = Convert.ToString(Me.oGrid.DataTable.GetValue("ProjectCode", i))
            Dim str8 As String = Convert.ToString(Me.oGrid.DataTable.GetValue("Phase", i))
            Dim str As String = Convert.ToString(Me.oGrid.DataTable.GetValue("U_Z_ItemCode", i))
            Dim str7 As String = Convert.ToString(Me.oGrid.DataTable.GetValue("LineId", i))
            If modVariables.blnIsHana Then
                str11 = "select T0.""U_Z_Code"" ,T1.""U_Z_Type"",T0.""U_Z_ItemCode"" ""Parent"",T1.""U_Z_ItemCode"" ,T1.""U_Z_ItemName"" ,T1.""U_Z_BaseQty"",ifnull(T1.""U_Z_BoMRef"",'') ""BoMRef"" from ""@Z_OPRPH"" T0 Inner Join ""@Z_PRPH1"" T1 on T1.""DocEntry""=T0.""DocEntry"""
            Else
                str11 = "select T0.""U_Z_Code"" ,T1.""U_Z_Type"",T0.""U_Z_ItemCode"" ""Parent"",T1.""U_Z_ItemCode"" ,T1.""U_Z_ItemName"" ,T1.""U_Z_BaseQty"",isnull(T1.""U_Z_BoMRef"",'') ""BoMRef"" from ""@Z_OPRPH"" T0 Inner Join ""@Z_PRPH1"" T1 on T1.""DocEntry""=T0.""DocEntry"""
            End If
            str11 = (str11 & " where   T0.""U_Z_Code""='" & str & "'")
            businessObject.DoQuery(str11)
            Dim str9 As String = ""
            Dim lineNum As Integer = 0
            Dim flag2 As Boolean = False
            Dim orders As SAPbobsCOM.ProductionOrders = DirectCast(modVariables.oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders), SAPbobsCOM.ProductionOrders)
            Dim num7 As Integer = (businessObject.RecordCount - 1)
            Dim j As Integer = 0
            Do While (j <= num7)
                orders.UserFields.Fields.Item("U_Z_PrjCode").Value = str10
                orders.UserFields.Fields.Item("U_Z_SubPrj").Value = str8
                orders.UserFields.Fields.Item("U_Z_Phase").Value = str
                orders.UserFields.Fields.Item("U_Z_EstNo").Value = str6
                orders.Project = str10
                orders.PlannedQuantity = Convert.ToDouble(Me.oGrid.DataTable.GetValue("U_Z_Qty", i))
                orders.ItemNo = Convert.ToString(businessObject.Fields.Item("Parent").Value)
                orders.PostingDate = dateTimeValue
                orders.DueDate = dateTimeValue.AddMonths(1)
                orders.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
                If Me.ValidateNonBOM(Convert.ToString(businessObject.Fields.Item("U_Z_ItemCode").Value), Convert.ToString(businessObject.Fields.Item("U_Z_Type").Value)) Then
                    If (lineNum > 0) Then
                        orders.Lines.Add()
                    End If
                    orders.Lines.SetCurrentLine(lineNum)
                    If businessObject.Fields.Item("U_Z_Type").Value = "4" Then
                        orders.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                    ElseIf businessObject.Fields.Item("U_Z_Type").Value = "290" Then
                        orders.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                    End If
                    orders.Lines.ItemNo = Convert.ToString(businessObject.Fields.Item("U_Z_ItemCode").Value)
                    orders.Lines.BaseQuantity = Convert.ToDouble(businessObject.Fields.Item("U_Z_BaseQty").Value)
                    flag2 = True
                    lineNum += 1
                End If
                businessObject.MoveNext()
                j += 1
            Loop

            If flag2 Then
                Dim str13 As String
                If (orders.Add <> 0) Then
                    modVariables.oApplication.Utilities.Message(modVariables.oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                modVariables.oApplication.Company.GetNewObjectCode(str13)
                If (str9 = "") Then
                    str9 = str13
                Else
                    str9 = (str9 & " ," & str13)
                End If
            End If



            'First Level BOM

            If modVariables.blnIsHana Then
                str11 = "select T0.""U_Z_Code"" ,T1.""U_Z_Type"",T1.""U_Z_ItemCode"" ,T1.""U_Z_ItemName"" ,T1.""U_Z_BaseQty"",ifnull(T1.""U_Z_BoMRef"",'') ""BoMRef"" from ""@Z_OPRPH"" T0 Inner Join ""@Z_PRPH1"" T1 on T1.""DocEntry""=T0.""DocEntry"""
            Else
                str11 = "select T0.""U_Z_Code"" ,T1.""U_Z_Type"",T1.""U_Z_ItemCode"" ,T1.""U_Z_ItemName"" ,T1.""U_Z_BaseQty"",isnull(T1.""U_Z_BoMRef"",'') ""BoMRef"" from ""@Z_OPRPH"" T0 Inner Join ""@Z_PRPH1"" T1 on T1.""DocEntry""=T0.""DocEntry"""
            End If
            If blnIsHana = True Then
                str11 = (str11 & " where ifnull(T1.""U_Z_BoMRef"",'')<>'' and  T1.""U_Z_Type""='4' and  T0.""U_Z_Code""='" & str & "'")
            Else
                str11 = (str11 & " where isnull(T1.""U_Z_BoMRef"",'')<>'' and  T1.""U_Z_Type""='4' and  T0.""U_Z_Code""='" & str & "'")
            End If
            businessObject.DoQuery(str11)
            flag2 = False
            Dim num8 As Integer = (businessObject.RecordCount - 1)
            Dim k As Integer = 0
            Do While (k <= num8)
                flag2 = False
                orders = DirectCast(modVariables.oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders), SAPbobsCOM.ProductionOrders)
                orders.UserFields.Fields.Item("U_Z_PrjCode").Value = str10
                orders.UserFields.Fields.Item("U_Z_SubPrj").Value = str8
                orders.UserFields.Fields.Item("U_Z_Phase").Value = str
                orders.UserFields.Fields.Item("U_Z_EstNo").Value = str6
                orders.Project = str10
                orders.PlannedQuantity = Convert.ToDouble(Me.oGrid.DataTable.GetValue("U_Z_Qty", i))
                orders.ItemNo = Convert.ToString(businessObject.Fields.Item("U_Z_ItemCode").Value)
                orders.PostingDate = dateTimeValue
                orders.DueDate = dateTimeValue.AddMonths(1)
                orders.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
                If Me.ValidateBoMItem(businessObject.Fields.Item("U_Z_ItemCode").Value) Then
                    Dim str3 As String
                    If (businessObject.Fields.Item("BoMRef").Value <> "") Then
                        If blnIsHana = True Then
                            str3 = "Select ""U_Z_Type"",""U_Z_ItemCode"",""U_Z_BaseQty"",""U_Z_Cost"",""U_Z_WhsCode"",""U_Z_UoM"",""U_Z_PlnList"",""U_Z_PHSRef"" from ""@Z_PRPH2"" where ifnull(""U_Z_PHSRef"",'')='' and  ""U_Z_PHRef""='" & businessObject.Fields.Item("BoMRef").Value & "'"
                        Else
                            str3 = "Select ""U_Z_Type"",""U_Z_ItemCode"",""U_Z_BaseQty"",""U_Z_Cost"",""U_Z_WhsCode"",""U_Z_UoM"",""U_Z_PlnList"",""U_Z_PHSRef"" from ""@Z_PRPH2"" where isnull(""U_Z_PHSRef"",'')='' and  ""U_Z_PHRef""='" & businessObject.Fields.Item("BoMRef").Value & "'"

                        End If
                         'recordset2.DoQuery(str3)
                        'If recordset2.Fields.Item("U_Z_PHSRef").Value <> "" Then
                        '    str3 = "Select ""U_Z_Type"",""U_Z_ItemCode"",""U_Z_BaseQty"",""U_Z_Cost"",""U_Z_WhsCode"",""U_Z_UoM"",""U_Z_PlnList"" from ""@Z_PRPH3"" where ""U_Z_PHRef""='" & recordset2.Fields.Item("U_Z_PHSRef").Value & "'"
                        'Else
                        '    str3 = "Select ""U_Z_Type"",""U_Z_ItemCode"",""U_Z_BaseQty"",""U_Z_Cost"",""U_Z_WhsCode"",""U_Z_UoM"",""U_Z_PlnList"" from ""@Z_PRPH2"" where ""U_Z_PHRef""='" & businessObject.Fields.Item("BoMRef").Value & "'"
                        'End If
                    Else
                        str3 = "Select * from ITT1 where ""Father""='" & businessObject.Fields.Item("U_Z_ItemCode").Value & "'"
                        str3 = "select ""Type"",""Code"",""Quantity"",""OrigPrice"",""Warehouse"",""Uom"",""PriceList""  from ITT1  where ""Father""='" & businessObject.Fields.Item("U_Z_ItemCode").Value & "'"
                    End If
                    recordset2.DoQuery(str3)
                    Dim num9 As Integer = (recordset2.RecordCount - 1)
                    Dim m As Integer = 0
                    Do While (m <= num9)
                        If (m > 0) Then
                            orders.Lines.Add()
                        End If
                        orders.Lines.SetCurrentLine(m)
                        If recordset2.Fields.Item("U_Z_Type").Value = "4" Then
                            orders.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                        ElseIf recordset2.Fields.Item("U_Z_Type").Value = "290" Then
                            orders.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                        End If
                        orders.Lines.ItemNo = Convert.ToString(recordset2.Fields.Item(1).Value)
                        orders.Lines.BaseQuantity = Convert.ToDouble(recordset2.Fields.Item(2).Value)
                        orders.Lines.PlannedQuantity = Convert.ToDouble(Me.oGrid.DataTable.GetValue("U_Z_Qty", i))
                        orders.Lines.Warehouse = Convert.ToString(recordset2.Fields.Item(4).Value)
                        flag2 = True
                        recordset2.MoveNext()
                        m += 1
                    Loop
                End If
                If flag2 Then
                    Dim str14 As String
                    If (orders.Add <> 0) Then
                        modVariables.oApplication.Utilities.Message(modVariables.oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                    modVariables.oApplication.Company.GetNewObjectCode(str14)
                    If (str9 = "") Then
                        str9 = str14
                    Else
                        str9 = (str9 & " ," & str14)
                    End If
                End If

                'Second Level BoM
                flag2 = False
                orders = DirectCast(modVariables.oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders), SAPbobsCOM.ProductionOrders)
                orders.UserFields.Fields.Item("U_Z_PrjCode").Value = str10
                orders.UserFields.Fields.Item("U_Z_SubPrj").Value = str8
                orders.UserFields.Fields.Item("U_Z_Phase").Value = str
                orders.UserFields.Fields.Item("U_Z_EstNo").Value = str6
                orders.Project = str10
                orders.PlannedQuantity = Convert.ToDouble(Me.oGrid.DataTable.GetValue("U_Z_Qty", i))
                orders.ItemNo = Convert.ToString(businessObject.Fields.Item("U_Z_ItemCode").Value)
                orders.PostingDate = dateTimeValue
                orders.DueDate = dateTimeValue.AddMonths(1)
                orders.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
                Dim strPOItem As String = Convert.ToString(businessObject.Fields.Item("U_Z_ItemCode").Value)
                Dim dblPOQty As Double = Convert.ToDouble(Me.oGrid.DataTable.GetValue("U_Z_Qty", i))
                If Me.ValidateBoMItem(businessObject.Fields.Item("U_Z_ItemCode").Value) Then
                    Dim str3 As String
                    If (businessObject.Fields.Item("BoMRef").Value <> "") Then
                        ' str3 = "Select ""U_Z_Type"",""U_Z_ItemCode"",""U_Z_BaseQty"",""U_Z_Cost"",""U_Z_WhsCode"",""U_Z_UoM"",""U_Z_PlnList"",""U_Z_PHSRef"" from ""@Z_PRPH2"" where ""U_Z_PHSRef""<>'' and  ""U_Z_PHRef""='" & businessObject.Fields.Item("BoMRef").Value & "'"
                        If blnIsHana = True Then
                            str3 = "Select ""U_Z_Type"",""U_Z_ItemCode"",""U_Z_BaseQty"",""U_Z_Cost"",""U_Z_WhsCode"",""U_Z_UoM"",""U_Z_PlnList"",""U_Z_PHSRef"" from ""@Z_PRPH2"" where ifnull(""U_Z_PHSRef"",'')<>'' and  ""U_Z_PHRef""='" & businessObject.Fields.Item("BoMRef").Value & "'"
                        Else
                            str3 = "Select ""U_Z_Type"",""U_Z_ItemCode"",""U_Z_BaseQty"",""U_Z_Cost"",""U_Z_WhsCode"",""U_Z_UoM"",""U_Z_PlnList"",""U_Z_PHSRef"" from ""@Z_PRPH2"" where isnull(""U_Z_PHSRef"",'')<>'' and  ""U_Z_PHRef""='" & businessObject.Fields.Item("BoMRef").Value & "'"

                        End If
                        recordset2.DoQuery(str3)
                        If recordset2.Fields.Item("U_Z_PHSRef").Value <> "" Then
                            strPOItem = recordset2.Fields.Item("U_Z_ItemCode").Value
                            dblPOQty = recordset2.Fields.Item("U_Z_BaseQty").Value
                            str3 = "Select ""U_Z_Type"",""U_Z_ItemCode"",""U_Z_BaseQty"",""U_Z_Cost"",""U_Z_WhsCode"",""U_Z_UoM"",""U_Z_PlnList"" from ""@Z_PRPH3"" where ""U_Z_PHRef""='" & recordset2.Fields.Item("U_Z_PHSRef").Value & "'"
                        Else
                            str3 = "Select ""U_Z_Type"",""U_Z_ItemCode"",""U_Z_BaseQty"",""U_Z_Cost"",""U_Z_WhsCode"",""U_Z_UoM"",""U_Z_PlnList"",""U_Z_PHSRef"" from ""@Z_PRPH2"" where ""U_Z_PHSRef""<>'' and  ""U_Z_PHRef""='" & businessObject.Fields.Item("BoMRef").Value & "'"
                        End If
                    Else
                        str3 = "Select * from ITT1 where ""Father""='" & businessObject.Fields.Item("U_Z_ItemCode").Value & "'"
                        str3 = "select ""Type"",""Code"",""Quantity"",""OrigPrice"",""Warehouse"",""Uom"",""PriceList""  from ITT1  where ""Father""='" & businessObject.Fields.Item("U_Z_ItemCode").Value & "'"
                    End If
                    recordset2.DoQuery(str3)
                    Dim num9 As Integer = (recordset2.RecordCount - 1)
                    Dim m As Integer = 0
                    Do While (m <= num9)
                        orders.PlannedQuantity = dblPOQty ' Convert.ToDouble(Me.oGrid.DataTable.GetValue("U_Z_Qty", i))
                        orders.ItemNo = strPOItem ' Convert.ToString(businessObject.Fields.Item("U_Z_ItemCode").Value)
                        If (m > 0) Then
                            orders.Lines.Add()
                        End If
                        orders.Lines.SetCurrentLine(m)
                        If recordset2.Fields.Item("U_Z_Type").Value = "4" Then
                            orders.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                        ElseIf recordset2.Fields.Item("U_Z_Type").Value = "290" Then
                            orders.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                        End If
                        orders.Lines.ItemNo = Convert.ToString(recordset2.Fields.Item(1).Value)
                        orders.Lines.BaseQuantity = Convert.ToDouble(recordset2.Fields.Item(2).Value)
                        orders.Lines.PlannedQuantity = Convert.ToDouble(Me.oGrid.DataTable.GetValue("U_Z_Qty", i))
                        orders.Lines.Warehouse = Convert.ToString(recordset2.Fields.Item(4).Value)
                        flag2 = True
                        recordset2.MoveNext()
                        m += 1
                    Loop
                End If
                If flag2 Then
                    Dim str14 As String
                    If (orders.Add <> 0) Then
                        modVariables.oApplication.Utilities.Message(modVariables.oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                    modVariables.oApplication.Company.GetNewObjectCode(str14)
                    If (str9 = "") Then
                        str9 = str14
                    Else
                        str9 = (str9 & " ," & str14)
                    End If
                End If

                businessObject.MoveNext()
                k += 1
            Loop

            'Second Level BoM



            'If modVariables.blnIsHana Then
            '    str11 = "select T0.""U_Z_Code"" ,T1.""U_Z_Type"",T1.""U_Z_ItemCode"" ,T1.""U_Z_ItemName"" ,T1.""U_Z_BaseQty"",ifnull(T1.""U_Z_BoMRef"",'') ""BoMRef"" from ""@Z_OPRPH"" T0 Inner Join ""@Z_PRPH1"" T1 on T1.""DocEntry""=T0.""DocEntry"""
            'Else
            '    str11 = "select T0.""U_Z_Code"" ,T1.""U_Z_Type"",T1.""U_Z_ItemCode"" ,T1.""U_Z_ItemName"" ,T1.""U_Z_BaseQty"",isnull(T1.""U_Z_BoMRef"",'') ""BoMRef"" from ""@Z_OPRPH"" T0 Inner Join ""@Z_PRPH1"" T1 on T1.""DocEntry""=T0.""DocEntry"""
            'End If
            ''str11 = (str11 & " where T1.""U_Z_Type""='4' and  T0.""U_Z_Code""='" & str & "'")
            'If blnIsHana = True Then
            '    str11 = (str11 & " where ifnull(T1.""U_Z_BoMRef"",'')='' and  T1.""U_Z_Type""='4' and  T0.""U_Z_Code""='" & str & "'")
            'Else
            '    str11 = (str11 & " where isnull(T1.""U_Z_BoMRef"",'')='' and  T1.""U_Z_Type""='4' and  T0.""U_Z_Code""='" & str & "'")
            'End If
            'businessObject.DoQuery(str11)
            'flag2 = False
            'Dim num8 As Integer = (businessObject.RecordCount - 1)
            'Dim k As Integer = 0
            'Do While (k <= num8)
            '    flag2 = False
            '    orders = DirectCast(modVariables.oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders), SAPbobsCOM.ProductionOrders)
            '    orders.UserFields.Fields.Item("U_Z_PrjCode").Value = str10
            '    orders.UserFields.Fields.Item("U_Z_SubPrj").Value = str8
            '    orders.UserFields.Fields.Item("U_Z_Phase").Value = str
            '    orders.UserFields.Fields.Item("U_Z_EstNo").Value = str6
            '    orders.Project = str10
            '    orders.PlannedQuantity = Convert.ToDouble(Me.oGrid.DataTable.GetValue("U_Z_Qty", i))
            '    orders.ItemNo = Convert.ToString(businessObject.Fields.Item("U_Z_ItemCode").Value)
            '    orders.PostingDate = dateTimeValue
            '    orders.DueDate = dateTimeValue.AddMonths(1)
            '    orders.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
            '    If Me.ValidateBoMItem(businessObject.Fields.Item("U_Z_ItemCode").Value) Then
            '        Dim str3 As String
            '        If (businessObject.Fields.Item("BoMRef").Value <> "") Then
            '            str3 = "Select ""U_Z_Type"",""U_Z_ItemCode"",""U_Z_BaseQty"",""U_Z_Cost"",""U_Z_WhsCode"",""U_Z_UoM"",""U_Z_PlnList"",""U_Z_PHSRef"" from ""@Z_PRPH2"" where ""U_Z_PHRef""='" & businessObject.Fields.Item("BoMRef").Value & "'"
            '            recordset2.DoQuery(str3)
            '            If recordset2.Fields.Item("U_Z_PHSRef").Value <> "" Then
            '                str3 = "Select ""U_Z_Type"",""U_Z_ItemCode"",""U_Z_BaseQty"",""U_Z_Cost"",""U_Z_WhsCode"",""U_Z_UoM"",""U_Z_PlnList"" from ""@Z_PRPH3"" where ""U_Z_PHRef""='" & recordset2.Fields.Item("U_Z_PHSRef").Value & "'"
            '            Else
            '                str3 = "Select ""U_Z_Type"",""U_Z_ItemCode"",""U_Z_BaseQty"",""U_Z_Cost"",""U_Z_WhsCode"",""U_Z_UoM"",""U_Z_PlnList"" from ""@Z_PRPH2"" where ""U_Z_PHRef""='" & businessObject.Fields.Item("BoMRef").Value & "'"
            '            End If
            '        Else
            '            str3 = "Select * from ITT1 where ""Father""='" & businessObject.Fields.Item("U_Z_ItemCode").Value & "'"
            '            str3 = "select ""Type"",""Code"",""Quantity"",""OrigPrice"",""Warehouse"",""Uom"",""PriceList""  from ITT1  where ""Father""='" & businessObject.Fields.Item("U_Z_ItemCode").Value & "'"
            '        End If
            '        recordset2.DoQuery(str3)
            '        Dim num9 As Integer = (recordset2.RecordCount - 1)
            '        Dim m As Integer = 0
            '        Do While (m <= num9)
            '            If (m > 0) Then
            '                orders.Lines.Add()
            '            End If
            '            orders.Lines.SetCurrentLine(m)
            '            If recordset2.Fields.Item("U_Z_Type").Value = "4" Then
            '                orders.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Item
            '            ElseIf recordset2.Fields.Item("U_Z_Type").Value = "290" Then
            '                orders.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
            '            End If
            '            orders.Lines.ItemNo = Convert.ToString(recordset2.Fields.Item(1).Value)
            '            orders.Lines.BaseQuantity = Convert.ToDouble(recordset2.Fields.Item(2).Value)
            '            orders.Lines.PlannedQuantity = Convert.ToDouble(Me.oGrid.DataTable.GetValue("U_Z_Qty", i))
            '            orders.Lines.Warehouse = Convert.ToString(recordset2.Fields.Item(4).Value)
            '            flag2 = True
            '            recordset2.MoveNext()
            '            m += 1
            '        Loop
            '    End If
            '    If flag2 Then
            '        Dim str14 As String
            '        If (orders.Add <> 0) Then
            '            modVariables.oApplication.Utilities.Message(modVariables.oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '            Return False
            '        End If
            '        modVariables.oApplication.Company.GetNewObjectCode(str14)
            '        If (str9 = "") Then
            '            str9 = str14
            '        Else
            '            str9 = (str9 & " ," & str14)
            '        End If
            '    End If
            '    businessObject.MoveNext()
            '    k += 1
            'Loop
            recordset2.DoQuery(String.Concat(New String() {"Update ""@Z_QUT1"" set ""U_Z_PONO""='", str9, "' where ""DocEntry""=", str6, " and ""LineId""=", str7}))
            i += 1
        Loop
        modVariables.oApplication.Utilities.Message("Operation completed successfuly....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Return True
    End Function
    'Private Function AddtoDocument_old(aform As SAPbouiCOM.Form) As Boolean
    '    Dim oMatrix As SAPbouiCOM.Matrix
    '    Dim oProduction As SAPbobsCOM.ProductionOrders
    '    Dim oRec, oRec1 As SAPbobsCOM.Recordset
    '    Dim strPhase, strSubProject, strEstimation, strBoM, strBoMRef, strQuery, strProject, strActivity, strBomLineQuery, strLineNo, strPONo, strDate As String
    '    Dim dtPostingdate As Date
    '    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    oGrid = aform.Items.Item("9").Specific
    '    strDate = oApplication.Utilities.getEditTextvalue(aform, "17")
    '    dtPostingdate = oApplication.Utilities.GetDateTimeValue(strDate)
    '    For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
    '        strEstimation = oGrid.DataTable.GetValue("DocEntry", intRow)
    '        strProject = oGrid.DataTable.GetValue("ProjectCode", intRow)
    '        strPhase = oGrid.DataTable.GetValue("Phase", intRow)
    '        strActivity = oGrid.DataTable.GetValue("U_Z_ItemCode", intRow)
    '        strLineNo = oGrid.DataTable.GetValue("LineId", intRow)
    '        If blnIsHana = True Then
    '            strQuery = "select T0.""U_Z_Code"" ,T1.""U_Z_ItemCode"" ,T1.""U_Z_ItemName"" ,T1.""U_Z_BaseQty"",ifnull(T1.""U_Z_BoMRef"",'') ""BoMRef"" from ""@Z_OPRPH"" T0 Inner Join ""@Z_PRPH1"" T1 on T1.""DocEntry""=T0.""DocEntry"""
    '        Else
    '            strQuery = "select T0.""U_Z_Code"" ,T1.""U_Z_ItemCode"" ,T1.""U_Z_ItemName"" ,T1.""U_Z_BaseQty"",isnull(T1.""U_Z_BoMRef"",'') ""BoMRef"" from ""@Z_OPRPH"" T0 Inner Join ""@Z_PRPH1"" T1 on T1.""DocEntry""=T0.""DocEntry"""
    '        End If
    '        strQuery = strQuery & " where T0.""U_Z_Code""='" & strActivity & "'"
    '        oRec.DoQuery(strQuery)
    '        strPONo = ""
    '        Dim str9 As String = ""
    '        Dim lineNum As Integer = 0
    '        Dim flag2 As Boolean = False
    '        For intMain As Integer = 0 To oRec.RecordCount - 1
    '            Dim blnLineExits As Boolean = False
    '            oProduction = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
    '            oProduction.UserFields.Fields.Item("U_Z_PrjCode").Value = strProject
    '            oProduction.UserFields.Fields.Item("U_Z_SubPrj").Value = strPhase
    '            oProduction.UserFields.Fields.Item("U_Z_Phase").Value = strActivity
    '            oProduction.UserFields.Fields.Item("U_Z_EstNo").Value = strEstimation
    '            oProduction.Project = strProject
    '            oProduction.PlannedQuantity = oGrid.DataTable.GetValue("U_Z_Qty", intRow)
    '            oProduction.ItemNo = oRec.Fields.Item("U_Z_ItemCode").Value
    '            oProduction.PostingDate = dtPostingdate
    '            oProduction.DueDate = dtPostingdate.AddMonths(1)
    '            oProduction.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned

    '            If ValidateNonBOM(oRec.Fields.Item("U_Z_ItemCode").Value, oRec.Fields.Item("U_Z_Type").Value) Then
    '                If lineNum > 0 Then
    '                    oProduction.Lines.Add()
    '                End If
    '                oProduction.Lines.SetCurrentLine(lineNum)

    '                If oRec.Fields.Item("U_Z_Type").Value = "4" Then
    '                    oProduction.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Item
    '                ElseIf oRec.Fields.Item("U_Z_Type").Value = "290" Then
    '                    oProduction.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
    '                End If
    '                oProduction.Lines.ItemNo = oRec.Fields.Item("U_Z_ItemCode").Value
    '                oProduction.Lines.BaseQuantity = oRec.Fields.Item("U_Z_Quantity").Value
    '                lineNum = lineNum + 1
    '                flag2 = True
    '            End If
    '            oRec.MoveNext()
    '        Next
    '        If flag2 Then
    '            Dim str13 As String
    '            If oProduction.Add <> 0 Then
    '                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Else
    '                oApplication.Company.GetNewObjectCode(str13)
    '                If (str9 = "") Then
    '                    str9 = str13
    '                Else
    '                    str9 = (str9 & " ," & str13)
    '                End If
    '            End If
    '        End If


    '        Dim str11 As String
    '        If modVariables.blnIsHana Then
    '            str11 = "select T0.""U_Z_Code"" ,T1.""U_Z_Type"",T1.""U_Z_ItemCode"" ,T1.""U_Z_ItemName"" ,T1.""U_Z_BaseQty"",ifnull(T1.""U_Z_BoMRef"",'') ""BoMRef"" from ""@Z_OPRPH"" T0 Inner Join ""@Z_PRPH1"" T1 on T1.""DocEntry""=T0.""DocEntry"""
    '        Else
    '            str11 = "select T0.""U_Z_Code"" ,T1.""U_Z_Type"",T1.""U_Z_ItemCode"" ,T1.""U_Z_ItemName"" ,T1.""U_Z_BaseQty"",isnull(T1.""U_Z_BoMRef"",'') ""BoMRef"" from ""@Z_OPRPH"" T0 Inner Join ""@Z_PRPH1"" T1 on T1.""DocEntry""=T0.""DocEntry"""
    '        End If
    '        str11 = (str11 & " where T1.""U_Z_Type""='4' and  T0.""U_Z_Code""='" & strActivity & "'")
    '        oRec.DoQuery(str11)
    '        flag2 = False
    '        Dim num8 As Integer = (oRec.RecordCount - 1)
    '        Dim k As Integer = 0
    '        For intRow1 As Integer = 0 To oRec.RecordCount - 1
    '            If oRec.Fields.Item("BoMRef").Value <> "" Then
    '                strBomLineQuery = "Select ""U_Z_Type"",""U_Z_ItemCode"",""U_Z_BaseQty"",""U_Z_Cost"",""U_Z_WhsCode"",""U_Z_UoM"",""U_Z_PlnList"" from ""@Z_PRPH2"" where ""U_Z_PHRef""='" & oRec.Fields.Item("BoMRef").Value & "'"
    '            Else
    '                strBomLineQuery = "Select * from ITT1 where ""Father""='" & oRec.Fields.Item("U_Z_ItemCode").Value & "'"
    '                strBomLineQuery = "select ""Type"",""Code"",""Quantity"",""OrigPrice"",""Warehouse"",""Uom"",""PriceList""  from ITT1  where ""Father""='" & oRec.Fields.Item("U_Z_ItemCode").Value & "'"
    '            End If
    '            oRec1.DoQuery(strBomLineQuery)
    '            For intloop As Integer = 0 To oRec1.RecordCount - 1
    '                If intloop > 0 Then
    '                    oProduction.Lines.Add()
    '                End If
    '                oProduction.Lines.SetCurrentLine(intloop)
    '                If oRec1.Fields.Item(0).Value = "4" Then
    '                    oProduction.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Item
    '                ElseIf oRec1.Fields.Item(0).Value = "290" Then
    '                    oProduction.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
    '                End If
    '                oProduction.Lines.ItemNo = oRec1.Fields.Item(1).Value
    '                oProduction.Lines.BaseQuantity = oRec1.Fields.Item(2).Value
    '                oProduction.Lines.PlannedQuantity = oGrid.DataTable.GetValue("U_Z_Qty", intRow)
    '                oProduction.Lines.Warehouse = oRec1.Fields.Item(4).Value
    '                blnLineExits = True
    '                oRec1.MoveNext()
    '            Next
    '            If blnLineExits = True Then
    '                If oProduction.Add <> 0 Then
    '                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                    Return False
    '                Else
    '                    Dim strDocno As String
    '                    oApplication.Company.GetNewObjectCode(strDocno)
    '                    If strPONo = "" Then
    '                        strPONo = strDocno
    '                    Else
    '                        strPONo = strPONo & " ," & strDocno
    '                    End If
    '                End If
    '            End If
    '            oRec1.DoQuery("Update ""@Z_QUT1"" set ""U_Z_PONO""='" & strPONo & "' where ""DocEntry""=" & strEstimation & " and ""LineId""=" & strLineNo)
    '            oRec.MoveNext()
    '        Next
    '    Next
    '    Return True

    'End Function
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_PO_Wizard Then
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
                                            If oForm.PaneLevel = "1" Then
                                                oApplication.Utilities.Message("Posting Date  missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
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
                                        If oApplication.SBO_Application.MessageBox("Do you want to create the Production Order for selected estimations?", , "Continue", "Cancel") = 2 Then
                                            Exit Sub
                                        End If
                                        If AddtoDocument(oForm) = True Then
                                            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            oForm.PaneLevel = 2
                                            PopulateEstimationsDetails(oForm, "Header")
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
                                    If 1 = 1 Then 'pVal.ItemUID = "4" Or pVal.ItemUID = "6" Then
                                        strValue = oDataTable.GetValue(CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).ChooseFromListAlias, 0)
                                        Try
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                        Catch ex As Exception
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                        End Try
                                        If pVal.ItemUID = "15" Then
                                            oApplication.Utilities.setEdittextvalue(oForm, "13", strValue)
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
                Case mnu_POWizard
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
