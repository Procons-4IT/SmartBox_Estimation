Public Class clsBomReference
    Inherits clsBase

    Private oMatrix As SAPbouiCOM.Matrix
    Dim oStatic As SAPbouiCOM.StaticText
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private ogrid As SAPbouiCOM.Grid
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
    Public Sub LoadForm(aCode As String, aRefno As String, aItemName As String)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_BOMRef, frm_BOMRef)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oApplication.Utilities.setEdittextvalue(oForm, "4", aCode)
            oApplication.Utilities.setEdittextvalue(oForm, "5", aItemName)
            oApplication.Utilities.setEdittextvalue(oForm, "7", aRefno)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            DataBind(oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub DataBind(aform As SAPbouiCOM.Form)
        ogrid = aform.Items.Item("8").Specific
        Dim s, aRefNo, aItemCode As String
        aRefNo = oApplication.Utilities.getEditTextvalue(aform, "7")
        aItemCode = oApplication.Utilities.getEditTextvalue(aform, "4")
        s = "SELECT T0.""Code"", T0.""Name"", T0.""U_Z_Type"", T0.""U_Z_ItemCode"", T0.""U_Z_ItemName"",T0.""U_Z_BaseQty"", T0.""U_Z_UoM"", T0.""U_Z_WhsCode"", T0.""U_Z_PlnList"",  T0.""U_Z_Cost"", T0.""U_Z_TotalCost"", T0.""U_Z_Remarks"", T0.""U_Z_PHRef"" FROM ""dbo"".""@Z_PRPH2""  T0"
        s = s & " where ""U_Z_PHRef""='" & aRefno & "'"
        ogrid.DataTable.ExecuteQuery(s)
        FormatGrid(ogrid)
        oApplication.Utilities.AssignRowNo(ogrid)
    End Sub
    Private Sub FormatGrid(aGrid As SAPbouiCOM.Grid)
        ogrid = aGrid
        ogrid.Columns.Item("Code").Visible = False
        ogrid.Columns.Item("Name").Visible = False
        ogrid.Columns.Item("U_Z_Type").TitleObject.Caption = "Type"
        ogrid.Columns.Item("U_Z_Type").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        oComboColumn = ogrid.Columns.Item("U_Z_Type")
        oComboColumn.ValidValues.Add("4", "Item")
        oComboColumn.ValidValues.Add("290", "Resource")
        oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
        ogrid.Columns.Item("U_Z_ItemCode").TitleObject.Caption = "Item Code"
        oEditTextColumn = ogrid.Columns.Item("U_Z_ItemCode")
        oEditTextColumn.ChooseFromListUID = "CFL_2"
        oEditTextColumn.ChooseFromListAlias = "ItemCode"
        oEditTextColumn.LinkedObjectType = "290"
        ogrid.Columns.Item("U_Z_ItemName").TitleObject.Caption = "Item Name"
        ogrid.Columns.Item("U_Z_ItemName").Editable = False
        ogrid.Columns.Item("U_Z_BaseQty").TitleObject.Caption = "Quantity"
        ogrid.Columns.Item("U_Z_Cost").TitleObject.Caption = "Unit Price"
        ogrid.Columns.Item("U_Z_Cost").Editable = True
        oEditTextColumn = ogrid.Columns.Item("U_Z_Cost")
        oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
        ogrid.Columns.Item("U_Z_WhsCode").TitleObject.Caption = "Warehouse"

        oEditTextColumn = ogrid.Columns.Item("U_Z_WhsCode")
        oEditTextColumn.ChooseFromListUID = "CFL_3"
        oEditTextColumn.ChooseFromListAlias = "WhsCode"
        oEditTextColumn.LinkedObjectType = "4"
        ogrid.Columns.Item("U_Z_UoM").TitleObject.Caption = "UoM"
        ogrid.Columns.Item("U_Z_UoM").Editable = False
        ogrid.Columns.Item("U_Z_TotalCost").TitleObject.Caption = "Total"
        ogrid.Columns.Item("U_Z_TotalCost").Editable = False
        oEditTextColumn = ogrid.Columns.Item("U_Z_TotalCost")
        oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
        ogrid.Columns.Item("U_Z_Remarks").TitleObject.Caption = "Comments"
        ogrid.Columns.Item("U_Z_PlnList").TitleObject.Caption = "PriceList"
        ogrid.Columns.Item("U_Z_PlnList").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        Dim Otest As SAPbobsCOM.Recordset
        Otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Otest.DoQuery("Select ""ListNum"",""ListName"" from OPLN order by ""ListNum"" ")
        oComboColumn = ogrid.Columns.Item("U_Z_PlnList")
        For intRow As Integer = 0 To Otest.RecordCount - 1
            oComboColumn.ValidValues.Add(Otest.Fields.Item(0).Value, Otest.Fields.Item(1).Value)
            Otest.MoveNext()
        Next
        oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
        ogrid.Columns.Item("U_Z_PHRef").Visible = False

        ogrid.AutoResizeColumns()
        ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
    Public Sub addcontrols(ByVal aforma As SAPbouiCOM.Form)
        Try
            oApplication.Utilities.AddControls(oForm, "_301", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", 0, 0, "2", "View Summary", 120)

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BOMRef Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "2" Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to Cancel the changes?", , "Ok", "Cancel") = 2 Then
                                        Exit Sub
                                    End If
                                    Dim strRef As String = oApplication.Utilities.getEditTextvalue(oForm, "7")
                                    Dim oTemp As SAPbobsCOM.Recordset
                                    oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oTemp.DoQuery("Update ""@Z_PRPH2"" set ""Name""=""Code"" where ""U_Z_PHRef""='" & strRef & "'")
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val, val2, val3, val4 As String
                                Dim intChoice As Integer
                                Dim codebar As String
                                If pVal.ItemUID = "8" Then
                                    If pVal.ColUID = "U_Z_ItemCode" Then
                                        ogrid = oForm.Items.Item("8").Specific
                                        If ogrid.DataTable.GetValue("U_Z_Type", pVal.Row) = "4" Then
                                            oEditTextColumn = ogrid.Columns.Item("U_Z_ItemCode")
                                            oEditTextColumn.ChooseFromListUID = "CFL_2"
                                            oEditTextColumn.ChooseFromListAlias = "ItemCode"
                                        Else
                                            oEditTextColumn = ogrid.Columns.Item("U_Z_ItemCode")
                                            oEditTextColumn.ChooseFromListUID = "CFL_4"
                                            oEditTextColumn.ChooseFromListAlias = "VisResCode"
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                If pVal.ItemUID = "8" Then
                                    If pVal.ColUID = "U_Z_ItemCode" Then
                                        ogrid = oForm.Items.Item("8").Specific
                                        If ogrid.DataTable.GetValue("U_Z_Type", pVal.Row) = "4" Then
                                            oEditTextColumn = ogrid.Columns.Item("U_Z_ItemCode")
                                            oEditTextColumn.LinkedObjectType = "4"
                                        Else
                                            oEditTextColumn = ogrid.Columns.Item("U_Z_ItemCode")
                                            oEditTextColumn.LinkedObjectType = "290"
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                If pVal.ItemUID = "8" Then
                                    If pVal.ColUID = "U_Z_PlnList" Then
                                        ogrid = oForm.Items.Item("8").Specific
                                        If ogrid.DataTable.GetValue("U_Z_Type", pVal.Row) = "4" Then
                                           
                                        Else
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                If pVal.ItemUID = "8" Then
                                    If pVal.ColUID = "U_Z_PlnList" Then
                                        ogrid = oForm.Items.Item("8").Specific
                                        If ogrid.DataTable.GetValue("U_Z_Type", pVal.Row) = "4" Then
                                            Dim Otest As SAPbobsCOM.Recordset
                                            oForm.Freeze(True)
                                            Otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            Otest.DoQuery("Select ""Price"" from ITM1 where ""ItemCode""='" & ogrid.DataTable.GetValue("U_Z_ItemCode", pVal.Row) & "' and pricelist=" & ogrid.DataTable.GetValue("U_Z_PlnList", pVal.Row))
                                            ogrid.DataTable.SetValue("U_Z_Cost", pVal.Row, Otest.Fields.Item(0).Value)
                                            Dim dblUnitPrice, dblQuantity, dblPercentage As Double
                                            dblUnitPrice = ogrid.DataTable.GetValue("U_Z_Cost", pVal.Row)
                                            dblQuantity = ogrid.DataTable.GetValue("U_Z_BaseQty", pVal.Row)
                                            dblPercentage = dblQuantity * dblUnitPrice
                                            ogrid.DataTable.SetValue("U_Z_TotalCost", pVal.Row, dblPercentage)
                                            oForm.Freeze(False)
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "8" And (pVal.ColUID = "U_Z_BaseQty" Or pVal.ColUID = "U_Z_Cost") And pVal.CharPressed = 9 Then
                                    ogrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    oForm.Freeze(True)
                                    Dim dblUnitPrice, dblQuantity, dblPercentage As Double
                                    dblUnitPrice = ogrid.DataTable.GetValue("U_Z_Cost", pVal.Row)
                                    dblQuantity = ogrid.DataTable.GetValue("U_Z_BaseQty", pVal.Row)
                                    dblPercentage = dblQuantity * dblUnitPrice
                                    ogrid.DataTable.SetValue("U_Z_TotalCost", pVal.Row, dblPercentage)
                                    oForm.Freeze(False)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "_301" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    'Dim oOBj As New clsBoMSummary
                                    'frm_SourceBoM = oForm
                                    'oOBj.LoadForm(oApplication.Utilities.getEditTextvalue(oForm, "4"))
                                End If

                                If pVal.ItemUID = "3" Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to submit the changes?", , "Ok", "Cancel") = 2 Then
                                        Exit Sub
                                    End If
                                    If AddtoUDT_Initialize(oForm) = True Then
                                        Dim strRef As String = oApplication.Utilities.getEditTextvalue(oForm, "7")
                                        oMatrix = frm_SourceProjectPhase.Items.Item("14").Specific
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", frm_ProjectPhaseRow, oApplication.Utilities.getEditTextvalue(oForm, "7"))
                                        Dim oTem1 As SAPbobsCOM.Recordset
                                        oTem1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oTem1.DoQuery("Select sum(""U_Z_TotalCost"") from ""@Z_PRPH2"" where ""U_Z_PHRef""='" & strRef & "'")
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", frm_ProjectPhaseRow, oTem1.Fields.Item(0).Value)
                                        Dim dblUnitPrice, dblQuantity, dblPercentage As Double
                                        dblUnitPrice = oApplication.Utilities.getMatrixValues(oMatrix, "V_3", frm_ProjectPhaseRow)
                                        dblQuantity = oApplication.Utilities.getMatrixValues(oMatrix, "V_2", frm_ProjectPhaseRow)
                                        dblPercentage = oApplication.Utilities.getMatrixValues(oMatrix, "V_4", frm_ProjectPhaseRow)
                                        dblQuantity = (dblUnitPrice * dblQuantity)
                                        dblQuantity = dblQuantity + (dblQuantity * dblPercentage / 100)
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", frm_ProjectPhaseRow, dblQuantity)
                                        oForm.Close()
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val, val2, val3, val4 As String
                                Dim intChoice As Integer
                                Dim codebar As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        If oCFL.ObjectType = "4" Then
                                            val = oDataTable.GetValue("ItemCode", 0)
                                            val1 = oDataTable.GetValue("ItemName", 0)
                                            ogrid = oForm.Items.Item(pVal.ItemUID).Specific
                                            ogrid.DataTable.SetValue("U_Z_ItemCode", pVal.Row, val)
                                            ogrid.DataTable.SetValue("U_Z_ItemName", pVal.Row, val1)
                                            ogrid.DataTable.SetValue("U_Z_UoM", pVal.Row, oDataTable.GetValue("InvntryUom", 0))
                                        End If
                                        If oCFL.ObjectType = "64" Then
                                            val = oDataTable.GetValue("WhsCode", 0)
                                            ogrid = oForm.Items.Item(pVal.ItemUID).Specific
                                            ogrid.DataTable.SetValue("U_Z_WhsCode", pVal.Row, val)
                                        End If

                                        If oCFL.ObjectType = "290" Then
                                            val = oDataTable.GetValue("VisResCode", 0)
                                            val1 = oDataTable.GetValue("ResName", 0)
                                            ogrid = oForm.Items.Item(pVal.ItemUID).Specific
                                            ogrid.DataTable.SetValue("U_Z_ItemCode", pVal.Row, val)
                                            ogrid.DataTable.SetValue("U_Z_ItemName", pVal.Row, val1)
                                        End If

                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    oForm.Freeze(False)
                                End Try
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region
    Private Function AddtoUDT_Initialize(aform As SAPbouiCOM.Form) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim otemp, otemp1 As SAPbobsCOM.Recordset
        Dim strqry, strCode, strqry1, strProCode, ProName, strGLAcc, ItemCode, aCHoice As String
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ItemCode = oApplication.Utilities.getEditTextvalue(aform, "4")
        aCHoice = oApplication.Utilities.getEditTextvalue(aform, "7")
        If 1 = 1 Then
            strCode = oApplication.Utilities.getMaxCode("@Z_PRES", "Code")
            oUserTable = oApplication.Company.UserTables.Item("Z_PRPH2")
            ogrid = aform.Items.Item("8").Specific
            For intLoop As Integer = 0 To ogrid.DataTable.Rows.Count - 1
                If ogrid.DataTable.GetValue("Code", intLoop) <> "" Then
                    strCode = ogrid.DataTable.GetValue("Code", intLoop)
                    oUserTable.GetByKey(strCode)
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_RItemCode").Value = ItemCode
                    oUserTable.UserFields.Fields.Item("U_Z_PHRef").Value = aCHoice

                    '  otemp1.DoQuery("Select *  from ITT1 T0  Inner Join  OITT T1 on T0.""Father"" = T1.""Code""   where ""Father"" ='" & ItemCode & "'")
                    oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ogrid.DataTable.GetValue("U_Z_ItemCode", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_ItemName").Value = ogrid.DataTable.GetValue("U_Z_ItemName", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_Type").Value = ogrid.DataTable.GetValue("U_Z_Type", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_BaseQty").Value = ogrid.DataTable.GetValue("U_Z_BaseQty", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_PlnList").Value = ogrid.DataTable.GetValue("U_Z_PlnList", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_WhsCode").Value = ogrid.DataTable.GetValue("U_Z_WhsCode", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = ogrid.DataTable.GetValue("U_Z_Cost", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_TotalCost").Value = ogrid.DataTable.GetValue("U_Z_BaseQty", intLoop) * ogrid.DataTable.GetValue("U_Z_Cost", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_Remarks").Value = ogrid.DataTable.GetValue("U_Z_Remarks", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_UoM").Value = ogrid.DataTable.GetValue("U_Z_UoM", intLoop)
                    If oUserTable.Update <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                Else
                    strCode = ogrid.DataTable.GetValue("Code", intLoop)
                    strCode = oApplication.Utilities.getMaxCode("@Z_PRPH2", "Code")
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_RItemCode").Value = ItemCode
                    oUserTable.UserFields.Fields.Item("U_Z_PHRef").Value = aCHoice
                    '  otemp1.DoQuery("Select *  from ITT1 T0  Inner Join  OITT T1 on T0.""Father"" = T1.""Code""   where ""Father"" ='" & ItemCode & "'")
                    oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ogrid.DataTable.GetValue("U_Z_ItemCode", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_ItemName").Value = ogrid.DataTable.GetValue("U_Z_ItemName", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_Type").Value = ogrid.DataTable.GetValue("U_Z_Type", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_BaseQty").Value = ogrid.DataTable.GetValue("U_Z_BaseQty", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_PlnList").Value = ogrid.DataTable.GetValue("U_Z_PlnList", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_WhsCode").Value = ogrid.DataTable.GetValue("U_Z_WhsCode", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = ogrid.DataTable.GetValue("U_Z_Cost", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_TotalCost").Value = ogrid.DataTable.GetValue("U_Z_BaseQty", intLoop) * ogrid.DataTable.GetValue("U_Z_Cost", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_Remarks").Value = ogrid.DataTable.GetValue("U_Z_Remarks", intLoop)
                    oUserTable.UserFields.Fields.Item("U_Z_UoM").Value = ogrid.DataTable.GetValue("U_Z_UoM", intLoop)
                    If oUserTable.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If

                End If
            Next
        End If
        otemp1.DoQuery("Delete  from ""@Z_PRPH2"" where ""Name"" Like '%_XD' and ""U_Z_PHRef""='" & aCHoice & "'")
        Return True
    End Function


#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        ogrid = oForm.Items.Item("8").Specific
                        If ogrid.DataTable.GetValue("U_Z_ItemCode", ogrid.DataTable.Rows.Count - 1) <> "" Then
                            ogrid.DataTable.Rows.Add()
                            oApplication.Utilities.AssignRowNo(ogrid)
                        End If
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = True Then
                        ogrid = oForm.Items.Item("8").Specific
                        For intRow As Integer = 0 To ogrid.DataTable.Rows.Count - 1
                            If ogrid.Rows.IsSelected(intRow) Then
                                Dim otemp As SAPbobsCOM.Recordset
                                otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                If blnIsHana = True Then
                                    otemp.DoQuery("Update ""@Z_PRPH2"" set ""Name""=""Name"" || '_XD' where ""Code""='" & ogrid.DataTable.GetValue("Code", intRow) & "'")
                                Else
                                    otemp.DoQuery("Update ""@Z_PRPH2"" set ""Name""=""Name"" + '_XD' where ""Code""='" & ogrid.DataTable.GetValue("Code", intRow) & "'")
                                End If

                                ogrid.DataTable.Rows.Remove(intRow)
                                oApplication.Utilities.AssignRowNo(ogrid)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Next
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
