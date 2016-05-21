Public Class clsBoMWizard_PO
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


    Public Sub LoadForm(aCode As String, Optional aSlpCode As String = "-1")
        Try
            oForm = oApplication.Utilities.LoadForm(xml_BoM_Wizard_PO, frm_BoM_Wizard_PO)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            aCode = ""
            aSlpCode = ""
            oApplication.Utilities.setEdittextvalue(oForm, "12", aCode)
            oApplication.Utilities.setEdittextvalue(oForm, "13", aSlpCode)
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
                strDocNum = strDocNum & "," & oGrid.DataTable.GetValue("Code", intRow)
            End If

        Next
        Return strDocNum
    End Function

    Private Function getParentItems(aform As SAPbouiCOM.Form) As String
        Dim strDocNum As String = "'xxxxxx'"
        oGrid = aform.Items.Item("7").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oCheck = oGrid.Columns.Item("Select")
            If oCheck.IsChecked(intRow) Then
                strDocNum = strDocNum & ",'" & oGrid.DataTable.GetValue("Code", intRow) & "'"
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
            '  oGrid.DataTable.ExecuteQuery("SELECT T0.[DocEntry], T0.[DocNum], T0.[CreateDate],T0.U_Z_CardCode,T0.U_Z_SlpCode, T0.[U_Z_Desc], T0.[U_Z_Remarks], ' ' 'Select' FROM [dbo].[@Z_OQUT]  T0 where T0.U_Z_DocStatus='C' and T0.U_Z_AppStatus='A' and T0.U_Z_SlpCode='" & oApplication.Utilities.getEditTextvalue(aform, "13") & "' and  T0.U_Z_CardCode='" & oApplication.Utilities.getEditTextvalue(aform, "12") & "' order by DocEntry Desc")
            'oGrid.DataTable.ExecuteQuery("SELECT T0.[DocEntry], T0.[DocNum], T0.[CreateDate],T0.U_Z_CardCode,T0.U_Z_SlpCode, T0.[U_Z_Desc], T0.[U_Z_Remarks], ' ' 'Select' FROM [dbo].[@Z_OQUT]  T0 where T0.U_Z_DocStatus='C' and T0.U_Z_AppStatus='A'  order by DocEntry Desc")

            oGrid.DataTable.ExecuteQuery("select  ' ' As ""Select"",T0.""Code"" from OITT T0 ")

            oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("Select").TitleObject.Caption = "Select"
            oGrid.Columns.Item("Select").Editable = True
            oGrid.Columns.Item("Code").TitleObject.Caption = "BoM Parent"
            oGrid.Columns.Item("Code").Editable = False
            oEditTextColumn = oGrid.Columns.Item("Code")
            oEditTextColumn.LinkedObjectType = "4"

            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
        Else
            Dim strDocNumbers As String = getParentItems(aform)
            oGrid = aform.Items.Item("9").Specific
            strItem = "SELECT T0.""DocEntry"",T1.""DocNum"", T0.""LineId"",T0.""U_Z_ItemCode"", T0.""U_Z_ItemDesc"", T0.""U_Z_Spec"",T0.""U_Z_Size"" AS ""Size"", T0.""U_Z_Price"", T0.""U_Z_Qty"", T0.""U_Z_Total"", ' ' As ""Select"" FROM ""@Z_QUT1""  T0"
            strItem = strItem & " inner Join ""@Z_OQUT"" T1 on T1.""DocEntry""=T0.""DocEntry"" where T0.""U_Z_ItemCode""<>'' AND T0.""DocEntry"" in (" & strDocNumbers & ")"
            strItem = "select ' ' As ""Select"",T0.""Code"",T1.""Code"" As ""ItemCode"",T2.""ItemName"",T1.""Quantity"" ,T1.""Warehouse"" As ""WhsCode"" ,T1.U_AVGCOST"",T1.""U_MARKUP"",T1.""Price"" from OITT T0 Inner Join ITT1 T1 on T1.""Father""=T0.""Code"" inner Join OITM T2 on T2.""ItemCode""=T1.""Code"" where T2.""ItmsGrpCod""<>112"
            strItem = strItem & " and T0.""Code"" in (" & strDocNumbers & ")"
            oGrid.DataTable.ExecuteQuery(strItem)
            oGrid.Columns.Item("Code").TitleObject.Caption = "Parent Item"
            oGrid.Columns.Item("Code").Editable = False
            oEditTextColumn = oGrid.Columns.Item("Code")
            oEditTextColumn.LinkedObjectType = "4"
            oGrid.Columns.Item("ItemCode").TitleObject.Caption = "Item Code"
            oGrid.Columns.Item("ItemCode").Editable = False
            oEditTextColumn = oGrid.Columns.Item("ItemCode")
            oEditTextColumn.LinkedObjectType = "4"
            oGrid.Columns.Item("ItemName").TitleObject.Caption = "Description"
            oGrid.Columns.Item("ItemName").Editable = False


            oGrid.Columns.Item("Quantity").TitleObject.Caption = "Quantity"
            oGrid.Columns.Item("Quantity").Editable = False
            oGrid.Columns.Item("WhsCode").TitleObject.Caption = "Warehouse"
            oGrid.Columns.Item("WhsCode").Editable = False
            oEditTextColumn = oGrid.Columns.Item("WhsCode")
            oEditTextColumn.LinkedObjectType = "64"
            oGrid.Columns.Item("Price").TitleObject.Caption = "Sales Price"
            oGrid.Columns.Item("Price").Editable = False
            oGrid.Columns.Item("U_AVGCOST").TitleObject.Caption = "Avg.Cost"
            oGrid.Columns.Item("U_AVGCOST").Editable = False

            oGrid.Columns.Item("U_MARKUP").TitleObject.Caption = "MarkUp %"
            oGrid.Columns.Item("U_MARKUP").Editable = False

            oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("Select").TitleObject.Caption = "Select"
            oGrid.Columns.Item("Select").Editable = True
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
    Private Sub AddtoDocument(aform As SAPbouiCOM.Form)
        oGrid = aform.Items.Item("9").Specific
        Dim oMatrix As SAPbouiCOM.Matrix
        If frm_SourceQuotation.TypeEx = frm_GoodsIssue Then
            oMatrix = frm_SourceQuotation.Items.Item("13").Specific
        Else
            oMatrix = frm_SourceQuotation.Items.Item("38").Specific
        End If

        oMatrix.Clear()
        frm_SourceQuotation.Select()
        '  frm_SourceQuotation.Freeze(True)
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oCheck = oGrid.Columns.Item("Select")
            If oCheck.IsChecked(intRow) Then
                If oMatrix.RowCount < 1 Then
                    oMatrix.AddRow()
                End If
                oApplication.Utilities.SetMatrixValues(oMatrix, "1", oMatrix.RowCount, oGrid.DataTable.GetValue("ItemCode", intRow))
                'oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Spec", oMatrix.RowCount - 1, oGrid.DataTable.GetValue("U_Z_Spec", intRow))
                'oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_EstDocNum", oMatrix.RowCount - 1, oGrid.DataTable.GetValue("DocNum", intRow))
                'oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_EstLineId", oMatrix.RowCount - 1, oGrid.DataTable.GetValue("LineId", intRow))
                If frm_SourceQuotation.TypeEx = frm_GoodsIssue Then
                    oApplication.Utilities.SetMatrixValues(oMatrix, "9", oMatrix.RowCount - 1, oGrid.DataTable.GetValue("Quantity", intRow))
                Else
                    oApplication.Utilities.SetMatrixValues(oMatrix, "11", oMatrix.RowCount - 1, oGrid.DataTable.GetValue("Quantity", intRow))
                End If

                If frm_SourceQuotation.TypeEx = frm_GoodsIssue Then
                    oApplication.Utilities.SetMatrixValues(oMatrix, "15", oMatrix.RowCount - 1, oGrid.DataTable.GetValue("WhsCode", intRow))
                Else
                    oApplication.Utilities.SetMatrixValues(oMatrix, "24", oMatrix.RowCount - 1, oGrid.DataTable.GetValue("WhsCode", intRow))
                End If
                'Try
                '    oApplication.Utilities.SetMatrixValues(oMatrix, "163", oMatrix.RowCount - 1, oGrid.DataTable.GetValue("Size", intRow))

                'Catch ex As Exception

                'End Try
                Try
                    oApplication.Utilities.SetMatrixValues(oMatrix, "14", oMatrix.RowCount - 1, oGrid.DataTable.GetValue("Price", intRow))
                Catch ex As Exception

                End Try
                'oApplication.Utilities.SetMatrixValues(oMatrix, "11", oMatrix.RowCount - 1, oGrid.DataTable.GetValue("U_Z_Qty", intRow))

            End If
        Next
        aform.Close()
        ' frm_SourceQuotation.Freeze(False)
    End Sub
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BoM_Wizard_PO Then
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
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                        If oForm.PaneLevel = 2 Then
                                            PopulateEstimationsDetails(oForm, "Header")
                                        End If
                                        If oForm.PaneLevel = 3 Then
                                            PopulateEstimationsDetails(oForm, "Trans")
                                        End If
                                    Case "10"
                                        SelectAll(oForm, True)
                                    Case "11"
                                        SelectAll(oForm, False)

                                    Case "5"
                                        AddtoDocument(oForm)
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
                'Case mnu_BarCode
                '    If pVal.BeforeAction = False Then
                '        LoadForm()
                '    End If

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
