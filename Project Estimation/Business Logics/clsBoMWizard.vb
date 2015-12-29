Public Class clsBoMWizard
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
            oForm = oApplication.Utilities.LoadForm(xml_BoM_Wizard, frm_BoM_Wizard)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
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
            oGrid.DataTable.ExecuteQuery("SELECT T0.""DocEntry"", T0.""DocNum"", T0.""CreateDate"",T0.""U_Z_PrjCode"",T0.""U_Z_PrjName"",""U_Z_SupPrjCode"",""U_Z_SupPrjName"",""U_Z_GLAcc"",""U_Z_FreeText"",T0.""U_Z_CardCode"",T0.""U_Z_SlpCode"",T0.""U_Z_TotalCost"" ""Total Cost"", T0.""U_Z_Remarks"", ' ' ""Select"" FROM ""@Z_OQUT""  T0 where T0.""U_Z_DocStatus""='A' and T0.""U_Z_AppStatus""='A' and T0.""U_Z_SlpCode""='" & oApplication.Utilities.getEditTextvalue(aform, "13") & "' and  T0.""U_Z_CardCode""='" & oApplication.Utilities.getEditTextvalue(aform, "12") & "' order by ""DocEntry"" Desc")

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
            oRecordSet.DoQuery("select SlpCode,SlpName  from OSLP order by SlpCode")
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
            oGrid.Columns.Item("U_Z_SupPrjCode").Visible = False
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
            strItem = "SELECT T0.""DocEntry"",T1.""DocNum"", T0.""LineId"",T0.""U_Z_ItemCode"", T0.""U_Z_ItemDesc"", T0.""U_Z_Price"", T0.""U_Z_Qty"", T0.""U_Z_Total"", 'Y' ""Select"" FROM ""@Z_QUT1""  T0"
            strItem = strItem & " inner Join ""@Z_OQUT"" T1 on T1.""DocEntry""=T0.""DocEntry"" where T0.""U_Z_ItemCode""<>'' AND T0.""DocEntry"" in (" & strDocNumbers & ")"
            oGrid.DataTable.ExecuteQuery(strItem)
            oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Estimation No"
            oGrid.Columns.Item("DocEntry").Editable = False
            oGrid.Columns.Item("DocNum").Editable = False
            oGrid.Columns.Item("LineId").Visible = False


            oGrid.Columns.Item("U_Z_ItemCode").TitleObject.Caption = "Item Code"
            oGrid.Columns.Item("U_Z_ItemCode").Editable = False
            oEditTextColumn = oGrid.Columns.Item("U_Z_ItemCode")
            oEditTextColumn.LinkedObjectType = "4"
            oGrid.Columns.Item("U_Z_ItemDesc").TitleObject.Caption = "Description"
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
    Private Sub AddtoDocument(aform As SAPbouiCOM.Form)

        Dim oMatrix As SAPbouiCOM.Matrix
        oCombobox = frm_SourceQuotation.Items.Item("3").Specific
        If oCombobox.Selected.Value = "S" Then
            oMatrix = frm_SourceQuotation.Items.Item("39").Specific
            oMatrix.Clear()
            frm_SourceQuotation.Select()
            oGrid = aform.Items.Item("7").Specific
            '  frm_SourceQuotation.Freeze(True)
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oCheck = oGrid.Columns.Item("Select")
                If oCheck.IsChecked(intRow) Then
                    If oMatrix.RowCount < 1 Then
                        oMatrix.AddRow()
                    End If
                    oApplication.Utilities.SetMatrixValues(oMatrix, "2", oMatrix.RowCount, oGrid.DataTable.GetValue("U_Z_GLAcc", intRow))
                    oApplication.Utilities.SetMatrixValues(oMatrix, "12", oMatrix.RowCount - 1, oGrid.DataTable.GetValue("Total Cost", intRow))
                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_EstDocNum", oMatrix.RowCount - 1, oGrid.DataTable.GetValue("DocNum", intRow))
                End If
            Next
        Else
            oMatrix = frm_SourceQuotation.Items.Item("38").Specific
            oMatrix.Clear()
            frm_SourceQuotation.Select()
            oGrid = aform.Items.Item("9").Specific
            '  frm_SourceQuotation.Freeze(True)
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oCheck = oGrid.Columns.Item("Select")
                If oCheck.IsChecked(intRow) Then
                    If oMatrix.RowCount < 1 Then
                        oMatrix.AddRow()
                    End If
                    oApplication.Utilities.SetMatrixValues(oMatrix, "1", oMatrix.RowCount, oGrid.DataTable.GetValue("U_Z_ItemCode", intRow))
                    oApplication.Utilities.SetMatrixValues(oMatrix, "14", oMatrix.RowCount - 1, oGrid.DataTable.GetValue("U_Z_Total", intRow))

                End If
            Next
        End If

       
        aform.Close()
        ' frm_SourceQuotation.Freeze(False)
    End Sub
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BoM_Wizard Then
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
