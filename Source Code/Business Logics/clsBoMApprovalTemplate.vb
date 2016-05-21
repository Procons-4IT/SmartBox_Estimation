Public Class clsBoMApprovalTemplate
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oMatrix, oMatrix1, oMatrix2 As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oComboBox, oComboBox1 As SAPbouiCOM.ComboBox
    Private oCheckBox, oCheckBox1 As SAPbouiCOM.CheckBox
    Private InvForConsumedItems, count As Integer
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLineZ_1 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLineZ_2, oDataSrc_Line As SAPbouiCOM.DBDataSource
    Public MatrixId As String
    Public intSelectedMatrixrow As Integer = 0
    Private RowtoDelete As Integer
    Private oRecordSet As SAPbobsCOM.Recordset
    Private dtValidFrom, dtValidTo As Date
    Private strQuery As String

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub


    Public Sub LoadForm()
        Try

            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_BoM_Template) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oForm = oApplication.Utilities.LoadForm(xml_BoM_Template, frm_BoM_Template)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            enableControls(oForm, True)
            FillDocType(oForm)
            oMatrix = oForm.Items.Item("9").Specific
            oMatrix.AutoResizeColumns()
            oMatrix = oForm.Items.Item("10").Specific
            oMatrix.AutoResizeColumns()
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, False)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub FillDocType(ByVal aForm As SAPbouiCOM.Form)
        Dim oTempRec As SAPbobsCOM.Recordset
        oComboBox = aForm.Items.Item("17").Specific
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        For intRow As Integer = oComboBox.ValidValues.Count - 1 To 0 Step -1
            oComboBox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        oComboBox.ValidValues.Add("", "")
        oComboBox.ValidValues.Add("B", "Bill of Material")
        oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        aForm.Items.Item("17").DisplayDesc = True
        oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
    End Sub


#Region "Item Event"

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BoM_Template Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        Else
                                            If validation(oForm) = False Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "26" And (oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE) Then
                                    'oCheckBox = oForm.Items.Item("26").Specific
                                    'oComboBox = oForm.Items.Item("17").Specific
                                    '    If RemoveValidation(oComboBox.Selected.Value, oApplication.Utilities.getEdittextvalue(oForm, "12")) = False Then
                                    '        oApplication.Utilities.Message("Some documents pending for approval. You can not inactive", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    '        BubbleEvent = False
                                    '        Exit Sub
                                    '    End If
                                End If
                                '    oComboBox = oForm.Items.Item("17").Specific
                                If pVal.ItemUID = "10" Or pVal.ItemUID = "9" Then
                                    If oApplication.Utilities.getEditTextvalue(oForm, "4") = "" Then
                                        BubbleEvent = False
                                        oApplication.SBO_Application.SetStatusBarMessage("Select Code to Proceed...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    ElseIf oApplication.Utilities.getEditTextvalue(oForm, "6") = "" Then
                                        BubbleEvent = False
                                        oApplication.SBO_Application.SetStatusBarMessage("Select Name to Proceed...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    ElseIf oComboBox.Selected.Value = "" Then
                                        BubbleEvent = False
                                        oApplication.SBO_Application.SetStatusBarMessage("Select Document Type to Proceed...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        oForm.Items.Item("17").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    End If
                                End If
                                If pVal.ItemUID = "9" And pVal.Row > 0 Then
                                    oMatrix = oForm.Items.Item("9").Specific
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = "9"
                                    frmSourceMatrix = oMatrix
                                End If
                                If pVal.ItemUID = "10" And pVal.Row > 0 And (oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE) Then
                                    oMatrix = oForm.Items.Item("10").Specific
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = "10"
                                    frmSourceMatrix = oMatrix

                                    If pVal.ColUID = "V_4" Then
                                        oCheckBox = oMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific
                                        oComboBox = oForm.Items.Item("17").Specific
                                        If oComboBox.Selected.Value <> "" Then
                                            If oCheckBox.Checked = True Then
                                                'If ValidateAuthorizer(oComboBox.Selected.Value, oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row)) = False Then
                                                '    oApplication.Utilities.Message("There is a pending request for this authorizer. You can not inactive", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                '    BubbleEvent = False
                                                '    Exit Sub
                                                'End If
                                            End If
                                        End If
                                    End If
                                End If

                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "17" Then
                                    oMatrix = oForm.Items.Item("9").Specific
                                    oMatrix1 = oForm.Items.Item("10").Specific
                                    oMatrix.Clear()
                                    oMatrix1.Clear()
                                    oComboBox = oForm.Items.Item("17").Specific
                                    oApplication.Utilities.setEdittextvalue(oForm, "19", oComboBox.Selected.Description)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "13"
                                        AddRow(oForm)
                                    Case "14"
                                        RefereshDeleteRow(oForm)
                                    Case "7"
                                        oForm.PaneLevel = 1
                                    Case "8"
                                        oForm.PaneLevel = 3
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim val1, val, Val2 As String
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If (pVal.ItemUID = "10" Or pVal.ItemUID = "9") And pVal.ColUID = "V_0" Then
                                            val1 = oDataTable.GetValue("USER_CODE", 0)
                                            val = oDataTable.GetValue("U_NAME", 0)
                                            oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                            Try
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", pVal.Row, val)
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", pVal.Row, val1)
                                            Catch ex As Exception
                                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                            End Try
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
                Case mnu_BoM_Template
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then

                    End If
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        AddRow(oForm)
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        RefereshDeleteRow(oForm)
                    End If
                Case mnu_ADD
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        enableControls(oForm, True)
                    End If
                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        enableControls(oForm, True)
                    End If
                Case "1283"
                    If pVal.BeforeAction = True Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        oComboBox = oForm.Items.Item("17").Specific
                        If oApplication.SBO_Application.MessageBox("Do you want to remove approval template?", , "Yes", "No") = 2 Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        'If RemoveValidation(oComboBox.Selected.Value, oApplication.Utilities.getEdittextvalue(oForm, "12")) = False Then
                        '    oApplication.Utilities.Message("Some documents pending for approval. You can not remove the template", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If
                    End If

            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

#End Region

#Region "Data Events"

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            If oForm.TypeEx = frm_BoM_Template Then
                Select Case BusinessObjectInfo.BeforeAction
                    Case True

                    Case False
                        Select Case BusinessObjectInfo.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@P_OAPPT")
                                enableControls(oForm, False)
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Function"




#Region "Methods"
    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Select Case aForm.PaneLevel
                Case "1"
                    oMatrix = aForm.Items.Item("9").Specific
                    oDBDataSourceLineZ_1 = oForm.DataSources.DBDataSources.Item("@P_APPT1")
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLineZ_1.Size
                        oDBDataSourceLineZ_1.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
                Case "3"
                    oMatrix = aForm.Items.Item("10").Specific
                    oDBDataSourceLineZ_2 = oForm.DataSources.DBDataSources.Item("@P_APPT2")
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLineZ_2.Size
                        oDBDataSourceLineZ_2.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

#End Region

    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)

            Select Case aForm.PaneLevel
                Case "1"
                    oMatrix = aForm.Items.Item("9").Specific
                    oDBDataSourceLineZ_1 = oForm.DataSources.DBDataSources.Item("@P_APPT1")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    End If
                    oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                    Try
                        If oEditText.Value <> "" Then
                            oMatrix.AddRow()
                            oMatrix.ClearRowData(oMatrix.RowCount)
                        End If
                    Catch ex As Exception
                        aForm.Freeze(False)
                    End Try
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLineZ_1.Size
                        oDBDataSourceLineZ_1.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    AssignLineNo(aForm)
                Case "3"
                    oMatrix = aForm.Items.Item("10").Specific
                    oDBDataSourceLineZ_2 = oForm.DataSources.DBDataSources.Item("@P_APPT2")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    End If
                    oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                    Try
                        If oEditText.Value <> "" Then 'And oCheckBox.Checked = False Then
                            oMatrix.AddRow()
                            oMatrix.ClearRowData(oMatrix.RowCount)
                        End If
                    Catch ex As Exception
                        aForm.Freeze(False)
                    End Try
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLineZ_2.Size
                        oDBDataSourceLineZ_2.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    AssignLineNo(aForm)
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

#Region "Delete Row"


    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            oDBDataSourceLineZ_1 = oForm.DataSources.DBDataSources.Item("@P_APPT1")
            oDBDataSourceLineZ_2 = oForm.DataSources.DBDataSources.Item("@P_APPT2")
            If Me.MatrixId = "9" Then
                oMatrix = aForm.Items.Item("9").Specific
                Me.RowtoDelete = intSelectedMatrixrow
                oDBDataSourceLineZ_1.RemoveRecord(Me.RowtoDelete - 1)
                oMatrix.LoadFromDataSource()
                oMatrix.FlushToDataSource()
                For count = 1 To oDBDataSourceLineZ_1.Size
                    oDBDataSourceLineZ_1.SetValue("LineId", count - 1, count)
                Next
            ElseIf (Me.MatrixId = "10") Then
                oMatrix = aForm.Items.Item("10").Specific
                Me.RowtoDelete = intSelectedMatrixrow
                oDBDataSourceLineZ_2.RemoveRecord(Me.RowtoDelete - 1)
                oMatrix.LoadFromDataSource()
                oMatrix.FlushToDataSource()
                For count = 1 To oDBDataSourceLineZ_2.Size
                    oDBDataSourceLineZ_2.SetValue("LineId", count - 1, count)
                Next
            End If
            oMatrix.LoadFromDataSource()
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub
#End Region

#Region "Validations"
    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            aForm.Freeze(True)
            oComboBox = aForm.Items.Item("17").Specific
            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oApplication.Utilities.getEditTextvalue(aForm, "4") = "" Then
                oApplication.Utilities.Message("Enter Code...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "6") = "" Then
                oApplication.Utilities.Message("Enter Name...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            ElseIf oComboBox.Selected.Value = "" Then
                oApplication.Utilities.Message("Select Document Type...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If
            oMatrix = aForm.Items.Item("10").Specific
            If oMatrix.RowCount = 0 Then
                oApplication.Utilities.Message("Authorizer Row Cannot be Empty...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select ""DocEntry"" From ""@P_OAPPT"""
            strQuery += " Where "
            strQuery += " ""U_Z_DocType"" = '" & oComboBox.Selected.Value & "' And ""DocEntry"" <> '" & oApplication.Utilities.getEditTextvalue(aForm, "12") & "'"
            ' oRecordSet.DoQuery(strQuery)
            'If Not oRecordSet.EoF Then
            '    '  oApplication.Utilities.Message("Document Type Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    '  aForm.Freeze(False)
            '    '  Return False
            'End If


            oMatrix = aForm.Items.Item("10").Specific
            Dim blnflag As Boolean = False
            Dim blnActive As Boolean = False
            Dim oCheck1 As SAPbouiCOM.CheckBox
            For intRow As Integer = 1 To oMatrix.RowCount
                oCheckBox = oMatrix.Columns.Item("V_3").Cells.Item(intRow).Specific
                oCheck1 = oMatrix.Columns.Item("V_4").Cells.Item(intRow).Specific
                If oCheck1.Checked = True Then
                    blnActive = True
                End If
                If oCheckBox.Checked = True Then
                    If oCheck1.Checked = False Then
                        oApplication.Utilities.Message("Only Active Authorizer will be set as final authorizer...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        aForm.Freeze(False)
                        Return False
                    End If
                    blnflag = True
                End If
            Next

            If blnActive = False Then
                oApplication.Utilities.Message("Atlease one  Authorizer should be active...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If
            If blnflag = False Then
                oApplication.Utilities.Message("Select Final Authorizer...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If

            Dim strECode, strECode1, strEname, strEname1 As String
            oMatrix = aForm.Items.Item("9").Specific
            For intRow As Integer = 1 To oMatrix.RowCount
                strECode = CType(oMatrix.Columns.Item("V_0").Cells.Item(intRow).Specific, SAPbouiCOM.EditText).Value
                For intInnerLoop As Integer = intRow To oMatrix.RowCount
                    strECode1 = CType(oMatrix.Columns.Item("V_0").Cells.Item(intInnerLoop).Specific, SAPbouiCOM.EditText).Value
                    If strECode = strECode1 And intRow <> intInnerLoop Then
                        oApplication.Utilities.Message("User Duplicated in Row : " + intInnerLoop.ToString() + "...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oMatrix.Columns.Item("V_0").Cells.Item(intInnerLoop).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        aForm.Freeze(False)
                        Return False
                    End If
                Next
            Next

            oMatrix = aForm.Items.Item("10").Specific
            For intRow As Integer = 1 To oMatrix.RowCount
                strECode = CType(oMatrix.Columns.Item("V_0").Cells.Item(intRow).Specific, SAPbouiCOM.EditText).Value
                oCheckBox = oMatrix.Columns.Item("V_3").Cells.Item(intRow).Specific
                For intInnerLoop As Integer = intRow To oMatrix.RowCount
                    strECode1 = CType(oMatrix.Columns.Item("V_0").Cells.Item(intInnerLoop).Specific, SAPbouiCOM.EditText).Value
                    oCheckBox1 = oMatrix.Columns.Item("V_3").Cells.Item(intInnerLoop).Specific
                    If strECode = strECode1 And intRow <> intInnerLoop Then
                        oApplication.Utilities.Message("Authorizer Duplicated in Row : " + intInnerLoop.ToString() + "...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oMatrix.Columns.Item("V_0").Cells.Item(intInnerLoop).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        aForm.Freeze(False)
                        Return False
                    ElseIf oCheckBox.Checked = True And oCheckBox1.Checked = True And intRow <> intInnerLoop Then
                        oApplication.Utilities.Message("Select Only one final Authorizer. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oMatrix.Columns.Item("V_0").Cells.Item(intInnerLoop).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        aForm.Freeze(False)
                        Return False
                    End If
                Next
            Next

            oMatrix = aForm.Items.Item("9").Specific
            oMatrix1 = aForm.Items.Item("10").Specific
            For intRow As Integer = 1 To oMatrix.RowCount
                strECode = CType(oMatrix.Columns.Item("V_0").Cells.Item(intRow).Specific, SAPbouiCOM.EditText).Value
                For intInnerLoop As Integer = 1 To oMatrix1.RowCount
                    strECode1 = CType(oMatrix1.Columns.Item("V_0").Cells.Item(intInnerLoop).Specific, SAPbouiCOM.EditText).Value
                    If strECode = strECode1 Then
                        oApplication.Utilities.Message("User is duplicated in User and Authorizer. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oMatrix.Columns.Item("V_0").Cells.Item(intInnerLoop).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        aForm.Freeze(False)
                        Return False
                    End If
                Next
            Next

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select 1 As ""Return"",""DocEntry"" From ""@P_OAPPT"""
            strQuery += " Where "
            strQuery += " ""U_Z_Code"" = '" & oApplication.Utilities.getEditTextvalue(aForm, "4") & "' And ""DocEntry"" <> '" & oApplication.Utilities.getEditTextvalue(aForm, "12") & "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oApplication.Utilities.Message("Code Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If
            AssignLineNo(aForm)
            aForm.Freeze(False)
            Return True
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Function
#End Region

#Region "Disable Controls"

    Private Sub enableControls(ByVal aForm As SAPbouiCOM.Form, ByVal blnEnable As Boolean)
        Try
            'oForm.Items.Item("12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            aForm.Items.Item("4").Enabled = blnEnable
            aForm.Items.Item("6").Enabled = blnEnable
            aForm.Items.Item("17").Enabled = blnEnable
            ' oComboBox = aForm.Items.Item("17").Specific
            ' oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#End Region
End Class

