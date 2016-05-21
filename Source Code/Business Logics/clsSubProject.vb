Public Class clsSubProject
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

            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_SubProject) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oForm = oApplication.Utilities.LoadForm(xml_SubProject, frm_SubProject)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.DataBrowser.BrowseBy = "4"
            enableControls(oForm, True)
            addChooseFromListConditions(oForm)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub addChooseFromListConditions(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList

            oCFLs = oForm.ChooseFromLists

            oCFL = oCFLs.Item("CFL_4")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)


        Catch ex As Exception
            Throw ex
        End Try
    End Sub


#Region "Item Event"

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_SubProject Then
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
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "13"
                                        '   AddRow(oForm)
                                    Case "14"
                                        '  RefereshDeleteRow(oForm)
                                    Case "7"
                                        ' oForm.PaneLevel = 1
                                    Case "8"
                                        'oForm.PaneLevel = 3
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
                                        'If (pVal.ItemUID = "10" Or pVal.ItemUID = "9") And pVal.ColUID = "V_0" Then
                                        '    val1 = oDataTable.GetValue("USER_CODE", 0)
                                        '    val = oDataTable.GetValue("U_NAME", 0)
                                        '    oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                        '    Try
                                        '        oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", pVal.Row, val)
                                        '        oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", pVal.Row, val1)
                                        '    Catch ex As Exception
                                        '        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        '            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        '        End If
                                        '    End Try
                                        'End If
                                        If pVal.ItemUID = "8" Then
                                            val = oDataTable.GetValue("FormatCode", 0)
                                            val1 = oDataTable.GetValue("AcctName", 0)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "9", val1)
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            Catch ex As Exception
                                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                            End Try
                                        End If

                                        If pVal.ItemUID = "15" Then
                                            val = oDataTable.GetValue("PrjCode", 0)
                                            val1 = oDataTable.GetValue("PrjName", 0)
                                            Try
                                                ' oApplication.Utilities.setEdittextvalue(oForm, "9", val1)
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            Catch ex As Exception
                                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                            End Try
                                        End If

                                        If pVal.ItemUID = "17" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            val1 = oDataTable.GetValue("CardName", 0)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "18", val1)
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
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
                Case mnu_SubProject
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then

                    End If
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        '  AddRow(oForm)
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        ' RefereshDeleteRow(oForm)
                    End If
                Case mnu_ADD
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        Dim strCode As String = oApplication.Utilities.getMaxCode("@Z_OSUP", "DocEntry")
                        oApplication.Utilities.setEdittextvalue(oForm, "11", strCode)
                        enableControls(oForm, True)
                        oForm.Items.Item("13").Enabled = True
                        oForm.Items.Item("13").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oApplication.Utilities.setEdittextvalue(oForm, "13", "t")
                        oApplication.SBO_Application.SendKeys("{TAB}")
                        oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oForm.Items.Item("13").Enabled = False
                    End If
                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        enableControls(oForm, True)
                    End If
                Case "1283"
                    If pVal.BeforeAction = True Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        oComboBox = oForm.Items.Item("17").Specific
                        If 1 = 1 Then 'oApplication.SBO_Application.MessageBox("Do you want to remove approval template?", , "Yes", "No") = 2 Then
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
            If oForm.TypeEx = frm_SubProject Then
                Select Case BusinessObjectInfo.BeforeAction
                    Case True

                    Case False
                        Select Case BusinessObjectInfo.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OSUP")
                                enableControls(oForm, False)
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                'MsgBox(BusinessObjectInfo.ObjectKey)
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


#Region "Validations"
    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            aForm.Freeze(True)
            '  oComboBox = aForm.Items.Item("17").Specific
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
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "8") = "" Then
                oApplication.Utilities.Message("GL Account Missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select ""DocEntry"" From ""@Z_OSUP"""
            strQuery += " Where "
            strQuery += " ""U_Z_Code"" = '" & oApplication.Utilities.getEditTextvalue(aForm, "4") & "' And ""DocNum"" <> '" & oApplication.Utilities.getEditTextvalue(aForm, "11") & "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oApplication.Utilities.Message("This Entry Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If
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
            aForm.Items.Item("8").Enabled = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#End Region
End Class
