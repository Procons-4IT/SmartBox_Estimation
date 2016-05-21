Imports System.IO
Public Class clsAppHisDetails
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oCombo As SAPbouiCOM.ComboBoxColumn
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtDocumentList As SAPbouiCOM.DataTable
    Private dtHistoryList As SAPbouiCOM.DataTable
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm(ByVal oForm As SAPbouiCOM.Form, ByVal DocNo As String)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_AppHisDetails, frm_AppHisDetails)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Items.Item("3").Visible = True
            LoadViewHistory(oForm, DocNo)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Public Sub assignMatrixLineno(ByVal aGrid As SAPbouiCOM.Grid, ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        For intNo As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            aGrid.RowHeaders.SetText(intNo, intNo + 1)
        Next
        aGrid.Columns.Item("RowsHeader").TitleObject.Caption = "#"
        aform.Freeze(False)
    End Sub
    Public Sub LoadViewHistory(ByVal aForm As SAPbouiCOM.Form, ByVal strDocEntry As String)
        Try
            aForm.Freeze(True)
            Dim sQuery As String
            oGrid = aForm.Items.Item("3").Specific

            sQuery = " Select ""DocEntry"",""U_Z_DocEntry"",""U_Z_DocType"",""U_Z_EmpId"",""U_Z_EmpName"",""U_Z_ApproveBy"",""CreateDate "",""CreateTime,""UpdateDate"",""UpdateTime"",""U_Z_AppStatus"",""U_Z_Remarks"" From ""@P_APHIS"" "
            sQuery += " Where ""U_Z_DocType"" = 'B'"
            sQuery += " And ""U_Z_DocEntry"" = '" + strDocEntry + "'"
            oGrid.DataTable.ExecuteQuery(sQuery)
            formatHistory(aForm)
            assignMatrixLineno(oGrid, aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub
    Private Sub formatHistory(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid
            Dim oComboBox As SAPbouiCOM.ComboBox
            Dim oGridCombo As SAPbouiCOM.ComboBoxColumn
            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            oGrid = aForm.Items.Item("3").Specific
            oGrid.Columns.Item("DocEntry").Visible = False
            oGrid.Columns.Item("U_Z_DocEntry").TitleObject.Caption = "Reference No."
            oGrid.Columns.Item("U_Z_DocEntry").Visible = False
            oGrid.Columns.Item("U_Z_DocType").Visible = False
            oGrid.Columns.Item("U_Z_EmpId").TitleObject.Caption = "Employee ID"
            oEditTextColumn = oGrid.Columns.Item("U_Z_EmpId")
            oEditTextColumn.LinkedObjectType = "171"
            oGrid.Columns.Item("U_Z_EmpId").Visible = False
            oGrid.Columns.Item("U_Z_EmpName").TitleObject.Caption = "Employee Name"
            oGrid.Columns.Item("U_Z_EmpName").Visible = False
            oGrid.Columns.Item("U_Z_ApproveBy").TitleObject.Caption = "Approved By"
            oGrid.Columns.Item("U_Z_AppStatus").TitleObject.Caption = "Approved Status"
            oGrid.Columns.Item("U_Z_AppStatus").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oGridCombo = oGrid.Columns.Item("U_Z_AppStatus")
            oGridCombo.ValidValues.Add("A", "Approved")
            oGridCombo.ValidValues.Add("R", "Rejected")
            oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            oGrid.Columns.Item("U_Z_Remarks").TitleObject.Caption = "Remarks"
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oGrid.AutoResizeColumns()
            aForm.Freeze(False)
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("U_Z_ApproveBy", intRow) = oApplication.Company.UserName Then
                    oGrid.Columns.Item("RowsHeader").Click(intRow, False, False)
                    aForm.Freeze(False)
                    Exit Sub
                End If
            Next
            aForm.Items.Item("8").Enabled = True
            aForm.Items.Item("10").Enabled = True
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_AppHisDetails Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_CLICK

                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_CLICK

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Then
                                    'oApplication.Utilities.Resize(oForm)
                                End If
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
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Class
