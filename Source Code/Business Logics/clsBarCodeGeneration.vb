Public Class clsBarCodeGeneration
    Inherits clsBase

    Private oMatrix As SAPbouiCOM.Matrix
    Dim oStatic As SAPbouiCOM.StaticText
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
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
            oForm = oApplication.Utilities.LoadForm(xml_BarCode, frm_BarCode)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Items.Item("12").Visible = False
            Dim otest As SAPbobsCOM.Recordset
            otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otest.DoQuery("Select isnull(""U_SEASON"",'') ,Count(*) from OITM group by isnull(""U_SEASON"",'')")
            oCombobox = oForm.Items.Item("8").Specific
            For intRow As Integer = 0 To otest.RecordCount - 1
                oCombobox.ValidValues.Add(otest.Fields.Item(0).Value, otest.Fields.Item(0).Value)
                otest.MoveNext()
            Next
            oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            otest.DoQuery("Select isnull(""U_BRAND"",'') ,Count(*) from OITM group by isnull(""U_BRAND"",'')")
            oCombobox = oForm.Items.Item("10").Specific
            For intRow As Integer = 0 To otest.RecordCount - 1
                oCombobox.ValidValues.Add(otest.Fields.Item(0).Value, otest.Fields.Item(0).Value)
                otest.MoveNext()
            Next
            oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
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

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BarCode Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                blnFormLoaded = True
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to generate BarCodes?", , "Continue", "Cancel") = 2 Then
                                        Exit Sub
                                    End If
                                    If oApplication.Utilities.generateBarCodes(oForm) = True Then
                                        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    End If


                                End If
                                If pVal.ItemUID = "13" Then
                                    If oApplication.SBO_Application.MessageBox("Confirm to process generated BarCode ?", , "Continue", "Cancel") = 2 Then
                                        Exit Sub
                                    End If
                                    Dim oGrid As SAPbouiCOM.Grid
                                    oGrid = oForm.Items.Item("12").Specific
                                    For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                        oStatic = oForm.Items.Item("11").Specific
                                        oStatic.Caption = "Processing Item Code : " & oGrid.DataTable.GetValue(0, intRow)
                                        oApplication.Utilities.AddBarCode(oGrid.DataTable.GetValue(0, intRow), oGrid.DataTable.GetValue(2, intRow), 0)
                                    Next
                                    oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    oForm.Close()
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
                Case mnu_BarCode
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
            strFBank = CType(oForm.Items.Item("8").Specific, SAPbouiCOM.ComboBox).Selected.Value
            strTBank = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.ComboBox).Selected.Value

            If strFBank = "" Then
                oApplication.Utilities.Message("Select From Bank ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strTBank = "" Then
                oApplication.Utilities.Message("Select To Bank ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

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
