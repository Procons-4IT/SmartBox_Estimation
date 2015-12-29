
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Threading
Imports System.Collections.Generic
Public Class clsBarCodePrinting
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
    Dim cryRpt As New ReportDocument
    Private ds As New Barcode       '(dataset)
    Private oDRow As DataRow

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub


    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm("frm_BarCodePrinting.xml", "frm_PrintBarCode")
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

    
    Private Sub openFileDialog()
        Dim objPL As New frmReportViewer
        objPL.iniViewer = AddressOf objPL.GenerateReport
        objPL.rptViewer.ReportSource = cryRpt
        objPL.rptViewer.Refresh()
        objPL.WindowState = FormWindowState.Maximized
        objPL.ShowDialog()
        System.Threading.Thread.CurrentThread.Abort()
    End Sub
    Private Sub addCrystal(ByVal ds1 As DataSet, ByVal aChoice As String)
        Dim strFilename, stfilepath As String
        Dim strReportFileName As String
        If aChoice = "BarCode" Then
            strReportFileName = "rptBarcode.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\BarCode"
        ElseIf aChoice = "Agreement" Then
            strReportFileName = "Agreement.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Rental_Agreement"
        Else
            strReportFileName = "AcctStatement.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\AccountStatement"
        End If
        strReportFileName = strReportFileName
        strFilename = strFilename & ".pdf"
        stfilepath = System.Windows.Forms.Application.StartupPath & "\Reports\" & strReportFileName
        If File.Exists(stfilepath) = False Then
            oApplication.Utilities.Message("Report does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        If File.Exists(strFilename) Then
            File.Delete(strFilename)
        End If
        ' If ds1.Tables.Item("AccountBalance").Rows.Count > 0 Then
        If 1 = 1 Then
            cryRpt.Load(System.Windows.Forms.Application.StartupPath & "\Reports\" & strReportFileName)
            cryRpt.SetDataSource(ds1)
            If "T" = "T" Then
                Dim mythread As New System.Threading.Thread(AddressOf OpenFileDialog)
                mythread.SetApartmentState(ApartmentState.STA)
                mythread.Start()
                mythread.Join()
                ds1.Clear()
            Else
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New  _
                DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
                CrDiskFileDestinationOptions.DiskFileName = strFilename
                CrExportOptions = cryRpt.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                cryRpt.Export()
                cryRpt.Close()
                Dim x As System.Diagnostics.ProcessStartInfo
                x = New System.Diagnostics.ProcessStartInfo
                x.UseShellExecute = True
                x.FileName = strFilename
                System.Diagnostics.Process.Start(x)
                x = Nothing
                ' objUtility.ShowSuccessMessage("Report exported into PDF File")
            End If

        Else
            ' objUtility.ShowWarningMessage("No data found")
        End If

    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = "frm_PrintBarCode" Then
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
                                    'If oApplication.SBO_Application.MessageBox("Do you want to generate BarCodes?", , "Continue", "Cancel") = 2 Then
                                    '    Exit Sub
                                    'End If
                                    If oApplication.Utilities.PrintBarCode(oForm) = True Then
                                        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    End If


                                End If
                                If pVal.ItemUID = "13" Then
                                    If oApplication.SBO_Application.MessageBox("Confirm to Print BarCodes ?", , "Continue", "Cancel") = 2 Then
                                        Exit Sub
                                    End If
                                    oApplication.Utilities.PrintbarCode_Report(oForm)
                                    'oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    'oForm.Close()
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
                Case "Menu_B02"
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
