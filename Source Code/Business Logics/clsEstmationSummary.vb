Public Class clsEstmationSummary
    Inherits clsBase

    Private oMatrix As SAPbouiCOM.Matrix
    Dim oStatic As SAPbouiCOM.StaticText
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oBankGrid As SAPbouiCOM.Grid
    Private oGrid As SAPbouiCOM.Grid
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
    Public Sub LoadForm(aCode As String)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_BoM_Summary, frm_BoM_Summary)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oApplication.Utilities.setEdittextvalue(oForm, "4", aCode)
            oForm.Freeze(True)
            DataBind(oForm)
            oForm.Items.Item("3").Visible = False
            oStatic = oForm.Items.Item("1").Specific
            oForm.Title = "Estimation BoM Summary"
            oStatic.Caption = "Document No"
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Public Sub LoadForm_Estimation(aCode As String, aChoice As String)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_BoM_Summary, frm_BoM_Summary)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oApplication.Utilities.setEdittextvalue(oForm, "4", aCode)
            oForm.Freeze(True)
            DataBind(oForm)
            oForm.Items.Item("3").Visible = False
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub DataBind(aForm As SAPbouiCOM.Form)
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select * from ""@Z_OITT1"" where ""U_Z_ItemCode""='" & oApplication.Utilities.getEditTextvalue(aForm, "4") & "'")
        If oRec.RecordCount <= 0 Then
            AddtoUDT(aForm, oApplication.Utilities.getEditTextvalue(aForm, "4"), "Add")
        Else
            AddtoUDT_Initialize_Update(oApplication.Utilities.getEditTextvalue(aForm, "4"), "Update")
        End If
        oGrid = aForm.Items.Item("5").Specific
        oGrid.DataTable.ExecuteQuery("Select * from ""@Z_OITT1"" where ""U_Z_ItemCode""='" & oApplication.Utilities.getEditTextvalue(aForm, "4") & "'")
        FormatGrid(oGrid)
        oGrid.AutoResizeColumns()

    End Sub

    Private Sub FormatGrid(aGrid As SAPbouiCOM.Grid)
        aGrid.Columns.Item("Code").Visible = False
        aGrid.Columns.Item("Name").Visible = False
        aGrid.Columns.Item("U_Z_ItemCode").Visible = False
        aGrid.Columns.Item("U_Z_Type").TitleObject.Caption = "Type"
        aGrid.Columns.Item("U_Z_Type").Editable = False
        aGrid.Columns.Item("U_Z_Cost").TitleObject.Caption = "Cost"
        aGrid.Columns.Item("U_Z_Cost").Editable = False
        aGrid.Columns.Item("U_Z_Markup").TitleObject.Caption = "Markup %"
        aGrid.Columns.Item("U_Z_Markup").Editable = False
        aGrid.Columns.Item("U_Z_Price").TitleObject.Caption = "Sales Price"
        aGrid.Columns.Item("U_Z_Price").Editable = False
        oEditTextColumn = aGrid.Columns.Item("U_Z_Price")
        oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

        oEditTextColumn = aGrid.Columns.Item("U_Z_Cost")
        oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
        aGrid.Columns.Item("U_Z_AvgMarkUp").TitleObject.Caption = "Avg.Markup"
        oEditTextColumn = aGrid.Columns.Item("U_Z_AvgMarkUp")
        oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
        aGrid.Columns.Item("U_Z_AvgMarkUp").Visible = False
        oApplication.Utilities.AssignRowNo(aGrid)

    End Sub
    Private Function AddtoUDT(ByVal aform As SAPbouiCOM.Form, ByVal ItemCode As String, aChoice As String) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim otemp, otemp1 As SAPbobsCOM.Recordset
        Dim strqry, strCode, strqry1, strProCode, ProName, strGLAcc As String
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oUserTable = oApplication.Company.UserTables.Item("Z_OITT1")
        If aChoice = "Add" Then
            AddtoUDT_Initialize(ItemCode, "Add")

        Else
            Validate(oForm)
            oGrid = aform.Items.Item("5").Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oUserTable.GetByKey(oGrid.DataTable.GetValue("Code", intRow))
                oUserTable.Code = oGrid.DataTable.GetValue("Code", intRow)
                oUserTable.Name = oGrid.DataTable.GetValue("Name", intRow)
                oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
                oUserTable.UserFields.Fields.Item("U_Z_Type").Value = oGrid.DataTable.GetValue("U_Z_Type", intRow)
                oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = oGrid.DataTable.GetValue("U_Z_Cost", intRow)
                oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = oGrid.DataTable.GetValue("U_Z_Markup", intRow)
                oUserTable.UserFields.Fields.Item("U_Z_Price").Value = oGrid.DataTable.GetValue("U_Z_Price", intRow)
                oUserTable.Update()
            Next
            UpdateBom(ItemCode)
        End If
        Return True
    End Function

    Private Function AddtoUDT_Initialize(ByVal ItemCode As String, aChoice As String) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim otemp, otemp1 As SAPbobsCOM.Recordset
        Dim strqry, strCode, strqry1, strProCode, ProName, strGLAcc As String
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oUserTable = oApplication.Company.UserTables.Item("Z_OITT1")
        Dim sItemCode As String = "SELECT ""U_Z_ItemCode""  FROM ""@Z_OQUT""  T0 inner join ""@Z_QUT1""  T1 on T1.""DocEntry""=T0.""DocEntry"" where T0.""DocNum""=" & ItemCode

        If aChoice = "Add" Then
            strCode = oApplication.Utilities.getMaxCode("@Z_OITT1", "Code")
            oUserTable.Code = strCode
            oUserTable.Name = strCode
            oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
            oUserTable.UserFields.Fields.Item("U_Z_Type").Value = "Material"
            'otemp1.DoQuery("Select Sum(U_AVGCOST) from ITT1 T0 where Father='" & ItemCode & "' and Type=4")
            otemp1.DoQuery("Select Sum(""U_AVGCOST""*""Quantity""),AVG(""U_MARKUP"") from ITT1 T0  Inner Join  OITM T1 on T1.""ItemCode""=T0.""Code""  INNER JOIN OITB T2 ON T1.""ItmsGrpCod"" = T2.""ItmsGrpCod"" where ""Father"" in ( " & sItemCode & ") and ""Type""=4 and T1.""ItmsGrpCod""<>112")

            oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = otemp1.Fields.Item(0).Value
            oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
            oUserTable.UserFields.Fields.Item("U_Z_Price").Value = otemp1.Fields.Item(0).Value
            oUserTable.Add()

            strCode = oApplication.Utilities.getMaxCode("@Z_OITT1", "Code")
            oUserTable.Code = strCode
            oUserTable.Name = strCode
            oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
            oUserTable.UserFields.Fields.Item("U_Z_Type").Value = "Labor"
            Dim intResourceCode As Integer
            Dim strTemp As String
            strTemp = "SELECT  Top 1 T1.""ResGrpCod"" FROM ORSC T0  INNER JOIN ORSB T1 ON T0.""ResGrpCod"" = T1.""ResGrpCod"" where T1.""ResGrpNam"" like 'Labour%'"
            strTemp = "SELECT Sum(""U_AVGCOST"" * ""Quantity"") ,AVG(""U_MARKUP"")  FROM ITT1 T0  inner Join  ORSC T1 on T1.""VisResCode""=T0.""Code"" where ""Father"" in ( " & sItemCode & ") and T0.""Type""=290 and T1.""ResGrpCod"" =(" & strTemp & ")"
            ' otemp1.DoQuery("Select Sum(U_AVGCOST) from ITT1 T0 where Father='" & ItemCode & "' and Type=4")
            otemp1.DoQuery(strTemp)
            oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = otemp1.Fields.Item(0).Value
            oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
            oUserTable.UserFields.Fields.Item("U_Z_Price").Value = otemp1.Fields.Item(0).Value
            oUserTable.Add()

            strCode = oApplication.Utilities.getMaxCode("@Z_OITT1", "Code")
            oUserTable.Code = strCode
            oUserTable.Name = strCode
            oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
            oUserTable.UserFields.Fields.Item("U_Z_Type").Value = "Machine"

            strTemp = "SELECT  Top 1 T1.""ResGrpCod"" FROM ORSC T0  INNER JOIN ORSB T1 ON T0.""ResGrpCod"" = T1.""ResGrpCod"" where T1.""ResGrpNam"" like 'Machine%'"
            strTemp = "SELECT Sum(""U_AVGCOST"" * ""Quantity""),Avg(""U_MARKUP"")   FROM ITT1 T0  inner Join  ORSC T1 on T1.""VisResCode""=T0.""Code"" where ""Father"" in ( " & sItemCode & ") and T0.""Type""=290 and T1.""ResGrpCod"" =(" & strTemp & ")"
            ' otemp1.DoQuery("Select Sum(U_AVGCOST) from ITT1 T0 where Father='" & ItemCode & "' and Type=4")
            otemp1.DoQuery(strTemp)
            oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = otemp1.Fields.Item(0).Value
            oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
            oUserTable.UserFields.Fields.Item("U_Z_Price").Value = otemp1.Fields.Item(0).Value
            oUserTable.Add()

            strCode = oApplication.Utilities.getMaxCode("@Z_OITT1", "Code")
            oUserTable.Code = strCode
            oUserTable.Name = strCode
            oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
            oUserTable.UserFields.Fields.Item("U_Z_Type").Value = "OutSource"
            strTemp = "SELECT  Top 1 T1.""ResGrpCod"" FROM ORSC T0  INNER JOIN ORSB T1 ON T0.""ResGrpCod"" = T1.""ResGrpCod"" where T1.""ResGrpNam"" like 'OutSource%'"
            strTemp = "SELECT Sum(""U_AVGCOST"" * ""Quantity""),AVG(""U_MARKUP"")   FROM ITT1 T0  inner Join  ORSC T1 on T1.""VisResCode""=T0.""Code"" where  ""Father"" in ( " & sItemCode & ") and T0.""Type""=290 and T1.""ResGrpCod"" =(" & strTemp & ")"
            ' otemp1.DoQuery("Select Sum(U_AVGCOST) from ITT1 T0 where Father='" & ItemCode & "' and Type=4")
            otemp1.DoQuery(strTemp)
            oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = otemp1.Fields.Item(0).Value
            oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
            oUserTable.UserFields.Fields.Item("U_Z_Price").Value = otemp1.Fields.Item(0).Value
            oUserTable.Add()

            strCode = oApplication.Utilities.getMaxCode("@Z_OITT1", "Code")
            oUserTable.Code = strCode
            oUserTable.Name = strCode
            oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
            oUserTable.UserFields.Fields.Item("U_Z_Type").Value = "LED Material"
            otemp1.DoQuery("Select Sum(""U_AVGCOST"" * ""Quantity""),AVG(""U_MARKUP"") from ITT1 T0  Inner Join  OITM T1 on T1.""ItemCode""=T0.""Code""  INNER JOIN OITB T2 ON T1.""ItmsGrpCod"" = T2.""ItmsGrpCod"" where ""Father"" in ( " & sItemCode & ") and ""Type""=4 and T1.""ItmsGrpCod""=112")
            oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = otemp1.Fields.Item(0).Value
            oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
            oUserTable.UserFields.Fields.Item("U_Z_Price").Value = otemp1.Fields.Item(0).Value
            oUserTable.Add()
        End If
        otemp1.DoQuery("Update ""@Z_OITT1"" set ""U_Z_Price""=""U_Z_Cost"" * ""U_Z_Markup"" where ""U_Z_ItemCode""='" & ItemCode & "'")
        Return True
    End Function

    Private Function AddtoUDT_Initialize_Update(ByVal ItemCode As String, aChoice As String) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim otemp, otemp1 As SAPbobsCOM.Recordset
        Dim strqry, strCode, strqry1, strProCode, ProName, strGLAcc As String
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oUserTable = oApplication.Company.UserTables.Item("Z_OITT1")
        Dim sItemCode As String = "SELECT ""U_Z_ItemCode"" FROM ""@Z_OQUT""  T0 inner join ""@Z_QUT1""  T1 on T1.""DocEntry""=T0.""DocEntry"" where T0.""DocNum""=" & ItemCode

        Dim dblAvgCost As Double
        If aChoice = "Update" Then
            otemp1.DoQuery("Select ""Code"",* from ""@Z_OITT1"" where ""U_Z_ItemCode""='" & ItemCode & "' and ""U_Z_Type""='Material'")
            If otemp1.RecordCount > 0 Then
                strCode = otemp1.Fields.Item("Code").Value
                oUserTable.GetByKey(strCode)
                oUserTable.Name = strCode
                oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
                oUserTable.UserFields.Fields.Item("U_Z_Type").Value = "Material"

                otemp1.DoQuery("Select Sum(""U_AVGCOST"" * ""Quantity""),AVG(""U_MARKUP"") from ITT1 T0  Inner Join  OITM T1 on T1.""ItemCode""=T0.""Code""  INNER JOIN OITB T2 ON T1.""ItmsGrpCod"" = T2.""ItmsGrpCod"" where ""Father"" in ( " & sItemCode & ") and ""Type"" = 4 and T1.""ItmsGrpCod"" <> 112")
                dblAvgCost = otemp1.Fields.Item(0).Value
                oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
                oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = dblAvgCost ' otemp1.Fields.Item(0).Value
                oUserTable.Update()
            End If
            Dim intResourceCode As Integer
            Dim strTemp As String

            otemp1.DoQuery("Select ""Code"",* from ""@Z_OITT1"" where ""U_Z_ItemCode"" = '" & ItemCode & "' and ""U_Z_Type"" = 'Labor'")
            If otemp1.RecordCount > 0 Then
                strCode = otemp1.Fields.Item("Code").Value
                oUserTable.GetByKey(strCode)
                oUserTable.Name = strCode
                oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
                oUserTable.UserFields.Fields.Item("U_Z_Type").Value = "Labor"
                strTemp = "SELECT  Top 1 T1.""ResGrpCod"" FROM ORSC T0  INNER JOIN ORSB T1 ON T0""ResGrpCod"" = T1.""ResGrpCod"" where T1.""ResGrpNam"" like 'Labour%'"
                strTemp = "SELECT Sum(""U_AVGCOST"" * ""Quantity"" ),AVG(""U_MARKUP"")   FROM ITT1 T0  inner Join  ORSC T1 on T1.""VisResCode"" = T0.""Code"" where ""Father"" in ( " & sItemCode & ") and T0.""Type"" = 290 and T1.""ResGrpCod"" = (" & strTemp & ")"
                otemp1.DoQuery(strTemp)
                dblAvgCost = otemp1.Fields.Item(0).Value

                oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = dblAvgCost
                oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
                oUserTable.Update()
            End If

            otemp1.DoQuery("Select ""Code"",* from ""@Z_OITT1"" where ""U_Z_ItemCode"" = '" & ItemCode & "' and ""U_Z_Type"" = 'Machine'")
            If otemp1.RecordCount > 0 Then
                strCode = otemp1.Fields.Item("Code").Value
                oUserTable.GetByKey(strCode)
                oUserTable.Name = strCode
                oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
                oUserTable.UserFields.Fields.Item("U_Z_Type").Value = "Machine"
                strTemp = "SELECT  Top 1 T1.""ResGrpCod"" FROM ORSC T0  INNER JOIN ORSB T1 ON T0.""ResGrpCod"" = T1.""ResGrpCod"" where T1.""ResGrpNam"" like 'Machine%'"
                strTemp = "SELECT Sum(""U_AVGCOST"" * ""Quantity""),AVG(""U_MARKUP"") FROM ITT1 T0  inner Join  ORSC T1 on T1.""VisResCode""=T0.""Code"" where  ""Father"" in ( " & sItemCode & ") and T0.""Type"" = 290 and T1.""ResGrpCod"" =(" & strTemp & ")"
                ' otemp1.DoQuery("Select Sum(U_AVGCOST) from ITT1 T0 where Father='" & ItemCode & "' and Type=4")
                otemp1.DoQuery(strTemp)
                dblAvgCost = otemp1.Fields.Item(0).Value

                oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = dblAvgCost
                oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
                oUserTable.Update()
            End If

            otemp1.DoQuery("Select ""Code"",* from ""@Z_OITT1"" where ""U_Z_ItemCode"" = '" & ItemCode & "' and ""U_Z_Type"" = 'OutSource'")
            If otemp1.RecordCount > 0 Then
                strCode = otemp1.Fields.Item("Code").Value
                oUserTable.GetByKey(strCode)
                oUserTable.Name = strCode
                oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
                oUserTable.UserFields.Fields.Item("U_Z_Type").Value = "OutSource"
                '   oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item("U_Z_Markup").Value
                strTemp = "SELECT  Top 1 T1.""ResGrpCod"" FROM ORSC T0  INNER JOIN ORSB T1 ON T0.""ResGrpCod"" = T1.""ResGrpCod"" where T1.""ResGrpNam"" like 'OutSource%'"
                strTemp = "SELECT Sum(""U_AVGCOST"" * ""Quantity""),AVG(""U_MARKUP"")   FROM ITT1 T0  inner Join  ORSC T1 on T1.""VisResCode"" = T0.""Code"" where ""Father"" in ( " & sItemCode & ")  and T0.""Type""=290 and T1.""ResGrpCod"" =(" & strTemp & ")"
                ' otemp1.DoQuery("Select Sum(U_AVGCOST) from ITT1 T0 where Father='" & ItemCode & "' and Type=4")
                otemp1.DoQuery(strTemp)
                dblAvgCost = otemp1.Fields.Item(0).Value
                oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
                oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = dblAvgCost
                oUserTable.Update()
            End If
            otemp1.DoQuery("Select ""Code"",* from ""@Z_OITT1"" where ""U_Z_ItemCode"" = '" & ItemCode & "' and ""U_Z_Type"" = 'LED Material'")
            If otemp1.RecordCount > 0 Then
                strCode = otemp1.Fields.Item("Code").Value
                oUserTable.GetByKey(strCode)
                oUserTable.Name = strCode
                oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
                oUserTable.UserFields.Fields.Item("U_Z_Type").Value = "LED Material"
                ' oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item("U_Z_Markup").Value
                otemp1.DoQuery("Select Sum(""U_AVGCOST""*""Quantity""),AVG(""U_MARKUP"") from ITT1 T0  Inner Join  OITM T1 on T1.""ItemCode""=T0.""Code""  INNER JOIN OITB T2 ON T1.""ItmsGrpCod"" = T2.""ItmsGrpCod"" where ""Father"" in ( " & sItemCode & ") and ""Type"" = 4 and T1.""ItmsGrpCod"" = 112")
                oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = otemp1.Fields.Item(0).Value
                dblAvgCost = otemp1.Fields.Item(0).Value
                oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = dblAvgCost
                oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
                oUserTable.Update()
            End If
        End If
        otemp1.DoQuery("Update ""@Z_OITT1"" set ""U_Z_Price"" = ""U_Z_Cost"" * ""U_Z_Markup"" where ""U_Z_ItemCode"" = '" & ItemCode & "'")
        Return True
    End Function
    Private Sub UpdateBom(aCode As String)
        Dim oTemp, otemp1 As SAPbobsCOM.Recordset
        Dim strTemp As String
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select * from ""@Z_OITT1"" where ""U_Z_ItemCode"" = '" & aCode & "' order by ""Code"" ")
        For intRow As Integer = 0 To oTemp.RecordCount - 1
            If oTemp.Fields.Item("U_Z_Type").Value = "Material" Then
                strTemp = "Update  ITT1 set ""U_Markup"" =" & oTemp.Fields.Item("U_Z_Markup").Value & " where ""Father"" = '" & aCode & "' and ""Type"" = 4 and ""Code"" in (Select ""ItemCode"" from OITM T1 inner Join OITB T2 on T2.""ItmsGrpCod"" = T1.""ItmsGrpCod"" where T2.""ItmsGrpCod"" <> 112)"
                otemp1.DoQuery(strTemp)
            ElseIf oTemp.Fields.Item("U_Z_Type").Value = "Labor" Then
                '  strTemp = "SELECT Sum(U_AVGCOST)   FROM ITT1 T0  inner Join  ORSC T1 on T1.VisResCode=T0.Code where Father='" & ItemCode & "'and T0.Type=290 and T1.ResGrpCod =(" & strTemp & ")"
                strTemp = "SELECT  ""VisResCode"" FROM ORSC T0  INNER JOIN ORSB T1 ON T0.""ResGrpCod"" = T1.""ResGrpCod"" where T1.""ResGrpNam"" like 'Labour%'"
                'strTemp = "(Select VisResCode from ORSC where "
                strTemp = "Update  ITT1 set  ""U_Markup"" =" & oTemp.Fields.Item("U_Z_MarkUp").Value & " where ""Father"" = '" & aCode & "' and ""Type"" = 290 and ""Code"" in (" & strTemp & ")"
                otemp1.DoQuery(strTemp)
            ElseIf oTemp.Fields.Item("U_Z_Type").Value = "Machine" Then
                strTemp = "SELECT  ""VisResCode"" FROM ORSC T0  INNER JOIN ORSB T1 ON T0.""ResGrpCod"" = T1.""ResGrpCod"" where T1.""ResGrpNam"" like 'Machine%'"
                strTemp = "Update  ITT1 set  ""U_Markup"" =" & oTemp.Fields.Item("U_Z_MarkUp").Value & " where ""Father"" = '" & aCode & "' and ""Type""= 290 and ""Code"" in (" & strTemp & ")"
                otemp1.DoQuery(strTemp)
            ElseIf oTemp.Fields.Item("U_Z_Type").Value = "OutSource" Then
                strTemp = "SELECT  ""VisResCode"" FROM ORSC T0  INNER JOIN ORSB T1 ON T0.""ResGrpCod"" = T1.""ResGrpCod"" where T1.""ResGrpNam"" like 'OutSource%'"
                strTemp = "Update  ITT1 set ""U_Markup"" =" & oTemp.Fields.Item("U_Z_MarkUp").Value & " where ""Father"" = '" & aCode & "' and ""Type"" = 290 and ""Code"" in(" & strTemp & ")"
                otemp1.DoQuery(strTemp)
            ElseIf oTemp.Fields.Item("U_Z_Type").Value = "LED Material" Then
                strTemp = "Update ITT1 set ""U_Markup"" =" & oTemp.Fields.Item("U_Z_Markup").Value & " where ""Father"" = '" & aCode & "' and ""Type"" = 4 and ""Code"" in (Select ""ItemCode"" from OITM T1 inner Join OITB T2 on T2.""ItmsGrpCod"" = T1.""ItmsGrpCod"" where T2.""ItmsGrpCod"" = 112)"
                otemp1.DoQuery(strTemp)
            End If
            oTemp.MoveNext()
        Next
        otemp1.DoQuery("Update ITT1 set ""Price""= ""U_AVGCOST"" * ""U_MARKUP"" where ""Father"" ='" & aCode & "'")

    End Sub
    Private Sub Validate(aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oGrid = aForm.Items.Item("5").Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.DataTable.SetValue("U_Z_Price", intRow, (oGrid.DataTable.GetValue("U_Z_Cost", intRow) * oGrid.DataTable.GetValue("U_Z_Markup", intRow)))
            Next
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
    Public Sub addcontrols(ByVal aforma As SAPbouiCOM.Form)
        Try
            '   oApplication.Utilities.AddControls(oForm, "_301", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", 0, 0, "2", "View Summary", 120)

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BoM_Summary Then
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
                                '  addcontrols(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" Then
                                    If AddtoUDT(oForm, oApplication.Utilities.getEditTextvalue(oForm, "4"), "Update") = True Then
                                        oForm.Close()
                                        frm_SourceBoM.Select()
                                        oApplication.SBO_Application.ActivateMenuItem("1304")
                                    End If
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
