Imports System.IO

Public Class clsBoMEstimation1
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
    Dim oDBDataSourceLineZ_1 As SAPbouiCOM.DBDataSource
    Private InvForConsumedItems, count As Integer
    Private blnFlag As Boolean = False
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim blnFormLoaded As Boolean = False
    Public MatrixId As String
    Public intSelectedMatrixrow As Integer = 0
    Private RowtoDelete As Integer
    Dim oDataSrc_Line, oDataSrc_Line3 As SAPbouiCOM.DBDataSource

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub


    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_BoM_Estimation, frm_BoM_Estimation)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.DataBrowser.BrowseBy = "4"
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm.Items.Item("4").Enabled = False
                oForm.Items.Item("6").Enabled = False
            Else
                oForm.Items.Item("4").Enabled = True
                oForm.Items.Item("6").Enabled = True
            End If
            addChooseFromListConditions(oForm)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select ""series"",""SeriesName""  from NNM1 where ""ObjectCode""='Z_OQUT'")
            oCombobox = oForm.Items.Item("20").Specific
            For introw As Integer = 0 To oRecordSet.RecordCount - 1
                oCombobox.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
                oRecordSet.MoveNext()
            Next
            oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            oRecordSet.DoQuery("select ""SlpCode"",""SlpName""  from OSLP order by ""SlpCode""")
            oCombobox = oForm.Items.Item("38").Specific
            For introw As Integer = 0 To oRecordSet.RecordCount - 1
                oCombobox.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
                oRecordSet.MoveNext()
            Next
            oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oMatrix = oForm.Items.Item("17").Specific
            Dim oColumn As SAPbouiCOM.Column
            oColumn = oMatrix.Columns.Item("V_5")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.PaneLevel = 1
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub LoadForm_View(aDocNum As String, Optional aSlpCode As String = "-1")
        Try
            oForm = oApplication.Utilities.LoadForm(xml_BoM_Estimation, frm_BoM_Estimation)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.DataBrowser.BrowseBy = "4"
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            'oForm.EnableMenu(mnu_ADD_ROW, True)
            ' oForm.EnableMenu(mnu_DELETE_ROW, True)
            If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm.Items.Item("4").Enabled = False
                oForm.Items.Item("6").Enabled = False
            Else
                oForm.Items.Item("4").Enabled = True
                oForm.Items.Item("6").Enabled = True
            End If
            addChooseFromListConditions(oForm)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select ""series"",""SeriesName""  from NNM1 where ""ObjectCode""='Z_OQUT'")
            oCombobox = oForm.Items.Item("20").Specific
            For introw As Integer = 0 To oRecordSet.RecordCount - 1
                oCombobox.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
                oRecordSet.MoveNext()
            Next
            oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            oRecordSet.DoQuery("select ""SlpCode"",""SlpName""  from OSLP order by ""SlpCode""")
            oCombobox = oForm.Items.Item("38").Specific
            For introw As Integer = 0 To oRecordSet.RecordCount - 1
                oCombobox.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
                oRecordSet.MoveNext()
            Next
            oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            oCombobox.Select(aSlpCode, SAPbouiCOM.BoSearchKey.psk_ByValue)

            oMatrix = oForm.Items.Item("17").Specific
            Dim oColumn As SAPbouiCOM.Column
            oColumn = oMatrix.Columns.Item("V_5")
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '  oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oApplication.Utilities.setEdittextvalue(oForm, "4", aDocNum)
            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            oForm.PaneLevel = 1
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Function GetUnitPrice(aCode As String) As Double
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim dblPrice, dblQuanity As Double
        If blnIsHana Then
            oTest.DoQuery("Select IFNull(Sum(""Price"" * ""Quantity""),0) from ITT1 where ""Father""='" & aCode & "'")
        Else
            oTest.DoQuery("Select isnull(Sum(Price * Quantity),0) from ITT1 where Father='" & aCode & "'")
        End If

        dblPrice = oTest.Fields.Item(0).Value

        oTest.DoQuery("Select ""Quantity"" from OITT where ""Code""='" & aCode & "'")
        dblQuanity = oTest.Fields.Item(0).Value
        dblPrice = dblPrice / dblQuanity

        Return dblPrice

    End Function

    Private Sub addChooseFromListConditions(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList

            oCFLs = oForm.ChooseFromLists

            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "TreeType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_3")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)


        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub AddMode(ByVal aform As SAPbouiCOM.Form)
        Dim strCode As String
        strCode = oApplication.Utilities.getMaxCode("@Z_OQUT", "DocEntry")
        aform.Items.Item("6").Enabled = True
        aform.Items.Item("4").Enabled = True
        aform.Items.Item("20").Enabled = True
        aform.Items.Item("22").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oApplication.Utilities.setEdittextvalue(aform, "4", strCode)
        aform.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oApplication.Utilities.setEdittextvalue(aform, "6", "t")
        oApplication.SBO_Application.SendKeys("{TAB}")
        aform.Items.Item("8").Enabled = True
        '  aform.Items.Item("8").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        aform.Items.Item("6").Enabled = False
        aform.Items.Item("4").Enabled = False
        oCombobox = aform.Items.Item("20").Specific
        Try
            oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        Catch ex As Exception
        End Try
        oForm.Items.Item("1").Enabled = True

    End Sub


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
            otemp1.DoQuery("Select Sum(""U_AVGCOST""),AVG(""U_MARKUP"") from ITT1 T0  Inner Join  OITM T1 on T1.""ItemCode""=T0.""Code""  INNER JOIN OITB T2 ON T1.""ItmsGrpCod"" = T2.""ItmsGrpCod"" where ""Father"" in ( " & sItemCode & ") and ""Type""=4 and T1.""ItmsGrpCod""<>112")

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
            strTemp = "SELECT Sum(""U_AVGCOST"") ,AVG(""U_MARKUP"")  FROM ITT1 T0  inner Join  ORSC T1 on T1.""VisResCode""=T0.""Code"" where ""Father"" in ( " & sItemCode & ") and T0.""Type""=290 and T1.""ResGrpCod"" =(" & strTemp & ")"
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
            strTemp = "SELECT Sum(""U_AVGCOST""),Avg(""U_MARKUP"")   FROM ITT1 T0  inner Join  ORSC T1 on T1.""VisResCode""=T0.""Code"" where ""Father"" in ( " & sItemCode & ") and T0.""Type""=290 and T1.""ResGrpCod"" =(" & strTemp & ")"
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
            strTemp = "SELECT Sum(""U_AVGCOST""),AVG(""U_MARKUP"")   FROM ITT1 T0  inner Join  ORSC T1 on T1.""VisResCode""=T0.""Code"" where  ""Father"" in ( " & sItemCode & ") and T0.""Type""=290 and T1.""ResGrpCod"" =(" & strTemp & ")"
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
            otemp1.DoQuery("Select Sum(""U_AVGCOST""),AVG(""U_MARKUP"") from ITT1 T0  Inner Join  OITM T1 on T1.""ItemCode""=T0.""Code""  INNER JOIN OITB T2 ON T1.""ItmsGrpCod"" = T2.""ItmsGrpCod"" where ""Father"" in ( " & sItemCode & ") and ""Type""=4 and T1.""ItmsGrpCod""=112")
            oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = otemp1.Fields.Item(0).Value
            oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
            oUserTable.UserFields.Fields.Item("U_Z_Price").Value = otemp1.Fields.Item(0).Value
            oUserTable.Add()
        End If

        Return True
    End Function

    Private Function AddtoUDT_Initialize_Update(ByVal ItemCode As String, aChoice As String) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim otemp, otemp1 As SAPbobsCOM.Recordset
        Dim strqry, strCode, strqry1, strProCode, ProName, strGLAcc As String
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oUserTable = oApplication.Company.UserTables.Item("Z_OITT1")
        Dim sItemCode As String = "SELECT ""U_Z_ItemCode""  FROM ""@Z_OQUT""  T0 inner join ""@Z_QUT1""  T1 on T1.""DocEntry""=T0.""DocEntry"" where T0.""DocNum""=" & ItemCode

        Dim dblAvgCost As Double
        If aChoice = "Update" Then
            otemp1.DoQuery("Select ""Code"",* from ""@Z_OITT1"" where ""U_Z_ItemCode""='" & ItemCode & "' and ""U_Z_Type""='Material'")
            If otemp1.RecordCount > 0 Then
                strCode = otemp1.Fields.Item("Code").Value
                oUserTable.GetByKey(strCode)
                oUserTable.Name = strCode
                oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
                oUserTable.UserFields.Fields.Item("U_Z_Type").Value = "Material"

                otemp1.DoQuery("Select Sum(""U_AVGCOST""),AVG(""U_MARKUP"") from ITT1 T0  Inner Join  OITM T1 on T1.""ItemCode""=T0.""Code""  INNER JOIN OITB T2 ON T1.""ItmsGrpCod"" = T2.""ItmsGrpCod"" where ""Father"" in ( " & sItemCode & ") and ""Type""=4 and T1.""ItmsGrpCod""<>112")
                dblAvgCost = otemp1.Fields.Item(0).Value
                oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
                oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = dblAvgCost ' otemp1.Fields.Item(0).Value
                oUserTable.Update()
            End If
            Dim intResourceCode As Integer
            Dim strTemp As String

            otemp1.DoQuery("Select ""Code"",* from ""@Z_OITT1"" where ""U_Z_ItemCode""='" & ItemCode & "' and ""U_Z_Type""='Labor'")
            If otemp1.RecordCount > 0 Then
                strCode = otemp1.Fields.Item("Code").Value
                oUserTable.GetByKey(strCode)
                oUserTable.Name = strCode
                oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
                oUserTable.UserFields.Fields.Item("U_Z_Type").Value = "Labor"
                strTemp = "SELECT  Top 1 T1.""ResGrpCod"" FROM ORSC T0  INNER JOIN ORSB T1 ON T0.""ResGrpCod"" = T1.""ResGrpCod"" where T1.""ResGrpNam"" like 'Labour%'"
                strTemp = "SELECT Sum(""U_AVGCOST""),AVG(""U_MARKUP"")   FROM ITT1 T0  inner Join  ORSC T1 on T1.""VisResCode""=T0.""Code"" where ""Father"" in ( " & sItemCode & ") and T0.""Type""=290 and T1.""ResGrpCod"" =(" & strTemp & ")"
                otemp1.DoQuery(strTemp)
                dblAvgCost = otemp1.Fields.Item(0).Value

                oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = dblAvgCost
                oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
                oUserTable.Update()
            End If

            otemp1.DoQuery("Select ""Code"",* from ""@Z_OITT1"" where ""U_Z_ItemCode""='" & ItemCode & "' and ""U_Z_Type""='Machine'")
            If otemp1.RecordCount > 0 Then
                strCode = otemp1.Fields.Item("Code").Value
                oUserTable.GetByKey(strCode)
                oUserTable.Name = strCode
                oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
                oUserTable.UserFields.Fields.Item("U_Z_Type").Value = "Machine"
                strTemp = "SELECT  Top 1 T1.""ResGrpCod"" FROM ORSC T0  INNER JOIN ORSB T1 ON T0.""ResGrpCod"" = T1.""ResGrpCod"" where T1.""ResGrpNam"" like 'Machine%'"
                strTemp = "SELECT Sum(""U_AVGCOST""),AVG(""U_MARKUP"")   FROM ITT1 T0  inner Join  ORSC T1 on T1.""VisResCode""=T0.""Code"" where  ""Father"" in ( " & sItemCode & ") and T0.""Type""=290 and T1.""ResGrpCod"" =(" & strTemp & ")"
                ' otemp1.DoQuery("Select Sum(U_AVGCOST) from ITT1 T0 where Father='" & ItemCode & "' and Type=4")
                otemp1.DoQuery(strTemp)
                dblAvgCost = otemp1.Fields.Item(0).Value

                oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = dblAvgCost
                oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
                oUserTable.Update()
            End If

            otemp1.DoQuery("Select ""Code"",* from ""@Z_OITT1"" where ""U_Z_ItemCode""='" & ItemCode & "' and ""U_Z_Type""='OutSource'")
            If otemp1.RecordCount > 0 Then
                strCode = otemp1.Fields.Item("Code").Value
                oUserTable.GetByKey(strCode)
                oUserTable.Name = strCode
                oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
                oUserTable.UserFields.Fields.Item("U_Z_Type").Value = "OutSource"
                '   oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item("U_Z_Markup").Value
                strTemp = "SELECT  Top 1 T1.""ResGrpCod"" FROM ORSC T0  INNER JOIN ORSB T1 ON T0.""ResGrpCod"" = T1.""ResGrpCod"" where T1.""ResGrpNam"" like 'OutSource%'"
                strTemp = "SELECT Sum(""U_AVGCOST""),AVG(""U_MARKUP"")   FROM ITT1 T0  inner Join  ORSC T1 on T1.""VisResCode""=T0.""Code"" where ""Father"" in ( " & sItemCode & ")  and T0.""Type""=290 and T1.""ResGrpCod"" =(" & strTemp & ")"
                ' otemp1.DoQuery("Select Sum(U_AVGCOST) from ITT1 T0 where Father='" & ItemCode & "' and Type=4")
                otemp1.DoQuery(strTemp)
                dblAvgCost = otemp1.Fields.Item(0).Value
                oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
                oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = dblAvgCost
                oUserTable.Update()
            End If
            otemp1.DoQuery("Select ""Code"",* from ""@Z_OITT1"" where ""U_Z_ItemCode""='" & ItemCode & "' and ""U_Z_Type""='LED Material'")
            If otemp1.RecordCount > 0 Then
                strCode = otemp1.Fields.Item("Code").Value
                oUserTable.GetByKey(strCode)
                oUserTable.Name = strCode
                oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = ItemCode
                oUserTable.UserFields.Fields.Item("U_Z_Type").Value = "LED Material"
                ' oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item("U_Z_Markup").Value
                otemp1.DoQuery("Select Sum(""U_AVGCOST""),AVG(""U_MARKUP"") from ITT1 T0  Inner Join  OITM T1 on T1.""ItemCode""=T0.""Code""  INNER JOIN OITB T2 ON T1.""ItmsGrpCod"" = T2.""ItmsGrpCod"" where ""Father"" in ( " & sItemCode & ") and ""Type""=4 and T1.""ItmsGrpCod""=112")
                oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = otemp1.Fields.Item(0).Value
                dblAvgCost = otemp1.Fields.Item(0).Value
                oUserTable.UserFields.Fields.Item("U_Z_Cost").Value = dblAvgCost
                oUserTable.UserFields.Fields.Item("U_Z_Markup").Value = otemp1.Fields.Item(1).Value
                oUserTable.Update()
            End If
        End If
        'otemp1.DoQuery("Update [@Z_OITT1] set U_Z_Price=U_Z_Cost * U_Z_Markup where U_Z_ItemCode='" & ItemCode & "'")
        Return True
    End Function

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BoM_Estimation Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    If Validation_Form(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        UpdateAttachment(oForm)
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                If (pVal.ItemUID = "17" Or pVal.ItemUID = "100") And pVal.Row > 0 Then
                                    oMatrix = oForm.Items.Item("17").Specific
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = pVal.ItemUID
                                    frmSourceMatrix = oMatrix
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                blnFormLoaded = True
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "17" And (pVal.ColUID = "V_3" Or pVal.ColUID = "V_4") And pVal.CharPressed = 9 Then
                                    ' oForm.Freeze(True)
                                    Dim dblUnitPrice, dblQuantity As Double
                                    oMatrix = oForm.Items.Item("17").Specific
                                    dblUnitPrice = oApplication.Utilities.getMatrixValues(oMatrix, "V_4", pVal.Row)
                                    dblQuantity = oApplication.Utilities.getMatrixValues(oMatrix, "V_3", pVal.Row)
                                    dblQuantity = dblUnitPrice * dblQuantity
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", pVal.Row, dblQuantity)
                                    ' oForm.Freeze(False)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "20" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet.DoQuery("select ""series"",""SeriesName"",""NextNumber""   from NNM1 where ""ObjectCode""='Z_OQUT'")
                                    oApplication.Utilities.setEdittextvalue(oForm, "4", oRecordSet.Fields.Item(2).Value)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "28"
                                        oForm.PaneLevel = 2

                                    Case "27"
                                        oForm.PaneLevel = 1

                                    Case "1"
                                        'If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        '    UpdateAttachment(oForm)
                                        'End If
                                    Case "18"
                                        AddRow(oForm)
                                    Case "19"
                                        RefereshDeleteRow(oForm)
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                    Case "32"
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            Dim oOBj As New clsEstmationSummary
                                            frm_SourceBoM = oForm
                                            oOBj.LoadForm(oApplication.Utilities.getEditTextvalue(oForm, "4"))
                                        End If
                                    Case "34"
                                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            Exit Sub
                                        End If
                                        'deleterow(oForm)
                                        'RefereshDeleteRow(oForm)
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                    Case "35"
                                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            Exit Sub
                                        End If
                                        LoadFiles(oForm)
                                    Case "33"
                                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            Exit Sub
                                        End If
                                        fillopen()
                                        If strSelectedFilepath <> "" Then
                                            oMatrix = oForm.Items.Item("100").Specific
                                            AddRow(oForm)
                                            Try
                                                oForm.Freeze(True)
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, strSelectedFilepath)
                                                Dim strDate As String
                                                Dim dtdate As Date
                                                dtdate = Now.Date
                                                strDate = Date.Today().ToString
                                                ''  strdate=
                                                Dim oColumn As SAPbouiCOM.Column
                                                oColumn = oMatrix.Columns.Item("V_1")
                                                ' oColumn.Editable = True
                                                oColumn.Editable = True
                                                oEditText = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
                                                oEditText.String = "t"
                                                oApplication.SBO_Application.SendKeys("{TAB}")
                                                oForm.Items.Item("24").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                oColumn.Editable = False
                                                'oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, dtdate)
                                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                                oForm.Freeze(False)
                                            Catch ex As Exception
                                                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                oForm.Freeze(False)

                                            End Try
                                        End If


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
                                        If pVal.ItemUID = "30" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            val1 = oDataTable.GetValue("CardName", 0)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "36", val1)
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            Catch ex As Exception
                                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                            End Try
                                            val1 = oDataTable.GetValue("SlpCode", 0)
                                            oCombobox = oForm.Items.Item("38").Specific
                                            Try
                                                oCombobox.Select(val1, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                            Catch ex As Exception

                                            End Try
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            End If
                                        End If
                                        If pVal.ItemUID = "17" And pVal.ColUID = "V_0" Then
                                            val1 = oDataTable.GetValue("ItemCode", 0)
                                            val = oDataTable.GetValue("ItemName", 0)
                                            oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                            Try
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", pVal.Row, val)
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", pVal.Row, val1)
                                            Catch ex As Exception
                                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                            End Try

                                            Dim dblUnitPrice, dblQuantity As Double
                                            dblUnitPrice = GetUnitPrice(val1)
                                            dblQuantity = 1
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", pVal.Row, dblUnitPrice)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", pVal.Row, dblQuantity)
                                            dblQuantity = dblUnitPrice * dblQuantity
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", pVal.Row, dblQuantity)

                                        End If
                                    End If
                                Catch ex As Exception
                                End Try

                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_BoM_Estimation
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        RefereshDeleteRow(oForm)
                    End If
                Case mnu_ADD
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        oForm.Items.Item("4").Enabled = False
                        oForm.Items.Item("6").Enabled = False
                    End If
                    If pVal.BeforeAction = False Then
                        AddMode(oForm)
                    End If
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        AddRow(oForm)
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
            If BusinessObjectInfo.BeforeAction = False And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If oForm.TypeEx = frm_BoM_Estimation Then
                    oForm.Items.Item("24").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oForm.Items.Item("4").Enabled = False
                    oForm.Items.Item("6").Enabled = False
                    oForm.Items.Item("20").Enabled = False
                    oCombobox = oForm.Items.Item("8").Specific
                    If oCombobox.Selected.Value <> "D" Then
                        oForm.Items.Item("1").Enabled = False
                    Else
                        oForm.Items.Item("1").Enabled = True
                    End If
                End If
            End If

            If BusinessObjectInfo.BeforeAction = False And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If oForm.TypeEx = frm_BoM_Estimation Then
                    oForm.Items.Item("24").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oForm.Items.Item("4").Enabled = False
                    oForm.Items.Item("6").Enabled = False
                    oForm.Items.Item("20").Enabled = True
                    AddMode(oForm)
                End If
            End If
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                Dim strdocnum, strQuery As String
                Dim stXML As String = BusinessObjectInfo.ObjectKey
                stXML = stXML.Replace("<?xml version=""1.0"" encoding=""UTF-16"" ?><EstimationParams><DocEntry>", "")
                stXML = stXML.Replace("</DocEntry></EstimationParams>", "")
                stXML = stXML.Replace("<?xml version=""1.0"" encoding=""UTF-16"" ?><EstimationParams><DocEntry>", "")
                stXML = stXML.Replace("</DocEntry></EstimationParams>", "")
                Dim otest As SAPbobsCOM.Recordset
                otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If stXML <> "" Then

                    otest.DoQuery("select * from ""@Z_OQUT""  where ""DocEntry""=" & stXML)
                    If otest.RecordCount > 0 Then
                        If otest.Fields.Item("U_Z_DocStatus").Value = "C" Then
                            Dim intTempID As String = GetTemplateID(oForm, "B")
                            If intTempID <> "0" Then
                                'UpdateApprovalRequired("@Z_OQUT", "DocEntry", otest.Fields.Item("DocEntry").Value, "Y", intTempID)
                                ' InitialMessage("Estimation approval", otest.Fields.Item("DocEntry").Value, "P", intTempID, "B")
                            Else
                                'UpdateApprovalRequired("@Z_OQUT", "DocEntry", otest.Fields.Item("DocEntry").Value, "N", intTempID)
                                strQuery = "Update ""@Z_OQUT"" set ""U_Z_AppStatus""='A' where ""DocEntry""='" & otest.Fields.Item("DocEntry").Value & "'"
                                otest.DoQuery(strQuery)
                            End If
                        End If

                    End If

                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub UpdateApprovalRequired(ByVal strTable As String, ByVal sColumn As String, ByVal StrCode As String, ByVal ReqValue As String, ByVal AppTempId As String)
        Try
            Dim strQuery As String
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Update """ & strTable & """ set ""U_Z_AppRequired""='" & ReqValue & "',""U_Z_AppReqDate""=getdate()"
            strQuery += "  where """ & sColumn & """='" & StrCode & "'"
            oRecordSet.DoQuery(strQuery)
        Catch ex As Exception
            MsgBox(oApplication.Company.GetLastErrorDescription)
        End Try
    End Sub
#End Region

#Region "Validations"

    Private Function Validation_Form(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Dim strQuery As String
        Try
            oCombobox = oForm.Items.Item("8").Specific
            If oApplication.Utilities.getEditTextvalue(oForm, "22") = "" Then
                oApplication.Utilities.Message("Estimation Description is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If oApplication.Utilities.getEditTextvalue(oForm, "30") = "" Then
                oApplication.Utilities.Message("Customer Code is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            oCombobox = oForm.Items.Item("8").Specific
            If oCombobox.Selected.Value = "C" Then
                Dim intTempID As String = GetTemplateID(oForm, "B")
                If intTempID <> "0" Then
                    If oApplication.SBO_Application.MessageBox("Generating this document requires the authorization of other users.Do You want to continue?", , "Continue", "Cancel") = 2 Then
                        Return False
                    End If


                End If
            End If

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#Region "Validations"
    Private Function Validation(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Dim strQuery As String
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCombobox = oForm.Items.Item("8").Specific
            If oCombobox.Selected.Value = "C" Then
                Dim intTempID As String = GetTemplateID(oForm, "B")
                If intTempID <> "0" Then
                    InitialMessage("estimation approval", oApplication.Utilities.getEditTextvalue(oForm, "4"), "P", intTempID, "B")
                Else
                    strQuery = "Update ""@Z_OQUT"" set ""U_Z_AppStatus""='A' where ""DocEntry""='" & oApplication.Utilities.getEditTextvalue(oForm, "4") & "'"
                    oRecordSet.DoQuery(strQuery)
                End If
            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Public Function GetTemplateID(ByVal aForm As SAPbouiCOM.Form, ByVal DocType As String) As String
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim strQuery As String = ""
            Dim Status As String = ""
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If blnIsHana Then
                strQuery = "Select * from ""@P_OAPPT"" T0 left join ""@P_APPT1"" T1 on T0.""DocEntry""=T1.""DocEntry"" where IfNull(T0.""U_Z_Active"",'N')='Y' and T0.""U_Z_DocType""='" & DocType.ToString() & "' and T1.""U_Z_OUser""='" & oApplication.Company.UserName & "' "
            Else
                strQuery = "Select * from ""@P_OAPPT"" T0 left join ""@P_APPT1"" T1 on T0.""DocEntry""=T1.""DocEntry"" where isnull(T0.""U_Z_Active"",'N')='Y' and T0.""U_Z_DocType""='" & DocType.ToString() & "' and T1.""U_Z_OUser""='" & oApplication.Company.UserName & "' "
            End If

            oRecordSet.DoQuery(strQuery)
            If oRecordSet.RecordCount > 0 Then
                Status = oRecordSet.Fields.Item("DocEntry").Value
            Else
                Status = "0"
            End If
            Return Status
        Catch ex As Exception
            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Public Sub InitialMessage(ByVal strReqType As String, ByVal strReqNo As String, ByVal strAppStatus As String _
        , ByVal strTemplateNo As String, ByVal enDocType As String)
        Try
            Dim strQuery As String
            Dim strMessageUser As String
            Dim oRecordSet, oTemp As SAPbobsCOM.Recordset
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim oMessageService As SAPbobsCOM.MessagesService
            Dim oMessage As SAPbobsCOM.Message
            Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns
            Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn
            Dim oLines As SAPbobsCOM.MessageDataLines
            Dim oLine As SAPbobsCOM.MessageDataLine
            Dim oRecipientCollection As SAPbobsCOM.RecipientCollection
            oCmpSrv = oApplication.Company.GetCompanyService()
            oMessageService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService)
            oMessage = oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If blnIsHana Then
                strQuery = "Select Top 1 ""U_Z_AUser"" From ""@P_APPT2"" Where ""DocEntry"" = '" + strTemplateNo + "'  and IFNull(""U_Z_AMan"",'')='Y' Order By ""LineId"" Asc "
            Else
                strQuery = "Select Top 1 U_Z_AUser From [@P_APPT2] Where DocEntry = '" + strTemplateNo + "'  and isnull(U_Z_AMan,'')='Y' Order By LineId Asc "
            End If

            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                strMessageUser = oRecordSet.Fields.Item(0).Value
                oMessage.Subject = strReqType
                Dim strMessage As String = ""
                Select Case enDocType
                    Case "B"
                        strQuery = "Select * from  ""@Z_OQUT"" where ""DocEntry""='" & strReqNo & "'"
                        oTemp.DoQuery(strQuery)
                        strMessage = " Requested by  :" & oApplication.Company.UserName & ": Documnet Number : " & strReqNo & " and Description :" & oTemp.Fields.Item("U_Z_Desc").Value & ""
                End Select
                Select Case enDocType
                    Case "B"
                        strQuery = "Update ""@Z_OQUT"" set ""U_Z_CurApprover""='" & strMessageUser & "',""U_Z_NxtApprover""='" & strMessageUser & "' where ""DocEntry""='" & strReqNo & "'"
                        oTemp.DoQuery(strQuery)
                End Select

                oMessage.Text = strReqType + " " + strMessage + " Needs Your Approval "
                oRecipientCollection = oMessage.RecipientCollection

                oRecipientCollection.Add()
                oRecipientCollection.Item(0).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
                oRecipientCollection.Item(0).UserCode = strMessageUser
                pMessageDataColumns = oMessage.MessageDataColumns

                pMessageDataColumn = pMessageDataColumns.Add()
                pMessageDataColumn.ColumnName = "Document Number"
                oLines = pMessageDataColumn.MessageDataLines()
                oLine = oLines.Add()
                oLine.Value = strReqNo
                oMessageService.SendMessage(oMessage)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#End Region


    Private Sub UpdateAttachment(ByVal aForm As SAPbouiCOM.Form)
        Try
            oMatrix = aForm.Items.Item("100").Specific
            For i As Integer = 1 To oMatrix.RowCount
                Dim oRec As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim strQry = "Select ""AttachPath"" From OADP"
                oRec.DoQuery(strQry)
                Dim SPath As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", i) ' oOfferGrid.DataTable.GetValue("U_Z_Attachment", i).ToString()
                If SPath = "" Then
                Else
                    Dim DPath As String = ""
                    If Not oRec.EoF Then
                        DPath = oRec.Fields.Item("AttachPath").Value.ToString()
                    End If
                    If Not Directory.Exists(DPath) Then
                        Directory.CreateDirectory(DPath)
                    End If
                    Dim file = New FileInfo(SPath)
                    Dim Filename As String = Path.GetFileName(SPath)
                    Dim SavePath As String = Path.Combine(DPath, Filename)
                    If System.IO.File.Exists(SavePath) Then
                    Else
                        file.CopyTo(Path.Combine(DPath, file.Name), True)
                    End If
                End If
            Next


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub LoadFiles(ByVal aform As SAPbouiCOM.Form)
        oMatrix = aform.Items.Item("100").Specific
        For intRow As Integer = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(intRow) Then
                Dim strFilename As String
                strFilename = oMatrix.Columns.Item("V_0").Cells.Item(intRow).Specific.value
                Dim x As System.Diagnostics.ProcessStartInfo
                x = New System.Diagnostics.ProcessStartInfo
                x.UseShellExecute = True
                x.FileName = strFilename
                If File.Exists(strFilename) Then
                    System.Diagnostics.Process.Start(x)
                Else
                    Dim oRec As SAPbobsCOM.Recordset
                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim strQry = "Select ""AttachPath"" From OADP"
                    oRec.DoQuery(strQry)
                    Dim SPath As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow) ' oOfferGrid.DataTable.GetValue("U_Z_Attachment", i).ToString()
                    If 1 = 2 Then
                    Else
                        Dim DPath As String = ""
                        If Not oRec.EoF Then
                            DPath = oRec.Fields.Item("AttachPath").Value.ToString()
                        End If
                        If Not Directory.Exists(DPath) Then
                            Directory.CreateDirectory(DPath)
                        End If
                        Dim file = New FileInfo(SPath)
                        Dim Filename As String = Path.GetFileName(SPath)
                        Dim SavePath As String = Path.Combine(DPath, Filename)
                        If System.IO.File.Exists(SavePath) Then
                            x.FileName = SavePath
                            System.Diagnostics.Process.Start(x)
                        Else

                        End If
                    End If
                End If

                x = Nothing
                Exit Sub
            End If
        Next
        oApplication.Utilities.Message("No file has been selected...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    End Sub
    Private Sub fillopen()
        Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
        mythr.SetApartmentState(Threading.ApartmentState.STA)
        mythr.Start()
        mythr.Join()

    End Sub

    Private Sub ShowFileDialog()
        Dim oDialogBox As New OpenFileDialog
        Dim strFileName, strMdbFilePath As String
        Dim oProcesses() As Process
        Try
            oProcesses = Process.GetProcessesByName("SAP Business One")
            If oProcesses.Length <> 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    Dim MyWindow As New clsListener.WindowWrapper(oProcesses(i).MainWindowHandle)
                    If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                        strMdbFilePath = oDialogBox.FileName
                        strSelectedFilepath = oDialogBox.FileName
                        strFileName = strSelectedFilepath
                        strSelectedFolderPath = strFileName
                    Else
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub
    'Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
    '    Try
    '        aForm.Freeze(True)
    '        oMatrix = aForm.Items.Item("31").Specific
    '        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_HR_EXFORM3")
    '        oMatrix.FlushToDataSource()
    '        For count = 1 To oDataSrc_Line.Size
    '            oDataSrc_Line.SetValue("LineId", count - 1, count)
    '        Next
    '        oMatrix.LoadFromDataSource()
    '        aForm.Freeze(False)
    '    Catch ex As Exception
    '        oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        aForm.Freeze(False)
    '    End Try
    'End Sub
#Region "Add Row/ Delete Row"
    'Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
    '    Try
    '        aForm.Freeze(True)

    '        Select Case aForm.PaneLevel
    '            Case "4"
    '                oMatrix = aForm.Items.Item("31").Specific
    '                oDataSrc_Line = aForm.DataSources.DBDataSources.Item("@Z_HR_EXFORM3")
    '                If oMatrix.RowCount <= 0 Then
    '                    oMatrix.AddRow()
    '                End If
    '                oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
    '                Try
    '                    If oEditText.Value <> "" Then
    '                        oMatrix.AddRow()
    '                        Select Case aForm.PaneLevel
    '                            Case "4"
    '                                oMatrix.ClearRowData(oMatrix.RowCount)
    '                        End Select
    '                    End If

    '                Catch ex As Exception
    '                    aForm.Freeze(False)
    '                    'oMatrix.AddRow()
    '                End Try
    '                oMatrix.FlushToDataSource()
    '                For count = 1 To oDataSrc_Line.Size
    '                    oDataSrc_Line.SetValue("LineId", count - 1, count)
    '                Next
    '                oMatrix.LoadFromDataSource()
    '                oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
    '                AssignLineNo(aForm)
    '        End Select


    '        aForm.Freeze(False)
    '    Catch ex As Exception
    '        oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        aForm.Freeze(False)

    '    End Try
    'End Sub
    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
        Select Case aForm.PaneLevel
            Case "2"
                oMatrix = aForm.Items.Item("100").Specific
                oDataSrc_Line = aForm.DataSources.DBDataSources.Item("@Z_QUT2")
            Case "1"
                oMatrix = aForm.Items.Item("17").Specific
                oDataSrc_Line = aForm.DataSources.DBDataSources.Item("@Z_QUT1")
        End Select

        '  oMatrix = aForm.Items.Item("16").Specific
        oMatrix.FlushToDataSource()
        For introw As Integer = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(introw) Then
                oMatrix.DeleteRow(introw)
                oDataSrc_Line.RemoveRecord(introw - 1)
                'oMatrix = frmSourceMatrix
                For count As Integer = 1 To oDataSrc_Line.Size
                    oDataSrc_Line.SetValue("LineId", count - 1, count)
                Next
                Select Case aForm.PaneLevel
                    Case "2"
                        oMatrix = aForm.Items.Item("100").Specific
                        oDataSrc_Line = aForm.DataSources.DBDataSources.Item("@Z_QUT2")
                        AssignLineNo(aForm)
                    Case "1"
                        oMatrix = aForm.Items.Item("17").Specific
                        oDataSrc_Line = aForm.DataSources.DBDataSources.Item("@Z_QUT1")
                        AssignLineNo(aForm)
                End Select
                oMatrix.LoadFromDataSource()
                If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
                Exit Sub
            End If
        Next

    End Sub
    'Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
    '    If Me.MatrixId = "31" Then
    '        oDataSrc_Line = aForm.DataSources.DBDataSources.Item("@Z_HR_EXFORM3")
    '    End If
    '    'oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
    '    If intSelectedMatrixrow <= 0 Then
    '        Exit Sub

    '    End If
    '    Me.RowtoDelete = intSelectedMatrixrow
    '    oDataSrc_Line.RemoveRecord(Me.RowtoDelete - 1)
    '    oMatrix = frmSourceMatrix
    '    oMatrix.FlushToDataSource()
    '    For count = 1 To oDataSrc_Line.Size - 1
    '        oDataSrc_Line.SetValue("LineId", count - 1, count)
    '    Next
    '    oMatrix.LoadFromDataSource()
    '    If oMatrix.RowCount > 0 Then
    '        oMatrix.DeleteRow(oMatrix.RowCount)
    '    End If
    'End Sub
#End Region

#Region "Function"

    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            If aForm.PaneLevel = 1 Then
                oMatrix = aForm.Items.Item("17").Specific
                oDBDataSourceLineZ_1 = oForm.DataSources.DBDataSources.Item("@Z_QUT1")
            Else
                oMatrix = aForm.Items.Item("100").Specific
                oDBDataSourceLineZ_1 = oForm.DataSources.DBDataSources.Item("@Z_QUT2")
            End If
            '   oMatrix = aForm.Items.Item("17").Specific

            If 1 = 1 Then ' Me.MatrixId = "17" Then
                Me.RowtoDelete = intSelectedMatrixrow
                oDBDataSourceLineZ_1.RemoveRecord(Me.RowtoDelete - 1)
                oMatrix.LoadFromDataSource()
                oMatrix.FlushToDataSource()
                For count = 1 To oDBDataSourceLineZ_1.Size
                    oDBDataSourceLineZ_1.SetValue("LineId", count - 1, count)
                Next
            End If
            oMatrix.LoadFromDataSource()
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            If aForm.PaneLevel = 1 Then
                oMatrix = aForm.Items.Item("17").Specific
                oDBDataSourceLineZ_1 = oForm.DataSources.DBDataSources.Item("@Z_QUT1")
            Else
                oMatrix = aForm.Items.Item("100").Specific
                oDBDataSourceLineZ_1 = oForm.DataSources.DBDataSources.Item("@Z_QUT2")
            End If

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
            Try
                oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Catch ex As Exception

            End Try

            AssignLineNo(aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            If aForm.PaneLevel = 1 Then
                oMatrix = aForm.Items.Item("17").Specific
                oDBDataSourceLineZ_1 = oForm.DataSources.DBDataSources.Item("@Z_QUT1")
            Else
                oMatrix = aForm.Items.Item("100").Specific
                oDBDataSourceLineZ_1 = oForm.DataSources.DBDataSources.Item("@Z_QUT2")
            End If

            oMatrix.FlushToDataSource()
            For count = 1 To oDBDataSourceLineZ_1.Size
                oDBDataSourceLineZ_1.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub











#End Region
End Class
