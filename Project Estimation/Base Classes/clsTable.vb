Public NotInheritable Class clsTable

#Region "Private Functions"
    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Tables in DB. This function shall be called by 
    '                     public functions to create a table
    '**************************************************************************************************************
    Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try

            oUserTablesMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            'Adding Table
            If Not oUserTablesMD.GetByKey(strTab) Then
                oUserTablesMD.TableName = strTab
                oUserTablesMD.TableDescription = strDesc
                oUserTablesMD.TableType = nType
                If oUserTablesMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
            oUserTablesMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddFields
    'Parameter          : SstrTab As String,strCol As String,
    '                     strDesc As String,nType As Integer,i,nEditSize,nSubType As Integer
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Fields in DB Tables. This function shall be called by 
    '                     public functions to create a Field
    '**************************************************************************************************************
    Private Sub AddFields(ByVal strTab As String, _
                            ByVal strCol As String, _
                                ByVal strDesc As String, _
                                    ByVal nType As SAPbobsCOM.BoFieldTypes, _
                                        Optional ByVal i As Integer = 0, _
                                            Optional ByVal nEditSize As Integer = 10, _
                                                Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, _
                                                    Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO)
        Dim oUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            If Not (strTab = "OITM" Or strTab = "OITM" Or strTab = "OADM" Or strTab = "QUT1" Or strTab = "OUSR" Or strTab = "OITW" Or strTab = "RDR1" Or strTab = "DSC1" Or strTab = "OPRJ") Then
                strTab = "@" + strTab
            End If

            If Not IsColumnExists(strTab, strCol) Then
                oUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

                oUserFieldMD.Description = strDesc
                oUserFieldMD.Name = strCol
                oUserFieldMD.Type = nType
                oUserFieldMD.SubType = nSubType
                oUserFieldMD.TableName = strTab
                oUserFieldMD.EditSize = nEditSize
                oUserFieldMD.Mandatory = Mandatory
                If oUserFieldMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)

            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddFields
    'Parameter          : SstrTab As String,strCol As String,
    '                     strDesc As String,nType As Integer,i,nEditSize,nSubType As Integer
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Fields in DB Tables. This function shall be called by 
    '                     public functions to create a Field
    '**************************************************************************************************************
    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If


            If (Not IsColumnExists(TableName, ColumnName)) Then
                objUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                For intLoop = 0 To strValue.GetLength(0) - 1
                    objUserFieldMD.ValidValues.Value = strValue(intLoop)
                    objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                    objUserFieldMD.ValidValues.Add()
                Next
                If (objUserFieldMD.Add() <> 0) Then
                    MsgBox(oApplication.Company.GetLastErrorDescription)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            Else
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            objUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()

        End Try


    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : IsColumnExists
    'Parameter          : ByVal Table As String, ByVal Column As String
    'Return Value       : Boolean
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Function to check if the Column already exists in Table
    '**************************************************************************************************************
    Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            strSQL = "SELECT COUNT(*) FROM CUFD WHERE ""TableID"" = '" & Table & "' AND ""AliasID"" = '" & Column & "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strSQL)

            If oRecordSet.Fields.Item(0).Value = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()
        End Try
    End Function

    Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        Try
            '// The meta-data object must be initialized with a
            '// regular UserKeys object
            oUserKeysMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

            If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                '// Set the table name and the key name
                oUserKeysMD.TableName = strTab
                oUserKeysMD.KeyName = strKey

                '// Set the column's alias
                oUserKeysMD.Elements.ColumnAlias = strColumn
                oUserKeysMD.Elements.Add()
                oUserKeysMD.Elements.ColumnAlias = "RentFac"

                '// Determine whether the key is unique or not
                oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                '// Add the key
                If oUserKeysMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
            oUserKeysMD = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Sub

    '********************************************************************
    'Type		            :   Function    
    'Name               	:	AddUDO
    'Parameter          	:   
    'Return Value       	:	Boolean
    'Author             	:	
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To Add a UDO for Transaction Tables
    '********************************************************************
    Private Sub AddUDO1(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                        Optional ByVal strChildTbl As String = "", Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document)

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) = 0 Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES

                If sFind1 <> "" And sFind2 <> "" Then
                    oUserObjectMD.FindColumns.ColumnAlias = sFind1
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(1)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind2
                    oUserObjectMD.FindColumns.Add()
                End If

                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.LogTableName = "A" & strUDO
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ExtensionName = ""

                If strChildTbl <> "" Then
                    oUserObjectMD.ChildTables.TableName = strChildTbl
                End If

                oUserObjectMD.Code = strUDO
                oUserObjectMD.Name = strDesc
                oUserObjectMD.ObjectType = nObjectType
                oUserObjectMD.TableName = strTable

                If oUserObjectMD.Add() <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try

    End Sub
    Private Sub AddUDO(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                        Optional ByVal strChildTbl As String = "", Optional ByVal strChildTb2 As String = "", Optional ByVal strChildTb3 As String = "", Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document, Optional ByVal blnCanArchive As Boolean = False, Optional ByVal strLogName As String = "")

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) = 0 Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES

                If sFind1 <> "" And sFind2 <> "" Then
                    oUserObjectMD.FindColumns.ColumnAlias = sFind1
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(1)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind2
                    oUserObjectMD.FindColumns.Add()
                End If

                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.LogTableName = "A" & strUDO
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ExtensionName = ""

                If strChildTbl <> "" Then
                    oUserObjectMD.ChildTables.TableName = strChildTbl
                End If
                If strChildTb2 <> "" Then
                    If strChildTbl <> "" Then
                        oUserObjectMD.ChildTables.Add()
                    End If
                    oUserObjectMD.ChildTables.TableName = strChildTb2
                End If
                If strChildTb3 <> "" Then
                    If strChildTb2 <> "" Then
                        If strChildTbl <> "" Then
                            oUserObjectMD.ChildTables.Add()
                        End If
                        oUserObjectMD.ChildTables.TableName = strChildTb3
                    End If
                End If

                '  oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.Code = strUDO
                oUserObjectMD.Name = strDesc
                oUserObjectMD.ObjectType = nObjectType
                oUserObjectMD.TableName = strTable

                If oUserObjectMD.Add() <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try

    End Sub
#End Region

#Region "Public Functions"
    '*************************************************************************************************************
    'Type               : Public Function
    'Name               : CreateTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Creating Tables by calling the AddTables & AddFields Functions
    '**************************************************************************************************************
    Public Sub CreateTables()
        Try

            oApplication.SBO_Application.StatusBar.SetText("Initializing Database...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oApplication.Company.StartTransaction()
            '---- User Defined Fields

            'Project Estimation Reference Tables
            AddTables("Z_PRES", "Project Estimation Reference", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_PRES", "Z_Type", "DocType", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)

            'SubProject
            AddTables("Z_OSUP", "Sub_Project", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OSUP", "Z_Code", "Sub Project Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OSUP", "Z_Name", "SubProject Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_OSUP", "Z_GLAcc", "G/L Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OSUP", "Z_GLName", "G/L Account Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            'Project Phase 
            AddTables("Z_OPRPH", "Project Phase ", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("Z_PRPH1", "Project Phase Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("Z_PRPH2", "Project Phase BOM Child", SAPbobsCOM.BoUTBTableType.bott_NoObject)


            AddFields("Z_OPRPH", "Z_Code", "Project Phase Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OPRPH", "Z_Name", "Project Phase Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_OPRPH", "Z_TotalCost", "Total Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_OPRPH", "Z_Margin", "Margin %", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            addField("@Z_OPRPH", "Z_Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")
            AddFields("Z_OPRPH", "Z_Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_OPRPH", "Z_UnitPrice", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)


            AddFields("Z_PRPH1", "Z_ItemCode", "BoM Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PRPH1", "Z_ItemName", "BoM Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PRPH1", "Z_BaseQty", "Base Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_PRPH1", "Z_Cost", "Base Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_PRPH1", "Z_Margin", "Margin %", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("Z_PRPH1", "Z_TotalCost", "Total Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_PRPH1", "Z_BoMRef", "BOM Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

          

            addField("@Z_PRPH2", "Z_Type", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "4,290", "Item,Resource", "4")
            AddFields("Z_PRPH2", "Z_ItemCode", "BoM Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PRPH2", "Z_ItemName", "BoM Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PRPH2", "Z_BaseQty", "Base Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_PRPH2", "Z_Cost", "Base Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_PRPH2", "Z_WhsCode", "Warehouse Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PRPH2", "Z_UoM", "UoM Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_PRPH2", "Z_PlnList", "Price List", SAPbobsCOM.BoFieldTypes.db_Alpha, , 2)
            AddFields("Z_PRPH2", "Z_TotalCost", "Total Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_PRPH2", "Z_Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_PRPH2", "Z_RItemCode", "Project Phase Item", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PRPH2", "Z_RItemName", "Project Phase ItemName", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PRPH2", "Z_PHRef", "Project Phase BOM Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

            'Approval Template


            AddTables("P_OAPPT", "Approval Template", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("P_APPT2", "Approval Authorizer", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("P_APPT1", "Approval Orginator", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            AddFields("P_OAPPT", "Z_Code", "Approval Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("P_OAPPT", "Z_Name", "Approval Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("P_OAPPT", "Z_DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("P_OAPPT", "Z_DocDesc", "Document Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("P_OAPPT", "Z_Active", "Active Template", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)

            AddFields("P_APPT1", "Z_OUser", "Orginator Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("P_APPT1", "Z_OName", "Orginator Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            AddFields("P_APPT2", "Z_AUser", "Authorizer Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("P_APPT2", "Z_AName", "Authorizer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("P_APPT2", "Z_AMan", "Mandatory", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("P_APPT2", "Z_AFinal", "Final Stage", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)

            'Approval History
            AddTables("P_APHIS", "Approval History", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("P_APHIS", "Z_DocEntry", "Document Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("P_APHIS", "Z_DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("P_APHIS", "Z_EmpId", "Employee Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("P_APHIS", "Z_EmpName", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@P_APHIS", "Z_AppStatus", "Approved Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,A,R", "Pending,Approved,Rejected", "P")
            AddFields("P_APHIS", "Z_Remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("P_APHIS", "Z_ApproveBy", "Approved By", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("P_APHIS", "Z_Approvedt", "Approver Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("P_APHIS", "Z_ADocEntry", "Template DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("P_APHIS", "Z_ALineId", "Template LineId", SAPbobsCOM.BoFieldTypes.db_Numeric)


            'Estimation Related 
            ''  AddFields("OITT", "Z_BoMRef", "BoM Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            'AddTables("Z_OITT", "Summary Details", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            'AddFields("Z_OITT", "Z_ItemCode", "BoM ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            'AddFields("Z_OITT", "Z_Type", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            'AddFields("Z_OITT", "Z_Cost", "Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            'AddFields("Z_OITT", "Z_Markup", "Markup %", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            'AddFields("Z_OITT", "Z_Price", "Sales Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            'AddFields("Z_OITT", "Z_AvgMarkUp", "Average markup", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)


            'AddTables("Z_OITT1", "Estimation Summary Details", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            'AddFields("Z_OITT1", "Z_ItemCode", "DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            'AddFields("Z_OITT1", "Z_Type", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            'AddFields("Z_OITT1", "Z_Cost", "Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            'AddFields("Z_OITT1", "Z_Markup", "Markup %", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            'AddFields("Z_OITT1", "Z_Price", "Sales Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            'AddFields("Z_OITT1", "Z_AvgMarkUp", "Average markup", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)

            AddTables("Z_OQUT", "Estimation Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("Z_QUT1", "Estimation Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_OQUT", "Z_PrjCode", "Project Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OQUT", "Z_PrjName", "Project Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_OQUT", "Z_Desc", "Project Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 230)
            AddFields("Z_OQUT", "Z_CardCode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OQUT", "Z_CardName", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_OQUT", "Z_SlpCode", "Sales Person Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_OQUT", "Z_SupPrjCode", "Sub Project Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OQUT", "Z_SupPrjName", "Sub Project Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_OQUT", "Z_FreeText", "Free Text", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_OQUT", "Z_Remarks", "ReMarks", SAPbobsCOM.BoFieldTypes.db_Memo)
            addField("@Z_OQUT", "Z_DocStatus", "Document Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, SAPbobsCOM.BoFldSubTypes.st_Address, "P,R,A,Re,C", "Planned,Released,Approved,Rejected,Canceled", "P")
            addField("@Z_OQUT", "Z_AppStatus", "Approved Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,A,R", "Pending,Approved,Rejected", "P")
            AddFields("Z_OQUT", "Z_CurApprover", "Current Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_OQUT", "Z_NxtApprover", "Next Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_OQUT", "Z_AppReqDate", "Required Date", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("@Z_OQUT", "Z_AppRequired", "Approval Required", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_OQUT", "Z_TotalCost", "Total Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_OQUT", "Z_GLAcc", "G/L Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddFields("Z_QUT1", "Z_ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_QUT1", "Z_ItemDesc", "Item Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_QUT1", "Z_Details", "Details", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_QUT1", "Z_Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_QUT1", "Z_Price", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_QUT1", "Z_Margin", "Margin %", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("Z_QUT1", "Z_Total", "Total Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            'AddFields("QUT1", "Z_Spec", "Specification", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("QUT1", "Z_EstDocNum", "Estimation Base Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("QUT1", "Z_EstLineId", "Estimation Base Line", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)


            AddTables("Z_QUT2", "Estimation Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_QUT2", "FileName", "File Name", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_QUT2", "AttDate", "Attachment Date", SAPbobsCOM.BoFieldTypes.db_Date)
            ' AddFields("Z_QUT2", "AttName", "File Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            AddTables("Z_QUT3", "Estimation Free Text", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_QUT3", "Z_Text1", "Free Text1", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_QUT3", "Z_Text2", "Free Text2", SAPbobsCOM.BoFieldTypes.db_Memo)

            AddFields("OPRJ", "Z_CardCode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("OPRJ", "Z_CardName", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            '---- User Defined Object
            CreateUDO()

            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            oApplication.SBO_Application.StatusBar.SetText("Database creation completed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Catch ex As Exception
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            Throw ex
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub
    Public Sub CreateUDO()
        Try
            AddUDO("P_OSUP", "SubProject_Master", "Z_OSUP", "DocEntry", "U_Z_Code", , , , SAPbobsCOM.BoUDOObjType.boud_Document, False)
            AddUDO("P_OPRPH", "Project_Phase_Master", "Z_OPRPH", "DocEntry", "U_Z_Code", "Z_PRPH1", , , SAPbobsCOM.BoUDOObjType.boud_Document, True, "AZ_OPRPH")
            AddUDO("P_APHIS", "Approval History", "P_APHIS", "DocEntry", "U_Z_DocEntry", , , , SAPbobsCOM.BoUDOObjType.boud_Document, True, "AP_APHIS")
            AddUDO("P_OAPPT", "Approval Template", "P_OAPPT", "DocEntry", "U_Z_Code", "P_APPT1", "P_APPT2", , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("P_OQUT", "Project_Estimation", "Z_OQUT", "DocEntry", "U_Z_PrjCode", "Z_QUT1", "Z_QUT2", "Z_QUT3", SAPbobsCOM.BoUDOObjType.boud_Document)

            'Update UDO

            UpdateUDO_1("P_OQUT", "Project_Estimation", "Z_OQUT", "DocEntry", "U_Z_PrjCode", True, "Z_QUT1,Z_QUT2,Z_QUT3", SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

    Private Sub UpdateUDO_1(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                        Optional ByVal blnMultiChild As Boolean = False, _
                                        Optional ByVal strChildTbl As String = "", _
                                        Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document)

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Dim blnUpdate As Boolean = False
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) Then

                If oUserObjectMD.Name <> strDesc Then
                    oUserObjectMD.Name = strDesc
                    blnUpdate = True
                End If

                If Not blnMultiChild Then
                    If strChildTbl <> "" Then
                        oUserObjectMD.ChildTables.TableName = strChildTbl
                    End If
                Else
                    Dim strChild As String()
                    strChild = strChildTbl.Split(",")

                    For Each strTabl As String In strChild
                        Dim blnTableExists As Boolean = False
                        For index As Integer = 0 To oUserObjectMD.ChildTables.Count - 1
                            oUserObjectMD.ChildTables.SetCurrentLine(index)
                            If oUserObjectMD.ChildTables.TableName = strTabl Then
                                blnTableExists = True
                            End If
                        Next
                        If Not blnTableExists Then
                            blnUpdate = True
                            oUserObjectMD.ChildTables.Add()
                            oUserObjectMD.ChildTables.SetCurrentLine(oUserObjectMD.ChildTables.Count - 1)
                            oUserObjectMD.ChildTables.TableName = strTabl
                        End If
                    Next
                End If

                If blnUpdate Then
                    If oUserObjectMD.Update() <> 0 Then
                        Throw New Exception(oApplication.Company.GetLastErrorDescription)
                    End If
                End If

            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try

    End Sub



End Class
