Public Class clsStart
    
    Shared Sub Main()
        'Dim i As Integer
        'Dim oMenuItem As SAPbouiCOM.MenuItem
        Try
            Try
                oApplication = New clsListener
                oApplication.Utilities.Connect()
                oApplication.SetFilter()

                With oApplication.Company.GetCompanyService
                    CompanyDecimalSeprator = .GetAdminInfo.DecimalSeparator
                    CompanyThousandSeprator = .GetAdminInfo.ThousandsSeparator
                End With
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Exit Sub
            End Try
            If oApplication.Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                blnIsHana = True
            Else
                blnIsHana = False
            End If
            oApplication.Utilities.CreateTables()
            oApplication.Utilities.AddRemoveMenus("Menu.xml")
            oApplication.Utilities.createPayrollMainAuthorization()
            oApplication.Utilities.AuthorizationCreation()
            oApplication.Utilities.Message("Project Estimation Addon Connected Successfully..", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oApplication.Utilities.NotifyAlert()
            System.Windows.Forms.Application.Run()

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

End Class
