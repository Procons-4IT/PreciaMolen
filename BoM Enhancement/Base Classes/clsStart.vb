Imports System.IO
Public Class clsStart

    Shared Sub Main()
        Dim oRead As System.IO.StreamReader
        Dim LineIn, strUsr, strPwd As String
        Dim i As Integer
        Dim strQuery As String = String.Empty
        ' Dim objValidate As New KeyValidator.Validator

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
            oApplication.Utilities.CreateTables()
            Dim oMenuItem As SAPbouiCOM.MenuItem
            oApplication.Utilities.AddRemoveMenus("Menu.xml")
            '  oMenuItem = oApplication.SBO_Application.Menus.Item("DABT_411")

            Dim strPath As String = System.Windows.Forms.Application.StartupPath & "\Script.txt"
            strQuery = File.ReadAllText(strPath)
            Dim oRec_ExeSP As SAPbobsCOM.Recordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec_ExeSP.DoQuery(strQuery)
            ' oMenuItem.Image = Application.StartupPath & "\Rental.jpg"
            oApplication.Utilities.Message("PreciaMolen Integration  Addon Connected successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oApplication.Utilities.NotifyAlert()
            System.Windows.Forms.Application.Run()

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

End Class
