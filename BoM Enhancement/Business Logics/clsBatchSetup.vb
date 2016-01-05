Public Class clsBatchSetup
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Function AddControls(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            aForm.Freeze(True)
            oApplication.Utilities.AddControls(aForm, "BtnAuto", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", 0, 0, , "Create Batches", 150)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
        End Try
    End Function
#Region "Assign Serianumbers"
    Private Sub AssignBatchNumber(ByVal aForm As SAPbouiCOM.Form)
        aForm.Freeze(True)
        Try
            Dim oRowsMatrix, oSerialMatrix As SAPbouiCOM.Matrix
            Dim dblSelectedqty, MatQuantity, Quantity, diffQuantity As Double
            Dim strItemCode, strwhs, strQry, strBatchNumber, strMaxBatch As String
            Dim BatchNumber As Integer
            Dim strBatchPrefix As String = ""
            Dim strmaxCode As String = ""
            Dim oSerialRec, oTemp1, oTemp, oTemp2 As SAPbobsCOM.Recordset
            oSerialRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRowsMatrix = aForm.Items.Item("35").Specific
            oSerialMatrix = aForm.Items.Item("3").Specific
            For intRow As Integer = 1 To oRowsMatrix.RowCount
                oRowsMatrix = aForm.Items.Item("35").Specific
                oRowsMatrix.Columns.Item("0").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                strItemCode = oRowsMatrix.Columns.Item("5").Cells.Item(intRow).Specific.value
                dblSelectedqty = oRowsMatrix.Columns.Item("39").Cells.Item(intRow).Specific.value
                strwhs = oRowsMatrix.Columns.Item("40").Cells.Item(intRow).Specific.value
                If dblSelectedqty > 0 Then
                    strQry = "SELECT U_Z_BATCHST FROM OITM where ItemCode='" & strItemCode & "'"
                    oTemp.DoQuery(strQry)
                    If oTemp.RecordCount > 0 Then
                        strBatchPrefix = oTemp.Fields.Item("U_Z_BATCHST").Value
                    Else
                        strBatchPrefix = "B"
                    End If
                    If strBatchPrefix = "" Then
                        strBatchPrefix = "B"
                    End If
                    strQry = "SELECT isnull(MAX(CAST(sysNumber AS numeric)),0) FROM [OBTN] where ItemCode='" & strItemCode & "'"
                    oTemp1.DoQuery(strQry)
                    If oTemp1.RecordCount > 0 Then
                        strQry = "SELECT DistNumber FROM [OBTN] where sysNumber='" & oTemp1.Fields.Item(0).Value & "' and ItemCode='" & strItemCode & "'"
                        oTemp2.DoQuery(strQry)
                        If oTemp2.RecordCount > 0 Then
                            strBatchNumber = oTemp2.Fields.Item("DistNumber").Value
                            BatchNumber = strBatchNumber.Replace(strBatchPrefix, 0)
                            strmaxCode = Format(BatchNumber + 1, "000000")
                        Else
                            strmaxCode = Format(1, "000000")
                        End If
                    End If
                End If
                strMaxBatch = strBatchPrefix & "" & strmaxCode
                Dim strBatch As String = strMaxBatch
                Dim intRow1 As Integer = 1
                If oSerialMatrix.RowCount > 1 Then
                    intRow1 = oSerialMatrix.RowCount - 1
                End If
                For intloop1 As Integer = intRow1 To intRow1
                    oApplication.Utilities.SetMatrixValues(oSerialMatrix, "2", intloop1, strMaxBatch)
                    oApplication.Utilities.SetMatrixValues(oSerialMatrix, "5", intloop1, dblSelectedqty)
                Next
                If aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    aForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                End If
            Next
            If aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                aForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            End If
            aForm.Freeze(False)
            'If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
            '    aForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'End If

        Catch ex As Exception
            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#End Region


#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BatchSetup Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                AddControls(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "BtnAuto" Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to Create the batches Automatically?", , "Yes", "No") = 2 Then
                                        Exit Sub
                                    End If
                                    AssignBatchNumber(oForm)
                                    oApplication.Utilities.Message("Operation Completed Successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
                Case "5946"
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        ' AssignBatchNumber(oForm)
                    End If
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
