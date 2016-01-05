Public Class clsInvoice
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private ocolumn As SAPbouiCOM.Column
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
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "Z_OPHS"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "Z_OSEQ"
            oCFLCreationParams.UniqueID = "CFL2"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()
            If oForm.TypeEx = frm_BoM Then
                oMatrix = oForm.Items.Item("3").Specific
            Else
                oMatrix = oForm.Items.Item("37").Specific
            End If

            ocolumn = oMatrix.Columns.Item("U_Z_Phase")
            ocolumn.ChooseFromListUID = "CFL1"
            ocolumn.ChooseFromListAlias = "U_Z_Code"
            ocolumn = oMatrix.Columns.Item("U_Z_Seq")
            ocolumn.ChooseFromListUID = "CFL2"
            ocolumn.ChooseFromListAlias = "U_Z_Code"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim oRecSet As SAPbobsCOM.Recordset
        Dim strFields, strValues, strQuery, strItem, strPhase, strSeq, strDocNum As String
        Dim dblQty, dblPlannedQty As Double
        Try
            strDocNum = oApplication.Utilities.getEdittextvalue(aForm, "18")
            oMatrix = aForm.Items.Item("37").Specific
            aForm.Freeze(True)
            oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSet.DoQuery("Delete  from [Z_PreciaMolen_Temp] where DocNum=" & oApplication.Utilities.getDocumentQuantity(strDocNum))

            dblPlannedQty = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "12"))
            For intRow As Integer = 1 To oMatrix.RowCount
                strItem = oApplication.Utilities.getMatrixValues(oMatrix, "4", intRow)
                If strItem <> "" Then
                    oRecSet.DoQuery("Select * from OITM where InvntItem='Y' and ItemCode='" & strItem & "'")
                    If oRecSet.RecordCount > 0 Then
                        If oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Phase", intRow) = "" Then
                            oApplication.Utilities.Message("Phase detail missing.. Line No : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oMatrix.Columns.Item("U_Z_Phase").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            aForm.Freeze(False)
                            Return False
                        Else
                            strPhase = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Phase", intRow)
                        End If

                        If oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Seq", intRow) = "" Then
                            oApplication.Utilities.Message("Sequence detail missing.. Line No : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oMatrix.Columns.Item("U_Z_Phase").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            aForm.Freeze(False)
                            Return False
                        Else
                            strSeq = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Seq", intRow)
                        End If
                        If oApplication.Utilities.getMatrixValues(oMatrix, "14", intRow) <> "" Then
                            dblQty = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "14", intRow))
                        Else
                            dblQty = 0
                        End If
                        strFields = "DocNum,ItemCode,U_Phase,U_Seq,ProducedQty"
                        strValues = "'" & strDocNum & "','" & strItem & "','" & strPhase & "','" & strSeq & "','" & dblQty & "'"

                        strQuery = "Insert into [Z_PreciaMolen_Temp] (" & strFields & ") Values (" & strValues & ")"
                        oRecSet.DoQuery(strQuery)
                    End If
                End If
            Next
            Dim strMandatoryPhase As String = ""

            strQuery = "Select U_Z_Code from [@Z_OPHS] where isnull(U_Z_Mandatory,'N') ='Y'"
            oRecSet.DoQuery(strQuery)
            For intLoop As Integer = 0 To oRecSet.RecordCount - 1
                strMandatoryPhase = strMandatoryPhase & "," & oRecSet.Fields.Item(0).Value
                oRecSet.MoveNext()
            Next
            strQuery = "Select U_Z_Code from [@Z_OPHS] where isnull(U_Z_Mandatory,'N') ='Y'"
            strQuery = "Select * from [Z_PreciaMolen_Temp] where DocNum=" & strDocNum & " and  U_Phase in (" & strQuery & ")"
            oRecSet.DoQuery(strQuery)
            If oRecSet.RecordCount <= 0 Then
                oApplication.Utilities.Message("Mandatory phase :" & strMandatoryPhase & " :  is missing..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If

            strQuery = "Select U_Phase,U_Seq,Count(*) from [Z_PreciaMolen_Temp] where DocNum=" & strDocNum & "  group by U_Phase,U_Seq having Count(*)>1"
            oRecSet.DoQuery(strQuery)
            If oRecSet.RecordCount > 0 Then
                oApplication.Utilities.Message(" phase :" & oRecSet.Fields.Item(0).Value & " and  Sequence: " & oRecSet.Fields.Item(1).Value & " is already defined.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If

            Dim oRecSet1 As SAPbobsCOM.Recordset
            oRecSet1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select U_Phase ,Count(*) from [Z_PreciaMolen_Temp]  where DocNum=" & strDocNum & "  group by U_Phase "
            oRecSet.DoQuery(strQuery)
            For intRow As Integer = 0 To oRecSet.RecordCount - 1
                oRecSet1.DoQuery("Select * from [Z_PreciaMolen_Temp] where  DocNum=" & strDocNum & "  and U_Phase='" & oRecSet.Fields.Item(0).Value & "' order by convert(numeric,U_Seq)")
                For intLoop As Integer = 0 To oRecSet1.RecordCount - 1
                    If intLoop = 0 Then
                        If oRecSet1.Fields.Item("U_Seq").Value <> intLoop + 1 Then
                            oApplication.Utilities.Message("Phase : " & oRecSet.Fields.Item(0).Value & " should be start with Sequence 1", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            aForm.Freeze(False)
                            Return False
                        End If
                    Else
                        If oRecSet1.Fields.Item("U_Seq").Value <> intLoop + 1 Then
                            oApplication.Utilities.Message("Phase : " & oRecSet.Fields.Item(0).Value & " Sequence " & intLoop + 1 & " is missing", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            aForm.Freeze(False)
                            Return False
                        End If
                    End If
                    oRecSet1.MoveNext()
                Next
                oRecSet.MoveNext()
            Next

            strQuery = "Select isnull(Sum(isnull(ProducedQty,0)),0) from [Z_PreciaMolen_Temp]  where DocNum=" & strDocNum
            oRecSet.DoQuery(strQuery)
            dblQty = oRecSet.Fields.Item(0).Value
            If dblPlannedQty < dblQty Then
                oApplication.Utilities.Message("Material Total Planned quantity :" & dblQty & " :  should be equal to Produced Quantity :" & dblPlannedQty, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If
            aForm.Freeze(False)
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
            Return False
        End Try
    End Function
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BoM Or frm_ProductionOrder Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                AddChooseFromList(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val, val2 As String
                                Dim intChoice As Integer
                                Dim codebar As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "3" And pVal.FormTypeEx = frm_BoM And (pVal.ColUID = "U_Z_Phase" Or pVal.ColUID = "U_Z_Seq") Then
                                            val1 = oDataTable.GetValue("U_Z_Code", 0)
                                            Try
                                                oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                                oApplication.Utilities.SetMatrixValues(oMatrix, pVal.ColUID, pVal.Row, val1)
                                            Catch ex As Exception
                                            End Try
                                        End If

                                        If pVal.ItemUID = "37" And pVal.FormTypeEx = frm_ProductionOrder And (pVal.ColUID = "U_Z_Phase" Or pVal.ColUID = "U_Z_Seq") Then
                                            val1 = oDataTable.GetValue("U_Z_Code", 0)
                                            Try
                                                oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                                oApplication.Utilities.SetMatrixValues(oMatrix, pVal.ColUID, pVal.Row, val1)
                                            Catch ex As Exception
                                            End Try
                                        End If
                                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception

                                    oForm.Freeze(False)
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

    Private Function AddtoUDT(ByVal aform As SAPbouiCOM.Form, ByVal ItemCode As String, ByVal ItemName As String) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim otemp, otemp1 As SAPbobsCOM.Recordset
        Dim strqry, strCode, strqry1, strProCode, ProName, strGLAcc As String
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oUserTable = oApplication.Company.UserTables.Item("Z_OCBS")

        If ItemCode <> "" Then
            strCode = ItemCode
            oUserTable.GetByKey(strCode)
            oUserTable.Code = strCode
            oUserTable.Name = strCode
            oUserTable.Update()
        Else
            strCode = oApplication.Utilities.getMaxCode("@Z_OCBS", "Code")
            oUserTable.Code = strCode
            oUserTable.Name = strCode
            oUserTable.Add()
        End If
        oApplication.Utilities.SetEditText(aform, "edCR", strCode)

        Return True
    End Function

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "CR1"
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    Dim ItemCode, ItemName As String
                    If pVal.BeforeAction = False Then
                        If oForm.TypeEx = frm_BoM Then
                            If oApplication.Utilities.getEdittextvalue(oForm, "4") <> "" Then
                                ItemCode = oApplication.Utilities.getEdittextvalue(oForm, "4")
                                ItemName = oApplication.Utilities.getEdittextvalue(oForm, "4")
                                If 1 = 1 = True Then
                                    ItemCode = oApplication.Utilities.getEdittextvalue(oForm, "4")
                                    ItemName = "BOM" 'oApplication.Utilities.getEdittextvalue(oForm, "4")
                                    Dim objct As New clsQCMaster
                                    objct.LoadForm(ItemCode, ItemName, "SalesOrder")
                                End If
                            Else
                                oApplication.Utilities.Message("Select the Customer... ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        ElseIf oForm.TypeEx = frm_ProductionOrder Then
                            If oApplication.Utilities.getEdittextvalue(oForm, "6") <> "" Then
                                ItemCode = oApplication.Utilities.getEdittextvalue(oForm, "6")
                                ItemName = oApplication.Utilities.getEdittextvalue(oForm, "6")
                                If 1 = 1 = True Then
                                    ItemCode = oApplication.Utilities.getEdittextvalue(oForm, "18")
                                    ItemName = "Production" ' oApplication.Utilities.getEdittextvalue(oForm, "18")
                                    Dim objct As New clsQCMaster
                                    objct.LoadForm(ItemCode, ItemName, "Production")
                                End If
                            Else
                                oApplication.Utilities.Message("Select the Customer... ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Else

                            ItemCode = oApplication.Utilities.getEdittextvalue(oForm, "edCR")
                            If ItemCode = "" Then
                                oApplication.Utilities.Message("Crew Briefing Sheet details are not available", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Sub
                            Else
                                Dim objct As New clsQCMaster
                                objct.LoadForm(ItemCode, "", "DelInvoice")
                            End If
                        End If


                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        'If eventInfo.FormUID = "RightClk" Then
        If oForm.TypeEx = frm_BoM Or frm_ProductionOrder Then
            If (eventInfo.BeforeAction = True) Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                        oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = "CR1"
                        oCreationPackage.String = "View Bill of Material"
                        oCreationPackage.Enabled = True
                        oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                        oMenus = oMenuItem.SubMenus
                        oMenus.AddEx(oCreationPackage)


                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Else
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        oApplication.SBO_Application.Menus.RemoveEx("CR1")
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End If



    End Sub

    Private Sub UpdateTable(ByVal aDocNum As String)
        Dim oRec As SAPbobsCOM.Recordset
        Dim strSQL As String
        strSQL = "insert into Z_PreciaMolen select T0.DocNum,isnull(T1.U_Z_Phase,''), isnull(T1.U_Z_Seq,''),T0.ItemCode,T0.PlannedQty,T2.ItemName,t0.Warehouse,T1.ItemCode,T4.ItemName,T1.PlannedQty,T1.PlannedQty-T1.PlannedQty,NULL ,'' batchno,'0' from OWOR T0 inner Join WOR1 T1 on T1.DocEntry=T0.DocEntry  inner Join OITM T2 on T2.ItemCode=T0.ItemCode  inner Join OITM T4 on T4.ItemCode=T1.ItemCode and T4.InvntItem='Y'  where T0.DocNum =" & aDocNum
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery(strSQL)
    End Sub
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                Dim oPro As SAPbobsCOM.ProductionOrders
                oPro = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                If oPro.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                    If oPro.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned Then
                        If oApplication.SBO_Application.MessageBox("Do you want to update Production Phase / Sequence from BOM?", , "Yes", "No") = 2 Then
                        Else
                            oApplication.Utilities.populateBoMDetaisl(oForm, oPro.AbsoluteEntry, oPro.ItemNo)
                        End If
                        UpdateTable(oPro.DocumentNumber)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
