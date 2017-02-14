Option Explicit On
Imports System.Windows.Forms

Namespace AE_ITRG_AO01
    Public Class clsEventHandler
        Dim WithEvents SBO_Application As SAPbouiCOM.Application ' holds connection with SBO
        Dim p_oDICompany As New SAPbobsCOM.Company
        Public Sub New(ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company)
            Dim sFuncName As String = String.Empty
            Try
                sFuncName = "Class_Initialize()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retriving SBO Application handle", sFuncName)
                SBO_Application = oApplication
                p_oDICompany = oCompany

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Call WriteToLogFile(exc.Message, sFuncName)
            End Try
        End Sub

        Public Function SetApplication(ByRef sErrDesc As String) As Long
            ' **********************************************************************************
            '   Function   :    SetApplication()
            '   Purpose    :    This function will be calling to initialize the default settings
            '                   such as Retrieving the Company Default settings, Creating Menus, and
            '                   Initialize the Event Filters
            '               
            '   Parameters :    ByRef sErrDesc AS string
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            ' **********************************************************************************
            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "SetApplication()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetMenus()", sFuncName)
                If SetMenus(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetFilters()", sFuncName)
                If SetFilters(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                SetApplication = RTN_SUCCESS
            Catch exc As Exception
                sErrDesc = exc.Message
                Call WriteToLogFile(exc.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                SetApplication = RTN_ERROR
            End Try
        End Function

        Private Function SetMenus(ByRef sErrDesc As String) As Long
            ' **********************************************************************************
            '   Function   :    SetMenus()
            '   Purpose    :    This function will be gathering to create the customized menu
            '               
            '   Parameters :    ByRef sErrDesc AS string
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            ' **********************************************************************************
            Dim sFuncName As String = String.Empty
            ' Dim oMenuItem As SAPbouiCOM.MenuItem
            Try
                sFuncName = "SetMenus()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                SetMenus = RTN_SUCCESS
            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                SetMenus = RTN_ERROR
            End Try
        End Function

        Private Function SetFilters(ByRef sErrDesc As String) As Long

            ' **********************************************************************************
            '   Function   :    SetFilters()
            '   Purpose    :    This function will be gathering to declare the event filter 
            '                   before starting the AddOn Application
            '               
            '   Parameters :    ByRef sErrDesc AS string
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            ' **********************************************************************************

            Dim oFilters As SAPbouiCOM.EventFilters
            Dim oFilter As SAPbouiCOM.EventFilter
            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "SetFilters()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing EventFilters object", sFuncName)
                oFilters = New SAPbouiCOM.EventFilters



                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding filters", sFuncName)
                SBO_Application.SetFilter(oFilters)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                SetFilters = RTN_SUCCESS
            Catch exc As Exception
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                SetFilters = RTN_ERROR
            End Try
        End Function

        Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
            ' **********************************************************************************
            '   Function   :    SBO_Application_AppEvent()
            '   Purpose    :    This function will be handling the SAP Application Event
            '               
            '   Parameters :    ByVal EventType As SAPbouiCOM.BoAppEventTypes
            '                       EventType = set the SAP UI Application Eveny Object        
            ' **********************************************************************************
            Dim sFuncName As String = String.Empty
            Dim sErrDesc As String = String.Empty
            Dim sMessage As String = String.Empty

            Try
                sFuncName = "SBO_Application_AppEvent()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                Select Case EventType
                    Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                        sMessage = String.Format("Please wait for a while to disconnect the AddOn {0} ....", System.Windows.Forms.Application.ProductName)
                        p_oSBOApplication.SetStatusBarMessage(sMessage, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        End
                End Select

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch ex As Exception
                sErrDesc = ex.Message
                WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                ShowErr(sErrDesc)
            Finally
                GC.Collect()  'Forces garbage collection of all generations.
            End Try
        End Sub

        Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
            ' **********************************************************************************
            '   Function   :    SBO_Application_MenuEvent()
            '   Purpose    :    This function will be handling the SAP Menu Event
            '               
            '   Parameters :    ByRef pVal As SAPbouiCOM.MenuEvent
            '                       pVal = set the SAP UI MenuEvent Object
            '                   ByRef BubbleEvent As Boolean
            '                       BubbleEvent = set the True/False        
            ' **********************************************************************************
            ' Dim oForm As SAPbouiCOM.Form = Nothing
            Dim sErrDesc As String = String.Empty
            Dim sFuncName As String = String.Empty
            Dim oForm As SAPbouiCOM.Form = Nothing
            Dim oMatrix As SAPbouiCOM.Matrix = Nothing
            Dim flg As Boolean = False
            Try
                sFuncName = "SBO_Application_MenuEvent()"
                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If Not p_oDICompany.Connected Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                    If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If

                If pVal.BeforeAction = False Then
                    Select Case pVal.MenuUID

                        Case "MDRU"
                            Try
                                LoadFromXML("MasterSync.srf", SBO_Application)
                                oForm = p_oSBOApplication.Forms.Item("MDR")
                                oForm.Freeze(True)
                                oMatrix = oForm.Items.Item("5").Specific
                                oForm.Items.Item("t_docdate").Specific.value = Format(Now.Date, "yyyyMMdd")
                                oForm.Items.Item("chk_post").Visible = False
                                oForm.Items.Item("Chk_exch").Visible = False
                                oForm.Items.Item("Chk_Prices").Visible = False
                                oForm.Items.Item("Chk_bin").Visible = False
                                oForm.ActiveItem = "c_masttype"
                                oForm.Visible = True
                                oMatrix.AutoResizeColumns()
                                oForm.Freeze(False)
                                Exit Try
                            Catch ex As Exception
                                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                            End Try
                            Exit Sub
                        Case "TDBL"
                            Try
                                Dim strMenuId As String
                                Dim oRecSet As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecSet.DoQuery("SELECT count(*) as ""count"" FROM ""OUTB"" WHERE ""TableName"" <= 'AE_TB004_TARCRE' and  ""ObjectType"" = 0")
                                strMenuId = 51200 + oRecSet.Fields.Item("count").Value
                                p_oSBOApplication.ActivateMenuItem(strMenuId)
                                Dim oForm1 As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                Dim oMat As SAPbouiCOM.Matrix = oForm1.Items.Item("3").Specific
                                Dim ColnHeader As SAPbouiCOM.ColumnTitle = oMat.Columns.Item("Name").TitleObject
                                ColnHeader.Caption = "Target Database Name"
                            Catch ex As Exception
                                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                            End Try
                        Case "BPMS"
                            Try
                                flg = False
                                Dim strMenuId As String
                                Dim oRecSet As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecSet.DoQuery("SELECT count(*) as ""count"" FROM ""OUTB"" WHERE ""TableName"" <= 'AE_TB001_BPSETUP' and  ""ObjectType"" = 0")
                                strMenuId = 51200 + oRecSet.Fields.Item("count").Value

                                Dim UDT As SAPbobsCOM.UserTable
                                UDT = p_oDICompany.UserTables.Item("AE_TB001_BPSETUP")
                                Dim sqr As String = "select ""Code"", ""Name"" from ""@AE_TB004_TARCRE"""
                                Dim Rse As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Rse.DoQuery(sqr)
                                If Rse.RecordCount > 0 Then
                                    Rse.MoveFirst()
                                    For I As Integer = 0 To Rse.RecordCount - 1
                                        If UDT.GetByKey(Rse.Fields.Item(0).Value) = True Then
                                            flg = True
                                            Exit For
                                        End If
                                    Next
                                End If
                                If flg = False Then
                                    If Rse.RecordCount > 0 Then
                                        Rse.MoveFirst()
                                        For I As Integer = 0 To Rse.RecordCount - 1
                                            UDT.Code = Rse.Fields.Item(0).Value
                                            UDT.Name = Rse.Fields.Item(0).Value
                                            UDT.UserFields.Fields.Item("U_TargetDB").Value = Rse.Fields.Item(0).Value
                                            UDT.Add()
                                            Rse.MoveNext()
                                        Next
                                    End If
                                End If
                                p_oSBOApplication.ActivateMenuItem(strMenuId)
                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText("Manu Event Failed." & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                        Case "ITMS"
                            Try
                                flg = False
                                Dim strMenuId As String
                                Dim oRecSet As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecSet.DoQuery("SELECT count(*) as ""count"" FROM ""OUTB"" WHERE ""TableName"" <= 'AE_TB002_ITEM' and  ""ObjectType"" = 0")
                                strMenuId = 51200 + oRecSet.Fields.Item("count").Value

                                Dim UDT As SAPbobsCOM.UserTable
                                UDT = p_oDICompany.UserTables.Item("AE_TB002_ITEM")
                                Dim sqr As String = "select ""Code"", ""Name"" from ""@AE_TB004_TARCRE"""
                                Dim Rse As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Rse.DoQuery(sqr)
                                If Rse.RecordCount > 0 Then
                                    Rse.MoveFirst()
                                    For I As Integer = 0 To Rse.RecordCount - 1
                                        If UDT.GetByKey(Rse.Fields.Item(0).Value) = True Then
                                            flg = True
                                            Exit For
                                        End If
                                    Next
                                End If
                                If flg = False Then
                                    If Rse.RecordCount > 0 Then
                                        Rse.MoveFirst()
                                        For I As Integer = 0 To Rse.RecordCount - 1
                                            UDT.Code = Rse.Fields.Item(0).Value
                                            UDT.Name = Rse.Fields.Item(0).Value
                                            UDT.UserFields.Fields.Item("U_TargetDB").Value = Rse.Fields.Item(0).Value
                                            UDT.Add()
                                            Rse.MoveNext()
                                        Next
                                    End If
                                End If
                                p_oSBOApplication.ActivateMenuItem(strMenuId)
                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText("Manu Event Failed." & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                        Case "FINS"
                            Try
                                flg = False
                                Dim strMenuId As String
                                Dim oRecSet As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecSet.DoQuery("SELECT count(*) as ""count"" FROM ""OUTB"" WHERE ""TableName"" <= 'AE_TB003_FIN' and  ""ObjectType"" = 0")
                                strMenuId = 51200 + oRecSet.Fields.Item("count").Value

                                Dim UDT As SAPbobsCOM.UserTable
                                UDT = p_oDICompany.UserTables.Item("AE_TB003_FIN")
                                Dim sqr As String = "select ""Code"", ""Name"" from ""@AE_TB004_TARCRE"""
                                Dim Rse As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Rse.DoQuery(sqr)
                                If Rse.RecordCount > 0 Then
                                    Rse.MoveFirst()
                                    For I As Integer = 0 To Rse.RecordCount - 1
                                        If UDT.GetByKey(Rse.Fields.Item(0).Value) = True Then
                                            flg = True
                                            Exit For
                                        End If
                                    Next
                                End If
                                If flg = False Then
                                    If Rse.RecordCount > 0 Then
                                        Rse.MoveFirst()
                                        For I As Integer = 0 To Rse.RecordCount - 1
                                            UDT.Code = Rse.Fields.Item(0).Value
                                            UDT.Name = Rse.Fields.Item(0).Value
                                            UDT.UserFields.Fields.Item("U_TargetDB").Value = Rse.Fields.Item(0).Value
                                            UDT.Add()
                                            Rse.MoveNext()
                                        Next
                                    End If
                                End If
                                p_oSBOApplication.ActivateMenuItem(strMenuId)
                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText("Manu Event Failed." & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                    End Select
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                BubbleEvent = False
                ShowErr(exc.Message)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                WriteToLogFile(Err.Description, sFuncName)
            End Try
        End Sub

        Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, _
                ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
            Dim sErrDesc As String = String.Empty
            Dim Errcode As String = String.Empty
            Dim sFuncName As String = String.Empty
            Dim oDSTemplateInformation As DataSet = Nothing
            Dim oDTGatheredInformation As New DataTable

            Try
                sFuncName = "SBO_Application_ItemEvent()"
                ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If Not IsNothing(p_oDICompany) Then
                    If Not p_oDICompany.Connected Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                        If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                End If

                If pVal.BeforeAction = False Then

                    Select Case pVal.FormUID

                        Case "MDR"
                            Select Case pVal.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    Select Case pVal.ItemUID
                                        Case "c_masttype"
                                            If pVal.ItemChanged = True Then
                                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                                Dim oMastertype As SAPbouiCOM.ComboBox = oForm.Items.Item("c_masttype").Specific
                                                If oMastertype.Selected.Value = "F" Then
                                                    oForm.Items.Item("chk_post").Visible = True
                                                    oForm.Items.Item("Chk_exch").Visible = True
                                                    oForm.Items.Item("Chk_Prices").Visible = False
                                                    oForm.Items.Item("Chk_bin").Visible = False
                                                ElseIf oMastertype.Selected.Value = "I" Then
                                                    oForm.Items.Item("Chk_Prices").Visible = True
                                                    oForm.Items.Item("Chk_bin").Visible = True
                                                    oForm.Items.Item("chk_post").Visible = False
                                                    oForm.Items.Item("Chk_exch").Visible = False
                                                Else
                                                    oForm.Items.Item("Chk_Prices").Visible = False
                                                    oForm.Items.Item("chk_post").Visible = False
                                                    oForm.Items.Item("Chk_exch").Visible = False
                                                    oForm.Items.Item("Chk_bin").Visible = False
                                                End If
                                                LoadReplicationDetails()
                                            End If
                                    End Select
                                Case SAPbouiCOM.BoEventTypes.et_CLICK
                                    Select Case pVal.ItemUID
                                        Case "Chk_Select"
                                            Try
                                                Dim fllg As Boolean = False
                                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                                Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("5").Specific
                                                oForm.Freeze(True)
                                                Dim SelectAll As SAPbouiCOM.CheckBox = oForm.Items.Item("Chk_Select").Specific
                                                If SelectAll.Checked = True Then fllg = True
                                                If oMatrix.VisualRowCount > 0 Then
                                                    For I As Integer = 1 To oMatrix.VisualRowCount
                                                        Dim LinSelect As SAPbouiCOM.CheckBox = oMatrix.Columns.Item("Select").Cells.Item(I).Specific
                                                        If fllg = True Then
                                                            LinSelect.Checked = False
                                                            LinSelect.ValOn = "Y"
                                                            LinSelect.ValOff = "N"
                                                        Else
                                                            LinSelect.Checked = True
                                                            LinSelect.ValOn = "Y"
                                                            LinSelect.ValOff = "N"
                                                        End If
                                                    Next
                                                End If
                                                oForm.Freeze(False)
                                            Catch ex As Exception

                                            End Try
                                    End Select
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    Select Case pVal.ItemUID

                                        Case "Refresh"
                                            Try
                                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                                If oForm.Items.Item("c_masttype").Specific.value.ToString.Trim() <> "" Then
                                                    LoadReplicationDetails()
                                                End If
                                            Catch ex As Exception

                                            End Try
                                        Case "Replicate"
                                            Try
                                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                                Dim oRset As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                Dim oRset1 As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                Dim sSQL As String = String.Empty
                                                dtTable = New DataTable
                                                Dim sCheck As String = String.Empty
                                                Dim oDICompany() As SAPbobsCOM.Company = Nothing
                                                Dim sMasterDataType As String = String.Empty
                                                Dim sMasterDataCodeF As String = String.Empty
                                                Dim sMasterDataCodeT As String = String.Empty
                                                Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("5").Specific
                                                Dim ErrerCode As Long
                                                Dim ErrerMsg As String = ""
                                                Dim Fllag As Boolean = False
                                                Try
                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Validation()", sFuncName)
                                                    oForm.Items.Item("Replicate").Enabled = False

                                                    oDT_ErrorMsg = New DataTable
                                                    oDT_ErrorMsg.Columns.Add("ErrorMsg", GetType(String))

                                                    SBO_Application.SetStatusBarMessage("Validation Process Started ........!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                    If HeaderValidation(oForm, sErrDesc) = RTN_ERROR Then
                                                        oForm.Items.Item("Replicate").Enabled = True
                                                        BubbleEvent = False
                                                        Exit Sub
                                                    End If

                                                    Dim cnt As Integer = oDT_Entities.Rows.Count
                                                    SBO_Application.SetStatusBarMessage("Validation Completed ........!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                    Dim oMasttype As SAPbouiCOM.ComboBox = oForm.Items.Item("c_masttype").Specific
                                                    If oMasttype.Selected.Value = "B" Then

                                                        Dim sqry As String = "Select T1.""Name"", T1.""U_UserName"", T1.""U_Password"", T0.""U_Customers"" ,T0.""U_Vendors"" , T0.""U_Leads"",T0.""U_PayTerms"" , T0.""U_BPGroups"" from ""@AE_TB001_BPSETUP"" T0 LEFT OUTER JOIN ""@AE_TB004_TARCRE"" T1 ON T0.""U_TargetDB"" = T1.""Code""   WHERE (T0.""U_Customers""  = 'Y' OR ""U_Vendors"" = 'Y' OR T0.""U_Leads"" = 'Y' OR T0.""U_PayTerms"" = 'Y' OR T0.""U_BPGroups""= 'Y');"
                                                        oRset1.DoQuery(sqry)
                                                        oDT_BPSetup = New DataTable
                                                        oDT_BPSetup = ConvertRecordset(oRset1)
                                                        Dim dtcount As Integer = oDT_BPSetup.Rows.Count
                                                        Dim oDV_BPSetup As New DataView(oDT_BPSetup)
                                                        Dim dvcount As Integer = oDV_BPSetup.Count
                                                        ' Dim dv As New DataView(dt)

                                                        Dim rrr As String = oMatrix.RowCount

                                                        If oDT_Entities.Rows.Count > 0 Then
                                                            For J As Integer = 0 To oDT_Entities.Rows.Count - 1
                                                                ''=========================================================================================================
                                                                '------------------------------- BP Master Data Replication ----------------------------------------------
                                                                ''=========================================================================================================
                                                                If (oDT_Entities.Rows(J).Item("TransType").ToString = "Customer" Or oDT_Entities.Rows(J).Item("TransType").ToString = "Lead" Or oDT_Entities.Rows(J).Item("TransType").ToString = "Suppliers") Then
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling BP Setup()", sFuncName)
                                                                    Dim rvcount As Integer = oDV_BPSetup.Count
                                                                    Fllag = False
                                                                    If oDT_Entities.Rows(J).Item("TransType").ToString = "Customer" Then
                                                                        oDV_BPSetup.RowFilter = "U_Customers ='Y'"
                                                                    ElseIf oDT_Entities.Rows(J).Item("TransType").ToString = "Suppliers" Then
                                                                        oDV_BPSetup.RowFilter = "U_Vendors ='Y'"
                                                                    ElseIf oDT_Entities.Rows(J).Item("TransType").ToString = "Lead" Then
                                                                        oDV_BPSetup.RowFilter = "U_Leads ='Y'"
                                                                    End If
                                                                    Dim Statu As String = oDT_Entities.Rows(J).Item("SyncStatus").ToString

                                                                    ReDim oDICompany(oDV_BPSetup.Count)
                                                                    Dim dvcount1 As Integer = oDV_BPSetup.Count
                                                                    If oDV_BPSetup.Count > 0 Then
                                                                        For S As Integer = 0 To oDV_BPSetup.Count - 1
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Connect To TargetCompany()", sFuncName)
                                                                            SBO_Application.SetStatusBarMessage("Connecting to the Target Company.." & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                            Dim targetent As String = oDV_BPSetup.Item(S).Item("Name").ToString

                                                                            If ConnectToTargetCompany(oDICompany(S), oDV_BPSetup.Item(S).Item("Name").ToString, oDV_BPSetup.Item(S).Item("U_UserName").ToString, oDV_BPSetup.Item(S).Item("U_Password").ToString, sErrDesc) <> RTN_SUCCESS Then
                                                                                GoTo 107
                                                                            End If
                                                                            SBO_Application.StatusBar.SetText("Connecting to the target company Successful.. " & oDICompany(S).CompanyDB, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction on company database " & oDICompany(S).CompanyDB, sFuncName)
                                                                            oDICompany(S).StartTransaction()
                                                                            Dim BPCODE As String = oDT_Entities.Rows(J).Item("Code").ToString
                                                                            If CreateBPMaster(oDT_Entities.Rows(J).Item("Code").ToString, oDICompany(S), sErrDesc) <> RTN_SUCCESS Then
107:
                                                                                sErrDesc = sErrDesc + " On.." + oDICompany(S).CompanyDB
                                                                                p_oSBOApplication.StatusBar.SetText("BP Replication Failed.." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                Fllag = False
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)

                                                                                Dim sqy As String = " Update ""INTEGRATION"" SET ""SYNCSTATUS"" = 'NO', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '" & sErrDesc.ToString.Replace("'", """") & "'  WHERE ""MASTERTYPE"" = 'BPMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                                Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                Run.DoQuery(sqy)

                                                                                Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                                oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "NO"
                                                                                oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                                oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = sErrDesc

                                                                                Dim erro As String = oDT_Entities.Rows(J).Item("SyncErrMsg").ToString

                                                                                For lCounter As Integer = 0 To UBound(oDICompany)
                                                                                    If Not oDICompany(lCounter) Is Nothing Then
                                                                                        If oDICompany(lCounter).Connected = True Then
                                                                                            If oDICompany(lCounter).InTransaction = True Then
                                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                                oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                                            End If
                                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                            oDICompany(lCounter).Disconnect()
                                                                                            oDICompany(lCounter) = Nothing
                                                                                        End If
                                                                                    End If
                                                                                Next
                                                                                Exit For
                                                                            Else
                                                                                Fllag = True
                                                                                SBO_Application.SetStatusBarMessage("'" & BPCODE & "' : BP Replicated Successfully on.. " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                            End If
                                                                        Next
                                                                        For lCounter As Integer = 0 To UBound(oDICompany)
                                                                            If Not oDICompany(lCounter) Is Nothing Then
                                                                                If oDICompany(lCounter).Connected = True Then
                                                                                    If oDICompany(lCounter).InTransaction = True Then
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                        oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                                                    End If
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                    oDICompany(lCounter).Disconnect()
                                                                                    oDICompany(lCounter) = Nothing
                                                                                End If
                                                                            End If
                                                                        Next
                                                                        If Fllag = True Then
                                                                            Dim sqy As String = " Update ""INTEGRATION"" SET ""SYNCSTATUS"" = 'YES', ""SYNCDATE"" = CURRENT_DATE , ""ERRORMSG""= '' WHERE ""MASTERTYPE"" = 'BPMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                            Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                            Run.DoQuery(sqy)

                                                                            Dim row As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                            oMatrix.Columns.Item("SyncStatus").Cells.Item(row).Specific.value() = "YES"
                                                                            oMatrix.Columns.Item("SyncDate").Cells.Item(row).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                            oMatrix.Columns.Item("ErrMsg").Cells.Item(row).Specific.value = "SUCCESS"
                                                                        End If
                                                                    End If
                                                                ElseIf oDT_Entities.Rows(J).Item("TransType").ToString = "BPGroup" Then
                                                                    ''=========================================================================================================
                                                                    '------------------------------- BP Group Replication ----------------------------------------------
                                                                    ''=========================================================================================================
                                                                    Try
                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling BP Setup() for BP Group Replication", sFuncName)
                                                                        Dim rvcount As Integer = oDV_BPSetup.Count
                                                                        Fllag = False
                                                                        oDV_BPSetup.RowFilter = "U_BPGroups ='Y'"
                                                                        sFuncName = "BP Group"

                                                                        ReDim oDICompany(oDV_BPSetup.Count)
                                                                        Dim dvcount1 As Integer = oDV_BPSetup.Count
                                                                        If oDV_BPSetup.Count > 0 Then
                                                                            For S As Integer = 0 To oDV_BPSetup.Count - 1
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Connect To TargetCompany() for BP Group", sFuncName)
                                                                                SBO_Application.SetStatusBarMessage("Connecting to the Target Company. " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                                Dim targetent As String = oDV_BPSetup.Item(S).Item("Name").ToString
                                                                                '------------------------------------
                                                                                'Connecting the target company.......
                                                                                '------------------------------------
                                                                                If ConnectToTargetCompany(oDICompany(S), oDV_BPSetup.Item(S).Item("Name").ToString, oDV_BPSetup.Item(S).Item("U_UserName").ToString, oDV_BPSetup.Item(S).Item("U_Password").ToString, sErrDesc) <> RTN_SUCCESS Then
                                                                                    GoTo 108
                                                                                End If
                                                                                SBO_Application.SetStatusBarMessage("Connecting to the target company Successfull On.." & 107, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to Target company Successful on. " & oDICompany(S).CompanyDB, sFuncName)

                                                                                oDICompany(S).StartTransaction()
                                                                                SBO_Application.StatusBar.SetText("Started BP Groups Synchronization on '" & oDV_BPSetup.Item(S).Item("Name").ToString & "' ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                                                                                '------------------------------------
                                                                                'Declaring the BP Groups objects for replication..........
                                                                                '------------------------------------
                                                                                Dim oBPGroup As SAPbobsCOM.BusinessPartnerGroups = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartnerGroups)
                                                                                Dim oTargetBPGroup As SAPbobsCOM.BusinessPartnerGroups = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartnerGroups)

                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
                                                                                Dim flg As Boolean = False
                                                                                Dim SucFlag As Boolean = False
                                                                                Dim groupcode As String = String.Empty

                                                                                '--------------------------------------------------------------------------------------
                                                                                'Check whether Group already exist or not.. if exist update.... else Add.........
                                                                                '--------------------------------------------------------------------------------------
                                                                                Dim orsGroup As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                Dim ss As String = "Select ""GroupCode"" from ""OCRG"" where ""GroupName"" = '" & oDT_Entities.Rows(J).Item("Name").ToString & "'"
                                                                                orsGroup.DoQuery(String.Format("Select ""GroupCode"" from ""OCRG"" where ""GroupName"" = '{0}'", oDT_Entities.Rows(J).Item("Name").ToString))
                                                                                If orsGroup.RecordCount = 1 Then
                                                                                    flg = True
                                                                                    groupcode = orsGroup.Fields.Item(0).Value
                                                                                End If

                                                                                If oBPGroup.GetByKey(oDT_Entities.Rows(J).Item("Code")) Then
                                                                                    If flg = True Then
                                                                                        If oTargetBPGroup.GetByKey(groupcode) = True Then
                                                                                            oTargetBPGroup.Name = oBPGroup.Name
                                                                                            oTargetBPGroup.Type = oBPGroup.Type
                                                                                        End If
                                                                                    Else
                                                                                        oTargetBPGroup.Name = oBPGroup.Name
                                                                                        oTargetBPGroup.Type = oBPGroup.Type
                                                                                    End If
                                                                                End If
                                                                                If flg = True Then
                                                                                    ErrerCode = oTargetBPGroup.Update
                                                                                    If ErrerCode <> 0 Then
                                                                                        oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating BP Group '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                        SucFlag = False
                                                                                    Else
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully updated BP Group '" & oDT_Entities.Rows(J).Item("Code") & "' on '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                        SucFlag = True
                                                                                    End If
                                                                                Else
                                                                                    ErrerCode = oTargetBPGroup.Add
                                                                                    If ErrerCode <> 0 Then
                                                                                        oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding BP Group '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                        SucFlag = False
                                                                                    Else
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully added BP Group '" & oDT_Entities.Rows(J).Item("Code") & "' on '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                        SucFlag = True
                                                                                    End If
                                                                                End If
                                                                                '--------------------------------------------------------------------------------------
                                                                                'Success flag... If Any error while replicating then Rollback and Disconnet DB connection... else Continue to next Target DB.........
                                                                                '--------------------------------------------------------------------------------------
                                                                                If SucFlag = False Then
108:
                                                                                    'oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                    sErrDesc = sErrDesc + " On.." + oDICompany(S).CompanyDB
                                                                                    Fllag = False
                                                                                    Dim sqy As String = " Update ""INTEGRATION"" SET ""SYNCSTATUS"" = 'NO', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '" & sErrDesc.ToString.Replace("'", """") & "'  WHERE ""MASTERTYPE"" = 'BPMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                                    Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                    Run.DoQuery(sqy)

                                                                                    Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                                    oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "NO"
                                                                                    oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                                    oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = sErrDesc

                                                                                    p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    For lCounter As Integer = 0 To UBound(oDICompany)
                                                                                        If Not oDICompany(lCounter) Is Nothing Then
                                                                                            If oDICompany(lCounter).Connected = True Then
                                                                                                If oDICompany(lCounter).InTransaction = True Then
                                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                                    oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                                                End If
                                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                                oDICompany(lCounter).Disconnect()
                                                                                                oDICompany(lCounter) = Nothing
                                                                                            End If
                                                                                        End If
                                                                                    Next
                                                                                    Exit For
                                                                                Else
                                                                                    Fllag = True
                                                                                    SBO_Application.SetStatusBarMessage("BP Group '" & oDT_Entities.Rows(J).Item("Code") & "' Replicated Successfully.. " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                                End If
                                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetBPGroup)
                                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBPGroup)
                                                                            Next
                                                                            '--------------------------------------------------------------------------------------
                                                                            'If No error while replicating then COMMIT and Disconnet DB connections...
                                                                            '--------------------------------------------------------------------------------------
                                                                            For lCounter As Integer = 0 To UBound(oDICompany)
                                                                                If Not oDICompany(lCounter) Is Nothing Then
                                                                                    If oDICompany(lCounter).Connected = True Then
                                                                                        If oDICompany(lCounter).InTransaction = True Then
                                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                            oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                                                        End If
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                        oDICompany(lCounter).Disconnect()
                                                                                        oDICompany(lCounter) = Nothing
                                                                                    End If
                                                                                End If
                                                                            Next
                                                                            '--------------------------------------------------------------------------------------
                                                                            'Updating success flag to UI Table and Integration TABLE.....
                                                                            '--------------------------------------------------------------------------------------
                                                                            If Fllag = True Then
                                                                                Dim sqy As String = " Update ""INTEGRATION"" SET ""SYNCSTATUS"" = 'YES', ""SYNCDATE"" = CURRENT_DATE , ""ERRORMSG""= '' WHERE ""MASTERTYPE"" = 'BPMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                                Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                Run.DoQuery(sqy)

                                                                                Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                                oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "YES"
                                                                                oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                                oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = "SUCCESS"
                                                                            End If

                                                                        End If

                                                                    Catch ex As Exception
                                                                        SBO_Application.SetStatusBarMessage("BP Groups Posting Failed..." & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                    Finally

                                                                    End Try
                                                                ElseIf oDT_Entities.Rows(J).Item("TransType").ToString = "PaymentTerms" Then
                                                                    ''=========================================================================================================
                                                                    '------------------------------- BP Payment Terms Type Replication ----------------------------------------------
                                                                    ''=========================================================================================================
                                                                    Try
                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling BP Setup for payment terms Setting", sFuncName)
                                                                        Dim rvcount As Integer = oDV_BPSetup.Count
                                                                        sFuncName = "Payment Terms()"
                                                                        Fllag = False
                                                                        oDV_BPSetup.RowFilter = "U_PayTerms ='Y'"
                                                                        ReDim oDICompany(oDV_BPSetup.Count)
                                                                        Dim dvcount1 As Integer = oDV_BPSetup.Count
                                                                        If oDV_BPSetup.Count > 0 Then
                                                                            For S As Integer = 0 To oDV_BPSetup.Count - 1
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Connect To TargetCompany for Payment Terms", sFuncName)
                                                                                SBO_Application.StatusBar.SetText("Connecting to the target company.. " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                Dim targetent As String = oDV_BPSetup.Item(S).Item("Name").ToString
                                                                                '------------------------------------
                                                                                'Connecting the target company.......
                                                                                '------------------------------------
                                                                                If ConnectToTargetCompany(oDICompany(S), oDV_BPSetup.Item(S).Item("Name").ToString, oDV_BPSetup.Item(S).Item("U_UserName").ToString, oDV_BPSetup.Item(S).Item("U_Password").ToString, sErrDesc) <> RTN_SUCCESS Then
                                                                                    '---------------------------------------------------------------------
                                                                                    ' Any error while connecting database then its will go to Rollback ...
                                                                                    '---------------------------------------------------------------------
                                                                                    GoTo 109
                                                                                End If
                                                                                SBO_Application.StatusBar.SetText("Connecting to the target company Successful " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the target company Successful " & oDICompany(S).CompanyDB, sFuncName)
                                                                                oDICompany(S).StartTransaction()

                                                                                SBO_Application.StatusBar.SetText("Started Master Data Synchronization " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                                                                                Dim oPaymentTerms As SAPbobsCOM.PaymentTermsTypes = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentTermsTypes)
                                                                                Dim oTargetPaymentTerms As SAPbobsCOM.PaymentTermsTypes = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentTermsTypes)
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                                                                                Dim flg1 As Boolean = False
                                                                                Dim SucFlag As Boolean = False
                                                                                Dim groupno As String = String.Empty

                                                                                Dim orsGroup As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                orsGroup.DoQuery(String.Format("Select ""GroupNum"" from ""OCTG"" where ""PymntGroup"" = '{0}'", oDT_Entities.Rows(J).Item("Name").ToString))
                                                                                If orsGroup.RecordCount = 1 Then
                                                                                    flg1 = True
                                                                                    groupno = orsGroup.Fields.Item(0).Value
                                                                                End If

                                                                                If oPaymentTerms.GetByKey(oDT_Entities.Rows(J).Item("Code")) Then
                                                                                    If flg1 = True Then
                                                                                        If oTargetPaymentTerms.GetByKey(groupno) Then
                                                                                            oTargetPaymentTerms.PaymentTermsGroupName = oPaymentTerms.PaymentTermsGroupName
                                                                                            oTargetPaymentTerms.BaselineDate = oPaymentTerms.BaselineDate
                                                                                            oTargetPaymentTerms.StartFrom = oPaymentTerms.StartFrom
                                                                                            oTargetPaymentTerms.NumberOfAdditionalDays = oPaymentTerms.NumberOfAdditionalDays
                                                                                            oTargetPaymentTerms.NumberOfAdditionalMonths = oPaymentTerms.NumberOfAdditionalMonths
                                                                                            oTargetPaymentTerms.OpenReceipt = oPaymentTerms.OpenReceipt
                                                                                            oTargetPaymentTerms.CreditLimit = oPaymentTerms.CreditLimit
                                                                                            oTargetPaymentTerms.DiscountCode = oPaymentTerms.DiscountCode
                                                                                            oTargetPaymentTerms.GeneralDiscount = oPaymentTerms.GeneralDiscount
                                                                                            oTargetPaymentTerms.LoadLimit = oPaymentTerms.LoadLimit
                                                                                            oTargetPaymentTerms.InterestOnArrears = oPaymentTerms.InterestOnArrears
                                                                                            oTargetPaymentTerms.NumberOfToleranceDays = oPaymentTerms.NumberOfToleranceDays
                                                                                            oTargetPaymentTerms.PriceListNo = oPaymentTerms.PriceListNo
                                                                                        End If
                                                                                    Else
                                                                                        oTargetPaymentTerms.PaymentTermsGroupName = oPaymentTerms.PaymentTermsGroupName
                                                                                        oTargetPaymentTerms.BaselineDate = oPaymentTerms.BaselineDate
                                                                                        oTargetPaymentTerms.StartFrom = oPaymentTerms.StartFrom
                                                                                        oTargetPaymentTerms.NumberOfAdditionalDays = oPaymentTerms.NumberOfAdditionalDays
                                                                                        oTargetPaymentTerms.NumberOfAdditionalMonths = oPaymentTerms.NumberOfAdditionalMonths
                                                                                        oTargetPaymentTerms.OpenReceipt = oPaymentTerms.OpenReceipt
                                                                                        oTargetPaymentTerms.CreditLimit = oPaymentTerms.CreditLimit
                                                                                        oTargetPaymentTerms.DiscountCode = oPaymentTerms.DiscountCode
                                                                                        oTargetPaymentTerms.GeneralDiscount = oPaymentTerms.GeneralDiscount
                                                                                        oTargetPaymentTerms.LoadLimit = oPaymentTerms.LoadLimit
                                                                                        oTargetPaymentTerms.InterestOnArrears = oPaymentTerms.InterestOnArrears
                                                                                        oTargetPaymentTerms.NumberOfToleranceDays = oPaymentTerms.NumberOfToleranceDays
                                                                                        oTargetPaymentTerms.PriceListNo = oPaymentTerms.PriceListNo
                                                                                    End If
                                                                                End If
                                                                                If flg1 = True Then
                                                                                    ErrerCode = oTargetPaymentTerms.Update()
                                                                                    If ErrerCode <> 0 Then
                                                                                        oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating Payment Terms '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                        SucFlag = False
                                                                                    Else
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully added Payment Terms '" & oDT_Entities.Rows(J).Item("Code") & "' on '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                        SucFlag = True
                                                                                    End If
                                                                                Else
                                                                                    ErrerCode = oTargetPaymentTerms.Add
                                                                                    If ErrerCode <> 0 Then
                                                                                        oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Payment Terms '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                        SucFlag = False
                                                                                    Else
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully added Payment Terms '" & oDT_Entities.Rows(J).Item("Code") & "' on '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                        SucFlag = True
                                                                                    End If
                                                                                End If

                                                                                If SucFlag = False Then
109:
                                                                                    sErrDesc = sErrDesc + " On.." + oDICompany(S).CompanyDB
                                                                                    p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    Fllag = False
                                                                                    Dim sqy As String = " Update ""INTEGRATION"" SET ""SYNCSTATUS"" = 'NO', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '" & sErrDesc.ToString.Replace("'", """") & "'  WHERE ""MASTERTYPE"" = 'BPMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                                    Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                    Run.DoQuery(sqy)

                                                                                    Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                                    oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "NO"
                                                                                    oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                                    oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = sErrDesc

                                                                                    For lCounter As Integer = 0 To UBound(oDICompany)
                                                                                        If Not oDICompany(lCounter) Is Nothing Then
                                                                                            If oDICompany(lCounter).Connected = True Then
                                                                                                If oDICompany(lCounter).InTransaction = True Then
                                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                                    oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                                                End If
                                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                                oDICompany(lCounter).Disconnect()
                                                                                                oDICompany(lCounter) = Nothing
                                                                                            End If
                                                                                        End If
                                                                                    Next
                                                                                    Exit For
                                                                                Else
                                                                                    Fllag = True
                                                                                    SBO_Application.SetStatusBarMessage("PayTerms '" & oDT_Entities.Rows(J).Item("Code") & "' Replicated Successfully.. " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PayTerms '" & oDT_Entities.Rows(J).Item("Code") & "' Replicated Successfully.. " & oDV_BPSetup.Item(S).Item("Name").ToString, sFuncName)
                                                                                End If
                                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetPaymentTerms)
                                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oPaymentTerms)
                                                                            Next
                                                                            For lCounter As Integer = 0 To UBound(oDICompany)
                                                                                If Not oDICompany(lCounter) Is Nothing Then
                                                                                    If oDICompany(lCounter).Connected = True Then
                                                                                        If oDICompany(lCounter).InTransaction = True Then
                                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                            oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                                                        End If
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                        oDICompany(lCounter).Disconnect()
                                                                                        oDICompany(lCounter) = Nothing
                                                                                    End If
                                                                                End If
                                                                            Next
                                                                            If Fllag = True Then
                                                                                Dim sqy As String = " Update ""INTEGRATION"" SET ""SYNCSTATUS"" = 'YES', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '' WHERE ""MASTERTYPE"" = 'BPMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                                Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                Run.DoQuery(sqy)

                                                                                Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                                oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "YES"
                                                                                oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                                oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = "SUCCESS"
                                                                            End If
                                                                        End If
                                                                    Catch ex As Exception
                                                                        SBO_Application.SetStatusBarMessage("Payment Terms Posting Failed...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                    End Try
                                                                End If
                                                            Next J
                                                        End If
                                                    ElseIf oMasttype.Selected.Value = "I" Then
                                                        Dim sqry As String = "Select T1.""Name"", T1.""U_UserName"", T1.""U_Password"", T0.""U_ItemGroups"" ,T0.""U_ItemCodes"" , T0.""U_BinLocatin"",T0.""U_BOM"" , T0.""U_PriceLists"", T0.""U_UOMGroups"" from ""@AE_TB002_ITEM"" T0  LEFT OUTER JOIN ""@AE_TB004_TARCRE"" T1 ON T0.""U_TargetDB"" = T1.""Code"" WHERE (T0.""U_ItemGroups"" = 'Y'  OR T0.""U_ItemCodes"" = 'Y' OR  T0.""U_BinLocatin""='Y' OR T0.""U_BOM""='Y' OR T0.""U_PriceLists"" ='Y' OR T0.""U_UOMGroups"" = 'Y');"
                                                        oRset1.DoQuery(sqry)
                                                        oDT_BPSetup = New DataTable
                                                        oDT_BPSetup = ConvertRecordset(oRset1)
                                                        Dim dtcount As Integer = oDT_BPSetup.Rows.Count
                                                        Dim oDV_BPSetup As New DataView(oDT_BPSetup)
                                                        Dim dvcount As Integer = oDV_BPSetup.Count
                                                        ' Dim dv As New DataView(dt)

                                                        Dim rrr As String = oMatrix.RowCount

                                                        If (oDT_Entities.Rows.Count > 0 And oDV_BPSetup.Count > 0) Then
                                                            For J As Integer = 0 To oDT_Entities.Rows.Count - 1

                                                                If (oDT_Entities.Rows(J).Item("TransType").ToString = "ItemGroups") Then
                                                                    ''=========================================================================================================
                                                                    '------------------------------- Item Groups Replication ----------------------------------------------
                                                                    ''=========================================================================================================
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Item Master Setup for Item Groups Settings", sFuncName)
                                                                    Dim rvcount As Integer = oDV_BPSetup.Count
                                                                    Fllag = False
                                                                    oDV_BPSetup.RowFilter = "U_ItemGroups ='Y'"
                                                                    ReDim oDICompany(oDV_BPSetup.Count)
                                                                    Dim dvcount1 As Integer = oDV_BPSetup.Count
                                                                    If oDV_BPSetup.Count > 0 Then
                                                                        For S As Integer = 0 To oDV_BPSetup.Count - 1
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Connect To TargetCompany() for Item Groups Posting", sFuncName)
                                                                            SBO_Application.StatusBar.SetText("Connecting to the target company " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                            Dim targetent As String = oDV_BPSetup.Item(S).Item("Name").ToString
                                                                            If ConnectToTargetCompany(oDICompany(S), oDV_BPSetup.Item(S).Item("Name").ToString, oDV_BPSetup.Item(S).Item("U_UserName").ToString, oDV_BPSetup.Item(S).Item("U_Password").ToString, sErrDesc) <> RTN_SUCCESS Then
                                                                                GoTo 110
                                                                            End If
                                                                            SBO_Application.StatusBar.SetText("Connecting to the target company Successfull " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the target company Successful on" & oDICompany(S).CompanyDB, sFuncName)
                                                                            oDICompany(S).StartTransaction()

                                                                            SBO_Application.SetStatusBarMessage("Started Master Data Synchronization " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                                            Dim oItemGroups As SAPbobsCOM.ItemGroups = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItemGroups)
                                                                            Dim oTargetItemGroup As SAPbobsCOM.ItemGroups = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItemGroups)
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                                                                            Dim flg1 As Boolean = False
                                                                            Dim SucFlag As Boolean = False
                                                                            Dim groupno As String = String.Empty

                                                                            Dim orsGroup As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                            orsGroup.DoQuery(String.Format("Select ""ItmsGrpCod"" from ""OITB"" where ""ItmsGrpNam"" = '{0}'", oDT_Entities.Rows(J).Item("Name").ToString))
                                                                            If orsGroup.RecordCount = 1 Then
                                                                                flg1 = True
                                                                                groupno = orsGroup.Fields.Item(0).Value
                                                                            End If

                                                                            If oItemGroups.GetByKey(oDT_Entities.Rows(J).Item("Code")) Then
                                                                                If flg1 = True Then
                                                                                    If oTargetItemGroup.GetByKey(groupno) Then
                                                                                        oTargetItemGroup.GroupName = oItemGroups.GroupName
                                                                                        oTargetItemGroup.PlanningSystem = oItemGroups.PlanningSystem
                                                                                        oTargetItemGroup.ProcurementMethod = oItemGroups.ProcurementMethod
                                                                                        oTargetItemGroup.OrderMultiple = oItemGroups.OrderMultiple
                                                                                        oTargetItemGroup.MinimumOrderQuantity = oItemGroups.MinimumOrderQuantity
                                                                                        oTargetItemGroup.LeadTime = oItemGroups.LeadTime
                                                                                        oTargetItemGroup.ToleranceDays = oItemGroups.ToleranceDays
                                                                                        oTargetItemGroup.InventorySystem = oItemGroups.InventorySystem
                                                                                        Dim cycle As String = oItemGroups.CycleCode
                                                                                        If oItemGroups.CycleCode <> 0 Then
                                                                                            oTargetItemGroup.OrderInterval = oItemGroups.OrderInterval
                                                                                        End If
                                                                                        Dim DefUOMGrp As Integer = oItemGroups.DefaultUoMGroup
                                                                                        Dim UOMGrpEnt As String = String.Empty
                                                                                        If DefUOMGrp <> 0 Then
                                                                                            Dim Sgrp As String = "Select ""UgpCode"" from OUGP where ""UgpEntry"" =  '" & DefUOMGrp & "'"
                                                                                            Dim Rgrp As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                            Rgrp.DoQuery(Sgrp)
                                                                                            If Rgrp.RecordCount > 0 Then UOMGrpEnt = Rgrp.Fields.Item(0).Value
                                                                                        End If
                                                                                        Dim oRGroup As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                        oRGroup.DoQuery(String.Format("Select ""UgpEntry"" from ""OUGP"" where ""UgpCode"" = '{0}'", UOMGrpEnt))
                                                                                        If oRGroup.RecordCount = 1 Then
                                                                                            oTargetItemGroup.DefaultUoMGroup = oRGroup.Fields.Item(0).Value
                                                                                        End If

                                                                                        Dim DefUOM As Integer = oItemGroups.DefaultInventoryUoM
                                                                                        Dim UOMEntry As String = String.Empty
                                                                                        If DefUOM <> 0 Then
                                                                                            Dim Sgrp As String = "Select ""UomCode"" from OUOM where ""UomEntry"" =  '" & DefUOM & "'"
                                                                                            Dim Rgrp As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                            Rgrp.DoQuery(Sgrp)
                                                                                            If Rgrp.RecordCount > 0 Then UOMEntry = Rgrp.Fields.Item(0).Value
                                                                                        End If
                                                                                        Dim oRGroup1 As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                        oRGroup1.DoQuery(String.Format("Select ""UomEntry"" from ""OUOM"" where ""UomCode"" = '{0}'", UOMEntry))
                                                                                        If oRGroup1.RecordCount = 1 Then
                                                                                            oTargetItemGroup.DefaultInventoryUoM = oRGroup1.Fields.Item(0).Value
                                                                                        End If
                                                                                    End If
                                                                                Else
                                                                                    oTargetItemGroup.GroupName = oItemGroups.GroupName
                                                                                    oTargetItemGroup.PlanningSystem = oItemGroups.PlanningSystem
                                                                                    oTargetItemGroup.ProcurementMethod = oItemGroups.ProcurementMethod
                                                                                    oTargetItemGroup.OrderMultiple = oItemGroups.OrderMultiple
                                                                                    oTargetItemGroup.MinimumOrderQuantity = oItemGroups.MinimumOrderQuantity
                                                                                    oTargetItemGroup.LeadTime = oItemGroups.LeadTime
                                                                                    oTargetItemGroup.ToleranceDays = oItemGroups.ToleranceDays
                                                                                    oTargetItemGroup.InventorySystem = oItemGroups.InventorySystem
                                                                                    If oItemGroups.CycleCode <> 0 Then
                                                                                        oTargetItemGroup.OrderInterval = oItemGroups.OrderInterval
                                                                                    End If
                                                                                    Dim DefUOMGrp As Integer = oItemGroups.DefaultUoMGroup
                                                                                    Dim UOMGRPEntry As String = String.Empty
                                                                                    If DefUOMGrp <> 0 Then
                                                                                        Dim Sgrp As String = "Select ""UgpCode"" from OUGP where ""UgpEntry"" =  '" & DefUOMGrp & "'"
                                                                                        Dim Rgrp As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                        Rgrp.DoQuery(Sgrp)
                                                                                        If Rgrp.RecordCount > 0 Then UOMGRPEntry = Rgrp.Fields.Item(0).Value
                                                                                    End If
                                                                                    Dim oRGroup As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                    oRGroup.DoQuery(String.Format("Select ""UgpEntry"" from ""OUGP"" where ""UgpCode"" = '{0}'", UOMGRPEntry))
                                                                                    If oRGroup.RecordCount = 1 Then
                                                                                        oTargetItemGroup.DefaultUoMGroup = oRGroup.Fields.Item(0).Value
                                                                                    End If

                                                                                    Dim DefUOM As Integer = oItemGroups.DefaultInventoryUoM
                                                                                    Dim UOMEntry As String = String.Empty
                                                                                    If DefUOM <> 0 Then
                                                                                        Dim Sgrp As String = "Select ""UomCode"" from OUOM where ""UomEntry"" =  '" & DefUOM & "'"
                                                                                        Dim Rgrp As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                        Rgrp.DoQuery(Sgrp)
                                                                                        If Rgrp.RecordCount > 0 Then UOMEntry = Rgrp.Fields.Item(0).Value
                                                                                    End If
                                                                                    Dim oRGroup1 As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                    oRGroup1.DoQuery(String.Format("Select ""UomEntry"" from ""OUOM"" where ""UomCode"" = '{0}'", UOMEntry))
                                                                                    If oRGroup1.RecordCount = 1 Then
                                                                                        oTargetItemGroup.DefaultInventoryUoM = oRGroup1.Fields.Item(0).Value
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                            If flg1 = True Then
                                                                                ErrerCode = oTargetItemGroup.Update()
                                                                                If ErrerCode <> 0 Then
                                                                                    oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                    p_oSBOApplication.StatusBar.SetText("Updating Item Group '" & oDT_Entities.Rows(J).Item("Name") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating BP Group '" & oDT_Entities.Rows(J).Item("Name") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                    SucFlag = False
                                                                                Else
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully Updated Item Group '" & oDT_Entities.Rows(J).Item("Name") & "' on '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                    p_oSBOApplication.StatusBar.SetText("Item Group '" & oDT_Entities.Rows(J).Item("Name") & "' updated Successfully on  '" & oDICompany(S).CompanyDB & "' ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                    SucFlag = True
                                                                                End If
                                                                            Else
                                                                                ErrerCode = oTargetItemGroup.Add
                                                                                If ErrerCode <> 0 Then
                                                                                    oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                    p_oSBOApplication.StatusBar.SetText("Adding Item Group '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating BP Group '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                    SucFlag = False
                                                                                Else
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully added Item Group '" & oDT_Entities.Rows(J).Item("Code") & "' on '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                    p_oSBOApplication.StatusBar.SetText("Item Group '" & oDT_Entities.Rows(J).Item("Name") & "' Created Successfully on  '" & oDICompany(S).CompanyDB & "' ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                    SucFlag = True
                                                                                End If
                                                                            End If

                                                                            If SucFlag = False Then
110:
                                                                                sErrDesc = sErrDesc + " On.." + oDICompany(S).CompanyDB
                                                                                p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                Fllag = False
                                                                                Dim sqy As String = " UPDATE ""INTEGRATION"" SET ""SYNCSTATUS"" = 'NO', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '" & sErrDesc.ToString.Replace("'", """") & "'  WHERE ""MASTERTYPE"" = 'ITEMMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                                Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                Run.DoQuery(sqy)

                                                                                Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                                oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "NO"
                                                                                oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                                oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = sErrDesc

                                                                                For lCounter As Integer = 0 To UBound(oDICompany)
                                                                                    If Not oDICompany(lCounter) Is Nothing Then
                                                                                        If oDICompany(lCounter).Connected = True Then
                                                                                            If oDICompany(lCounter).InTransaction = True Then
                                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                                oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                                            End If
                                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                            oDICompany(lCounter).Disconnect()
                                                                                            oDICompany(lCounter) = Nothing
                                                                                        End If
                                                                                    End If
                                                                                Next
                                                                                Exit For
                                                                            Else
                                                                                Fllag = True
                                                                                SBO_Application.SetStatusBarMessage("Item Group '" & oDT_Entities.Rows(J).Item("Name") & "' Replicated Successfully on.. " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Item Groups '" & oDT_Entities.Rows(J).Item("Name") & "' Replicated Successfully on .. " & oDV_BPSetup.Item(S).Item("Name").ToString, sFuncName)
                                                                            End If

                                                                            Try
                                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItemGroups)
                                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetItemGroup)
                                                                            Catch ex As Exception
                                                                            End Try
                                                                        Next
                                                                        For lCounter As Integer = 0 To UBound(oDICompany)
                                                                            If Not oDICompany(lCounter) Is Nothing Then
                                                                                If oDICompany(lCounter).Connected = True Then
                                                                                    If oDICompany(lCounter).InTransaction = True Then
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                        oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                                                    End If
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                    oDICompany(lCounter).Disconnect()
                                                                                    oDICompany(lCounter) = Nothing
                                                                                End If
                                                                            End If
                                                                        Next
                                                                        If Fllag = True Then
                                                                            Dim sqy As String = " Update ""INTEGRATION"" SET ""SYNCSTATUS"" = 'YES', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '' WHERE ""MASTERTYPE"" = 'ITEMMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                            Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                            Run.DoQuery(sqy)

                                                                            Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                            oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "YES"
                                                                            oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                            oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = "SUCCESS"
                                                                        End If
                                                                    End If
                                                                ElseIf (oDT_Entities.Rows(J).Item("TransType").ToString = "UOMGroups") Then
                                                                    ''=========================================================================================================
                                                                    '------------------------------- UOM Groups Replication ----------------------------------------------
                                                                    ''=========================================================================================================
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Item Setup for UOM Group Settings", sFuncName)
                                                                    Dim rvcount As Integer = oDV_BPSetup.Count
                                                                    Fllag = False
                                                                    oDV_BPSetup.RowFilter = "U_UOMGroups ='Y'"
                                                                    ReDim oDICompany(oDV_BPSetup.Count)
                                                                    Dim dvcount1 As Integer = oDV_BPSetup.Count
                                                                    If oDV_BPSetup.Count > 0 Then
                                                                        For S As Integer = 0 To oDV_BPSetup.Count - 1
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Connect To TargetCompany() for UOM Groups Posting", sFuncName)
                                                                            SBO_Application.StatusBar.SetText("Connecting to the Target Company" & oDICompany(S).CompanyDB, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                            Dim targetent As String = oDV_BPSetup.Item(S).Item("Name").ToString
                                                                            If ConnectToTargetCompany(oDICompany(S), oDV_BPSetup.Item(S).Item("Name").ToString, oDV_BPSetup.Item(S).Item("U_UserName").ToString, oDV_BPSetup.Item(S).Item("U_Password").ToString, sErrDesc) <> RTN_SUCCESS Then
                                                                                GoTo 111
                                                                            End If
                                                                            SBO_Application.SetStatusBarMessage("Connecting to the target company Successfull on  " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("onnecting to the target company Successfull on " & oDICompany(S).CompanyDB, sFuncName)
                                                                            oDICompany(S).StartTransaction()

                                                                            SBO_Application.SetStatusBarMessage("Started Master Data Synchronization " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                                            Dim oUOMGroupEntry As String = String.Empty

                                                                            Dim svrUOMGroups As SAPbobsCOM.UnitOfMeasurementGroupsService = p_oDICompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.UnitOfMeasurementGroupsService)
                                                                            Dim oUOMGroups As SAPbobsCOM.UnitOfMeasurementGroup = svrUOMGroups.GetDataInterface(SAPbobsCOM.UnitOfMeasurementGroupsServiceDataInterfaces.uomgsUnitOfMeasurementGroup)

                                                                            Dim TargetsvrUOMGroups As SAPbobsCOM.UnitOfMeasurementGroupsService = oDICompany(S).GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.UnitOfMeasurementGroupsService)
                                                                            Dim oTargetUOMGroups As SAPbobsCOM.UnitOfMeasurementGroup = TargetsvrUOMGroups.GetDataInterface(SAPbobsCOM.UnitOfMeasurementGroupsServiceDataInterfaces.uomgsUnitOfMeasurementGroup)

                                                                            Dim oUpdateUOMGroupParams As SAPbobsCOM.UnitOfMeasurementGroupParams = TargetsvrUOMGroups.GetDataInterface(SAPbobsCOM.UnitOfMeasurementGroupsServiceDataInterfaces.uomgsUnitOfMeasurementGroupParams)
                                                                            Dim oUpdateTargetUOMGroup As SAPbobsCOM.UnitOfMeasurementGroup = TargetsvrUOMGroups.GetDataInterface(SAPbobsCOM.UnitOfMeasurementGroupsServiceDataInterfaces.uomgsUnitOfMeasurementGroup)

                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                                                                            Dim flg1 As Boolean = False
                                                                            Dim SucFlag As Boolean = False
                                                                            Dim groupno As String = String.Empty

                                                                            Dim orsGroup As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                            orsGroup.DoQuery(String.Format("Select ""UgpCode"" from ""OUGP"" where ""UgpCode"" = '{0}'", oDT_Entities.Rows(J).Item("Code").ToString))
                                                                            If orsGroup.RecordCount = 1 Then
                                                                                flg1 = True
                                                                                groupno = orsGroup.Fields.Item(0).Value
                                                                            End If

                                                                            Dim SSQRY As String = "Select ""UgpCode"",""UgpName"", ""BaseUom"" from OUGP where ""UgpCode"" = '" & oDT_Entities.Rows(J).Item("Code").ToString & "'"
                                                                            Dim RsetUOMGrp As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                            RsetUOMGrp.DoQuery(SSQRY)
                                                                            If RsetUOMGrp.RecordCount > 0 Then
                                                                                If flg1 = False Then
                                                                                    oTargetUOMGroups.Code = RsetUOMGrp.Fields.Item("UgpCode").Value
                                                                                    oTargetUOMGroups.Name = RsetUOMGrp.Fields.Item("UgpName").Value
                                                                                    oTargetUOMGroups.BaseUoM = RsetUOMGrp.Fields.Item("BaseUom").Value
                                                                                Else
                                                                                    oUpdateUOMGroupParams.Code = groupno
                                                                                    Try
                                                                                        oUpdateTargetUOMGroup = TargetsvrUOMGroups.Get(oUpdateUOMGroupParams)
                                                                                    Catch ex As Exception
                                                                                        SucFlag = False
                                                                                    End Try
                                                                                    oUpdateTargetUOMGroup.Name = RsetUOMGrp.Fields.Item("UgpName").Value
                                                                                    oUpdateTargetUOMGroup.BaseUoM = RsetUOMGrp.Fields.Item("BaseUom").Value
                                                                                End If
                                                                            End If
                                                                            Dim oUOMGroupParams As SAPbobsCOM.UnitOfMeasurementGroupParams = TargetsvrUOMGroups.GetDataInterface(SAPbobsCOM.UnitOfMeasurementGroupsServiceDataInterfaces.uomgsUnitOfMeasurementGroupParams)

                                                                            If flg1 = False Then
                                                                                Try
                                                                                    oUOMGroupParams = TargetsvrUOMGroups.Add(oTargetUOMGroups)
                                                                                    oUOMGroupEntry = oUOMGroupParams.Code
                                                                                    SucFlag = True
                                                                                    p_oSBOApplication.StatusBar.SetText("Adding UOM Groups '" & oDT_Entities.Rows(J).Item("Code").ToString & "' Successful on.." & oDICompany(S).CompanyDB, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding UOM Group '" & oDT_Entities.Rows(J).Item("Code") & "' Successful on  '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                Catch ex As Exception
                                                                                    sErrDesc = ex.Message
                                                                                    SucFlag = False
                                                                                    p_oSBOApplication.StatusBar.SetText("Adding UOM Groups '" & oDT_Entities.Rows(J).Item("Code").ToString & "' Failed.." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding UOM Group '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                Finally
                                                                                End Try
                                                                            Else
                                                                                Try
                                                                                    TargetsvrUOMGroups.Update(oUpdateTargetUOMGroup)
                                                                                    SucFlag = True
                                                                                    p_oSBOApplication.StatusBar.SetText("Updating UOM Groups '" & oDT_Entities.Rows(J).Item("Code").ToString & "' Successful on.." & oDICompany(S).CompanyDB, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating UOM Group '" & oDT_Entities.Rows(J).Item("Code") & "' Successful on  '" & oDICompany(S).CompanyDB & "' ", sFuncName)
                                                                                Catch ex As Exception
                                                                                    sErrDesc = ex.Message
                                                                                    SucFlag = False
                                                                                    p_oSBOApplication.StatusBar.SetText("Updating UOM GRoups '" & oDT_Entities.Rows(J).Item("Code").ToString & "' Failed.." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating UOM Group '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                Finally
                                                                                End Try
                                                                            End If
                                                                            If SucFlag = False Then
111:
                                                                                sErrDesc = sErrDesc + " On.." + oDICompany(S).CompanyDB
                                                                                p_oSBOApplication.StatusBar.SetText("UOM GROUP Replication Failed" & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                Fllag = False
                                                                                Dim sqy As String = " UPDATE ""INTEGRATION"" SET ""SYNCSTATUS"" = 'NO', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '" & sErrDesc.ToString.Replace("'", """") & "'  WHERE ""MASTERTYPE"" = 'ITEMMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                                Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                Run.DoQuery(sqy)

                                                                                Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                                oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "NO"
                                                                                oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                                oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = sErrDesc

                                                                                For lCounter As Integer = 0 To UBound(oDICompany)
                                                                                    If Not oDICompany(lCounter) Is Nothing Then
                                                                                        If oDICompany(lCounter).Connected = True Then
                                                                                            If oDICompany(lCounter).InTransaction = True Then
                                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                                oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                                            End If
                                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                            oDICompany(lCounter).Disconnect()
                                                                                            oDICompany(lCounter) = Nothing
                                                                                        End If
                                                                                    End If
                                                                                Next
                                                                                Exit For
                                                                            Else
                                                                                Fllag = True
                                                                                SBO_Application.SetStatusBarMessage("UOM Group '" & oDT_Entities.Rows(J).Item("Code").ToString & "' Replicated Successfully on.. " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("UOM Groups '" & oDT_Entities.Rows(J).Item("Code").ToString & "' Replicated Successfully on .. " & oDV_BPSetup.Item(S).Item("Name").ToString, sFuncName)
                                                                            End If

                                                                            Try
                                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetUOMGroups)
                                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUpdateTargetUOMGroup)
                                                                            Catch ex As Exception
                                                                            End Try
                                                                        Next
                                                                        For lCounter As Integer = 0 To UBound(oDICompany)
                                                                            If Not oDICompany(lCounter) Is Nothing Then
                                                                                If oDICompany(lCounter).Connected = True Then
                                                                                    If oDICompany(lCounter).InTransaction = True Then
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                        oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                                                    End If
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                    oDICompany(lCounter).Disconnect()
                                                                                    oDICompany(lCounter) = Nothing
                                                                                End If
                                                                            End If
                                                                        Next
                                                                        If Fllag = True Then
                                                                            Dim sqy As String = " Update ""INTEGRATION"" SET ""SYNCSTATUS"" = 'YES', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '' WHERE ""MASTERTYPE"" = 'ITEMMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                            Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                            Run.DoQuery(sqy)

                                                                            Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                            oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "YES"
                                                                            oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                            oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = "SUCCESS"
                                                                        End If
                                                                    End If
                                                                ElseIf (oDT_Entities.Rows(J).Item("TransType").ToString = "BOM") Then
                                                                    ''=========================================================================================================
                                                                    '------------------------------- Bill of Material Replication ----------------------------------------------
                                                                    ''=========================================================================================================
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Item Master Setup for BOM Setting", sFuncName)
                                                                    Dim rvcount As Integer = oDV_BPSetup.Count
                                                                    Fllag = False
                                                                    oDV_BPSetup.RowFilter = "U_BOM ='Y'"
                                                                    ReDim oDICompany(oDV_BPSetup.Count)
                                                                    Dim dvcount1 As Integer = oDV_BPSetup.Count
                                                                    If oDV_BPSetup.Count > 0 Then
                                                                        For S As Integer = 0 To oDV_BPSetup.Count - 1
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Connect To TargetCompany for BOM Posting", sFuncName)
                                                                            SBO_Application.SetStatusBarMessage("Connecting to the target company On" & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                            Dim targetent As String = oDV_BPSetup.Item(S).Item("Name").ToString
                                                                            If ConnectToTargetCompany(oDICompany(S), oDV_BPSetup.Item(S).Item("Name").ToString, oDV_BPSetup.Item(S).Item("U_UserName").ToString, oDV_BPSetup.Item(S).Item("U_Password").ToString, sErrDesc) <> RTN_SUCCESS Then
                                                                                GoTo 112
                                                                            End If
                                                                            SBO_Application.SetStatusBarMessage("Connecting to the target company Successfull On" & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the target company Successfull On " & oDICompany(S).CompanyDB, sFuncName)
                                                                            oDICompany(S).StartTransaction()

                                                                            SBO_Application.SetStatusBarMessage("Started Master Data Synchronization " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                                            Dim oBOM As SAPbobsCOM.ProductTrees = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees)
                                                                            Dim oTargetBOM As SAPbobsCOM.ProductTrees = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees)
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                                                                            Dim flg1 As Boolean = False
                                                                            Dim SucFlag As Boolean = False
                                                                            Dim groupno As String = String.Empty

                                                                            If oTargetBOM.GetByKey(oDT_Entities.Rows(J).Item("Code")) = True Then
                                                                                flg1 = True
                                                                            End If

                                                                            If oBOM.GetByKey(oDT_Entities.Rows(J).Item("Code")) Then
                                                                                If flg1 = True Then
                                                                                    If oTargetBOM.GetByKey(oBOM.TreeCode) Then
                                                                                        oTargetBOM.Quantity = oBOM.Quantity
                                                                                        oTargetBOM.TreeType = oBOM.TreeType
                                                                                        oTargetBOM.Warehouse = oBOM.Warehouse
                                                                                        oTargetBOM.Project = oBOM.Project

                                                                                        Dim PList As Integer = oBOM.PriceList
                                                                                        Dim PListName As String = String.Empty
                                                                                        If PList <> 0 Then
                                                                                            Dim Sgrp As String = "Select ""ListName"" from OPLN where ""ListNum"" =  '" & PList & "'"
                                                                                            Dim Rgrp As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                            Rgrp.DoQuery(Sgrp)
                                                                                            If Rgrp.RecordCount > 0 Then PListName = Rgrp.Fields.Item(0).Value
                                                                                        End If
                                                                                        Dim oRGroup1 As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                        oRGroup1.DoQuery(String.Format("Select ""ListNum"" from ""OPLN"" where ""ListName"" = '{0}'", PListName))
                                                                                        If oRGroup1.RecordCount = 1 Then
                                                                                            oTargetBOM.PriceList = oRGroup1.Fields.Item(0).Value
                                                                                        End If

                                                                                        If oTargetBOM.Items.Count > 0 Then

                                                                                            Dim delete As Boolean = False
                                                                                            For i As Integer = 0 To oTargetBOM.Items.Count - 1
                                                                                                oTargetBOM.Items.SetCurrentLine(oTargetBOM.Items.Count - 1)
                                                                                                oTargetBOM.Items.Delete()
                                                                                                If oTargetBOM.Items.Count = 0 Then
                                                                                                    Exit For
                                                                                                End If
                                                                                            Next
                                                                                        End If
                                                                                        If oBOM.Items.Count > 0 Then
                                                                                            For i As Integer = 0 To oBOM.Items.Count - 1
                                                                                                oBOM.Items.SetCurrentLine(i)
                                                                                                oTargetBOM.Items.ItemType = oBOM.Items.ItemType
                                                                                                If oBOM.Items.ItemType <> SAPbobsCOM.ProductionItemType.pit_Text Then
                                                                                                    oTargetBOM.Items.ItemCode = oBOM.Items.ItemCode
                                                                                                    oTargetBOM.Items.Quantity = oBOM.Items.Quantity
                                                                                                    oTargetBOM.Items.Project = oBOM.Items.Project
                                                                                                    oTargetBOM.Items.IssueMethod = oBOM.Items.IssueMethod
                                                                                                    oTargetBOM.Items.Comment = oBOM.Items.Comment
                                                                                                    Dim PList1 As Integer = oBOM.Items.PriceList
                                                                                                    Dim PListName1 As String = String.Empty
                                                                                                    If PList1 <> 0 Then
                                                                                                        Dim Sgrp As String = "Select ""ListName"" from OPLN where ""ListNum"" =  '" & PList1 & "'"
                                                                                                        Dim Rgrp As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                                        Rgrp.DoQuery(Sgrp)
                                                                                                        If Rgrp.RecordCount > 0 Then PListName1 = Rgrp.Fields.Item(0).Value
                                                                                                    End If
                                                                                                    Dim oRGroup2 As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                                    oRGroup2.DoQuery(String.Format("Select ""ListNum"" from ""OPLN"" where ""ListName"" = '{0}'", PListName1))
                                                                                                    If oRGroup2.RecordCount = 1 Then
                                                                                                        oTargetBOM.Items.PriceList = oRGroup2.Fields.Item(0).Value
                                                                                                    End If
                                                                                                Else
                                                                                                    oTargetBOM.Items.LineText = oBOM.Items.LineText
                                                                                                End If

                                                                                                oTargetBOM.Items.Add()
                                                                                            Next
                                                                                        End If
                                                                                    End If
                                                                                Else
                                                                                    oTargetBOM.TreeCode = oBOM.TreeCode
                                                                                    oTargetBOM.Quantity = oBOM.Quantity
                                                                                    oTargetBOM.TreeType = oBOM.TreeType
                                                                                    oTargetBOM.Warehouse = oBOM.Warehouse
                                                                                    oTargetBOM.Project = oBOM.Project

                                                                                    Dim PList As Integer = oBOM.PriceList
                                                                                    Dim PListName As String = String.Empty
                                                                                    If PList <> 0 Then
                                                                                        Dim Sgrp As String = "Select ""ListName"" from OPLN where ""ListNum"" =  '" & PList & "'"
                                                                                        Dim Rgrp As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                        Rgrp.DoQuery(Sgrp)
                                                                                        If Rgrp.RecordCount > 0 Then PListName = Rgrp.Fields.Item(0).Value
                                                                                    End If
                                                                                    Dim oRGroup1 As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                    oRGroup1.DoQuery(String.Format("Select ""ListNum"" from ""OPLN"" where ""ListName"" = '{0}'", PListName))
                                                                                    If oRGroup1.RecordCount = 1 Then
                                                                                        oTargetBOM.PriceList = oRGroup1.Fields.Item(0).Value
                                                                                    End If

                                                                                    If oBOM.Items.Count > 0 Then
                                                                                        For i As Integer = 0 To oBOM.Items.Count - 1
                                                                                            oBOM.Items.SetCurrentLine(i)
                                                                                            oTargetBOM.Items.ItemType = oBOM.Items.ItemType
                                                                                            oTargetBOM.Items.ItemType = oBOM.Items.ItemType
                                                                                            If oBOM.Items.ItemType <> SAPbobsCOM.ProductionItemType.pit_Text Then
                                                                                                oTargetBOM.Items.ItemCode = oBOM.Items.ItemCode
                                                                                                oTargetBOM.Items.Quantity = oBOM.Items.Quantity
                                                                                                oTargetBOM.Items.Project = oBOM.Items.Project
                                                                                                oTargetBOM.Items.IssueMethod = oBOM.Items.IssueMethod
                                                                                                oTargetBOM.Items.Comment = oBOM.Items.Comment
                                                                                                Dim PList1 As Integer = oBOM.Items.PriceList
                                                                                                Dim PListName1 As String = String.Empty
                                                                                                If PList1 <> 0 Then
                                                                                                    Dim Sgrp As String = "Select ""ListName"" from OPLN where ""ListNum"" =  '" & PList1 & "'"
                                                                                                    Dim Rgrp As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                                    Rgrp.DoQuery(Sgrp)
                                                                                                    If Rgrp.RecordCount > 0 Then PListName1 = Rgrp.Fields.Item(0).Value
                                                                                                End If
                                                                                                Dim oRGroup2 As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                                oRGroup2.DoQuery(String.Format("Select ""ListNum"" from ""OPLN"" where ""ListName"" = '{0}'", PListName1))
                                                                                                If oRGroup2.RecordCount = 1 Then
                                                                                                    oTargetBOM.Items.PriceList = oRGroup2.Fields.Item(0).Value
                                                                                                End If
                                                                                            Else
                                                                                                oTargetBOM.Items.LineText = oBOM.Items.LineText
                                                                                            End If
                                                                                            oTargetBOM.Items.Add()
                                                                                        Next
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                            If flg1 = True Then
                                                                                ErrerCode = oTargetBOM.Update()
                                                                                If ErrerCode <> 0 Then
                                                                                    oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                    p_oSBOApplication.StatusBar.SetText("Updating BOM '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on '" & oDICompany(S).CompanyDB & "'.." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating BOM :'" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                    SucFlag = False
                                                                                Else
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully Updated BOM '" & oDT_Entities.Rows(J).Item("Code") & "' on '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                    SucFlag = True
                                                                                End If
                                                                            Else
                                                                                ErrerCode = oTargetBOM.Add()
                                                                                If ErrerCode <> 0 Then
                                                                                    oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                    p_oSBOApplication.StatusBar.SetText("Adding BOM '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on '" & oDICompany(S).CompanyDB & "'.." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding BOM :'" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                    SucFlag = False
                                                                                Else
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully Added BOM '" & oDT_Entities.Rows(J).Item("Code") & "' on '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                    SucFlag = True
                                                                                End If
                                                                            End If

                                                                            If SucFlag = False Then
112:
                                                                                sErrDesc = sErrDesc + " On.." + oDICompany(S).CompanyDB
                                                                                p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                Fllag = False
                                                                                Dim sqy As String = " UPDATE ""INTEGRATION"" SET ""SYNCSTATUS"" = 'NO', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '" & sErrDesc.ToString.Replace("'", """") & "'  WHERE ""MASTERTYPE"" = 'ITEMMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                                Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                Run.DoQuery(sqy)

                                                                                Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                                oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "NO"
                                                                                oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                                oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = sErrDesc

                                                                                For lCounter As Integer = 0 To UBound(oDICompany)
                                                                                    If Not oDICompany(lCounter) Is Nothing Then
                                                                                        If oDICompany(lCounter).Connected = True Then
                                                                                            If oDICompany(lCounter).InTransaction = True Then
                                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                                oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                                            End If
                                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                            oDICompany(lCounter).Disconnect()
                                                                                            oDICompany(lCounter) = Nothing
                                                                                        End If
                                                                                    End If
                                                                                Next
                                                                                Exit For
                                                                            Else
                                                                                Fllag = True
                                                                                SBO_Application.SetStatusBarMessage("BOM  '" & oDT_Entities.Rows(J).Item("Code") & "' Replicated Successfully on.. " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("BOM '" & oDT_Entities.Rows(J).Item("Code") & "' Replicated Successfully on .. " & oDV_BPSetup.Item(S).Item("Name").ToString, sFuncName)
                                                                            End If
                                                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBOM)
                                                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetBOM)
                                                                        Next
                                                                        For lCounter As Integer = 0 To UBound(oDICompany)
                                                                            If Not oDICompany(lCounter) Is Nothing Then
                                                                                If oDICompany(lCounter).Connected = True Then
                                                                                    If oDICompany(lCounter).InTransaction = True Then
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                        oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                                                    End If
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                    oDICompany(lCounter).Disconnect()
                                                                                    oDICompany(lCounter) = Nothing
                                                                                End If
                                                                            End If
                                                                        Next
                                                                        If Fllag = True Then
                                                                            Dim sqy As String = " Update ""INTEGRATION"" SET ""SYNCSTATUS"" = 'YES', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '' WHERE ""MASTERTYPE"" = 'ITEMMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                            Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                            Run.DoQuery(sqy)

                                                                            Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                            oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "YES"
                                                                            oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                            oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = "SUCCESS"
                                                                        End If
                                                                    End If
                                                                ElseIf (oDT_Entities.Rows(J).Item("TransType").ToString = "ItemCodes") Then
                                                                    ''=========================================================================================================
                                                                    '------------------------------- Item Master Data Replication ----------------------------------------------
                                                                    ''=========================================================================================================
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Item Master Setup for Item Master Setting", sFuncName)
                                                                    Dim rvcount As Integer = oDV_BPSetup.Count
                                                                    Fllag = False
                                                                    oDV_BPSetup.RowFilter = "U_ItemCodes ='Y'"
                                                                    ReDim oDICompany(oDV_BPSetup.Count)
                                                                    Dim dvcount1 As Integer = oDV_BPSetup.Count
                                                                    If oDV_BPSetup.Count > 0 Then
                                                                        For S As Integer = 0 To oDV_BPSetup.Count - 1
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Connect To Target Company() for Item Codes Posting", sFuncName)
                                                                            SBO_Application.SetStatusBarMessage("Connecting to the Target Company" & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                            Dim targetent As String = oDV_BPSetup.Item(S).Item("Name").ToString
                                                                            If ConnectToTargetCompany(oDICompany(S), oDV_BPSetup.Item(S).Item("Name").ToString, oDV_BPSetup.Item(S).Item("U_UserName").ToString, oDV_BPSetup.Item(S).Item("U_Password").ToString, sErrDesc) <> RTN_SUCCESS Then
                                                                                GoTo 113
                                                                            End If
                                                                            SBO_Application.SetStatusBarMessage("Connecting to the target company is Successful " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the target company is Successful " & oDICompany(S).CompanyDB, sFuncName)
                                                                            oDICompany(S).StartTransaction()

                                                                            SBO_Application.SetStatusBarMessage("Started Master Data Synchronization " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                                            Dim oItemMaster As SAPbobsCOM.Items = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                                                                            Dim oTargetItemMaster As SAPbobsCOM.Items = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                                                                            Dim flg1 As Boolean = False
                                                                            Dim SucFlag As Boolean = False
                                                                            Dim groupno As String = String.Empty

                                                                            If oTargetItemMaster.GetByKey(oDT_Entities.Rows(J).Item("Code")) = True Then
                                                                                flg1 = True
                                                                            End If
                                                                            If oItemMaster.GetByKey(oDT_Entities.Rows(J).Item("Code")) Then
                                                                                If flg1 = True Then
                                                                                    If oTargetItemMaster.GetByKey(oItemMaster.ItemCode) Then

                                                                                        oTargetItemMaster.ItemName = oItemMaster.ItemName
                                                                                        oTargetItemMaster.ItemType = oItemMaster.ItemType
                                                                                        oTargetItemMaster.ForeignName = oItemMaster.ForeignName
                                                                                        oTargetItemMaster.ItemsGroupCode = oItemMaster.ItemsGroupCode
                                                                                        oTargetItemMaster.InventoryItem = oItemMaster.InventoryItem
                                                                                        oTargetItemMaster.SalesItem = oItemMaster.SalesItem
                                                                                        oTargetItemMaster.PurchaseItem = oItemMaster.PurchaseItem
                                                                                        oTargetItemMaster.InventoryUOM = oItemMaster.InventoryUOM
                                                                                        oTargetItemMaster.GLMethod = oItemMaster.GLMethod
                                                                                        oTargetItemMaster.WTLiable = oItemMaster.WTLiable
                                                                                        oTargetItemMaster.PurchaseUnit = oItemMaster.PurchaseUnit

                                                                                        oTargetItemMaster.SupplierCatalogNo = oItemMaster.SupplierCatalogNo
                                                                                        oTargetItemMaster.Manufacturer = oItemMaster.Manufacturer
                                                                                        oTargetItemMaster.PurchaseUnit = oItemMaster.PurchaseUnit
                                                                                        oTargetItemMaster.PurchasePackagingUnit = oItemMaster.PurchasePackagingUnit
                                                                                        oTargetItemMaster.PurchaseQtyPerPackUnit = oItemMaster.PurchaseQtyPerPackUnit
                                                                                        oTargetItemMaster.PurchaseItemsPerUnit = oItemMaster.PurchaseItemsPerUnit

                                                                                        oTargetItemMaster.SalesUnit = oItemMaster.SalesUnit
                                                                                        oTargetItemMaster.SalesPackagingUnit = oItemMaster.SalesPackagingUnit
                                                                                        oTargetItemMaster.SalesQtyPerPackUnit = oItemMaster.SalesQtyPerPackUnit
                                                                                        oTargetItemMaster.SalesItemsPerUnit = oItemMaster.SalesItemsPerUnit

                                                                                        oTargetItemMaster.InventoryUoMEntry = oItemMaster.InventoryUoMEntry
                                                                                        oTargetItemMaster.InventoryUOM = oItemMaster.InventoryUOM
                                                                                        oTargetItemMaster.MaxInventory = oItemMaster.MaxInventory
                                                                                        oTargetItemMaster.MinInventory = oItemMaster.MinInventory
                                                                                        oTargetItemMaster.MinOrderQuantity = oItemMaster.MinOrderQuantity


                                                                                        'If oTargetItemMaster.WhsInfo.Count > 0 Then
                                                                                        '    Dim delete As Boolean = False
                                                                                        '    For i As Integer = 0 To oTargetItemMaster.WhsInfo.Count - 1
                                                                                        '        oTargetItemMaster.WhsInfo.SetCurrentLine(oTargetItemMaster.WhsInfo.Count - 1)
                                                                                        '        oTargetItemMaster.WhsInfo.Delete()
                                                                                        '        If oTargetItemMaster.WhsInfo.Count = 0 Then
                                                                                        '            Exit For
                                                                                        '        End If
                                                                                        '    Next
                                                                                        'End If

                                                                                        'For iLine As Integer = 0 To oItemMaster.WhsInfo.Count - 1
                                                                                        '    oItemMaster.WhsInfo.SetCurrentLine(iLine)
                                                                                        '    oTargetItemMaster.WhsInfo.WarehouseCode = oItemMaster.WhsInfo.WarehouseCode
                                                                                        '    oTargetItemMaster.WhsInfo.ExpensesAccount = oItemMaster.WhsInfo.ExpensesAccount
                                                                                        '    oTargetItemMaster.WhsInfo.ForeignExpensAcc = oItemMaster.WhsInfo.ForeignExpensAcc
                                                                                        '    oTargetItemMaster.WhsInfo.PurchaseCreditAcc = oItemMaster.WhsInfo.PurchaseCreditAcc
                                                                                        '    oTargetItemMaster.WhsInfo.ForeignPurchaseCreditAcc = oItemMaster.WhsInfo.ForeignPurchaseCreditAcc
                                                                                        '    oTargetItemMaster.WhsInfo.Add()
                                                                                        'Next
                                                                                        oTargetItemMaster.Employee = oItemMaster.Employee
                                                                                        oTargetItemMaster.Properties(1) = oItemMaster.Properties(1)
                                                                                        oTargetItemMaster.Properties(2) = oItemMaster.Properties(2)
                                                                                        oTargetItemMaster.Properties(3) = oItemMaster.Properties(3)
                                                                                        oTargetItemMaster.Properties(4) = oItemMaster.Properties(4)
                                                                                        oTargetItemMaster.Properties(5) = oItemMaster.Properties(5)
                                                                                        oTargetItemMaster.Properties(6) = oItemMaster.Properties(6)
                                                                                        oTargetItemMaster.Properties(7) = oItemMaster.Properties(7)
                                                                                        oTargetItemMaster.Properties(8) = oItemMaster.Properties(8)
                                                                                        oTargetItemMaster.Properties(9) = oItemMaster.Properties(9)
                                                                                        oTargetItemMaster.Properties(10) = oItemMaster.Properties(10)
                                                                                        oTargetItemMaster.Properties(11) = oItemMaster.Properties(11)
                                                                                        oTargetItemMaster.Properties(12) = oItemMaster.Properties(12)

                                                                                        oTargetItemMaster.User_Text = oItemMaster.User_Text
                                                                                        oTargetItemMaster.Frozen = SAPbobsCOM.BoYesNoEnum.tNO
                                                                                        oTargetItemMaster.Valid = SAPbobsCOM.BoYesNoEnum.tYES

                                                                                        oTargetItemMaster.FrozenFrom = oItemMaster.FrozenFrom
                                                                                        oTargetItemMaster.FrozenTo = oItemMaster.FrozenTo
                                                                                        oTargetItemMaster.ValidFrom = oItemMaster.ValidFrom
                                                                                        oTargetItemMaster.ValidTo = oItemMaster.ValidTo

                                                                                        oTargetItemMaster.PlanningSystem = oItemMaster.PlanningSystem
                                                                                        oTargetItemMaster.ProcurementMethod = oItemMaster.ProcurementMethod
                                                                                        'oTargetItemMaster.OrderIntervals = oItemMaster.OrderIntervals
                                                                                        oTargetItemMaster.OrderMultiple = oItemMaster.OrderMultiple
                                                                                        oTargetItemMaster.LeadTime = oItemMaster.LeadTime
                                                                                        oTargetItemMaster.ToleranceDays = oItemMaster.ToleranceDays
                                                                                        oTargetItemMaster.IssueMethod = oItemMaster.IssueMethod
                                                                                        oTargetItemMaster.ShipType = oItemMaster.ShipType
                                                                                        oTargetItemMaster.SWW = oItemMaster.SWW
                                                                                        oTargetItemMaster.CustomsGroupCode = oItemMaster.CustomsGroupCode
                                                                                        oTargetItemMaster.PurchaseVATGroup = oItemMaster.PurchaseVATGroup
                                                                                        oTargetItemMaster.SalesVATGroup = oItemMaster.SalesVATGroup
                                                                                        oTargetItemMaster.BarCode = oItemMaster.BarCode

                                                                                        If oTargetItemMaster.PreferredVendors.Count > 0 Then

                                                                                            Dim delete As Boolean = False
                                                                                            For i As Integer = 0 To oTargetItemMaster.PreferredVendors.Count - 1
                                                                                                oTargetItemMaster.PreferredVendors.SetCurrentLine(oTargetItemMaster.PreferredVendors.Count - 1)
                                                                                                oTargetItemMaster.PreferredVendors.Delete()
                                                                                                If oTargetItemMaster.PreferredVendors.Count = 0 Then
                                                                                                    Exit For
                                                                                                End If
                                                                                            Next
                                                                                        End If
                                                                                        If oItemMaster.PreferredVendors.Count > 0 Then
                                                                                            For i As Integer = 0 To oItemMaster.PreferredVendors.Count - 1
                                                                                                oItemMaster.PreferredVendors.SetCurrentLine(i)
                                                                                                Dim bpcode As String = oItemMaster.PreferredVendors.BPCode
                                                                                                oTargetItemMaster.PreferredVendors.BPCode = oItemMaster.PreferredVendors.BPCode
                                                                                                oTargetItemMaster.PreferredVendors.Add()
                                                                                            Next
                                                                                        End If
                                                                                    End If
                                                                                Else
                                                                                    oTargetItemMaster.ItemCode = oItemMaster.ItemCode
                                                                                    oTargetItemMaster.ItemName = oItemMaster.ItemName
                                                                                    oTargetItemMaster.ItemType = oItemMaster.ItemType
                                                                                    oTargetItemMaster.ForeignName = oItemMaster.ForeignName
                                                                                    oTargetItemMaster.ItemsGroupCode = oItemMaster.ItemsGroupCode
                                                                                    oTargetItemMaster.InventoryItem = oItemMaster.InventoryItem
                                                                                    oTargetItemMaster.SalesItem = oItemMaster.SalesItem
                                                                                    oTargetItemMaster.PurchaseItem = oItemMaster.PurchaseItem
                                                                                    oTargetItemMaster.InventoryUOM = oItemMaster.InventoryUOM
                                                                                    oTargetItemMaster.GLMethod = oItemMaster.GLMethod
                                                                                    oTargetItemMaster.WTLiable = oItemMaster.WTLiable
                                                                                    oTargetItemMaster.PurchaseUnit = oItemMaster.PurchaseUnit

                                                                                    oTargetItemMaster.SupplierCatalogNo = oItemMaster.SupplierCatalogNo
                                                                                    oTargetItemMaster.Manufacturer = oItemMaster.Manufacturer
                                                                                    oTargetItemMaster.PurchaseUnit = oItemMaster.PurchaseUnit
                                                                                    oTargetItemMaster.PurchasePackagingUnit = oItemMaster.PurchasePackagingUnit
                                                                                    oTargetItemMaster.PurchaseQtyPerPackUnit = oItemMaster.PurchaseQtyPerPackUnit
                                                                                    oTargetItemMaster.PurchaseItemsPerUnit = oItemMaster.PurchaseItemsPerUnit

                                                                                    oTargetItemMaster.SalesUnit = oItemMaster.SalesUnit
                                                                                    oTargetItemMaster.SalesPackagingUnit = oItemMaster.SalesPackagingUnit
                                                                                    oTargetItemMaster.SalesQtyPerPackUnit = oItemMaster.SalesQtyPerPackUnit
                                                                                    oTargetItemMaster.SalesItemsPerUnit = oItemMaster.SalesItemsPerUnit

                                                                                    oTargetItemMaster.InventoryUoMEntry = oItemMaster.InventoryUoMEntry
                                                                                    oTargetItemMaster.InventoryUOM = oItemMaster.InventoryUOM
                                                                                    oTargetItemMaster.MaxInventory = oItemMaster.MaxInventory
                                                                                    oTargetItemMaster.MinInventory = oItemMaster.MinInventory
                                                                                    oTargetItemMaster.MinOrderQuantity = oItemMaster.MinOrderQuantity

                                                                                    'For iLine As Integer = 0 To oItemMaster.WhsInfo.Count - 1
                                                                                    '    oItemMaster.WhsInfo.SetCurrentLine(iLine)
                                                                                    '    oTargetItemMaster.WhsInfo.WarehouseCode = oItemMaster.WhsInfo.WarehouseCode
                                                                                    '    oTargetItemMaster.WhsInfo.ExpensesAccount = oItemMaster.WhsInfo.ExpensesAccount
                                                                                    '    oTargetItemMaster.WhsInfo.ForeignExpensAcc = oItemMaster.WhsInfo.ForeignExpensAcc
                                                                                    '    oTargetItemMaster.WhsInfo.PurchaseCreditAcc = oItemMaster.WhsInfo.PurchaseCreditAcc
                                                                                    '    oTargetItemMaster.WhsInfo.ForeignPurchaseCreditAcc = oItemMaster.WhsInfo.ForeignPurchaseCreditAcc
                                                                                    '    oTargetItemMaster.WhsInfo.Add()
                                                                                    'Next
                                                                                    oTargetItemMaster.Employee = oItemMaster.Employee
                                                                                    oTargetItemMaster.Properties(1) = oItemMaster.Properties(1)
                                                                                    oTargetItemMaster.Properties(2) = oItemMaster.Properties(2)
                                                                                    oTargetItemMaster.Properties(3) = oItemMaster.Properties(3)
                                                                                    oTargetItemMaster.Properties(4) = oItemMaster.Properties(4)
                                                                                    oTargetItemMaster.Properties(5) = oItemMaster.Properties(5)
                                                                                    oTargetItemMaster.Properties(6) = oItemMaster.Properties(6)
                                                                                    oTargetItemMaster.Properties(7) = oItemMaster.Properties(7)
                                                                                    oTargetItemMaster.Properties(8) = oItemMaster.Properties(8)
                                                                                    oTargetItemMaster.Properties(9) = oItemMaster.Properties(9)
                                                                                    oTargetItemMaster.Properties(10) = oItemMaster.Properties(10)
                                                                                    oTargetItemMaster.Properties(11) = oItemMaster.Properties(11)
                                                                                    oTargetItemMaster.Properties(12) = oItemMaster.Properties(12)

                                                                                    oTargetItemMaster.User_Text = oItemMaster.User_Text
                                                                                    oTargetItemMaster.Frozen = SAPbobsCOM.BoYesNoEnum.tNO
                                                                                    oTargetItemMaster.Valid = SAPbobsCOM.BoYesNoEnum.tYES

                                                                                    oTargetItemMaster.FrozenFrom = oItemMaster.FrozenFrom
                                                                                    oTargetItemMaster.FrozenTo = oItemMaster.FrozenTo
                                                                                    oTargetItemMaster.ValidFrom = oItemMaster.ValidFrom
                                                                                    oTargetItemMaster.ValidTo = oItemMaster.ValidTo

                                                                                    oTargetItemMaster.PlanningSystem = oItemMaster.PlanningSystem
                                                                                    oTargetItemMaster.ProcurementMethod = oItemMaster.ProcurementMethod
                                                                                    'oTargetItemMaster.OrderIntervals = oItemMaster.OrderIntervals
                                                                                    oTargetItemMaster.OrderMultiple = oItemMaster.OrderMultiple
                                                                                    oTargetItemMaster.LeadTime = oItemMaster.LeadTime
                                                                                    oTargetItemMaster.ToleranceDays = oItemMaster.ToleranceDays
                                                                                    oTargetItemMaster.IssueMethod = oItemMaster.IssueMethod
                                                                                    oTargetItemMaster.BarCode = oItemMaster.BarCode
                                                                                    oTargetItemMaster.ShipType = oItemMaster.ShipType
                                                                                    oTargetItemMaster.SWW = oItemMaster.SWW
                                                                                    oTargetItemMaster.CustomsGroupCode = oItemMaster.CustomsGroupCode
                                                                                    oTargetItemMaster.PurchaseVATGroup = oItemMaster.PurchaseVATGroup
                                                                                    oTargetItemMaster.SalesVATGroup = oItemMaster.SalesVATGroup

                                                                                    If oItemMaster.PreferredVendors.Count > 0 Then
                                                                                        For i As Integer = 0 To oItemMaster.PreferredVendors.Count - 1
                                                                                            oItemMaster.PreferredVendors.SetCurrentLine(i)
                                                                                            oTargetItemMaster.PreferredVendors.BPCode = oItemMaster.PreferredVendors.BPCode
                                                                                            oTargetItemMaster.PreferredVendors.Add()
                                                                                        Next
                                                                                    End If
                                                                                End If

                                                                            End If
                                                                            If flg1 = True Then
                                                                                ErrerCode = oTargetItemMaster.Update
                                                                                If ErrerCode <> 0 Then
                                                                                    oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                    p_oSBOApplication.StatusBar.SetText("Updating Item Code '" & oDT_Entities.Rows(J).Item("Code") & "' Failed.'" & oDICompany(S).CompanyDB & "'. " & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating Item Code '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                    SucFlag = False
                                                                                Else
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully updated Item Code '" & oDT_Entities.Rows(J).Item("Code") & "' on '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                    SucFlag = True
                                                                                End If
                                                                            Else
                                                                                ErrerCode = oTargetItemMaster.Add

                                                                                If ErrerCode <> 0 Then
                                                                                    oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                    p_oSBOApplication.StatusBar.SetText("Adding Item Code '" & oDT_Entities.Rows(J).Item("Code") & "' Failed.'" & oDICompany(S).CompanyDB & "'. " & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Item Code '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                    SucFlag = False
                                                                                Else
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully added Item Code '" & oDT_Entities.Rows(J).Item("Code") & "' on '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                    SucFlag = True
                                                                                End If
                                                                            End If

                                                                            If SucFlag = False Then
113:
                                                                                sErrDesc = sErrDesc + " On.." + oDICompany(S).CompanyDB
                                                                                p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                Fllag = False
                                                                                Dim sqy As String = " UPDATE ""INTEGRATION"" SET ""SYNCSTATUS"" = 'NO', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '" & sErrDesc.ToString.Replace("'", """") & "'  WHERE ""MASTERTYPE"" = 'ITEMMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                                Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                Run.DoQuery(sqy)

                                                                                Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                                oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "NO"
                                                                                oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                                oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = sErrDesc

                                                                                For lCounter As Integer = 0 To UBound(oDICompany)
                                                                                    If Not oDICompany(lCounter) Is Nothing Then
                                                                                        If oDICompany(lCounter).Connected = True Then
                                                                                            If oDICompany(lCounter).InTransaction = True Then
                                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                                oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                                            End If
                                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                            oDICompany(lCounter).Disconnect()
                                                                                            oDICompany(lCounter) = Nothing
                                                                                        End If
                                                                                    End If
                                                                                Next
                                                                                Exit For
                                                                            Else
                                                                                Fllag = True
                                                                                SBO_Application.SetStatusBarMessage("Item Code '" & oDT_Entities.Rows(J).Item("Code") & "' Replicated Successfully on.. " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Item Code '" & oDT_Entities.Rows(J).Item("Code") & "' is Replicated Successfully on .. " & oDV_BPSetup.Item(S).Item("Name").ToString, sFuncName)
                                                                            End If
                                                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oItemMaster)
                                                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetItemMaster)
                                                                        Next
                                                                        For lCounter As Integer = 0 To UBound(oDICompany)
                                                                            If Not oDICompany(lCounter) Is Nothing Then
                                                                                If oDICompany(lCounter).Connected = True Then
                                                                                    If oDICompany(lCounter).InTransaction = True Then
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                        oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                                                    End If
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                    oDICompany(lCounter).Disconnect()
                                                                                    oDICompany(lCounter) = Nothing
                                                                                End If
                                                                            End If
                                                                        Next
                                                                        If Fllag = True Then
                                                                            Dim sqy As String = " Update ""INTEGRATION"" SET ""SYNCSTATUS"" = 'YES', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '' WHERE ""MASTERTYPE"" = 'ITEMMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                            Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                            Run.DoQuery(sqy)

                                                                            Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                            oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "YES"
                                                                            oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                            oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = "SUCCESS"
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next
                                                        End If
                                                        Dim oPricelist As SAPbouiCOM.CheckBox = oForm.Items.Item("Chk_Prices").Specific
                                                        If (oPricelist.Checked = True And oDV_BPSetup.Count > 0) Then
                                                            ''=========================================================================================================
                                                            '------------------------------- Price List Replication ----------------------------------------------
                                                            ''=========================================================================================================
                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Item Master Setup for Price List Setting", sFuncName)
                                                            Dim rvcount As Integer = oDV_BPSetup.Count
                                                            Fllag = False
                                                            Dim CheckFlag As Boolean = False
                                                            Dim TestFlg As Boolean = False
                                                            Dim PricelistNo As Integer
                                                            oDV_BPSetup.RowFilter = "U_PriceLists ='Y'"
                                                            Dim SucFlag As Boolean = False

                                                            Dim sqry2 As String = " select T0.""ItemCode"",T0.""PriceList"",T1.""ListName"", T0.""Price"",T0.""Currency"",T0.""BasePLNum"", T0.""Ovrwritten"" from ITM1 T0  INNER JOIN OPLN T1 ON T0.""PriceList"" = T1.""ListNum""  where ""Price"" > 0 and T0.""Ovrwritten"" = 'Y' ORDER BY T0.""ItemCode"";"
                                                            Dim oRset2 As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                            oRset2.DoQuery(sqry2)
                                                            oDT_ItemPricelist = New DataTable
                                                            oDT_ItemPricelist = ConvertRecordset(oRset2)
                                                            Dim dtcount1 As Integer = oDT_ItemPricelist.Rows.Count
                                                            Dim oDV_ItemPricelist As New DataView(oDT_ItemPricelist)
                                                            Dim dvcount2 As Integer = oDV_ItemPricelist.Count

                                                            Dim sqry3 As String = "select distinct ""ItemCode"" from ITM1 where  ""Price"" > 0 and ""ItemCode"" not in ('R00001','L10001') ORDER BY ""ItemCode"";"
                                                            Dim RItemSets As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                            RItemSets.DoQuery(sqry3)
                                                            oDT_Itemlists = New DataTable
                                                            oDT_Itemlists = ConvertRecordset(RItemSets)
                                                            Dim dtcount12 As Integer = oDT_Itemlists.Rows.Count

                                                            Dim sqry1 As String = "select distinct T0.""ItemCode"",T0.""BasePLNum"", T2.""ListName"", T1.""Price""from ITM1 T0 INNER JOIN OPLN T2 ON T0.""BasePLNum"" = T2.""ListNum"" INNER JOIN ITM1 T1 ON T1.""PriceList"" = T0.""BasePLNum"" and T1.""ItemCode"" = T0.""ItemCode"" where T0.""Price"" > 0 and T1.""Price""> 0 ORDER BY T0.""ItemCode"",T0.""BasePLNum"";"
                                                            Dim RPLSets As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                            RPLSets.DoQuery(sqry1)
                                                            oDT_Pricelists = New DataTable
                                                            oDT_Pricelists = ConvertRecordset(RPLSets)
                                                            Dim dtcount3 As Integer = oDT_Pricelists.Rows.Count
                                                            Dim oDV_PriceLists As New DataView(oDT_Pricelists)
                                                            Dim dvcount3 As Integer = oDV_PriceLists.Count

                                                            ReDim oDICompany(oDV_BPSetup.Count)
                                                            Dim dvcount1 As Integer = oDV_BPSetup.Count
                                                            If oDV_BPSetup.Count > 0 Then
                                                                For S As Integer = 0 To oDV_BPSetup.Count - 1
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Connect To TargetCompany() for Price Lists Posting", sFuncName)
                                                                    SBO_Application.SetStatusBarMessage("Connecting to the target company  " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                    Dim targetent As String = oDV_BPSetup.Item(S).Item("Name").ToString
                                                                    If ConnectToTargetCompany(oDICompany(S), oDV_BPSetup.Item(S).Item("Name").ToString, oDV_BPSetup.Item(S).Item("U_UserName").ToString, oDV_BPSetup.Item(S).Item("U_Password").ToString, sErrDesc) <> RTN_SUCCESS Then
                                                                        GoTo 114
                                                                    End If
                                                                    SBO_Application.SetStatusBarMessage("Connecting to the target company Successful " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the target company Successfull  " & oDICompany(S).CompanyDB, sFuncName)

                                                                    SBO_Application.SetStatusBarMessage("Started Master Data Synchronization " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                                    If CreatePricelistMaster(oDICompany(S), sErrDesc) <> False Then
                                                                        If oDICompany(S).InTransaction = False Then oDICompany(S).StartTransaction()

                                                                        Dim oPriceLists As SAPbobsCOM.Items = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                                                                        Dim oTargetPriceLists As SAPbobsCOM.Items = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                                                                        If oDT_Itemlists.Rows.Count > 0 Then
                                                                            For I As Integer = 0 To oDT_Itemlists.Rows.Count - 1
                                                                                If oTargetPriceLists.GetByKey(oDT_Itemlists.Rows(I).Item("ItemCode").ToString) Then
                                                                                    Dim oitm As String = oDT_Itemlists.Rows(I).Item("ItemCode").ToString
                                                                                    oDV_PriceLists.RowFilter = "ItemCode ='" & oDT_Itemlists.Rows(I).Item("ItemCode").ToString & "'"
                                                                                    Dim Itemcount As Integer = oDV_PriceLists.Count
                                                                                    p_oSBOApplication.StatusBar.SetText("Price lists are updating for the item..." & oDT_Itemlists.Rows(I).Item("ItemCode").ToString, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                    If oDV_PriceLists.Count > 0 Then
                                                                                        TestFlg = False
                                                                                        'Dim listname As String = oDV_ItemPricelist.Item(I).Item("ListName").ToString
                                                                                        'CheckFlag = False
                                                                                        For T As Integer = 0 To oDV_PriceLists.Count - 1
                                                                                            CheckFlag = False
                                                                                            Dim listname As String = oDV_PriceLists.Item(T).Item("ListName").ToString
                                                                                            Dim oChecking As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                            oChecking.DoQuery(String.Format("Select ""ListNum"" from ""OPLN"" where ""ListName"" = '{0}'", oDV_PriceLists.Item(T).Item("ListName").ToString))
                                                                                            If oChecking.RecordCount = 1 Then
                                                                                                TestFlg = True
                                                                                                CheckFlag = True
                                                                                                PricelistNo = oChecking.Fields.Item(0).Value
                                                                                            End If
                                                                                            If CheckFlag = True Then
                                                                                                For M As Integer = 0 To oTargetPriceLists.PriceList.Count - 1
                                                                                                    oTargetPriceLists.PriceList.SetCurrentLine(M)
                                                                                                    If oTargetPriceLists.PriceList.PriceList = PricelistNo Then
                                                                                                        Dim Prc1 As Double = Convert.ToDouble(oDV_PriceLists.Item(T).Item("Price"))
                                                                                                        oTargetPriceLists.PriceList.Price = Prc1
                                                                                                        oTargetPriceLists.PriceList.Currency = oTargetPriceLists.PriceList.Currency
                                                                                                        Exit For
                                                                                                    End If
                                                                                                Next
                                                                                            End If
                                                                                        Next
                                                                                    End If

                                                                                    oDV_ItemPricelist.RowFilter = "ItemCode ='" & oDT_Itemlists.Rows(I).Item("ItemCode").ToString & "' and Ovrwritten = 'Y'"
                                                                                    Dim Itemcount2 As Integer = oDV_PriceLists.Count
                                                                                    CheckFlag = False
                                                                                    If oDV_ItemPricelist.Count > 0 Then
                                                                                        TestFlg = False
                                                                                        'Dim listname As String = oDV_ItemPricelist.Item(I).Item("ListName").ToString
                                                                                        For U As Integer = 0 To oDV_ItemPricelist.Count - 1
                                                                                            Dim oChecking As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                            oChecking.DoQuery(String.Format("Select ""ListNum"" from ""OPLN"" where ""ListName"" = '{0}'", oDV_ItemPricelist.Item(U).Item("ListName").ToString))
                                                                                            If oChecking.RecordCount = 1 Then
                                                                                                TestFlg = True
                                                                                                CheckFlag = True
                                                                                                PricelistNo = oChecking.Fields.Item(0).Value
                                                                                            End If
                                                                                            If CheckFlag = True Then
                                                                                                For M As Integer = 0 To oTargetPriceLists.PriceList.Count - 1
                                                                                                    oTargetPriceLists.PriceList.SetCurrentLine(M)
                                                                                                    If oTargetPriceLists.PriceList.PriceList = PricelistNo Then
                                                                                                        Dim Prc As Double = Convert.ToDouble(oDV_ItemPricelist.Item(U).Item("Price"))
                                                                                                        oTargetPriceLists.PriceList.Price = Prc
                                                                                                        Exit For
                                                                                                    End If
                                                                                                Next
                                                                                            End If
                                                                                        Next
                                                                                    End If

                                                                                    SucFlag = False
                                                                                    If TestFlg = True Then
                                                                                        ErrerCode = oTargetPriceLists.Update()
                                                                                        If ErrerCode <> 0 Then
                                                                                            oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                            p_oSBOApplication.StatusBar.SetText("Updating Price Lists to Target Company '" & oDICompany(S).CompanyDB & "' Failed.." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating Price Lists to Target Company '" & oDICompany(S).CompanyDB & "' Failed.." + " - " + sErrDesc, sFuncName)
                                                                                            SucFlag = False
                                                                                            GoTo 114
                                                                                        Else
                                                                                            p_oSBOApplication.StatusBar.SetText("Updating Price Lists to Target Company '" & oDICompany(S).CompanyDB & "' Successful..", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating Price Lists to Target Company '" & oDICompany(S).CompanyDB & "' Successful..", sFuncName)
                                                                                            SucFlag = True
                                                                                        End If
                                                                                    End If

                                                                                End If
                                                                            Next
                                                                        End If
                                                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oPriceLists)
                                                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetPriceLists)
                                                                    End If
                                                                    If SucFlag = False Then
114:
                                                                        sErrDesc = sErrDesc + " On.." + oDICompany(S).CompanyDB
                                                                        p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                        Fllag = False
                                                                        Dim oEdit As SAPbouiCOM.StaticText = oForm.Items.Item("l_RepStat").Specific
                                                                        oEdit.Caption = sErrDesc
                                                                        oEdit.Item.FontSize = 10
                                                                        oEdit.Item.ForeColor = RGB(255, 0, 0)
                                                                        For lCounter As Integer = 0 To UBound(oDICompany)
                                                                            If Not oDICompany(lCounter) Is Nothing Then
                                                                                If oDICompany(lCounter).Connected = True Then
                                                                                    If oDICompany(lCounter).InTransaction = True Then
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                        oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                                    End If
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                    oDICompany(lCounter).Disconnect()
                                                                                    oDICompany(lCounter) = Nothing
                                                                                End If
                                                                            End If
                                                                        Next
                                                                        Exit For
                                                                    Else
                                                                        Fllag = True
                                                                        SBO_Application.SetStatusBarMessage("Price Lists Replicated Successfully on.. " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Price Lists Replicated Successfully on .. " & oDV_BPSetup.Item(S).Item("Name").ToString, sFuncName)
                                                                    End If
                                                                Next

                                                                For lCounter As Integer = 0 To UBound(oDICompany)
                                                                    If Not oDICompany(lCounter) Is Nothing Then
                                                                        If oDICompany(lCounter).Connected = True Then
                                                                            If oDICompany(lCounter).InTransaction = True Then
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                                            End If
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                            oDICompany(lCounter).Disconnect()
                                                                            oDICompany(lCounter) = Nothing
                                                                        End If
                                                                    End If
                                                                Next
                                                                If Fllag = True Then
                                                                    Dim oEdit1 As SAPbouiCOM.StaticText = oForm.Items.Item("l_RepStat").Specific
                                                                    oEdit1.Caption = "Price Lists are Replicated Successfully on Target Databases.."
                                                                    oEdit1.Item.FontSize = 10
                                                                    oEdit1.Item.ForeColor = RGB(0, 128, 0)
                                                                End If
                                                            End If
                                                        End If
                                                        Dim oBincheck As SAPbouiCOM.CheckBox = oForm.Items.Item("Chk_bin").Specific
                                                        If (oBincheck.Checked = True And oDV_BPSetup.Count > 0) Then
                                                            ''=========================================================================================================
                                                            '------------------------------- Bin Location Replication ----------------------------------------------
                                                            ''=========================================================================================================
                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Item Master Setup for BIN Location Setting", sFuncName)
                                                            Dim rvcount As Integer = oDV_BPSetup.Count
                                                            Fllag = False
                                                            Dim CheckFlag As Boolean = False
                                                            Dim TestFlg As Boolean = False
                                                            'Dim PricelistNo As Integer
                                                            oDV_BPSetup.RowFilter = "U_BinLocatin ='Y'"
                                                            Dim SucFlag As Boolean = False

                                                            Dim sqry3 As String = " select ""AbsEntry"",""BinCode"",""WhsCode"",""SL1Code"", ""SL2Code"",""SL3Code"",""BarCode"",""SL4Code"",""RtrictType"",""ReceiveBin"",""MinLevel"",""MaxLevel"",""Disabled"", ""Descr"", ""AltSortCod"",""ItmRtrictT"", ""SpcItmCode"", ""SpcItmGrpC"",""RtrictResn"",""NoAutoAllc"",""SngBatch"" from OBIN where ""SysBin"" = 'N';"
                                                            Dim RBinSets As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                            RBinSets.DoQuery(sqry3)
                                                            oDT_Binlists = New DataTable
                                                            oDT_Binlists = ConvertRecordset(RBinSets)
                                                            Dim BinCount As Integer = oDT_Binlists.Rows.Count

                                                            ReDim oDICompany(oDV_BPSetup.Count)
                                                            'Dim dvcount1 As Integer = oDV_BPSetup.Count
                                                            If oDV_BPSetup.Count > 0 Then
                                                                For S As Integer = 0 To oDV_BPSetup.Count - 1
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Connect To Target Company() for Bin Location Posting", sFuncName)
                                                                    SBO_Application.SetStatusBarMessage("Connecting to the Target Company" & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                    Dim targetent As String = oDV_BPSetup.Item(S).Item("Name").ToString
                                                                    If ConnectToTargetCompany(oDICompany(S), oDV_BPSetup.Item(S).Item("Name").ToString, oDV_BPSetup.Item(S).Item("U_UserName").ToString, oDV_BPSetup.Item(S).Item("U_Password").ToString, sErrDesc) <> RTN_SUCCESS Then
                                                                        GoTo 115
                                                                    End If
                                                                    SBO_Application.SetStatusBarMessage("Connecting to the target company Successful " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the target company Successful " & oDICompany(S).CompanyDB, sFuncName)
                                                                    oDICompany(S).StartTransaction()

                                                                    SBO_Application.SetStatusBarMessage("Started Bin Location Synchronization on " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                                    Dim oBinLocationEntry As String = String.Empty
                                                                    Dim svrBinLocation As SAPbobsCOM.BinLocationsService = p_oDICompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.BinLocationsService)
                                                                    Dim oBinLocation As SAPbobsCOM.BinLocation = svrBinLocation.GetDataInterface(SAPbobsCOM.BinLocationsServiceDataInterfaces.blcsBinLocation)

                                                                    Dim TargetsvrBinLocation As SAPbobsCOM.BinLocationsService = oDICompany(S).GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.BinLocationsService)
                                                                    Dim oTargetBinLocation As SAPbobsCOM.BinLocation = TargetsvrBinLocation.GetDataInterface(SAPbobsCOM.BinLocationsServiceDataInterfaces.blcsBinLocation)

                                                                    If oDT_Binlists.Rows.Count > 0 Then
                                                                        For Y As Integer = 0 To oDT_Binlists.Rows.Count - 1

                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
                                                                            Dim flg1 As Boolean = False

                                                                            Dim bincode As String = String.Empty
                                                                            Dim BnCd As String = oDT_Binlists.Rows(Y).Item("BinCode").ToString
                                                                            Dim orsGroup As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                            orsGroup.DoQuery(String.Format("Select ""BinCode"" from ""OBIN"" where ""BinCode"" = '{0}'", oDT_Binlists.Rows(Y).Item("BinCode").ToString))
                                                                            If orsGroup.RecordCount = 1 Then
                                                                                flg1 = True
                                                                                bincode = orsGroup.Fields.Item(0).Value
                                                                            End If
                                                                            SucFlag = False

                                                                            If flg1 = False Then
                                                                                SBO_Application.StatusBar.SetText("Adding Bin Location '" & oDT_Binlists.Rows(Y).Item("BinCode").ToString & "' in.." & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                                                                                oTargetBinLocation.Warehouse = oDT_Binlists.Rows(S).Item("WhsCode").ToString
                                                                                If oDT_Binlists.Rows(Y).Item("SL1Code").ToString <> String.Empty Then oTargetBinLocation.Sublevel1 = oDT_Binlists.Rows(Y).Item("SL1Code").ToString
                                                                                If oDT_Binlists.Rows(Y).Item("SL2Code").ToString <> String.Empty Then oTargetBinLocation.Sublevel2 = oDT_Binlists.Rows(Y).Item("SL2Code").ToString
                                                                                If oDT_Binlists.Rows(Y).Item("SL3Code").ToString <> String.Empty Then oTargetBinLocation.Sublevel3 = oDT_Binlists.Rows(Y).Item("SL3Code").ToString
                                                                                If oDT_Binlists.Rows(Y).Item("SL4Code").ToString <> String.Empty Then oTargetBinLocation.Sublevel4 = oDT_Binlists.Rows(Y).Item("SL4Code").ToString
                                                                                If oDT_Binlists.Rows(Y).Item("ReceiveBin").ToString = "Y" Then
                                                                                    oTargetBinLocation.ReceivingBinLocation = SAPbobsCOM.BoYesNoEnum.tYES
                                                                                Else
                                                                                    oTargetBinLocation.ReceivingBinLocation = SAPbobsCOM.BoYesNoEnum.tNO
                                                                                End If

                                                                                If oDT_Binlists.Rows(Y).Item("NoAutoAllc").ToString = "Y" Then
                                                                                    oTargetBinLocation.ExcludeAutoAllocOnIssue = SAPbobsCOM.BoYesNoEnum.tYES
                                                                                Else
                                                                                    oTargetBinLocation.ExcludeAutoAllocOnIssue = SAPbobsCOM.BoYesNoEnum.tNO
                                                                                End If

                                                                                If oDT_Binlists.Rows(Y).Item("Disabled").ToString = "Y" Then
                                                                                    oTargetBinLocation.Inactive = SAPbobsCOM.BoYesNoEnum.tYES
                                                                                Else
                                                                                    oTargetBinLocation.Inactive = SAPbobsCOM.BoYesNoEnum.tNO
                                                                                End If

                                                                                If oDT_Binlists.Rows(Y).Item("SngBatch").ToString = "Y" Then
                                                                                    oTargetBinLocation.BatchRestrictions = SAPbobsCOM.BinRestrictionBatchEnum.brbSingleBatch
                                                                                Else
                                                                                    oTargetBinLocation.BatchRestrictions = SAPbobsCOM.BinRestrictionBatchEnum.brbNoRestrictions
                                                                                End If

                                                                                oTargetBinLocation.Description = oDT_Binlists.Rows(Y).Item("Descr").ToString
                                                                                oTargetBinLocation.AlternativeSortCode = oDT_Binlists.Rows(Y).Item("AltSortCod").ToString
                                                                                oTargetBinLocation.BarCode = oDT_Binlists.Rows(Y).Item("BarCode").ToString

                                                                                If oDT_Binlists.Rows(Y).Item("ItmRtrictT").ToString = 0 Then
                                                                                    oTargetBinLocation.RestrictedItemType = SAPbobsCOM.BinRestrictItemEnum.briNone
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("ItmRtrictT").ToString = 1 Then
                                                                                    oTargetBinLocation.RestrictedItemType = SAPbobsCOM.BinRestrictItemEnum.briSpecificItem
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("ItmRtrictT").ToString = 2 Then
                                                                                    oTargetBinLocation.RestrictedItemType = SAPbobsCOM.BinRestrictItemEnum.briSingleItemOnly
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("ItmRtrictT").ToString = 3 Then
                                                                                    oTargetBinLocation.RestrictedItemType = SAPbobsCOM.BinRestrictItemEnum.briSpecificItemGroup
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("ItmRtrictT").ToString = 4 Then
                                                                                    oTargetBinLocation.RestrictedItemType = SAPbobsCOM.BinRestrictItemEnum.briSpecificItemGroupOnly
                                                                                End If

                                                                                If oDT_Binlists.Rows(Y).Item("RtrictType").ToString = 0 Then
                                                                                    oTargetBinLocation.RestrictedTransType = SAPbobsCOM.BinRestrictTransactionEnum.brtNoRestrictions
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("RtrictType").ToString = 1 Then
                                                                                    oTargetBinLocation.RestrictedTransType = SAPbobsCOM.BinRestrictTransactionEnum.brtAllTrans
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("RtrictType").ToString = 2 Then
                                                                                    oTargetBinLocation.RestrictedTransType = SAPbobsCOM.BinRestrictTransactionEnum.brtInboundTrans
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("RtrictType").ToString = 3 Then
                                                                                    oTargetBinLocation.RestrictedTransType = SAPbobsCOM.BinRestrictTransactionEnum.brtOutboundTrans
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("RtrictType").ToString = 4 Then
                                                                                    oTargetBinLocation.RestrictedTransType = SAPbobsCOM.BinRestrictTransactionEnum.brtAllExceptInventoryTrans
                                                                                End If

                                                                                oTargetBinLocation.RestrictionReason = oDT_Binlists.Rows(Y).Item("RtrictResn").ToString

                                                                                oTargetBinLocation.SpecificItem = oDT_Binlists.Rows(Y).Item("SpcItmCode").ToString
                                                                                oTargetBinLocation.SpecificItemGroup = oDT_Binlists.Rows(Y).Item("SpcItmGrpC").ToString

                                                                                oTargetBinLocation.MinimumQty = Convert.ToDouble(oDT_Binlists.Rows(Y).Item("MinLevel"))
                                                                                oTargetBinLocation.MaximumQty = Convert.ToDouble(oDT_Binlists.Rows(Y).Item("MaxLevel"))

                                                                                Dim oAddBinLocationParams As SAPbobsCOM.BinLocationParams = TargetsvrBinLocation.GetDataInterface(SAPbobsCOM.BinLocationsServiceDataInterfaces.blcsBinLocationParams)
                                                                                Try
                                                                                    oAddBinLocationParams = TargetsvrBinLocation.Add(oTargetBinLocation)
                                                                                    oBinLocationEntry = oAddBinLocationParams.BinCode
                                                                                    SucFlag = True
                                                                                    p_oSBOApplication.StatusBar.SetText("Adding Bin Location '" & oDT_Binlists.Rows(Y).Item("BinCode").ToString & "' Successful on.." & oDICompany(S).CompanyDB, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Bin Location '" & oDT_Binlists.Rows(Y).Item("BinCode").ToString & "' Successful on  '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                Catch ex As Exception
                                                                                    sErrDesc = ex.Message
                                                                                    SucFlag = False
                                                                                    p_oSBOApplication.StatusBar.SetText("Adding Bin Location '" & oDT_Binlists.Rows(Y).Item("BinCode").ToString & "' Failed. on '" & oDICompany(S).CompanyDB & "'." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Bin Location '" & oDT_Binlists.Rows(Y).Item("BinCode").ToString & "' Failed on  '" & oDICompany(S).CompanyDB & "'" & sErrDesc, sFuncName)
                                                                                    GoTo 115
                                                                                Finally
                                                                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBinLocation)
                                                                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetBinLocation)
                                                                                End Try
                                                                            Else
                                                                                SBO_Application.StatusBar.SetText("Updating Bin Location '" & bincode & "' in.." & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                Dim oUpdateBinLocationParams As SAPbobsCOM.BinLocationParams = TargetsvrBinLocation.GetDataInterface(SAPbobsCOM.BinLocationsServiceDataInterfaces.blcsBinLocationParams)
                                                                                Dim oUpdateTargetBinLocation As SAPbobsCOM.BinLocation = TargetsvrBinLocation.GetDataInterface(SAPbobsCOM.BinLocationsServiceDataInterfaces.blcsBinLocation)

                                                                                oUpdateBinLocationParams.BinCode = bincode

                                                                                Try
                                                                                    oUpdateTargetBinLocation = TargetsvrBinLocation.Get(oUpdateBinLocationParams)
                                                                                Catch ex As Exception
                                                                                    SucFlag = False
                                                                                End Try

                                                                                oUpdateTargetBinLocation.Warehouse = oDT_Binlists.Rows(Y).Item("WhsCode").ToString

                                                                                If oDT_Binlists.Rows(Y).Item("SL1Code").ToString <> String.Empty Then oUpdateTargetBinLocation.Sublevel1 = oDT_Binlists.Rows(Y).Item("SL1Code").ToString
                                                                                If oDT_Binlists.Rows(Y).Item("SL2Code").ToString <> String.Empty Then oUpdateTargetBinLocation.Sublevel2 = oDT_Binlists.Rows(Y).Item("SL2Code").ToString
                                                                                If oDT_Binlists.Rows(Y).Item("SL3Code").ToString <> String.Empty Then oUpdateTargetBinLocation.Sublevel3 = oDT_Binlists.Rows(Y).Item("SL3Code").ToString
                                                                                If oDT_Binlists.Rows(Y).Item("SL4Code").ToString <> String.Empty Then oUpdateTargetBinLocation.Sublevel4 = oDT_Binlists.Rows(Y).Item("SL4Code").ToString
                                                                                If oDT_Binlists.Rows(Y).Item("ReceiveBin").ToString = "Y" Then
                                                                                    oUpdateTargetBinLocation.ReceivingBinLocation = SAPbobsCOM.BoYesNoEnum.tYES
                                                                                Else
                                                                                    oUpdateTargetBinLocation.ReceivingBinLocation = SAPbobsCOM.BoYesNoEnum.tNO
                                                                                End If

                                                                                If oDT_Binlists.Rows(Y).Item("NoAutoAllc").ToString = "Y" Then
                                                                                    oUpdateTargetBinLocation.ExcludeAutoAllocOnIssue = SAPbobsCOM.BoYesNoEnum.tYES
                                                                                Else
                                                                                    oUpdateTargetBinLocation.ExcludeAutoAllocOnIssue = SAPbobsCOM.BoYesNoEnum.tNO
                                                                                End If

                                                                                If oDT_Binlists.Rows(Y).Item("Disabled").ToString = "Y" Then
                                                                                    oUpdateTargetBinLocation.Inactive = SAPbobsCOM.BoYesNoEnum.tYES
                                                                                Else
                                                                                    oUpdateTargetBinLocation.Inactive = SAPbobsCOM.BoYesNoEnum.tNO
                                                                                End If

                                                                                If oDT_Binlists.Rows(Y).Item("SngBatch").ToString = "Y" Then
                                                                                    oUpdateTargetBinLocation.BatchRestrictions = SAPbobsCOM.BinRestrictionBatchEnum.brbSingleBatch
                                                                                Else
                                                                                    oUpdateTargetBinLocation.BatchRestrictions = SAPbobsCOM.BinRestrictionBatchEnum.brbNoRestrictions
                                                                                End If

                                                                                oUpdateTargetBinLocation.Description = oDT_Binlists.Rows(Y).Item("Descr").ToString
                                                                                oUpdateTargetBinLocation.AlternativeSortCode = oDT_Binlists.Rows(Y).Item("AltSortCod").ToString
                                                                                oUpdateTargetBinLocation.BarCode = oDT_Binlists.Rows(Y).Item("BarCode").ToString

                                                                                If oDT_Binlists.Rows(Y).Item("ItmRtrictT").ToString = 0 Then
                                                                                    oUpdateTargetBinLocation.RestrictedItemType = SAPbobsCOM.BinRestrictItemEnum.briNone
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("ItmRtrictT").ToString = 1 Then
                                                                                    oUpdateTargetBinLocation.RestrictedItemType = SAPbobsCOM.BinRestrictItemEnum.briSpecificItem
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("ItmRtrictT").ToString = 2 Then
                                                                                    oUpdateTargetBinLocation.RestrictedItemType = SAPbobsCOM.BinRestrictItemEnum.briSingleItemOnly
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("ItmRtrictT").ToString = 3 Then
                                                                                    oUpdateTargetBinLocation.RestrictedItemType = SAPbobsCOM.BinRestrictItemEnum.briSpecificItemGroup
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("ItmRtrictT").ToString = 4 Then
                                                                                    oUpdateTargetBinLocation.RestrictedItemType = SAPbobsCOM.BinRestrictItemEnum.briSpecificItemGroupOnly
                                                                                End If

                                                                                If oDT_Binlists.Rows(Y).Item("RtrictType").ToString = 0 Then
                                                                                    oUpdateTargetBinLocation.RestrictedTransType = SAPbobsCOM.BinRestrictTransactionEnum.brtNoRestrictions
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("RtrictType").ToString = 1 Then
                                                                                    oUpdateTargetBinLocation.RestrictedTransType = SAPbobsCOM.BinRestrictTransactionEnum.brtAllTrans
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("RtrictType").ToString = 2 Then
                                                                                    oUpdateTargetBinLocation.RestrictedTransType = SAPbobsCOM.BinRestrictTransactionEnum.brtInboundTrans
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("RtrictType").ToString = 3 Then
                                                                                    oUpdateTargetBinLocation.RestrictedTransType = SAPbobsCOM.BinRestrictTransactionEnum.brtOutboundTrans
                                                                                ElseIf oDT_Binlists.Rows(Y).Item("RtrictType").ToString = 4 Then
                                                                                    oUpdateTargetBinLocation.RestrictedTransType = SAPbobsCOM.BinRestrictTransactionEnum.brtAllExceptInventoryTrans
                                                                                End If

                                                                                oUpdateTargetBinLocation.RestrictionReason = oDT_Binlists.Rows(Y).Item("RtrictResn").ToString

                                                                                oUpdateTargetBinLocation.SpecificItem = oDT_Binlists.Rows(Y).Item("SpcItmCode").ToString
                                                                                oUpdateTargetBinLocation.SpecificItemGroup = oDT_Binlists.Rows(Y).Item("SpcItmGrpC").ToString

                                                                                oUpdateTargetBinLocation.MinimumQty = Convert.ToDouble(oDT_Binlists.Rows(Y).Item("MinLevel"))
                                                                                oUpdateTargetBinLocation.MaximumQty = Convert.ToDouble(oDT_Binlists.Rows(Y).Item("MaxLevel"))

                                                                                Try
                                                                                    TargetsvrBinLocation.Update(oUpdateTargetBinLocation)
                                                                                    SucFlag = True
                                                                                    p_oSBOApplication.StatusBar.SetText("Updating Bin Location '" & oDT_Binlists.Rows(Y).Item("BinCode").ToString & "' Successful on.." & oDICompany(S).CompanyDB, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating Bin Location '" & oDT_Binlists.Rows(Y).Item("BinCode").ToString & "' Successful on  '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                Catch ex As Exception
                                                                                    sErrDesc = ex.Message
                                                                                    SucFlag = False
                                                                                    p_oSBOApplication.StatusBar.SetText("Updating Bin Location '" & oDT_Binlists.Rows(Y).Item("BinCode").ToString & "' Failed. on '" & oDICompany(S).CompanyDB & "'." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating Bin Location '" & oDT_Binlists.Rows(Y).Item("BinCode").ToString & "' Failed on  '" & oDICompany(S).CompanyDB & "'" & sErrDesc, sFuncName)
                                                                                    GoTo 115
                                                                                Finally
                                                                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUpdateTargetBinLocation)
                                                                                End Try
                                                                            End If
                                                                        Next
                                                                    Else
                                                                        SucFlag = True
                                                                    End If
                                                                    If SucFlag = False Then
115:
                                                                        sErrDesc = sErrDesc + " On.." + oDICompany(S).CompanyDB
                                                                        p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                        Fllag = False
                                                                        Dim oEdit As SAPbouiCOM.StaticText = oForm.Items.Item("l_RepStat1").Specific
                                                                        oEdit.Caption = sErrDesc
                                                                        oEdit.Item.FontSize = 10
                                                                        oEdit.Item.ForeColor = RGB(255, 0, 0)
                                                                        For lCounter As Integer = 0 To UBound(oDICompany)
                                                                            If Not oDICompany(lCounter) Is Nothing Then
                                                                                If oDICompany(lCounter).Connected = True Then
                                                                                    If oDICompany(lCounter).InTransaction = True Then
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                        oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                                    End If
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                    oDICompany(lCounter).Disconnect()
                                                                                    oDICompany(lCounter) = Nothing
                                                                                End If
                                                                            End If
                                                                        Next
                                                                        Exit For
                                                                    Else
                                                                        Fllag = True
                                                                        SBO_Application.SetStatusBarMessage("Bin Location Replicated Successfully on.. " & oDV_BPSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Bin Location Replicated Successfully on .. " & oDV_BPSetup.Item(S).Item("Name").ToString, sFuncName)
                                                                    End If
                                                                Next
                                                                For lCounter As Integer = 0 To UBound(oDICompany)
                                                                    If Not oDICompany(lCounter) Is Nothing Then
                                                                        If oDICompany(lCounter).Connected = True Then
                                                                            If oDICompany(lCounter).InTransaction = True Then
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                                            End If
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                            oDICompany(lCounter).Disconnect()
                                                                            oDICompany(lCounter) = Nothing
                                                                        End If
                                                                    End If
                                                                Next
                                                                If Fllag = True Then
                                                                    Dim oEdit1 As SAPbouiCOM.StaticText = oForm.Items.Item("l_RepStat1").Specific
                                                                    oEdit1.Caption = "Bin Locations Replicated Successfully on Target Databases.."
                                                                    oEdit1.Item.FontSize = 10
                                                                    oEdit1.Item.ForeColor = RGB(0, 128, 0)
                                                                End If
                                                            End If
                                                        End If
                                                    ElseIf oMasttype.Selected.Value = "F" Then
                                                        Dim sqry As String = "Select T1.""Name"", T1.""U_UserName"", T1.""U_Password"", T0.""U_Currency"" ,T0.""U_COA"" , T0.""U_PostPeriod"",T0.""U_ExcRates"" , T0.""U_CostCenter1"",T0.""U_CostCenter2"",T0.""U_CostCenter3"",T0.""U_CostCenter4"",T0.""U_CostCenter5"" from ""@AE_TB003_FIN"" T0  LEFT OUTER JOIN ""@AE_TB004_TARCRE"" T1 ON T0.""U_TargetDB"" = T1.""Code"" 	WHERE (T0.""U_Currency""  = 'Y' OR T0.""U_COA"" = 'Y' OR T0.""U_PostPeriod"" ='Y' OR T0.""U_ExcRates"" ='Y' OR	T0.""U_CostCenter1"" ='Y' OR T0.""U_CostCenter2"" ='Y' OR T0.""U_CostCenter3"" ='Y' OR T0.""U_CostCenter4"" ='Y' OR T0.""U_CostCenter5""='Y');"
                                                        oRset1.DoQuery(sqry)
                                                        oDT_FINSetup = New DataTable
                                                        oDT_FINSetup = ConvertRecordset(oRset1)
                                                        Dim dtcount As Integer = oDT_FINSetup.Rows.Count
                                                        Dim oDV_FINSetup As New DataView(oDT_FINSetup)
                                                        Dim dvcount As Integer = oDV_FINSetup.Count
                                                        ' Dim dv As New DataView(dt)

                                                        Dim rrr As String = oMatrix.RowCount

                                                        If (oDT_Entities.Rows.Count > 0 And oDV_FINSetup.Count > 0) Then
                                                            For J As Integer = 0 To oDT_Entities.Rows.Count - 1

                                                                If (oDT_Entities.Rows(J).Item("TransType").ToString = "COA") Then
                                                                    ''=========================================================================================================
                                                                    '------------------------------- Chart of Accounts Replication ----------------------------------------------
                                                                    ''=========================================================================================================
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Finance Master Setup for COA Settings", sFuncName)
                                                                    Dim rvcount As Integer = oDV_FINSetup.Count
                                                                    Fllag = False
                                                                    oDV_FINSetup.RowFilter = "U_COA ='Y'"
                                                                    ReDim oDICompany(oDV_FINSetup.Count)
                                                                    Dim dvcount1 As Integer = oDV_FINSetup.Count
                                                                    If oDV_FINSetup.Count > 0 Then
                                                                        For S As Integer = 0 To oDV_FINSetup.Count - 1
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Connect To TargetCompany() for COA Posting", sFuncName)
                                                                            SBO_Application.SetStatusBarMessage("Connecting to the Target Company" & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                            Dim targetent As String = oDV_FINSetup.Item(S).Item("Name").ToString
                                                                            If ConnectToTargetCompany(oDICompany(S), oDV_FINSetup.Item(S).Item("Name").ToString, oDV_FINSetup.Item(S).Item("U_UserName").ToString, oDV_FINSetup.Item(S).Item("U_Password").ToString, sErrDesc) <> RTN_SUCCESS Then
                                                                                GoTo 116
                                                                            End If
                                                                            SBO_Application.SetStatusBarMessage("Connecting to the target company Successful " & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the target company Successful " & oDICompany(S).CompanyDB, sFuncName)
                                                                            oDICompany(S).StartTransaction()

                                                                            SBO_Application.SetStatusBarMessage("Started Master Data Synchronization " & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                                            Dim oCOA As SAPbobsCOM.ChartOfAccounts = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts)
                                                                            Dim oTargetCOA As SAPbobsCOM.ChartOfAccounts = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts)
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                                                                            Dim flg1 As Boolean = False
                                                                            Dim SucFlag As Boolean = False
                                                                            Dim groupno As String = String.Empty

                                                                            If oCOA.GetByKey(oDT_Entities.Rows(J).Item("Code")) Then
                                                                                If oTargetCOA.GetByKey(oDT_Entities.Rows(J).Item("Code")) = True Then
                                                                                    flg1 = True
                                                                                    oTargetCOA.Name = oCOA.Name
                                                                                    oTargetCOA.ForeignName = oCOA.ForeignName
                                                                                    oTargetCOA.AccountType = oCOA.AccountType
                                                                                    oTargetCOA.ExternalCode = oCOA.ExternalCode
                                                                                    oTargetCOA.ActiveAccount = oCOA.ActiveAccount
                                                                                    oTargetCOA.FatherAccountKey = oCOA.FatherAccountKey
                                                                                    oTargetCOA.AcctCurrency = oCOA.AcctCurrency
                                                                                    oTargetCOA.CashAccount = oCOA.CashAccount
                                                                                    oTargetCOA.Protected = oCOA.Protected
                                                                                    oTargetCOA.LockManualTransaction = oCOA.LockManualTransaction
                                                                                    oTargetCOA.Details = oCOA.Details
                                                                                    oTargetCOA.ValidFor = oCOA.ValidFor
                                                                                    oTargetCOA.FrozenFor = oCOA.FrozenFor
                                                                                    oTargetCOA.ValidFrom = oCOA.ValidFrom
                                                                                    oTargetCOA.ValidTo = oCOA.ValidTo
                                                                                    oTargetCOA.ValidRemarks = oCOA.ValidRemarks
                                                                                    oTargetCOA.FrozenFrom = oCOA.FrozenFrom
                                                                                    oTargetCOA.FrozenTo = oCOA.FrozenTo
                                                                                    oTargetCOA.FrozenRemarks = oCOA.FrozenRemarks
                                                                                    oTargetCOA.AllowChangeVatGroup = oCOA.AllowChangeVatGroup
                                                                                    oTargetCOA.AllowMultipleLinking = oCOA.AllowMultipleLinking
                                                                                    oTargetCOA.BudgetAccount = oCOA.BudgetAccount
                                                                                    oTargetCOA.DataExportCode = oCOA.DataExportCode
                                                                                    oTargetCOA.LiableForAdvances = oCOA.LiableForAdvances
                                                                                    oTargetCOA.LoadingType = oCOA.LoadingType
                                                                                    oTargetCOA.PlanningLevel = oCOA.PlanningLevel
                                                                                    oTargetCOA.ProjectRelevant = oCOA.ProjectRelevant
                                                                                    oTargetCOA.ProjectCode = oCOA.ProjectCode
                                                                                    oTargetCOA.RateConversion = oCOA.RateConversion
                                                                                    oTargetCOA.ReconciledAccount = oCOA.ReconciledAccount
                                                                                    oTargetCOA.RevaluationCoordinated = oCOA.RevaluationCoordinated
                                                                                    oTargetCOA.TaxExemptAccount = oCOA.TaxExemptAccount
                                                                                    oTargetCOA.TaxLiableAccount = oCOA.TaxLiableAccount
                                                                                Else
                                                                                    oTargetCOA.Code = oCOA.Code
                                                                                    oTargetCOA.Name = oCOA.Name
                                                                                    oTargetCOA.ForeignName = oCOA.ForeignName
                                                                                    oTargetCOA.AccountType = oCOA.AccountType
                                                                                    oTargetCOA.ExternalCode = oCOA.ExternalCode
                                                                                    oTargetCOA.ActiveAccount = oCOA.ActiveAccount
                                                                                    oTargetCOA.FatherAccountKey = oCOA.FatherAccountKey
                                                                                    oTargetCOA.AcctCurrency = oCOA.AcctCurrency
                                                                                    oTargetCOA.CashAccount = oCOA.CashAccount
                                                                                    oTargetCOA.Protected = oCOA.Protected
                                                                                    oTargetCOA.LockManualTransaction = oCOA.LockManualTransaction
                                                                                    oTargetCOA.Details = oCOA.Details
                                                                                    oTargetCOA.ValidFor = oCOA.ValidFor
                                                                                    oTargetCOA.FrozenFor = oCOA.FrozenFor
                                                                                    oTargetCOA.ValidFrom = oCOA.ValidFrom
                                                                                    oTargetCOA.ValidTo = oCOA.ValidTo
                                                                                    oTargetCOA.ValidRemarks = oCOA.ValidRemarks
                                                                                    oTargetCOA.FrozenFrom = oCOA.FrozenFrom
                                                                                    oTargetCOA.FrozenTo = oCOA.FrozenTo
                                                                                    oTargetCOA.FrozenRemarks = oCOA.FrozenRemarks
                                                                                    oTargetCOA.AllowChangeVatGroup = oCOA.AllowChangeVatGroup
                                                                                    oTargetCOA.AllowMultipleLinking = oCOA.AllowMultipleLinking
                                                                                    oTargetCOA.BudgetAccount = oCOA.BudgetAccount
                                                                                    oTargetCOA.DataExportCode = oCOA.DataExportCode
                                                                                    oTargetCOA.LiableForAdvances = oCOA.LiableForAdvances
                                                                                    oTargetCOA.LoadingType = oCOA.LoadingType
                                                                                    oTargetCOA.PlanningLevel = oCOA.PlanningLevel
                                                                                    oTargetCOA.ProjectRelevant = oCOA.ProjectRelevant
                                                                                    oTargetCOA.ProjectCode = oCOA.ProjectCode
                                                                                    oTargetCOA.RateConversion = oCOA.RateConversion
                                                                                    oTargetCOA.ReconciledAccount = oCOA.ReconciledAccount
                                                                                    oTargetCOA.RevaluationCoordinated = oCOA.RevaluationCoordinated
                                                                                    oTargetCOA.TaxExemptAccount = oCOA.TaxExemptAccount
                                                                                    oTargetCOA.TaxLiableAccount = oCOA.TaxLiableAccount
                                                                                End If
                                                                            End If
                                                                            If flg1 = True Then
                                                                                ErrerCode = oTargetCOA.Update()
                                                                                If ErrerCode <> 0 Then
                                                                                    oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                    p_oSBOApplication.StatusBar.SetText("Updating COA '" & oDT_Entities.Rows(J).Item("Code") & "' to Target Company '" & oDICompany(S).CompanyDB & "' Failed" & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating COA '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                    SucFlag = False
                                                                                Else
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully added COA: '" & oDT_Entities.Rows(J).Item("Code") & "' on '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                    p_oSBOApplication.StatusBar.SetText("Updating COA '" & oDT_Entities.Rows(J).Item("Code") & "' to Target Company '" & oDICompany(S).CompanyDB & "' Successful", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                    SucFlag = True
                                                                                End If
                                                                            Else
                                                                                ErrerCode = oTargetCOA.Add
                                                                                If ErrerCode <> 0 Then
                                                                                    oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                    p_oSBOApplication.StatusBar.SetText("Adding COA '" & oDT_Entities.Rows(J).Item("Code") & "' to Target Company '" & oDICompany(S).CompanyDB & "' Failed" & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding COA '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                    SucFlag = False
                                                                                Else
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully added COA: '" & oDT_Entities.Rows(J).Item("Code") & "' on '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                    p_oSBOApplication.StatusBar.SetText("Adding COA '" & oDT_Entities.Rows(J).Item("Code") & "' to Target Company '" & oDICompany(S).CompanyDB & "' Successful", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                    SucFlag = True
                                                                                End If
                                                                            End If

                                                                            If SucFlag = False Then
116:
                                                                                sErrDesc = sErrDesc + " On.." + oDICompany(S).CompanyDB
                                                                                p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                Fllag = False
                                                                                Dim sqy As String = " UPDATE ""INTEGRATION"" SET ""SYNCSTATUS"" = 'NO', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '" & sErrDesc.ToString.Replace("'", """") & "'  WHERE ""MASTERTYPE"" = 'FINANCEMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                                Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                Run.DoQuery(sqy)

                                                                                Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                                oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "NO"
                                                                                oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                                oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = sErrDesc

                                                                                For lCounter As Integer = 0 To UBound(oDICompany)
                                                                                    If Not oDICompany(lCounter) Is Nothing Then
                                                                                        If oDICompany(lCounter).Connected = True Then
                                                                                            If oDICompany(lCounter).InTransaction = True Then
                                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                                oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                                            End If
                                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                            oDICompany(lCounter).Disconnect()
                                                                                            oDICompany(lCounter) = Nothing
                                                                                        End If
                                                                                    End If
                                                                                Next
                                                                                Exit For
                                                                            Else
                                                                                Fllag = True
                                                                                SBO_Application.SetStatusBarMessage("Chart of Account '" & oDT_Entities.Rows(J).Item("Code") & "' Replicated Successfully on.. " & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Chart of Account '" & oDT_Entities.Rows(J).Item("Code") & "' Replicated Successfully on .. " & oDV_FINSetup.Item(S).Item("Name").ToString, sFuncName)
                                                                            End If
                                                                            Try
                                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCOA)
                                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetCOA)
                                                                            Catch ex As Exception
                                                                            End Try
                                                                        Next
                                                                        For lCounter As Integer = 0 To UBound(oDICompany)
                                                                            If Not oDICompany(lCounter) Is Nothing Then
                                                                                If oDICompany(lCounter).Connected = True Then
                                                                                    If oDICompany(lCounter).InTransaction = True Then
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                        oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                                                    End If
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                    oDICompany(lCounter).Disconnect()
                                                                                    oDICompany(lCounter) = Nothing
                                                                                End If
                                                                            End If
                                                                        Next
                                                                        If Fllag = True Then
                                                                            Dim sqy As String = " Update ""INTEGRATION"" SET ""SYNCSTATUS"" = 'YES', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '' WHERE ""MASTERTYPE"" = 'FINANCEMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                            Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                            Run.DoQuery(sqy)

                                                                            Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                            oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "YES"
                                                                            oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                            oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = "SUCCESS"
                                                                        End If
                                                                    End If
                                                                ElseIf (oDT_Entities.Rows(J).Item("TransType").ToString = "Currency") Then
                                                                    ''=========================================================================================================
                                                                    '------------------------------- Currency Replication ----------------------------------------------
                                                                    ''=========================================================================================================
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Finance Master Setup for Currency Setup", sFuncName)
                                                                    Dim rvcount As Integer = oDV_FINSetup.Count
                                                                    Fllag = False
                                                                    oDV_FINSetup.RowFilter = "U_Currency ='Y'"
                                                                    ReDim oDICompany(oDV_FINSetup.Count)
                                                                    Dim dvcount1 As Integer = oDV_FINSetup.Count
                                                                    If oDV_FINSetup.Count > 0 Then
                                                                        For S As Integer = 0 To oDV_FINSetup.Count - 1
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Connect To Target Company() for Currency Posting", sFuncName)
                                                                            SBO_Application.SetStatusBarMessage("Connecting to the Target Company" & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                            Dim targetent As String = oDV_FINSetup.Item(S).Item("Name").ToString
                                                                            If ConnectToTargetCompany(oDICompany(S), oDV_FINSetup.Item(S).Item("Name").ToString, oDV_FINSetup.Item(S).Item("U_UserName").ToString, oDV_FINSetup.Item(S).Item("U_Password").ToString, sErrDesc) <> RTN_SUCCESS Then
                                                                                GoTo 117
                                                                            End If
                                                                            SBO_Application.SetStatusBarMessage("Connecting to the target company Successfull " & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the target company Successfull " & oDICompany(S).CompanyDB, sFuncName)
                                                                            oDICompany(S).StartTransaction()

                                                                            SBO_Application.SetStatusBarMessage("Started Master Data Synchronization " & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                                            Dim oCurrency As SAPbobsCOM.Currencies = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCurrencyCodes)
                                                                            Dim oTargetCurrency As SAPbobsCOM.Currencies = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCurrencyCodes)
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                                                                            Dim flg1 As Boolean = False
                                                                            Dim SucFlag As Boolean = False
                                                                            Dim groupno As String = String.Empty

                                                                            If oCurrency.GetByKey(oDT_Entities.Rows(J).Item("Code")) Then
                                                                                If oTargetCurrency.GetByKey(oDT_Entities.Rows(J).Item("Code")) = True Then
                                                                                    flg1 = True
                                                                                    oTargetCurrency.Name = oCurrency.Name
                                                                                    oTargetCurrency.DocumentsCode = oCurrency.DocumentsCode
                                                                                    oTargetCurrency.EnglishName = oCurrency.EnglishName
                                                                                    oTargetCurrency.EnglishHundredthName = oCurrency.EnglishHundredthName
                                                                                    oTargetCurrency.HundredthName = oCurrency.HundredthName
                                                                                    oTargetCurrency.InternationalDescription = oCurrency.InternationalDescription
                                                                                    oTargetCurrency.Rounding = oCurrency.Rounding
                                                                                    oTargetCurrency.RoundingInPayment = oCurrency.RoundingInPayment
                                                                                    oTargetCurrency.Decimals = oCurrency.Decimals
                                                                                    oTargetCurrency.MaxIncomingAmtDiff = oCurrency.MaxIncomingAmtDiff
                                                                                    oTargetCurrency.MaxIncomingAmtDiffPercent = oCurrency.MaxIncomingAmtDiffPercent
                                                                                    oTargetCurrency.MaxOutgoingAmtDiffPercent = oCurrency.MaxOutgoingAmtDiffPercent
                                                                                    oTargetCurrency.MaxOutgoingAmtDiff = oCurrency.MaxOutgoingAmtDiff
                                                                                Else
                                                                                    oTargetCurrency.Code = oCurrency.Code
                                                                                    oTargetCurrency.Name = oCurrency.Name
                                                                                    oTargetCurrency.DocumentsCode = oCurrency.DocumentsCode
                                                                                    oTargetCurrency.EnglishName = oCurrency.EnglishName
                                                                                    oTargetCurrency.EnglishHundredthName = oCurrency.EnglishHundredthName
                                                                                    oTargetCurrency.HundredthName = oCurrency.HundredthName
                                                                                    oTargetCurrency.InternationalDescription = oCurrency.InternationalDescription
                                                                                    oTargetCurrency.Rounding = oCurrency.Rounding
                                                                                    oTargetCurrency.RoundingInPayment = oCurrency.RoundingInPayment
                                                                                    oTargetCurrency.Decimals = oCurrency.Decimals
                                                                                    oTargetCurrency.MaxIncomingAmtDiff = oCurrency.MaxIncomingAmtDiff
                                                                                    oTargetCurrency.MaxIncomingAmtDiffPercent = oCurrency.MaxIncomingAmtDiffPercent
                                                                                    oTargetCurrency.MaxOutgoingAmtDiffPercent = oCurrency.MaxOutgoingAmtDiffPercent
                                                                                    oTargetCurrency.MaxOutgoingAmtDiff = oCurrency.MaxOutgoingAmtDiff
                                                                                End If
                                                                            End If
                                                                            If flg1 = True Then
                                                                                ErrerCode = oTargetCurrency.Update()
                                                                                If ErrerCode <> 0 Then
                                                                                    oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                    p_oSBOApplication.StatusBar.SetText("Updating Currency '" & oDT_Entities.Rows(J).Item("Code") & "' to Target Company '" & oDICompany(S).CompanyDB & "' Failed" & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Medium)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating Currency '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                    SucFlag = False
                                                                                Else
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully updated Currency: '" & oDT_Entities.Rows(J).Item("Code") & "' on '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                    p_oSBOApplication.StatusBar.SetText("Updating Currency '" & oDT_Entities.Rows(J).Item("Code") & "' to Target Company '" & oDICompany(S).CompanyDB & "' Successful", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                    SucFlag = True
                                                                                End If
                                                                            Else
                                                                                ErrerCode = oTargetCurrency.Add
                                                                                If ErrerCode <> 0 Then
                                                                                    oDICompany(S).GetLastError(ErrerCode, sErrDesc)
                                                                                    p_oSBOApplication.StatusBar.SetText("Adding Currency '" & oDT_Entities.Rows(J).Item("Code") & "' to Target Company '" & oDICompany(S).CompanyDB & "' Failed" & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Medium)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Currency '" & oDT_Entities.Rows(J).Item("Code") & "' Failed on  '" & oDICompany(S).CompanyDB & "' " + " - " + sErrDesc, sFuncName)
                                                                                    SucFlag = False
                                                                                Else
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully Added Currency: '" & oDT_Entities.Rows(J).Item("Code") & "' on '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                    p_oSBOApplication.StatusBar.SetText("Adding Currency '" & oDT_Entities.Rows(J).Item("Code") & "' to Target Company '" & oDICompany(S).CompanyDB & "' Successful", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                    SucFlag = True
                                                                                End If
                                                                            End If

                                                                            If SucFlag = False Then
117:
                                                                                sErrDesc = sErrDesc + " On.." + oDICompany(S).CompanyDB
                                                                                p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                Fllag = False
                                                                                Dim sqy As String = " UPDATE ""INTEGRATION"" SET ""SYNCSTATUS"" = 'NO', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '" & sErrDesc.ToString.Replace("'", """") & "'  WHERE ""MASTERTYPE"" = 'FINANCEMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                                Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                Run.DoQuery(sqy)

                                                                                Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                                oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "NO"
                                                                                oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                                oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = sErrDesc

                                                                                For lCounter As Integer = 0 To UBound(oDICompany)
                                                                                    If Not oDICompany(lCounter) Is Nothing Then
                                                                                        If oDICompany(lCounter).Connected = True Then
                                                                                            If oDICompany(lCounter).InTransaction = True Then
                                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                                oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                                            End If
                                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                            oDICompany(lCounter).Disconnect()
                                                                                            oDICompany(lCounter) = Nothing
                                                                                        End If
                                                                                    End If
                                                                                Next
                                                                                Exit For
                                                                            Else
                                                                                Fllag = True
                                                                                SBO_Application.SetStatusBarMessage("Currency '" & oDT_Entities.Rows(J).Item("Code") & "' Replicated Successfully on.. " & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Currency '" & oDT_Entities.Rows(J).Item("Code") & "' Replicated Successfully on .. " & oDV_FINSetup.Item(S).Item("Name").ToString, sFuncName)
                                                                            End If
                                                                            Try
                                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCurrency)
                                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetCurrency)
                                                                            Catch ex As Exception
                                                                            End Try
                                                                        Next
                                                                        For lCounter As Integer = 0 To UBound(oDICompany)
                                                                            If Not oDICompany(lCounter) Is Nothing Then
                                                                                If oDICompany(lCounter).Connected = True Then
                                                                                    If oDICompany(lCounter).InTransaction = True Then
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                        oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                                                    End If
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                    oDICompany(lCounter).Disconnect()
                                                                                    oDICompany(lCounter) = Nothing
                                                                                End If
                                                                            End If
                                                                        Next
                                                                        If Fllag = True Then
                                                                            Dim sqy As String = " Update ""INTEGRATION"" SET ""SYNCSTATUS"" = 'YES', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '' WHERE ""MASTERTYPE"" = 'FINANCEMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                            Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                            Run.DoQuery(sqy)

                                                                            Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                            oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "YES"
                                                                            oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                            oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = "SUCCESS"
                                                                        End If
                                                                    End If
                                                                ElseIf (oDT_Entities.Rows(J).Item("TransType").ToString = "CostCenter1" Or oDT_Entities.Rows(J).Item("TransType").ToString = "CostCenter2" Or oDT_Entities.Rows(J).Item("TransType").ToString = "CostCenter3" Or oDT_Entities.Rows(J).Item("TransType").ToString = "CostCenter4" Or oDT_Entities.Rows(J).Item("TransType").ToString = "CostCenter5") Then
                                                                    ''=========================================================================================================
                                                                    '------------------------------- Cost Center Replication ----------------------------------------------
                                                                    ''=========================================================================================================
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Finance Master Setup for Profit Center Settings", sFuncName)
                                                                    Dim rvcount As Integer = oDV_FINSetup.Count
                                                                    Fllag = False
                                                                    Dim SucFlag As Boolean = False
                                                                    Dim flag7 As Boolean = False
                                                                    If oDT_Entities.Rows(J).Item("TransType").ToString = "CostCenter1" Then
                                                                        oDV_FINSetup.RowFilter = "U_CostCenter1 ='Y'"
                                                                    ElseIf oDT_Entities.Rows(J).Item("TransType").ToString = "CostCenter2" Then
                                                                        oDV_FINSetup.RowFilter = "U_CostCenter2 ='Y'"
                                                                    ElseIf oDT_Entities.Rows(J).Item("TransType").ToString = "CostCenter3" Then
                                                                        oDV_FINSetup.RowFilter = "U_CostCenter3 ='Y'"
                                                                    ElseIf oDT_Entities.Rows(J).Item("TransType").ToString = "CostCenter4" Then
                                                                        oDV_FINSetup.RowFilter = "U_CostCenter4 ='Y'"
                                                                    ElseIf oDT_Entities.Rows(J).Item("TransType").ToString = "CostCenter5" Then
                                                                        oDV_FINSetup.RowFilter = "U_CostCenter5 ='Y'"
                                                                    End If
                                                                    ReDim oDICompany(oDV_FINSetup.Count)
                                                                    Dim dvcount1 As Integer = oDV_FINSetup.Count
                                                                    If oDV_FINSetup.Count > 0 Then
                                                                        For S As Integer = 0 To oDV_FINSetup.Count - 1
                                                                            Fllag = False
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Connect To Target Company() for Cost Center Posting", sFuncName)
                                                                            SBO_Application.SetStatusBarMessage("Connecting to the Target Company - " & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                            Dim targetent As String = oDV_FINSetup.Item(S).Item("Name").ToString
                                                                            If ConnectToTargetCompany(oDICompany(S), oDV_FINSetup.Item(S).Item("Name").ToString, oDV_FINSetup.Item(S).Item("U_UserName").ToString, oDV_FINSetup.Item(S).Item("U_Password").ToString, sErrDesc) <> RTN_SUCCESS Then
                                                                                GoTo 118
                                                                            End If
                                                                            SBO_Application.SetStatusBarMessage("Connecting to the target company Successfull -" & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the target company Successfull " & oDICompany(S).CompanyDB, sFuncName)
                                                                            oDICompany(S).StartTransaction()

                                                                            SBO_Application.SetStatusBarMessage("Started Master Data Synchronization " & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                                            Try

                                                                                Dim oCostCenterServices As SAPbobsCOM.IProfitCentersService = p_oDICompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ProfitCentersService)
                                                                                Dim oCostCenter1 As SAPbobsCOM.IProfitCenter = oCostCenterServices.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenter)

                                                                                Dim oTargetCostCenterServices As SAPbobsCOM.IProfitCentersService = oDICompany(S).GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ProfitCentersService)
                                                                                Dim oTargetCostCenter As SAPbobsCOM.IProfitCenter = oTargetCostCenterServices.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenter)

                                                                                Dim oUpdateProfitCenterParams As SAPbobsCOM.ProfitCenterParams = oTargetCostCenterServices.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenterParams)
                                                                                Dim oUpdateTargetProfitCenter As SAPbobsCOM.ProfitCenter = oTargetCostCenterServices.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenter)

                                                                                Dim RsetCostCenter As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                RsetCostCenter.DoQuery(String.Format("Select ""PrcCode"" from ""OPRC"" where ""PrcCode"" = '{0}'", oDT_Entities.Rows(J).Item("Code").ToString))
                                                                                If RsetCostCenter.RecordCount = 1 Then
                                                                                    flag7 = True
                                                                                End If

                                                                                Dim sss As String = "select ""PrcCode"", ""PrcName"", ""GrpCode"", ""DimCode"",""CCTypeCode"",""Active"", ""ValidFrom"", ""ValidTo"" from OPRC where ""PrcCode""= '" & oDT_Entities.Rows(J).Item("Code").ToString & "'"
                                                                                Dim RsetValues As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                RsetValues.DoQuery(sss)
                                                                                If RsetValues.RecordCount > 0 Then
                                                                                    If flag7 = False Then
                                                                                        oTargetCostCenter.CenterCode = RsetValues.Fields.Item("PrcCode").Value
                                                                                        oTargetCostCenter.CenterName = RsetValues.Fields.Item("PrcName").Value
                                                                                        oTargetCostCenter.CostCenterType = RsetValues.Fields.Item("CCTypeCode").Value
                                                                                        oTargetCostCenter.InWhichDimension = RsetValues.Fields.Item("DimCode").Value
                                                                                        oTargetCostCenter.GroupCode = RsetValues.Fields.Item("GrpCode").Value
                                                                                        If RsetValues.Fields.Item("Active").Value = "Y" Then
                                                                                            oTargetCostCenter.Active = SAPbobsCOM.BoYesNoEnum.tYES
                                                                                        Else
                                                                                            oTargetCostCenter.Active = SAPbobsCOM.BoYesNoEnum.tNO
                                                                                        End If
                                                                                        oTargetCostCenter.Effectivefrom = CDate(RsetValues.Fields.Item("ValidFrom").Value)
                                                                                        oTargetCostCenter.EffectiveTo = CDate(RsetValues.Fields.Item("ValidTo").Value)
                                                                                    Else
                                                                                        oUpdateProfitCenterParams.CenterCode = RsetValues.Fields.Item("PrcCode").Value
                                                                                        Try
                                                                                            oUpdateTargetProfitCenter = oTargetCostCenterServices.GetProfitCenter(oUpdateProfitCenterParams)
                                                                                        Catch ex As Exception

                                                                                        End Try
                                                                                        oUpdateTargetProfitCenter.CenterCode = RsetValues.Fields.Item("PrcCode").Value
                                                                                        oUpdateTargetProfitCenter.CenterName = RsetValues.Fields.Item("PrcName").Value
                                                                                        oUpdateTargetProfitCenter.CostCenterType = RsetValues.Fields.Item("CCTypeCode").Value
                                                                                        oUpdateTargetProfitCenter.InWhichDimension = RsetValues.Fields.Item("DimCode").Value
                                                                                        oUpdateTargetProfitCenter.GroupCode = RsetValues.Fields.Item("GrpCode").Value
                                                                                        If RsetValues.Fields.Item("Active").Value = "Y" Then
                                                                                            oUpdateTargetProfitCenter.Active = SAPbobsCOM.BoYesNoEnum.tYES
                                                                                        Else
                                                                                            oUpdateTargetProfitCenter.Active = SAPbobsCOM.BoYesNoEnum.tNO
                                                                                        End If
                                                                                        oUpdateTargetProfitCenter.Effectivefrom = CDate(RsetValues.Fields.Item("ValidFrom").Value)
                                                                                        oUpdateTargetProfitCenter.EffectiveTo = CDate(RsetValues.Fields.Item("ValidTo").Value)
                                                                                    End If
                                                                                End If
                                                                                If flag7 = False Then
                                                                                    Try
                                                                                        oTargetCostCenterServices.AddProfitCenter(DirectCast(oTargetCostCenter, SAPbobsCOM.ProfitCenter))
                                                                                        SucFlag = True
                                                                                        p_oSBOApplication.StatusBar.SetText("Adding Profit Center '" & oDT_Entities.Rows(J).Item("Code").ToString & "' Successful on.." & oDICompany(S).CompanyDB, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Profit Center '" & oDT_Entities.Rows(J).Item("Code").ToString & "' Successful on  '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                    Catch ex As Exception
                                                                                        sErrDesc = ex.Message
                                                                                        SucFlag = False
                                                                                        p_oSBOApplication.StatusBar.SetText("Adding Profit Center '" & oDT_Entities.Rows(J).Item("Code").ToString & "' Failed. on '" & oDICompany(S).CompanyDB & "'." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Profit Center'" & oDT_Entities.Rows(J).Item("Code").ToString & "' Failed on  '" & oDICompany(S).CompanyDB & "'" & sErrDesc, sFuncName)
                                                                                    End Try
                                                                                Else
                                                                                    Try
                                                                                        oTargetCostCenterServices.UpdateProfitCenter(DirectCast(oUpdateTargetProfitCenter, SAPbobsCOM.ProfitCenter))
                                                                                        SucFlag = True
                                                                                        p_oSBOApplication.StatusBar.SetText("Updating Profit Center '" & oDT_Entities.Rows(J).Item("Code").ToString & "' Successful on.." & oDICompany(S).CompanyDB, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating Profit Center '" & oDT_Entities.Rows(J).Item("Code").ToString & "' Successful on  '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                    Catch ex As Exception
                                                                                        sErrDesc = ex.Message
                                                                                        SucFlag = False
                                                                                        p_oSBOApplication.StatusBar.SetText("Updating Profit Center '" & oDT_Entities.Rows(J).Item("Code").ToString & "' Failed. on '" & oDICompany(S).CompanyDB & "'." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating Profit Center'" & oDT_Entities.Rows(J).Item("Code").ToString & "' Failed on  '" & oDICompany(S).CompanyDB & "'" & sErrDesc, sFuncName)
                                                                                    End Try
                                                                                End If
                                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetCostCenter)
                                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUpdateTargetProfitCenter)
                                                                            Catch ex As Exception
                                                                                p_oSBOApplication.StatusBar.SetText("Replicating Cost Centers Failed.." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                            End Try

                                                                            If SucFlag = False Then
118:
                                                                                sErrDesc = sErrDesc + " On.." + oDICompany(S).CompanyDB
                                                                                p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                Fllag = False
                                                                                Dim sqy As String = " UPDATE ""INTEGRATION"" SET ""SYNCSTATUS"" = 'NO', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '" & sErrDesc.ToString.Replace("'", """") & "'  WHERE ""MASTERTYPE"" = 'FINANCEMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                                Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                                Run.DoQuery(sqy)

                                                                                Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                                oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "NO"
                                                                                oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                                oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = sErrDesc

                                                                                For lCounter As Integer = 0 To UBound(oDICompany)
                                                                                    If Not oDICompany(lCounter) Is Nothing Then
                                                                                        If oDICompany(lCounter).Connected = True Then
                                                                                            If oDICompany(lCounter).InTransaction = True Then
                                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                                oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                                            End If
                                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                            oDICompany(lCounter).Disconnect()
                                                                                            oDICompany(lCounter) = Nothing
                                                                                        End If
                                                                                    End If
                                                                                Next
                                                                                Exit For
                                                                            Else
                                                                                Fllag = True
                                                                                SBO_Application.SetStatusBarMessage("Cost Center  '" & oDT_Entities.Rows(J).Item("Code").ToString & "' Replicated Successfully on.. " & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Cost Center  '" & oDT_Entities.Rows(J).Item("Code").ToString & "' Replicated Successfully on .. " & oDV_FINSetup.Item(S).Item("Name").ToString, sFuncName)
                                                                            End If
                                                                        Next
                                                                        For lCounter As Integer = 0 To UBound(oDICompany)
                                                                            If Not oDICompany(lCounter) Is Nothing Then
                                                                                If oDICompany(lCounter).Connected = True Then
                                                                                    If oDICompany(lCounter).InTransaction = True Then
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                        oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                                                    End If
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                    oDICompany(lCounter).Disconnect()
                                                                                    oDICompany(lCounter) = Nothing
                                                                                End If
                                                                            End If
                                                                        Next
                                                                        If Fllag = True Then
                                                                            Dim sqy As String = " Update ""INTEGRATION"" SET ""SYNCSTATUS"" = 'YES', ""SYNCDATE"" = CURRENT_DATE, ""ERRORMSG""= '' WHERE ""MASTERTYPE"" = 'FINANCEMASTER' AND ""UNIQUEID"" = '" & oDT_Entities.Rows(J).Item("UniqueNo").ToString & "'"
                                                                            Dim Run As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                            Run.DoQuery(sqy)

                                                                            Dim row1 As Integer = oDT_Entities.Rows(J).Item("SNo")
                                                                            oMatrix.Columns.Item("SyncStatus").Cells.Item(row1).Specific.value() = "YES"
                                                                            oMatrix.Columns.Item("SyncDate").Cells.Item(row1).Specific.value = Format(Now.Date, "yyyyMMdd")
                                                                            oMatrix.Columns.Item("ErrMsg").Cells.Item(row1).Specific.value = "SUCCESS"
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next
                                                        End If
                                                        Dim oPostingPeriod As SAPbouiCOM.CheckBox = oForm.Items.Item("chk_post").Specific
                                                        If (oPostingPeriod.Checked = True And oDV_FINSetup.Count > 0) Then
                                                            ''=========================================================================================================
                                                            '------------------------------- Posting Period Replication ----------------------------------------------
                                                            ''=========================================================================================================
                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Finance Setu for Posting Period Setting", sFuncName)
                                                            Dim rvcount As Integer = oDV_FINSetup.Count
                                                            Fllag = False
                                                            Dim NoEntries As Boolean = False
                                                            Dim CheckFlag As Boolean = False
                                                            Dim TestFlg As Boolean = False
                                                            Dim SucFlag As Boolean = False
                                                            oDV_FINSetup.RowFilter = "U_PostPeriod ='Y'"

                                                            Dim sqry3 As String = "select Distinct A.""AbsEntry"", ""PeriodCat"", ""FinancYear"", ""Year"",""PeriodName"", ""SubType"", ""PeriodNum""  from OACP A LEFT OUTER JOIN ""OFPR"" B ON B.""Category"" = A.""PeriodCat"" where B.""PeriodStat"" = 'N';"
                                                            Dim RPPSets As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                            RPPSets.DoQuery(sqry3)
                                                            oDT_PPlists = New DataTable
                                                            oDT_PPlists = ConvertRecordset(RPPSets)
                                                            Dim dtcount12 As Integer = oDT_PPlists.Rows.Count

                                                            ReDim oDICompany(oDV_FINSetup.Count)
                                                            Dim dvcount1 As Integer = oDV_FINSetup.Count
                                                            If oDV_FINSetup.Count > 0 Then
                                                                For S As Integer = 0 To oDV_FINSetup.Count - 1
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Connect To Target Company() for Posting Period Posting", sFuncName)
                                                                    SBO_Application.SetStatusBarMessage("Connecting to the Target Company" & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                    Dim targetent As String = oDV_FINSetup.Item(S).Item("Name").ToString
                                                                    If ConnectToTargetCompany(oDICompany(S), oDV_FINSetup.Item(S).Item("Name").ToString, oDV_FINSetup.Item(S).Item("U_UserName").ToString, oDV_FINSetup.Item(S).Item("U_Password").ToString, sErrDesc) <> RTN_SUCCESS Then
                                                                        GoTo 119
                                                                    End If
                                                                    SBO_Application.SetStatusBarMessage("Connecting to the target company Successfull " & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the target company Successfull " & oDICompany(S).CompanyDB, sFuncName)
                                                                    oDICompany(S).StartTransaction()

                                                                    SBO_Application.SetStatusBarMessage("Started Posting Periods Synchronization on" & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                                                                    Dim sboTargetCmpService As SAPbobsCOM.CompanyService = DirectCast(oDICompany(S).GetCompanyService(), SAPbobsCOM.CompanyService)
                                                                    Dim TargetPeriodCat As SAPbobsCOM.PeriodCategory
                                                                    Dim sboPeriodCatParams As SAPbobsCOM.PeriodCategoryParams
                                                                    'Dim sboFinancePeriods As SAPbobsCOM.FinancePeriods
                                                                    ' Dim sboFinancePeriod As SAPbobsCOM.FinancePeriod

                                                                    'Try
                                                                    ' Create an instance of the period category object
                                                                    TargetPeriodCat = DirectCast(sboTargetCmpService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiPeriodCategory), SAPbobsCOM.PeriodCategory)

                                                                    If oDT_PPlists.Rows.Count > 0 Then
                                                                        For Y As Integer = 0 To oDT_PPlists.Rows.Count - 1
                                                                            NoEntries = True
                                                                            Dim flg1 As Boolean = False
                                                                            Dim PeriodCat As String = String.Empty
                                                                            Dim BnCd As String = oDT_PPlists.Rows(Y).Item("PeriodCat").ToString
                                                                            Dim orsGroup As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                            orsGroup.DoQuery(String.Format("Select ""PeriodCat"" from ""OACP"" where ""PeriodCat"" = '{0}'", oDT_PPlists.Rows(Y).Item("PeriodCat").ToString))
                                                                            If orsGroup.RecordCount = 1 Then
                                                                                flg1 = True
                                                                                PeriodCat = orsGroup.Fields.Item(0).Value
                                                                                Continue For
                                                                            End If
                                                                            SucFlag = False

                                                                            If flg1 = False Then
                                                                                TargetPeriodCat.PeriodCategory = oDT_PPlists.Rows(Y).Item("PeriodCat").ToString
                                                                                TargetPeriodCat.PeriodName = oDT_PPlists.Rows(Y).Item("PeriodName").ToString
                                                                                Dim SubType As String = oDT_PPlists.Rows(Y).Item("SubType").ToString
                                                                                If SubType = "Y" Then
                                                                                    TargetPeriodCat.SubPeriodType = SAPbobsCOM.BoSubPeriodTypeEnum.spt_Year
                                                                                ElseIf SubType = "Q" Then
                                                                                    TargetPeriodCat.SubPeriodType = SAPbobsCOM.BoSubPeriodTypeEnum.spt_Quarters
                                                                                ElseIf SubType = "M" Then
                                                                                    TargetPeriodCat.SubPeriodType = SAPbobsCOM.BoSubPeriodTypeEnum.spt_Months
                                                                                ElseIf SubType = "D" Then
                                                                                    TargetPeriodCat.SubPeriodType = SAPbobsCOM.BoSubPeriodTypeEnum.spt_Days
                                                                                End If
                                                                                TargetPeriodCat.BeginningofFinancialYear = oDT_PPlists.Rows(Y).Item("FinancYear").ToString
                                                                            End If

                                                                            If flg1 = False Then
                                                                                Try
                                                                                    sboPeriodCatParams = sboTargetCmpService.CreatePeriod(TargetPeriodCat)
                                                                                    SucFlag = True
                                                                                    p_oSBOApplication.StatusBar.SetText("Adding Posting Period '" & oDT_PPlists.Rows(Y).Item("PeriodCat").ToString & "' Successful on.." & oDICompany(S).CompanyDB, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Posting Period '" & oDT_PPlists.Rows(Y).Item("PeriodCat").ToString & "' Successful on  '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                Catch ex As Exception
                                                                                    sErrDesc = ex.Message
                                                                                    SucFlag = False
                                                                                    p_oSBOApplication.StatusBar.SetText("Adding Posting Period'" & oDT_PPlists.Rows(Y).Item("PeriodCat").ToString & "' Failed. on '" & oDICompany(S).CompanyDB & "'." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Posting Period'" & oDT_PPlists.Rows(Y).Item("PeriodCat").ToString & "' Failed on  '" & oDICompany(S).CompanyDB & "'" & sErrDesc, sFuncName)
                                                                                End Try
                                                                            End If
                                                                        Next
                                                                    End If
                                                                    'System.Runtime.InteropServices.Marshal.ReleaseComObject(oBinLocation)
                                                                    'System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetBinLocation)
                                                                    If NoEntries = True Then
                                                                        If SucFlag = False Then
119:
                                                                            sErrDesc = sErrDesc + " On.." + oDICompany(S).CompanyDB
                                                                            p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                            Fllag = False
                                                                            Dim oEdit As SAPbouiCOM.StaticText = oForm.Items.Item("l_RepStat").Specific
                                                                            oEdit.Caption = sErrDesc
                                                                            oEdit.Item.FontSize = 10
                                                                            oEdit.Item.ForeColor = RGB(255, 0, 0)
                                                                            For lCounter As Integer = 0 To UBound(oDICompany)
                                                                                If Not oDICompany(lCounter) Is Nothing Then
                                                                                    If oDICompany(lCounter).Connected = True Then
                                                                                        If oDICompany(lCounter).InTransaction = True Then
                                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                            oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                                        End If
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                        oDICompany(lCounter).Disconnect()
                                                                                        oDICompany(lCounter) = Nothing
                                                                                    End If
                                                                                End If
                                                                            Next
                                                                            Exit For
                                                                        Else
                                                                            Fllag = True
                                                                            SBO_Application.SetStatusBarMessage("Posting Period Replicated Successfully on.. " & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Posting Period Replicated Successfully on .. " & oDV_FINSetup.Item(S).Item("Name").ToString, sFuncName)
                                                                        End If
                                                                    End If
                                                                Next
                                                                For lCounter As Integer = 0 To UBound(oDICompany)
                                                                    If Not oDICompany(lCounter) Is Nothing Then
                                                                        If oDICompany(lCounter).Connected = True Then
                                                                            If oDICompany(lCounter).InTransaction = True Then
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                                            End If
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                            oDICompany(lCounter).Disconnect()
                                                                            oDICompany(lCounter) = Nothing
                                                                        End If
                                                                    End If
                                                                Next
                                                                If Fllag = True Then
                                                                    Dim oEdit1 As SAPbouiCOM.StaticText = oForm.Items.Item("l_RepStat").Specific
                                                                    oEdit1.Caption = "Posting Period Replicated Successfully on Target Databases.."
                                                                    oEdit1.Item.FontSize = 10
                                                                    oEdit1.Item.ForeColor = RGB(0, 128, 0)
                                                                End If
                                                            End If
                                                        End If

                                                        Dim oExhRate As SAPbouiCOM.CheckBox = oForm.Items.Item("Chk_exch").Specific
                                                        If (oExhRate.Checked = True And oDV_FINSetup.Count > 0) Then
                                                            ''=========================================================================================================
                                                            '------------------------------- Exchagne Rates Replication ----------------------------------------------
                                                            ''=========================================================================================================
                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Finance Setup for Exchange Rates", sFuncName)
                                                            Dim rvcount As Integer = oDV_FINSetup.Count
                                                            Fllag = False
                                                            Dim NoRecords As Boolean = False
                                                            Dim CheckFlag As Boolean = False
                                                            Dim TestFlg As Boolean = False
                                                            Dim SucFlag As Boolean = False
                                                            oDV_FINSetup.RowFilter = "U_ExcRates ='Y'"

                                                            Dim sqry3 As String = "select ""RateDate"", ""Currency"",""Rate"" from ORTT where ""RateDate"" = CURRENT_DATE;"
                                                            Dim REPSets As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                            REPSets.DoQuery(sqry3)
                                                            oDT_ExchRates = New DataTable
                                                            oDT_ExchRates = ConvertRecordset(REPSets)
                                                            Dim dtcount12 As Integer = oDT_ExchRates.Rows.Count

                                                            ReDim oDICompany(oDV_FINSetup.Count)
                                                            Dim dvcount1 As Integer = oDV_FINSetup.Count
                                                            If oDV_FINSetup.Count > 0 Then
                                                                For S As Integer = 0 To oDV_FINSetup.Count - 1
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Connect To Target Company() for Posting Period Posting", sFuncName)
                                                                    SBO_Application.SetStatusBarMessage("Connecting to the Target Company" & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                    Dim targetent As String = oDV_FINSetup.Item(S).Item("Name").ToString
                                                                    If ConnectToTargetCompany(oDICompany(S), oDV_FINSetup.Item(S).Item("Name").ToString, oDV_FINSetup.Item(S).Item("U_UserName").ToString, oDV_FINSetup.Item(S).Item("U_Password").ToString, sErrDesc) <> RTN_SUCCESS Then
                                                                        GoTo 120
                                                                    End If
                                                                    SBO_Application.SetStatusBarMessage("Connecting to the target company Successful " & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the target company Successful " & oDICompany(S).CompanyDB, sFuncName)
                                                                    oDICompany(S).StartTransaction()

                                                                    SBO_Application.SetStatusBarMessage("Started Posting Periods Synchronization " & oDV_FINSetup.Item(S).Item("Name").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                                                                    Dim oSBObob As SAPbobsCOM.SBObob
                                                                    oSBObob = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)

                                                                    'If oDT_ExchRates.Rows.Count > 0 Then
                                                                    '    For Y As Integer = 0 To oDT_ExchRates.Rows.Count - 1
                                                                    '        NoRecords = True
                                                                    '        SucFlag = False
                                                                    '        Dim flg1 As Boolean = False

                                                                    '        Dim curr As String = oDT_ExchRates.Rows(Y).Item("Currency").ToString
                                                                    '        Dim currrate As Double = Convert.ToDouble(oDT_ExchRates.Rows(Y).Item("Rate").ToString)

                                                                    '        Dim RsetExchRates As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                    '        RsetExchRates.DoQuery(String.Format("Select ""CurrCode"" from ""OCRN"" where ""CurrCode"" = '{0}'", oDT_ExchRates.Rows(Y).Item("Currency").ToString))

                                                                    '        Try
                                                                    '            oSBObob.SetCurrencyRate(oDT_ExchRates.Rows(Y).Item("Currency").ToString, DateTime.Now, Convert.ToDouble(oDT_ExchRates.Rows(Y).Item("Rate")), True)
                                                                    '            SucFlag = True
                                                                    '            p_oSBOApplication.StatusBar.SetText("Adding Exchange Rate for Currency: '" & oDT_ExchRates.Rows(Y).Item("Currency").ToString & "' Successful on.." & oDICompany(S).CompanyDB, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                    '            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Exchange Rate for Currency: '" & oDT_ExchRates.Rows(Y).Item("Currency").ToString & "' Successful on  '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                    '        Catch ex As Exception
                                                                    '            sErrDesc = ex.Message
                                                                    '            SucFlag = False
                                                                    '            p_oSBOApplication.StatusBar.SetText("Adding Posting Period'" & oDT_ExchRates.Rows(Y).Item("Currency").ToString & "' Failed. on '" & oDICompany(S).CompanyDB & "'." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                    '            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Posting Period'" & oDT_ExchRates.Rows(Y).Item("Currency").ToString & "' Failed on  '" & oDICompany(S).CompanyDB & "'" & sErrDesc, sFuncName)
                                                                    '            GoTo 120
                                                                    '        End Try
                                                                    '    Next
                                                                    'End If
                                                                    If oDT_ExchRates.Rows.Count > 0 Then
                                                                        For Y As Integer = 0 To oDT_ExchRates.Rows.Count - 1
                                                                            NoRecords = True
                                                                            SucFlag = False
                                                                            Dim flg1 As Boolean = False

                                                                            Dim curr As String = oDT_ExchRates.Rows(Y).Item("Currency").ToString
                                                                            Dim currrate As Double = Convert.ToDouble(oDT_ExchRates.Rows(Y).Item("Rate").ToString)

                                                                            Dim RsetExchRates As SAPbobsCOM.Recordset = oDICompany(S).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                                            RsetExchRates.DoQuery(String.Format("Select ""CurrCode"" from ""OCRN"" where ""CurrCode"" = '{0}'", oDT_ExchRates.Rows(Y).Item("Currency").ToString))
                                                                            If RsetExchRates.RecordCount = 1 Then
                                                                                Try
                                                                                    oSBObob.SetCurrencyRate(oDT_ExchRates.Rows(Y).Item("Currency").ToString, DateTime.Now, Convert.ToDouble(oDT_ExchRates.Rows(Y).Item("Rate")), True)
                                                                                    SucFlag = True
                                                                                    p_oSBOApplication.StatusBar.SetText("Replicating Exchange Rate for Currency: '" & oDT_ExchRates.Rows(Y).Item("Currency").ToString & "' Successful on.." & oDICompany(S).CompanyDB, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Exchange Rate for Currency: '" & oDT_ExchRates.Rows(Y).Item("Currency").ToString & "' Successful on  '" & oDICompany(S).CompanyDB & "'", sFuncName)
                                                                                Catch ex As Exception
                                                                                    sErrDesc = ex.Message
                                                                                    SucFlag = False
                                                                                    p_oSBOApplication.StatusBar.SetText("Replicating Exchange Rate for Currency: '" & oDT_ExchRates.Rows(Y).Item("Currency").ToString & "' Failed. on '" & oDICompany(S).CompanyDB & "'." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Exchange Rate for Currency'" & oDT_ExchRates.Rows(Y).Item("Currency").ToString & "' Failed on  '" & oDICompany(S).CompanyDB & "'" & sErrDesc, sFuncName)
                                                                                    GoTo 120
                                                                                End Try
                                                                            Else
                                                                                p_oSBOApplication.StatusBar.SetText("Replication unsuccessful, Currency : '" & oDT_ExchRates.Rows(Y).Item("Currency").ToString & "' is not Exists on '" & oDICompany(S).CompanyDB & "'." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Replication unsuccessful, Currency '" & oDT_ExchRates.Rows(Y).Item("Currency").ToString & "' is not Exists on  '" & oDICompany(S).CompanyDB & "'" & sErrDesc, sFuncName)
                                                                            End If
                                                                        Next
                                                                    End If
                                                                    If NoRecords = True Then
                                                                        If SucFlag = False Then
120:
                                                                            sErrDesc = sErrDesc + " On.." + oDICompany(S).CompanyDB
                                                                            p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                                            Fllag = False
                                                                            Dim oEdit As SAPbouiCOM.StaticText = oForm.Items.Item("l_RepStat1").Specific
                                                                            oEdit.Caption = sErrDesc
                                                                            oEdit.Item.FontSize = 10
                                                                            oEdit.Item.ForeColor = RGB(255, 0, 0)
                                                                            For lCounter As Integer = 0 To UBound(oDICompany)
                                                                                If Not oDICompany(lCounter) Is Nothing Then
                                                                                    If oDICompany(lCounter).Connected = True Then
                                                                                        If oDICompany(lCounter).InTransaction = True Then
                                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                            oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                                        End If
                                                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                        oDICompany(lCounter).Disconnect()
                                                                                        oDICompany(lCounter) = Nothing
                                                                                    End If
                                                                                End If
                                                                            Next
                                                                            Exit For
                                                                        Else
                                                                            Fllag = True
                                                                            SBO_Application.SetStatusBarMessage("Exchange Rate Replicated Successfully on.. ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Exchange Rate Replicated Successfully on .. ", sFuncName)
                                                                        End If
                                                                    End If
                                                                Next

                                                                For lCounter As Integer = 0 To UBound(oDICompany)
                                                                    If Not oDICompany(lCounter) Is Nothing Then
                                                                        If oDICompany(lCounter).Connected = True Then
                                                                            If oDICompany(lCounter).InTransaction = True Then
                                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                                oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                                            End If
                                                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                            oDICompany(lCounter).Disconnect()
                                                                            oDICompany(lCounter) = Nothing
                                                                        End If
                                                                    End If
                                                                Next
                                                                If Fllag = True Then
                                                                    Dim oEdit1 As SAPbouiCOM.StaticText = oForm.Items.Item("l_RepStat1").Specific
                                                                    oEdit1.Caption = "Exchange Rate Replicated Successfully on Target Databases.."
                                                                    oEdit1.Item.FontSize = 10
                                                                    oEdit1.Item.ForeColor = RGB(0, 128, 0)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    oMatrix.AutoResizeColumns()
                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS .......", sFuncName)
                                                Catch ex As Exception
                                                    BubbleEvent = False
                                                    sErrDesc = ex.Message
                                                    p_oSBOApplication.SetStatusBarMessage("Replication Failed... " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium)
                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                                    WriteToLogFile(Err.Description, sFuncName)
                                                    ShowErr(sErrDesc)
                                                Finally
                                                    EndStatus(sErrDesc)
                                                    oForm.Items.Item("Replicate").Enabled = True
                                                End Try
                                            Catch ex As Exception

                                            End Try
                                    End Select
                            End Select
                    End Select
                End If
                If pVal.Before_Action = True Then
                    Select Case pVal.FormUID
                    End Select
                End If
                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                BubbleEvent = False
                sErrDesc = exc.Message
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                WriteToLogFile(Err.Description, sFuncName)
                ShowErr(sErrDesc)
            End Try

        End Sub

        Function LoadReplicationDetails()
            Dim FormUID As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
            Dim sFuncName As String = String.Empty
            Dim oMatrix As SAPbouiCOM.Matrix = Nothing
            Dim oRset As SAPbobsCOM.Recordset = Nothing
            Dim sQry As String = String.Empty

            Try
                FormUID.Freeze(True)
                sFuncName = "Load_Replication_Details()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Replication Details()", sFuncName)

                sQry = "select  ROW_NUMBER() OVER (ORDER BY T0.""UNIQUEID"") AS ""ROW"", T0.""UNIQUEID"", T0.""TRANSTYPE"", T0.""CODE"", T0.""NAME"", T0.""SYNCSTATUS"" from ""INTEGRATION"" T0 WHERE T0.""MASTERTYPE"" = case '" & FormUID.Items.Item("c_masttype").Specific.value & "' when 'B' THEN 'BPMASTER' WHEN 'I' THEN 'ITEMMASTER' WHEN 'F' THEN 'FINANCEMASTER' END AND T0.""SYNCSTATUS"" = 'NO'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL " & sQry, sFuncName)
                oMatrix = FormUID.Items.Item("5").Specific

                'Creating Data table to load the query results..
                Try
                    FormUID.DataSources.DataTables.Add("REP")
                Catch ex As Exception

                End Try
                ' Loading the query results into datatable.
                FormUID.DataSources.DataTables.Item("REP").ExecuteQuery(sQry)

                'Loading into matrix....
                oMatrix.Clear()
                FormUID.Items.Item("5").Specific.columns.item("LineID").databind.bind("REP", "ROW")
                FormUID.Items.Item("5").Specific.columns.item("UniqueId").databind.bind("REP", "UNIQUEID")
                FormUID.Items.Item("5").Specific.columns.item("TransType").databind.bind("REP", "TRANSTYPE")
                FormUID.Items.Item("5").Specific.columns.item("Code").databind.bind("REP", "CODE")
                FormUID.Items.Item("5").Specific.columns.item("Name").databind.bind("REP", "NAME")
                FormUID.Items.Item("5").Specific.columns.item("SyncStatus").databind.bind("REP", "SYNCSTATUS")
                FormUID.Items.Item("5").Specific.LoadFromDataSource()

                Dim oEdit As SAPbouiCOM.StaticText = FormUID.Items.Item("l_RepStat").Specific
                oEdit.Caption = ""
                Dim oEdit1 As SAPbouiCOM.StaticText = FormUID.Items.Item("l_RepStat1").Specific
                oEdit1.Caption = ""
                Dim oCmb As SAPbouiCOM.CheckBox = FormUID.Items.Item("Chk_Prices").Specific
                oCmb.Checked = False
                Dim oCmb1 As SAPbouiCOM.CheckBox = FormUID.Items.Item("chk_post").Specific
                oCmb1.Checked = False
                Dim oCmb2 As SAPbouiCOM.CheckBox = FormUID.Items.Item("Chk_exch").Specific
                oCmb2.Checked = False
                Dim oCmb3 As SAPbouiCOM.CheckBox = FormUID.Items.Item("Chk_bin").Specific
                oCmb3.Checked = False
                Dim oCmb4 As SAPbouiCOM.CheckBox = FormUID.Items.Item("Chk_Select").Specific
                oCmb4.Checked = False
               
                LoadReplicationDetails = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                oMatrix.AutoResizeColumns()
                FormUID.Freeze(False)
            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText("Loading Replication Details : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                LoadReplicationDetails = RTN_ERROR
                FormUID.Freeze(False)
            End Try
            Return RTN_SUCCESS
        End Function

        Sub AddMenuItems()
            Dim oMenus As SAPbouiCOM.Menus
            Dim oMenuItem As SAPbouiCOM.MenuItem
            oMenus = SBO_Application.Menus

            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = (SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams))
            oMenuItem = SBO_Application.Menus.Item("43520") 'Modules

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "MDR"
            oCreationPackage.String = "Master Data Replication"
            oCreationPackage.Enabled = True
            oCreationPackage.Position = -1

            oCreationPackage.Image = System.Windows.Forms.Application.StartupPath & "\MDR.bmp"
            oMenus = oMenuItem.SubMenus

            Try
                'If the manu already exists this code will fail
                If Not p_oSBOApplication.Menus.Exists("MDR") Then
                    oMenus.AddEx(oCreationPackage)
                End If

            Catch
            End Try


            Try
                'Get the menu collection of the newly added pop-up item
                oMenuItem = SBO_Application.Menus.Item("MDR")
                oMenus = oMenuItem.SubMenus

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
                oCreationPackage.UniqueID = "RS"
                oCreationPackage.String = "Replication Setup"
                oCreationPackage.Enabled = True
                oCreationPackage.Position = 0

                '' oCreationPackage.Image = System.Windows.Forms.Application.StartupPath & "\MDR.bmp"
                oMenus = oMenuItem.SubMenus


                'If the manu already exists this code will fail
                If Not p_oSBOApplication.Menus.Exists("RS") Then
                    oMenus.AddEx(oCreationPackage)
                End If


                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
                oCreationPackage.UniqueID = "MR"
                oCreationPackage.String = "Master Replication"
                oCreationPackage.Enabled = True
                oCreationPackage.Position = 1

                '' oCreationPackage.Image = System.Windows.Forms.Application.StartupPath & "\MDR.bmp"
                oMenus = oMenuItem.SubMenus


                'If the manu already exists this code will fail
                If Not p_oSBOApplication.Menus.Exists("MR") Then
                    oMenus.AddEx(oCreationPackage)
                End If



                oMenuItem = SBO_Application.Menus.Item("RS")
                oMenus = oMenuItem.SubMenus

                'Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "TDBL"
                oCreationPackage.String = "Target Database List"

                If Not p_oSBOApplication.Menus.Exists("TDBL") Then
                    oMenus.AddEx(oCreationPackage)
                End If

                'Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "BPMS"
                oCreationPackage.String = "BP Master Setup"


                If Not p_oSBOApplication.Menus.Exists("BPMS") Then
                    oMenus.AddEx(oCreationPackage)
                End If

                'Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "ITMS"
                oCreationPackage.String = "Item Master Setup"


                If Not p_oSBOApplication.Menus.Exists("ITMS") Then
                    oMenus.AddEx(oCreationPackage)
                End If

                'Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "FINS"
                oCreationPackage.String = "Financial Setup"


                If Not p_oSBOApplication.Menus.Exists("FINS") Then
                    oMenus.AddEx(oCreationPackage)
                End If

                oMenuItem = SBO_Application.Menus.Item("MR")
                oMenus = oMenuItem.SubMenus

                'Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "MDRU"
                oCreationPackage.String = "Master Data Replication"


                If Not p_oSBOApplication.Menus.Exists("MDRU") Then
                    oMenus.AddEx(oCreationPackage)
                End If


            Catch
                'Menu already exists
                SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End Try
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

        Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent

            Select Case BusinessObjectInfo.FormTypeEx

                Case "804"

                    If BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then ' Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then


                        Dim ival As Integer
                        Dim IsError As Boolean
                        Dim iErr As Integer = 0
                        Dim sErr As String = String.Empty

                        Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(804, p_FrmType)
                        Dim sAcctCode As String = String.Empty
                        sAcctCode = oForm.DataSources.DBDataSources.Item(0).GetValue("AcctCode", "0").ToString.Trim
                        Dim oChartofAccounts As SAPbobsCOM.ChartOfAccounts = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts)
                        If oChartofAccounts.GetByKey(sAcctCode) Then
                            oChartofAccounts.FrozenFor = SAPbobsCOM.BoYesNoEnum.tYES
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Set InActive for Account Code " & sAcctCode, sFuncName)
                            ival = oChartofAccounts.Update()
                            If ival <> 0 Then
                                IsError = True
                                p_oDICompany.GetLastError(iErr, sErr)
                                p_oSBOApplication.StatusBar.SetText(sErr, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                                Exit Sub
                            End If
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sAcctCode, sFuncName)
                        End If

                    End If


            End Select

        End Sub

        Private Sub SBO_Application_LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.LayoutKeyEvent

        End Sub

        Public Function ConvertRecordset(ByVal SAPRecordset As SAPbobsCOM.Recordset) As DataTable

            '\ This function will take an SAP recordset from the SAPbobsCOM library and convert it to a more
            '\ easily used ADO.NET datatable which can be used for data binding much easier.

            Dim dtTable As New DataTable
            Dim NewCol As DataColumn
            Dim NewRow As DataRow
            Dim ColCount As Integer

            Try
                For ColCount = 0 To SAPRecordset.Fields.Count - 1
                    NewCol = New DataColumn(SAPRecordset.Fields.Item(ColCount).Name)
                    dtTable.Columns.Add(NewCol)
                Next

                Do Until SAPRecordset.EoF

                    NewRow = dtTable.NewRow
                    'populate each column in the row we're creating
                    For ColCount = 0 To SAPRecordset.Fields.Count - 1

                        NewRow.Item(SAPRecordset.Fields.Item(ColCount).Name) = SAPRecordset.Fields.Item(ColCount).Value

                    Next

                    'Add the row to the datatable
                    dtTable.Rows.Add(NewRow)


                    SAPRecordset.MoveNext()
                Loop

                Return dtTable

            Catch ex As Exception
                MsgBox(ex.ToString & Chr(10) & "Error converting SAP Recordset to DataTable", MsgBoxStyle.Exclamation)
                Exit Function
            End Try
        End Function

        Public Function CreateBPMaster(CardCode As String, targetCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
            Dim errCode As Integer
            Dim ErrorCode As Long
            Dim errMess As String = ""
            Dim oBP As SAPbobsCOM.BusinessPartners
            Dim oTargetBP As SAPbobsCOM.BusinessPartners
            'Dim p_oDICompany As New SAPbobsCOM.Company
            Dim ors As SAPbobsCOM.Recordset = Nothing
            Dim orsTarget As SAPbobsCOM.Recordset = Nothing
            Dim GroupName As String = ""
            Dim oDVContact As DataView = Nothing
            Dim sFuncName = "CreateBPMaster()"
            oBP = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oTargetBP = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            If oBP.GetByKey(CardCode) Then
                Try
                    ors = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim orsB As SAPbobsCOM.Recordset = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    orsTarget = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


                    If oTargetBP.GetByKey(oBP.CardCode) Then
                        Dim sSQL As String = "select ROW_NUMBER() OVER(ORDER BY T1.""CntctCode"" ) -1 ""No"", T1.""CntctCode"" , T1.""Name"" , T1.""Position""  from" & _
                         """OCPR"" T1  where T1.""CardCode"" ='" & CardCode & "' order by T1.""CntctCode"" "
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Contact Info " & sSQL, sFuncName)
                        orsB.DoQuery(sSQL)
                        oDVContact = New DataView(ConvertRecordset(orsB))

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating BP " & CardCode, sFuncName)

                        oTargetBP.CardName = oBP.CardName
                        oTargetBP.CardType = oBP.CardType
                        oTargetBP.CardForeignName = oBP.CardForeignName
                        oTargetBP.CompanyPrivate = oBP.CompanyPrivate
                        oTargetBP.DiscountPercent = oBP.DiscountPercent
                        oTargetBP.Address = oBP.Address
                        oTargetBP.EmailAddress = oBP.EmailAddress
                        oTargetBP.Phone1 = oBP.Phone1
                        oTargetBP.Phone2 = oBP.Phone2
                        oTargetBP.Cellular = oBP.Cellular
                        oTargetBP.Fax = oBP.Fax
                        oTargetBP.Password = oBP.Password
                        oTargetBP.BusinessType = oBP.BusinessType
                        oTargetBP.AdditionalID = oBP.AdditionalID
                        oTargetBP.VatIDNum = oBP.VatIDNum
                        oTargetBP.FederalTaxID = oBP.FederalTaxID
                        oTargetBP.Notes = oBP.Notes
                        oTargetBP.FreeText = oBP.FreeText
                        oTargetBP.AliasName = oBP.AliasName
                        oTargetBP.GlobalLocationNumber = oBP.GlobalLocationNumber
                        oTargetBP.Valid = oBP.Valid
                        oTargetBP.Frozen = oBP.Frozen

                        oTargetBP.Website = oBP.Website
                        oTargetBP.UnifiedFederalTaxID = oBP.UnifiedFederalTaxID

                        Dim orsGroup As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        orsGroup.DoQuery(String.Format("Select ""GroupName"" from ""OCRG"" where ""GroupCode"" = {0}", oBP.GroupCode))
                        GroupName = orsGroup.Fields.Item(0).Value

                        orsGroup = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        orsGroup.DoQuery(String.Format("Select ""GroupCode"" from ""OCRG"" where ""GroupName"" = '{0}'", GroupName))
                        If orsGroup.RecordCount = 1 Then
                            oTargetBP.GroupCode = orsGroup.Fields.Item(0).Value
                        End If

                        If oTargetBP.Valid = SAPbobsCOM.BoYesNoEnum.tYES Then
                            oTargetBP.ValidFrom = oBP.ValidFrom
                            oTargetBP.ValidTo = oBP.ValidTo
                            oTargetBP.ValidRemarks = oBP.ValidRemarks
                        End If
                        If oTargetBP.Frozen = SAPbobsCOM.BoYesNoEnum.tYES Then
                            oTargetBP.FrozenFrom = oBP.FrozenFrom
                            oTargetBP.FrozenTo = oBP.FrozenTo
                            oTargetBP.FrozenRemarks = oBP.FrozenRemarks
                        End If
                        If oTargetBP.Addresses.Count > 0 Then

                            Dim delete As Boolean = False
                            For i As Integer = 0 To oTargetBP.Addresses.Count - 1
                                oTargetBP.Addresses.SetCurrentLine(oTargetBP.Addresses.Count - 1)
                                oTargetBP.Addresses.Delete()
                                If oTargetBP.Addresses.Count = 0 Then
                                    Exit For
                                End If
                            Next
                        End If
                        'Handle add update/add new Address
                        For i As Integer = 0 To oBP.Addresses.Count - 1
                            oBP.Addresses.SetCurrentLine(i)
                            oTargetBP.Addresses.AddressName = oBP.Addresses.AddressName
                            oTargetBP.Addresses.AddressName2 = oBP.Addresses.AddressName2
                            oTargetBP.Addresses.AddressName3 = oBP.Addresses.AddressName3
                            oTargetBP.Addresses.AddressType = oBP.Addresses.AddressType
                            oTargetBP.Addresses.Block = oBP.Addresses.Block
                            oTargetBP.Addresses.City = oBP.Addresses.City
                            oTargetBP.Addresses.County = oBP.Addresses.County
                            oTargetBP.Addresses.Country = oBP.Addresses.Country
                            oTargetBP.Addresses.StreetNo = oBP.Addresses.StreetNo
                            oTargetBP.Addresses.TypeOfAddress = oBP.Addresses.TypeOfAddress
                            oTargetBP.Addresses.State = oBP.Addresses.State
                            oTargetBP.Addresses.ZipCode = oBP.Addresses.ZipCode
                            oTargetBP.Addresses.Street = oBP.Addresses.Street
                            oTargetBP.Addresses.BuildingFloorRoom = oBP.Addresses.BuildingFloorRoom
                            oTargetBP.Addresses.GlobalLocationNumber = oBP.Addresses.GlobalLocationNumber
                            oTargetBP.Addresses.Add()
                        Next
                        oTargetBP.BilltoDefault = oBP.BilltoDefault
                        oTargetBP.ShipToDefault = oBP.ShipToDefault

                        oTargetBP.PayTermsGrpCode = oBP.PayTermsGrpCode
                        oTargetBP.IntrestRatePercent = oBP.IntrestRatePercent
                        oTargetBP.PriceListNum = oBP.PriceListNum
                        oTargetBP.DiscountPercent = oBP.DiscountPercent

                        oTargetBP.CreditLimit = oBP.CreditLimit
                        oTargetBP.MaxCommitment = oBP.MaxCommitment
                        oTargetBP.EffectiveDiscount = oBP.EffectiveDiscount

                        oTargetBP.HouseBank = oBP.HouseBank
                        oTargetBP.HouseBankAccount = oBP.HouseBankAccount
                        oTargetBP.HouseBankBranch = oBP.HouseBankBranch
                        oTargetBP.HouseBankCountry = oBP.HouseBankCountry
                        oTargetBP.BPPaymentMethods.PaymentMethodCode = oBP.BPPaymentMethods.PaymentMethodCode

                        For imjs As Integer = 1 To oBP.BPPaymentMethods.PaymentMethodCode.Count
                            oTargetBP.BPPaymentMethods.PaymentMethodCode = oBP.BPPaymentMethods.PaymentMethodCode
                            oTargetBP.BPPaymentMethods.Add()
                            'oRset_Tar.MoveNext()
                        Next imjs

                        ' ''BP Bank Details 
                        If oTargetBP.BPBankAccounts.Count > 0 Then

                            Dim delete As Boolean = False
                            For i As Integer = 0 To oTargetBP.BPBankAccounts.Count - 1
                                oTargetBP.BPBankAccounts.SetCurrentLine(oTargetBP.BPBankAccounts.Count - 1)
                                oTargetBP.BPBankAccounts.Delete()
                                If oTargetBP.BPBankAccounts.Count = 0 Then
                                    Exit For
                                End If
                            Next
                        End If
                        For i As Integer = 0 To oBP.BPBankAccounts.Count - 1
                            oBP.BPBankAccounts.SetCurrentLine(i)
                            'orsTarget.DoQuery(String.Format("SELECT ""BankCode"" from ""ODSC"" where ""BankCode"" =  '{0}'", oBP.BPBankAccounts.BankCode))
                            'If orsTarget.RecordCount = 1 Then
                            oTargetBP.BPBankAccounts.BankCode = oBP.BPBankAccounts.BankCode
                            oTargetBP.BPBankAccounts.Country = oBP.BPBankAccounts.Country
                            oTargetBP.BPBankAccounts.BPCode = oBP.BPBankAccounts.BPCode
                            oTargetBP.BPBankAccounts.AccountNo = oBP.BPBankAccounts.AccountNo
                            oTargetBP.BPBankAccounts.AccountName = oBP.BPBankAccounts.AccountName
                            oTargetBP.BPBankAccounts.Branch = oBP.BPBankAccounts.Branch
                            oTargetBP.BPBankAccounts.BICSwiftCode = oBP.BPBankAccounts.BICSwiftCode
                            oTargetBP.BPBankAccounts.InternalKey = oBP.BPBankAccounts.InternalKey
                            oTargetBP.BPBankAccounts.ControlKey = oBP.BPBankAccounts.ControlKey
                            oTargetBP.BPBankAccounts.IBAN = oBP.BPBankAccounts.IBAN
                            oTargetBP.BPBankAccounts.Street = oBP.BPBankAccounts.Street
                            oTargetBP.BPBankAccounts.State = oBP.BPBankAccounts.State
                            oTargetBP.BPBankAccounts.Block = oBP.BPBankAccounts.Block
                            oTargetBP.BPBankAccounts.BuildingFloorRoom = oBP.BPBankAccounts.BuildingFloorRoom
                            oTargetBP.BPBankAccounts.City = oBP.BPBankAccounts.City
                            oTargetBP.BPBankAccounts.MandateID = oBP.BPBankAccounts.MandateID

                            oTargetBP.BPBankAccounts.Add()

                        Next
                        orsTarget.DoQuery(String.Format("SELECT ""BankCode"",""CountryCod"" from ""ODSC"" where ""BankCode"" =  '{0}'", oBP.DefaultBankCode))
                        If orsTarget.RecordCount = 1 Then
                            Dim sBankcode As String = orsTarget.Fields.Item(0).Value
                            Dim sCtrycode As String = orsTarget.Fields.Item(1).Value

                            oTargetBP.DefaultBankCode = oBP.DefaultBankCode
                            oTargetBP.DefaultAccount = oBP.DefaultAccount

                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting Contact Employees", sFuncName)

                        If oTargetBP.ContactEmployees.Count = 1 Then
                            For imjs As Integer = 0 To oBP.ContactEmployees.Count - 1
                                oBP.ContactEmployees.SetCurrentLine(imjs)
                                'oTargetBP.ContactEmployees.Add()
                                oTargetBP.ContactEmployees.Name = oBP.ContactEmployees.Name
                                oTargetBP.ContactEmployees.Position = oBP.ContactEmployees.Position
                                oTargetBP.ContactEmployees.E_Mail = oBP.ContactEmployees.E_Mail
                                oTargetBP.ContactEmployees.Phone1 = oBP.ContactEmployees.Phone1
                                oTargetBP.ContactEmployees.Phone2 = oBP.ContactEmployees.Phone2

                                oTargetBP.ContactEmployees.FirstName = oBP.ContactEmployees.FirstName
                                oTargetBP.ContactEmployees.MiddleName = oBP.ContactEmployees.MiddleName
                                oTargetBP.ContactEmployees.LastName = oBP.ContactEmployees.LastName
                                oTargetBP.ContactEmployees.Title = oBP.ContactEmployees.Title
                                oTargetBP.ContactEmployees.Address = oBP.ContactEmployees.Address
                                oTargetBP.ContactEmployees.MobilePhone = oBP.ContactEmployees.MobilePhone
                                oTargetBP.ContactEmployees.Fax = oBP.ContactEmployees.Fax
                                oTargetBP.ContactEmployees.Pager = oBP.ContactEmployees.Pager
                                oTargetBP.ContactEmployees.Remarks1 = oBP.ContactEmployees.Remarks1
                                oTargetBP.ContactEmployees.Remarks2 = oBP.ContactEmployees.Remarks2
                                oTargetBP.ContactEmployees.Password = oBP.ContactEmployees.Password
                                oTargetBP.ContactEmployees.PlaceOfBirth = oBP.ContactEmployees.PlaceOfBirth
                                oTargetBP.ContactEmployees.DateOfBirth = oBP.ContactEmployees.DateOfBirth
                                oTargetBP.ContactEmployees.Gender = oBP.ContactEmployees.Gender
                                oTargetBP.ContactEmployees.Profession = oBP.ContactEmployees.Profession
                                oTargetBP.ContactEmployees.CityOfBirth = oBP.ContactEmployees.CityOfBirth
                                oTargetBP.ContactEmployees.Add()
                            Next
                        ElseIf oTargetBP.ContactEmployees.Count > 0 Then
                            For imjs As Integer = 0 To oBP.ContactEmployees.Count - 1
                                oBP.ContactEmployees.SetCurrentLine(imjs)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("oBP.ContactEmployees.Name " & oBP.ContactEmployees.Name, sFuncName)
                                If oBP.ContactEmployees.Name = "" Then Continue For
                                oDVContact.RowFilter = "Name='" & oBP.ContactEmployees.Name & "'"
                                If oDVContact.Count > 0 Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Index " & oDVContact(0)("No").ToString(), sFuncName)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Name " & oDVContact(0)("Name").ToString(), sFuncName)
                                    oTargetBP.ContactEmployees.SetCurrentLine(Convert.ToInt32(oDVContact(0)("No").ToString()))
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigned", sFuncName)
                                    oTargetBP.ContactEmployees.Name = oBP.ContactEmployees.Name
                                    oTargetBP.ContactEmployees.Position = oBP.ContactEmployees.Position
                                    oTargetBP.ContactEmployees.E_Mail = oBP.ContactEmployees.E_Mail
                                    oTargetBP.ContactEmployees.Phone1 = oBP.ContactEmployees.Phone1
                                    oTargetBP.ContactEmployees.Phone2 = oBP.ContactEmployees.Phone2

                                    oTargetBP.ContactEmployees.FirstName = oBP.ContactEmployees.FirstName
                                    oTargetBP.ContactEmployees.MiddleName = oBP.ContactEmployees.MiddleName
                                    oTargetBP.ContactEmployees.LastName = oBP.ContactEmployees.LastName
                                    oTargetBP.ContactEmployees.Title = oBP.ContactEmployees.Title
                                    oTargetBP.ContactEmployees.Address = oBP.ContactEmployees.Address
                                    oTargetBP.ContactEmployees.MobilePhone = oBP.ContactEmployees.MobilePhone
                                    oTargetBP.ContactEmployees.Fax = oBP.ContactEmployees.Fax
                                    oTargetBP.ContactEmployees.Pager = oBP.ContactEmployees.Pager
                                    oTargetBP.ContactEmployees.Remarks1 = oBP.ContactEmployees.Remarks1
                                    oTargetBP.ContactEmployees.Remarks2 = oBP.ContactEmployees.Remarks2
                                    oTargetBP.ContactEmployees.Password = oBP.ContactEmployees.Password
                                    oTargetBP.ContactEmployees.PlaceOfBirth = oBP.ContactEmployees.PlaceOfBirth
                                    oTargetBP.ContactEmployees.DateOfBirth = oBP.ContactEmployees.DateOfBirth
                                    oTargetBP.ContactEmployees.Gender = oBP.ContactEmployees.Gender
                                    oTargetBP.ContactEmployees.Profession = oBP.ContactEmployees.Profession
                                    oTargetBP.ContactEmployees.CityOfBirth = oBP.ContactEmployees.CityOfBirth
                                Else

                                    oTargetBP.ContactEmployees.Add()
                                    oTargetBP.ContactEmployees.Name = oBP.ContactEmployees.Name
                                    oTargetBP.ContactEmployees.Position = oBP.ContactEmployees.Position
                                    oTargetBP.ContactEmployees.E_Mail = oBP.ContactEmployees.E_Mail
                                    oTargetBP.ContactEmployees.Phone1 = oBP.ContactEmployees.Phone1
                                    oTargetBP.ContactEmployees.Phone2 = oBP.ContactEmployees.Phone2

                                    oTargetBP.ContactEmployees.FirstName = oBP.ContactEmployees.FirstName
                                    oTargetBP.ContactEmployees.MiddleName = oBP.ContactEmployees.MiddleName
                                    oTargetBP.ContactEmployees.LastName = oBP.ContactEmployees.LastName
                                    oTargetBP.ContactEmployees.Title = oBP.ContactEmployees.Title
                                    oTargetBP.ContactEmployees.Address = oBP.ContactEmployees.Address
                                    oTargetBP.ContactEmployees.MobilePhone = oBP.ContactEmployees.MobilePhone
                                    oTargetBP.ContactEmployees.Fax = oBP.ContactEmployees.Fax
                                    oTargetBP.ContactEmployees.Pager = oBP.ContactEmployees.Pager
                                    oTargetBP.ContactEmployees.Remarks1 = oBP.ContactEmployees.Remarks1
                                    oTargetBP.ContactEmployees.Remarks2 = oBP.ContactEmployees.Remarks2
                                    oTargetBP.ContactEmployees.Password = oBP.ContactEmployees.Password
                                    oTargetBP.ContactEmployees.PlaceOfBirth = oBP.ContactEmployees.PlaceOfBirth
                                    oTargetBP.ContactEmployees.DateOfBirth = oBP.ContactEmployees.DateOfBirth
                                    oTargetBP.ContactEmployees.Gender = oBP.ContactEmployees.Gender
                                    oTargetBP.ContactEmployees.Profession = oBP.ContactEmployees.Profession
                                    oTargetBP.ContactEmployees.CityOfBirth = oBP.ContactEmployees.CityOfBirth
                                End If
                                ''oTargetBP.ContactEmployees.Add()
                            Next
                        End If
                        oTargetBP.ContactPerson = oBP.ContactPerson
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating BP " & CardCode, sFuncName)
                        ErrorCode = oTargetBP.Update()
                        If ErrorCode <> 0 Then
                            targetCompany.GetLastError(ErrorCode, sErrDesc)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating BP Code '" & CardCode & "' Failed on '" & targetCompany.CompanyDB & "'" + " - " + sErrDesc, sFuncName)
                            Return RTN_ERROR
                        Else
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("BP Code: '" & CardCode & "' Created Successfully on '" & targetCompany.CompanyDB & "'", sFuncName)
                            Return RTN_SUCCESS
                        End If
                    Else
                        oTargetBP.CardCode = oBP.CardCode
                        oTargetBP.CardName = oBP.CardName
                        oTargetBP.CardType = oBP.CardType
                        oTargetBP.CardForeignName = oBP.CardForeignName

                        oTargetBP.CompanyPrivate = oBP.CompanyPrivate
                        If oBP.CardType = SAPbobsCOM.BoCardTypes.cCustomer Then
                            oTargetBP.Currency = "##"
                        End If
                        oTargetBP.DiscountPercent = oBP.DiscountPercent
                        oTargetBP.Address = oBP.Address
                        oTargetBP.EmailAddress = oBP.EmailAddress
                        oTargetBP.Phone1 = oBP.Phone1
                        oTargetBP.Phone2 = oBP.Phone2
                        oTargetBP.Cellular = oBP.Cellular
                        oTargetBP.Fax = oBP.Fax
                        oTargetBP.Password = oBP.Password
                        oTargetBP.BusinessType = oBP.BusinessType
                        oTargetBP.AdditionalID = oBP.AdditionalID
                        oTargetBP.VatIDNum = oBP.VatIDNum
                        oTargetBP.FederalTaxID = oBP.FederalTaxID
                        oTargetBP.Notes = oBP.Notes
                        oTargetBP.FreeText = oBP.FreeText
                        oTargetBP.AliasName = oBP.AliasName
                        oTargetBP.GlobalLocationNumber = oBP.GlobalLocationNumber
                        oTargetBP.Valid = oBP.Valid
                        oTargetBP.Frozen = oBP.Frozen
                        oTargetBP.Website = oBP.Website
                        oTargetBP.UnifiedFederalTaxID = oBP.UnifiedFederalTaxID

                        ors.DoQuery(String.Format("Select ""GroupName"" from ""OCRG"" where ""GroupCode"" = {0}", oBP.GroupCode))
                        GroupName = ors.Fields.Item(0).Value

                        ors = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ors.DoQuery(String.Format("Select ""GroupCode"" from ""OCRG"" where ""GroupName"" = '{0}'", GroupName))

                        If ors.RecordCount = 1 Then
                            oTargetBP.GroupCode = ors.Fields.Item(0).Value
                        End If

                        'oTargetBP.DebitorAccount = oBP.DebitorAccount
                        If oTargetBP.Valid = SAPbobsCOM.BoYesNoEnum.tYES Then
                            oTargetBP.ValidFrom = oBP.ValidFrom
                            oTargetBP.ValidTo = oBP.ValidTo
                            oTargetBP.ValidRemarks = oBP.ValidRemarks
                        End If
                        If oTargetBP.Frozen = SAPbobsCOM.BoYesNoEnum.tYES Then
                            oTargetBP.FrozenFrom = oBP.FrozenFrom
                            oTargetBP.FrozenTo = oBP.FrozenTo
                            oTargetBP.FrozenRemarks = oBP.FrozenRemarks
                        End If

                        For i As Integer = 0 To oBP.Addresses.Count - 1
                            oBP.Addresses.SetCurrentLine(i)
                            oTargetBP.Addresses.AddressName = oBP.Addresses.AddressName
                            oTargetBP.Addresses.AddressName2 = oBP.Addresses.AddressName2
                            oTargetBP.Addresses.AddressName3 = oBP.Addresses.AddressName3
                            oTargetBP.Addresses.AddressType = oBP.Addresses.AddressType
                            oTargetBP.Addresses.Block = oBP.Addresses.Block
                            oTargetBP.Addresses.City = oBP.Addresses.City
                            oTargetBP.Addresses.County = oBP.Addresses.County
                            oTargetBP.Addresses.Country = oBP.Addresses.Country
                            oTargetBP.Addresses.StreetNo = oBP.Addresses.StreetNo
                            oTargetBP.Addresses.TypeOfAddress = oBP.Addresses.TypeOfAddress
                            oTargetBP.Addresses.State = oBP.Addresses.State
                            oTargetBP.Addresses.ZipCode = oBP.Addresses.ZipCode
                            oTargetBP.Addresses.Street = oBP.Addresses.Street
                            oTargetBP.Addresses.BuildingFloorRoom = oBP.Addresses.BuildingFloorRoom
                            oTargetBP.Addresses.GlobalLocationNumber = oBP.Addresses.GlobalLocationNumber
                            oTargetBP.Addresses.Add()
                        Next

                        oTargetBP.BilltoDefault = oBP.BilltoDefault
                        oTargetBP.ShipToDefault = oBP.ShipToDefault
                        oTargetBP.PayTermsGrpCode = oBP.PayTermsGrpCode
                        oTargetBP.IntrestRatePercent = oBP.IntrestRatePercent
                        oTargetBP.PriceListNum = oBP.PriceListNum
                        oTargetBP.DiscountPercent = oBP.DiscountPercent

                        oTargetBP.CreditLimit = oBP.CreditLimit
                        oTargetBP.MaxCommitment = oBP.MaxCommitment
                        oTargetBP.EffectiveDiscount = oBP.EffectiveDiscount

                        oTargetBP.HouseBank = oBP.HouseBank
                        oTargetBP.HouseBankAccount = oBP.HouseBankAccount
                        oTargetBP.HouseBankBranch = oBP.HouseBankBranch
                        oTargetBP.HouseBankCountry = oBP.HouseBankCountry
                        oTargetBP.BPPaymentMethods.PaymentMethodCode = oBP.BPPaymentMethods.PaymentMethodCode

                        'BP Bank Details 
                        For i As Integer = 0 To oBP.BPBankAccounts.Count - 1
                            oBP.BPBankAccounts.SetCurrentLine(i)
                            'orsTarget.DoQuery(String.Format("SELECT ""BankCode"" from ""ODSC"" where ""BankCode"" =  '{0}'", oBP.BPBankAccounts.BankCode))
                            'If orsTarget.RecordCount = 1 Then
                            oTargetBP.BPBankAccounts.BankCode = oBP.BPBankAccounts.BankCode
                            oTargetBP.BPBankAccounts.Country = oBP.BPBankAccounts.Country
                            oTargetBP.BPBankAccounts.BPCode = oBP.BPBankAccounts.BPCode
                            oTargetBP.BPBankAccounts.AccountNo = oBP.BPBankAccounts.AccountNo
                            oTargetBP.BPBankAccounts.AccountName = oBP.BPBankAccounts.AccountName
                            oTargetBP.BPBankAccounts.Branch = oBP.BPBankAccounts.Branch
                            oTargetBP.BPBankAccounts.BICSwiftCode = oBP.BPBankAccounts.BICSwiftCode
                            oTargetBP.BPBankAccounts.InternalKey = oBP.BPBankAccounts.InternalKey
                            oTargetBP.BPBankAccounts.ControlKey = oBP.BPBankAccounts.ControlKey
                            oTargetBP.BPBankAccounts.IBAN = oBP.BPBankAccounts.IBAN
                            oTargetBP.BPBankAccounts.Street = oBP.BPBankAccounts.Street
                            oTargetBP.BPBankAccounts.State = oBP.BPBankAccounts.Street
                            oTargetBP.BPBankAccounts.Block = oBP.BPBankAccounts.Street
                            oTargetBP.BPBankAccounts.BuildingFloorRoom = oBP.BPBankAccounts.BuildingFloorRoom
                            oTargetBP.BPBankAccounts.City = oBP.BPBankAccounts.City
                            oTargetBP.BPBankAccounts.MandateID = oBP.BPBankAccounts.MandateID

                            oTargetBP.BPBankAccounts.Add()

                        Next
                        orsTarget.DoQuery(String.Format("SELECT ""BankCode"",""CountryCod"" from ""ODSC"" where ""BankCode"" =  '{0}'", oBP.DefaultBankCode))
                        If orsTarget.RecordCount = 1 Then
                            Dim sBankcode As String = orsTarget.Fields.Item(0).Value
                            Dim sCtrycode As String = orsTarget.Fields.Item(1).Value

                            oTargetBP.DefaultBankCode = oBP.DefaultBankCode
                            oTargetBP.DefaultAccount = oBP.DefaultAccount

                        End If
                        For i As Integer = 0 To oBP.ContactEmployees.Count - 1
                            oBP.ContactEmployees.SetCurrentLine(i)
                            oTargetBP.ContactEmployees.Name = oBP.ContactEmployees.Name
                            oTargetBP.ContactEmployees.FirstName = oBP.ContactEmployees.FirstName
                            oTargetBP.ContactEmployees.MiddleName = oBP.ContactEmployees.MiddleName
                            oTargetBP.ContactEmployees.LastName = oBP.ContactEmployees.LastName
                            oTargetBP.ContactEmployees.Title = oBP.ContactEmployees.Title
                            oTargetBP.ContactEmployees.Position = oBP.ContactEmployees.Position
                            oTargetBP.ContactEmployees.Address = oBP.ContactEmployees.Address
                            oTargetBP.ContactEmployees.Phone1 = oBP.ContactEmployees.Phone1
                            oTargetBP.ContactEmployees.Phone2 = oBP.ContactEmployees.Phone2
                            oTargetBP.ContactEmployees.MobilePhone = oBP.ContactEmployees.MobilePhone
                            oTargetBP.ContactEmployees.Fax = oBP.ContactEmployees.Fax
                            oTargetBP.ContactEmployees.E_Mail = oBP.ContactEmployees.E_Mail
                            oTargetBP.ContactEmployees.Pager = oBP.ContactEmployees.Pager
                            oTargetBP.ContactEmployees.Remarks1 = oBP.ContactEmployees.Remarks1
                            oTargetBP.ContactEmployees.Remarks2 = oBP.ContactEmployees.InternalCode
                            oTargetBP.ContactEmployees.CityOfBirth = oBP.ContactEmployees.CityOfBirth
                            oTargetBP.ContactEmployees.DateOfBirth = oBP.ContactEmployees.DateOfBirth
                            oTargetBP.ContactEmployees.Gender = oBP.ContactEmployees.Gender
                            oTargetBP.ContactEmployees.Profession = oBP.ContactEmployees.Profession
                            oTargetBP.ContactEmployees.Active = oBP.ContactEmployees.Active
                            oTargetBP.ContactEmployees.Add()
                        Next

                        oTargetBP.ContactPerson = oBP.ContactPerson
                        ErrorCode = oTargetBP.Add
                        If ErrorCode <> 0 Then
                            targetCompany.GetLastError(ErrorCode, sErrDesc)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding BP Code '" & CardCode & "' Failed on '" & targetCompany.CompanyDB & "'" + " - " + sErrDesc, sFuncName)
                            CreateBPMaster = RTN_ERROR
                        Else
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("BP Code: '" & CardCode & "' Created Successfully on '" & targetCompany.CompanyDB & "'", sFuncName)
                            CreateBPMaster = RTN_SUCCESS
                        End If
                    End If
                    'Return sErrDesc
                Catch ex As Exception
                    sErrDesc = ex.Message
                    WriteToLogFile(ex.Message, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("BP Replcation Failed", sFuncName)
                    CreateBPMaster = RTN_ERROR
                Finally
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetBP)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP)
                    oTargetBP = Nothing
                    oBP = Nothing
                End Try
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error: CardCode not found!!!!", sFuncName)
            End If
        End Function

        Public Function CreatePricelistMaster(targetCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
            Try
                Dim oPricelistMaster As SAPbobsCOM.PriceLists
                Dim oTargetPricelistsMaster As SAPbobsCOM.PriceLists
                Dim ors As SAPbobsCOM.Recordset = Nothing
                Dim orsTarget As SAPbobsCOM.Recordset = Nothing
                Dim ErrorCode As Long
                ' Dim Errmsg As String = ""
                Dim sFuncName = "CreatePriceListMaster()"
                targetCompany.StartTransaction()
                CreatePricelistMaster = False
                Dim CheckFlag As Boolean = False

                oPricelistMaster = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPriceLists)
                oTargetPricelistsMaster = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPriceLists)

                Dim sqry As String = "  select ""ListNum"",""ListName"", ""BASE_NUM"",""Factor"",""RoundSys"",""GroupCode"", ""ValidFor"" from OPLN ORDER BY ""ListNum"";"
                Dim RPLSets As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                RPLSets.DoQuery(sqry)
                oDT_PricelistsMaster = New DataTable
                oDT_PricelistsMaster = ConvertRecordset(RPLSets)
                Dim dtcount12 As Integer = oDT_PricelistsMaster.Rows.Count


                For T As Integer = 0 To oDT_PricelistsMaster.Rows.Count - 1
                    CheckFlag = False
                    Dim listname As String = oDT_PricelistsMaster.Rows(T).Item("ListName").ToString
                    Dim oChecking As SAPbobsCOM.Recordset = targetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oChecking.DoQuery(String.Format("Select ""ListNum"" from ""OPLN"" where ""ListName"" = '{0}'", oDT_PricelistsMaster.Rows(T).Item("ListName").ToString))
                    If oChecking.RecordCount = 0 Then
                        CreatePricelistMaster = False
                        Try
                            If oPricelistMaster.GetByKey(oDT_PricelistsMaster.Rows(T).Item("ListNum").ToString) Then
                                oTargetPricelistsMaster.PriceListName = oPricelistMaster.PriceListName
                                oTargetPricelistsMaster.BasePriceList = oPricelistMaster.BasePriceList
                                oTargetPricelistsMaster.Factor = oPricelistMaster.Factor
                                oTargetPricelistsMaster.RoundingMethod = oPricelistMaster.RoundingMethod
                                oTargetPricelistsMaster.RoundingRule = oPricelistMaster.RoundingRule
                                oTargetPricelistsMaster.GroupNum = oPricelistMaster.GroupNum
                                oTargetPricelistsMaster.Active = oPricelistMaster.Active

                                ErrorCode = oTargetPricelistsMaster.Add
                                If ErrorCode <> 0 Then
                                    targetCompany.GetLastError(ErrorCode, sErrDesc)
                                    p_oSBOApplication.StatusBar.SetText("Price List Creating Failed..'" & targetCompany.CompanyDB & "'.." & sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    CreatePricelistMaster = False
                                Else
                                    CreatePricelistMaster = True
                                    p_oSBOApplication.StatusBar.SetText("Price List Created Successfully on ..." & targetCompany.CompanyDB, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If
                            End If
                        Catch ex As Exception
                            sErrDesc = ex.Message
                            p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            CreatePricelistMaster = False
                        End Try
                    Else
                        CheckFlag = True
                    End If
                Next
                If CreatePricelistMaster = True Then
                    If targetCompany.InTransaction Then targetCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                End If
                If CheckFlag = True Then CreatePricelistMaster = True
                'CreatePricelistMaster = True
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oPricelistMaster)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetPricelistsMaster)
            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText("Price List Master Creation Failed...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                CreatePricelistMaster = False
            Finally
            End Try
        End Function
    End Class
End Namespace


