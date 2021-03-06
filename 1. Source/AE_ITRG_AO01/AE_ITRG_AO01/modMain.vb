﻿Option Explicit On
Imports SAPbouiCOM.framework

Namespace AE_ITRG_AO01
    Module modMain

        Public Structure Costcenter

            Public sCostcode As String
            Public sCostName As String
            Public iDimension As Integer
            Public dEffectiveFrom As Date
            Public dEffectiveTo As Date
            Public sEntity As String
            Public sReportCode As String
            Public sEntityName As String
            Public sNatCostType As String
        End Structure

        Public p_oApps As SAPbouiCOM.SboGuiApi
        Public p_oEventHandler As clsEventHandler
        Public WithEvents p_oSBOApplication As SAPbouiCOM.Application
        Public p_oDICompany As SAPbobsCOM.Company
        Public p_oUICompany As SAPbouiCOM.Company
        Public sFuncName As String
        Public sErrDesc As String
        Public oTableCreation As New TableCreation

        Public p_iDebugMode As Int16
        Public p_iErrDispMethod As Int16
        Public p_iDeleteDebugLog As Int16

        Public p_sSQLName As String = String.Empty
        Public p_sSQLPass As String = String.Empty

        Public Const RTN_SUCCESS As Int16 = 1
        Public Const RTN_ERROR As Int16 = 0

        Public Const DEBUG_ON As Int16 = 1
        Public Const DEBUG_OFF As Int16 = 0

        Public Const ERR_DISPLAY_STATUS As Int16 = 1
        Public Const ERR_DISPLAY_DIALOGUE As Int16 = 2
        Public format1 As New System.Globalization.CultureInfo("fr-FR", True)

        Public p_sSelectedFilepath As String = String.Empty
        Public dtTable As DataTable = Nothing
        Public oDT_Entities As DataTable = Nothing
        Public oDT_BPSetup As DataTable = Nothing
        Public oDV_BPSetup As DataView = Nothing
        Public oDT_FINSetup As DataTable = Nothing
        Public oDV_FINSetup As DataView = Nothing
        Public oDT_ItemPricelist As DataTable = Nothing
        Public oDV_ItemPricelist As DataView = Nothing
        Public oDT_Pricelists As DataTable = Nothing
        Public oDV_PriceLists As DataView = Nothing
        Public oDT_Itemlists As DataTable = Nothing
        Public oDT_Binlists As DataTable = Nothing
        Public oDT_PricelistsMaster As DataTable = Nothing
        Public oDT_PPlists As DataTable = Nothing
        Public oDT_ExchRates As DataTable = Nothing

        Public oDT_INTCompany As DataTable = Nothing
        Public oDV_INTCompany As DataView = Nothing

        Public p_FrmType As Integer
        Public p_oCostCenter As Costcenter
        Public p_sBudgetType As String = String.Empty
        Public oDT_ErrorMsg As DataTable = Nothing

        Public oFlag2 As Boolean = False
        Public oFlag3 As Boolean = False
        Public oFlag4 As Boolean = False
        Public oFlag5 As Boolean = False

        <STAThread()>
        Sub Main(ByVal args() As String)

            ''Dim oApp As Application
            Dim sconn As String = String.Empty
            ''If (args.Length < 1) Then
            ''    oApp = New Application
            ''Else
            ''    oApp = New Application(args(0))
            ''End If

            sFuncName = "Main()"
            Try
                p_iDebugMode = DEBUG_ON
                p_iErrDispMethod = ERR_DISPLAY_STATUS

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Addon startup function", sFuncName)
                p_oApps = New SAPbouiCOM.SboGuiApi
                'sconn = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
                'p_oApps.Connect(args(0))
                p_oApps.Connect(args(0))

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing public SBO Application object", sFuncName)
                p_oSBOApplication = p_oApps.GetApplication

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retriving SBO application company handle", sFuncName)
                p_oUICompany = p_oSBOApplication.Company


                p_oDICompany = New SAPbobsCOM.Company
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retrived SBO application company handle", sFuncName)
                ' p_oDICompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                If Not p_oDICompany.Connected Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                    If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If

                'Call WriteToLogFile_Debug("Calling DisplayStatus()", sFuncName)
                'Call DisplayStatus(Nothing, "Addon starting.....please wait....", sErrDesc)
                '--------------Hide for Table creation add-on
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Event handler class", sFuncName)
                p_oEventHandler = New clsEventHandler(p_oSBOApplication, p_oDICompany)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddMenuItems()", sFuncName)
                p_oEventHandler.AddMenuItems()
                '--------------------------------------------

                '' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Tables and Fields... ", sFuncName)
                '' oTableCreation.TableCreation()
                ''  
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")

                'Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
                ' Call EndStatus(sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing Recordset ", "Main()")

                p_oSBOApplication.StatusBar.SetText("Addon Started Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                System.Windows.Forms.Application.Run()

            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                p_oSBOApplication.StatusBar.SetText("Addon Connection Failed.... " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try


        End Sub

    End Module
End Namespace