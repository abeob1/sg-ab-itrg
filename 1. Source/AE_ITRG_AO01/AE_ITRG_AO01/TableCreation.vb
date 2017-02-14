Namespace AE_ITRG_AO01
    Public Class TableCreation
        Dim DocType As String(,) = New String(,) {{"Y", "Yes"}, {"N", "No"}}
        Dim v_RetVal As Long
        Dim v_ErrCode As Long
        Dim v_ErrMsg As String = ""
        Sub TableCreation()
            Try
                Me.Target_Database_Credentials_Table()
                Me.BP_Master_Setup_Table()
                Me.Item_Master_Setup_Table()
                Me.Financial_Master_Setup_Table()
                Me.Create_Integration_Table()
            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText("Table Creation Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Finally
            End Try
        End Sub
        Sub BP_Master_Setup_Table()
            Try
                CreateTable("AE_TB001_BPSETUP", "Business Partner Setup", SAPbobsCOM.BoUTBTableType.bott_NoObject)

                CreateUserFieldsComboBoxWithLinkedTable("@AE_TB001_BPSETUP", "TargetDB", "Target DB", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "AE_TB004_TARCRE")
                CreateUserFieldsComboBox("@AE_TB001_BPSETUP", "Customers", "Customers", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB001_BPSETUP", "Vendors", "Vendors", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB001_BPSETUP", "Leads", "Leads", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB001_BPSETUP", "PayTerms", "PayTerms", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB001_BPSETUP", "BPGroups", "BPGroups", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")

            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText("Business Parter Setup Creation Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Finally
            End Try
        End Sub

        Sub Item_Master_Setup_Table()
            Try
                CreateTable("AE_TB002_ITEM", "Item Master Setup", SAPbobsCOM.BoUTBTableType.bott_NoObject)

                CreateUserFieldsComboBoxWithLinkedTable("@AE_TB002_ITEM", "TargetDB", "Target DB", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "AE_TB004_TARCRE")
                CreateUserFieldsComboBox("@AE_TB002_ITEM", "ItemGroups", "ItemGroups", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB002_ITEM", "ItemCodes", "Item Codes", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB002_ITEM", "BinLocatin", "Bin Location", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB002_ITEM", "BOM", "Bill of Material", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB002_ITEM", "PriceLists", "Price Lists", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB002_ITEM", "UOMGroups", "UOM Groups", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText("Item Master Setup Creation Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Finally
            End Try
        End Sub

        Sub Financial_Master_Setup_Table()
            Try
                CreateTable("AE_TB003_FIN", "Financial Setup", SAPbobsCOM.BoUTBTableType.bott_NoObject)

                CreateUserFieldsComboBoxWithLinkedTable("@AE_TB003_FIN", "TargetDB", "Target DB", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "AE_TB004_TARCRE")
                CreateUserFieldsComboBox("@AE_TB003_FIN", "Currency", "Currencies", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB003_FIN", "COA", "COA", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB003_FIN", "PostPeriod", "Posting Periods", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB003_FIN", "ExcRates", "Exchange Rates", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB003_FIN", "CostCenter1", "Cost Center1", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB003_FIN", "CostCenter2", "Cost Center2", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB003_FIN", "CostCenter3", "Cost Center3", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB003_FIN", "CostCenter4", "Cost Center4", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
                CreateUserFieldsComboBox("@AE_TB003_FIN", "CostCenter5", "Cost Center5", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , DocType, "N")
            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText("Financial Setup Creation Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Finally
            End Try
        End Sub
        Sub Target_Database_Credentials_Table()
            Try
                CreateTable("AE_TB004_TARCRE", "Target DB Credentials Setup", SAPbobsCOM.BoUTBTableType.bott_NoObject)

                CreateUserFields("@AE_TB004_TARCRE", "UserName", "User Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
                CreateUserFields("@AE_TB004_TARCRE", "Password", "Password", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)

            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText("Target DB Credentials Table Creation Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Finally
            End Try
        End Sub
        Sub Create_Integration_Table()
            Try
                Dim IntTable As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oFlag As Boolean = True
                Dim sts As String = "CREATE COLUMN TABLE INTEGRATION (uniqueId INTEGER,MasterType VARCHAR(100),TransType VARCHAR(100),Code VARCHAR(100),Name VARCHAR(100),SyncStatus VARCHAR(100),Syncdate Date,ErrorMsg VARCHAR(254));"
                IntTable.DoQuery("CREATE COLUMN TABLE INTEGRATION (uniqueId INTEGER,MasterType VARCHAR(100),TransType VARCHAR(100),Code VARCHAR(100),Name VARCHAR(100),SyncStatus VARCHAR(100),Syncdate Date,ErrorMsg VARCHAR(254));")
            Catch ex As Exception
                '' p_oSBOApplication.StatusBar.SetText("Integration Table Creation Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Sub
        Function CreateTable(ByVal TableName As String, ByVal TableDesc As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
            CreateTable = False
            Dim v_RetVal As Long
            Dim v_ErrCode As Long
            Dim v_ErrMsg As String = ""
            Try
                If Not Me.TableExists(TableName) Then
                    Dim v_UserTableMD As SAPbobsCOM.UserTablesMD
                    p_oSBOApplication.StatusBar.SetText("Creating Table " + TableName + " ...................", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    v_UserTableMD = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                    v_UserTableMD.TableName = TableName
                    v_UserTableMD.TableDescription = TableDesc
                    v_UserTableMD.TableType = TableType
                    v_RetVal = v_UserTableMD.Add()
                    If v_RetVal <> 0 Then
                        p_oDICompany.GetLastError(v_ErrCode, v_ErrMsg)
                        p_oSBOApplication.StatusBar.SetText("Failed to Create Table " & TableDesc & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD)
                        v_UserTableMD = Nothing
                        Return False
                    Else
                        p_oSBOApplication.StatusBar.SetText("[" & TableName & "] - " & TableDesc & " Created Successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD)
                        v_UserTableMD = Nothing
                        Return True
                    End If
                Else
                    GC.Collect()
                    Return False
                End If
            Catch ex As Exception
                '-- p_oSBOApplication.StatusBar.SetText(AddOnName & ":> " & ex.Message & " @ " & ex.Source)
            End Try
        End Function

        Function TableExists(ByVal TableName As String) As Boolean
            Dim oTables As SAPbobsCOM.UserTablesMD
            Dim oFlag As Boolean
            oTables = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            oFlag = oTables.GetByKey(TableName)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oTables)
            Return oFlag
        End Function
        Function ColumnExists(ByVal TableName As String, ByVal FieldID As String) As Boolean
            Try
                Dim rs As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oFlag As Boolean = True
                Dim ss As String = "Select 1 from ""CUFD"" T0 Where ""TableID""='" & Trim(TableName) & "' and ""AliasID""='" & Trim(FieldID) & "'"
                rs.DoQuery("Select 1 from ""CUFD"" T0 Where ""TableID""='" & Trim(TableName) & "' and ""AliasID""='" & Trim(FieldID) & "'")
                If rs.EoF Then oFlag = False
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
                rs = Nothing
                GC.Collect()
                Return oFlag
            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText(ex.Message)
            End Try
        End Function
        Function UDFExists(ByVal TableName As String, ByVal FieldID As String) As Boolean
            Try
                Dim rs As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oFlag As Boolean = True
                Dim aa = "Select 1 from ""CUFD"" Where ""TableID""='" & Trim(TableName) & "' and ""AliasID""='" & Trim(FieldID) & "'"
                rs.DoQuery("Select 1 from ""CUFD"" Where ""TableID""='" & Trim(TableName) & "' and ""AliasID""='" & Trim(FieldID) & "'")
                If rs.EoF Then oFlag = False
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
                rs = Nothing
                GC.Collect()
                Return oFlag
            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText(ex.Message)
            End Try
        End Function
        Function CreateUserFields(ByVal TableName As String, ByVal FieldName As String, ByVal FieldDescription As String, ByVal type As SAPbobsCOM.BoFieldTypes, Optional ByVal size As Long = 0, Optional ByVal subType As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal LinkedTable As String = "", Optional ByVal DefaultValue As String = "") As Boolean
            Dim v_RetVal As Long
            Dim v_ErrCode As Long
            Dim v_ErrMsg As String = ""
            Try
                If TableName.StartsWith("@") = True Then
                    If Not Me.ColumnExists(TableName, FieldName) Then
                        Dim v_UserField As SAPbobsCOM.UserFieldsMD
                        v_UserField = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                        v_UserField.TableName = TableName
                        v_UserField.Name = FieldName
                        v_UserField.Description = FieldDescription
                        v_UserField.Type = type
                        If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                            If size <> 0 Then
                                v_UserField.Size = size
                            End If
                        End If
                        If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                            v_UserField.SubType = subType
                        End If
                        If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                        If DefaultValue <> "" Then v_UserField.DefaultValue = DefaultValue

                        v_RetVal = v_UserField.Add()
                        If v_RetVal <> 0 Then
                            p_oDICompany.GetLastError(v_ErrCode, v_ErrMsg)
                            p_oSBOApplication.StatusBar.SetText("Failed to add UserField masterid" & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                            v_UserField = Nothing
                            Return False
                        Else
                            p_oSBOApplication.StatusBar.SetText("[" & TableName & "] - " & FieldDescription & " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                            v_UserField = Nothing
                            Return True
                        End If
                    Else
                        Return False
                    End If
                End If

                If TableName.StartsWith("@") = False Then
                    If Not Me.UDFExists(TableName, FieldName) Then
                        Dim v_UserField As SAPbobsCOM.UserFieldsMD = p_oSBOApplication.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                        v_UserField.TableName = TableName
                        v_UserField.Name = FieldName
                        v_UserField.Description = FieldDescription
                        v_UserField.Type = type
                        If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                            If size <> 0 Then
                                v_UserField.Size = size
                            End If
                        End If
                        If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                            v_UserField.SubType = subType
                        End If
                        If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                        v_RetVal = v_UserField.Add()
                        If v_RetVal <> 0 Then
                            p_oDICompany.GetLastError(v_ErrCode, v_ErrMsg)
                            p_oSBOApplication.StatusBar.SetText("Failed to add UserField " & FieldDescription & " - " & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                            v_UserField = Nothing
                            Return False
                        Else
                            p_oSBOApplication.StatusBar.SetText(" & TableName & - " & FieldDescription & " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                            v_UserField = Nothing
                            Return True
                        End If

                    Else
                        Return False
                    End If
                End If
            Catch ex As Exception
                p_oSBOApplication.MessageBox(ex.Message)
            End Try
        End Function
        Function CreateUserFieldsComboBox(ByVal TableName As String, ByVal FieldName As String, ByVal FieldDescription As String, ByVal type As SAPbobsCOM.BoFieldTypes, Optional ByVal size As Long = 0, Optional ByVal subType As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal LinkedTable As String = "", Optional ByVal ComboValidValues As String(,) = Nothing, Optional ByVal DefaultValidValues As String = "") As Boolean
            Try
                'If TableName.StartsWith("@") = False Then
                If Not Me.UDFExists(TableName, FieldName) Then
                    Dim v_UserField As SAPbobsCOM.UserFieldsMD
                    v_UserField = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    v_UserField.TableName = TableName
                    v_UserField.Name = FieldName
                    v_UserField.Description = FieldDescription
                    v_UserField.Type = type
                    If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                        If size <> 0 Then
                            v_UserField.Size = size
                        End If
                    End If
                    If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                        v_UserField.SubType = subType
                    End If

                    For i As Int16 = 0 To ComboValidValues.GetLength(0) - 1
                        If i > 0 Then v_UserField.ValidValues.Add()
                        v_UserField.ValidValues.Value = ComboValidValues(i, 0)
                        v_UserField.ValidValues.Description = ComboValidValues(i, 1)
                    Next
                    If DefaultValidValues <> "" Then v_UserField.DefaultValue = DefaultValidValues

                    If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                    v_RetVal = v_UserField.Add()
                    If v_RetVal <> 0 Then
                        p_oDICompany.GetLastError(v_ErrCode, v_ErrMsg)
                        p_oSBOApplication.StatusBar.SetText("Failed to add UserField " & FieldDescription & " - " & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return False
                    Else
                        p_oSBOApplication.StatusBar.SetText(" & TableName & - " & FieldDescription & " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return True
                    End If
                Else
                    Return False
                End If
                ' End If
            Catch ex As Exception
                p_oSBOApplication.MessageBox(ex.Message)
            End Try
        End Function

        Function CreateUserFieldsComboBoxWithLinkedTable(ByVal TableName As String, ByVal FieldName As String, ByVal FieldDescription As String, ByVal type As SAPbobsCOM.BoFieldTypes, Optional ByVal size As Long = 0, Optional ByVal subType As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal LinkedTable As String = "", Optional ByVal ComboValidValues As String(,) = Nothing, Optional ByVal DefaultValidValues As String = "") As Boolean
            Try
                'If TableName.StartsWith("@") = False Then
                If Not Me.UDFExists(TableName, FieldName) Then
                    Dim v_UserField As SAPbobsCOM.UserFieldsMD
                    v_UserField = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    v_UserField.TableName = TableName
                    v_UserField.Name = FieldName
                    v_UserField.Description = FieldDescription
                    v_UserField.Type = type
                    If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                        If size <> 0 Then
                            v_UserField.Size = size
                        End If
                    End If
                    If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                        v_UserField.SubType = subType
                    End If

                    If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                    v_RetVal = v_UserField.Add()
                    If v_RetVal <> 0 Then
                        p_oDICompany.GetLastError(v_ErrCode, v_ErrMsg)
                        p_oSBOApplication.StatusBar.SetText("Failed to add UserField " & FieldDescription & " - " & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return False
                    Else
                        p_oSBOApplication.StatusBar.SetText(" & TableName & - " & FieldDescription & " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return True
                    End If
                Else
                    Return False
                End If
                ' End If
            Catch ex As Exception
                p_oSBOApplication.MessageBox(ex.Message)
            End Try
        End Function
    End Class
End Namespace
