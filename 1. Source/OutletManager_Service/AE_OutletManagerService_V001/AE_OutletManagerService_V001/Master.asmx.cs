using System.Collections.Generic;
using System.Web.Services;
using System.Web.Script.Serialization;
using System.Web.Script.Services;
using System.Data;
using System;
using System.Text.RegularExpressions;
using AE_OutletManagerService_BLL;
using System.Data.Odbc;
using System.Data.Common;

namespace AE_OutletManagerService_V001
{

    /// <summary>
    /// Summary description for Master
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class Master : System.Web.Services.WebService
    {
        #region Objects
        public string sFuncName = string.Empty;
        public string sSQL = string.Empty;
        public string sErrDesc = string.Empty;
        clsLog oLog = new clsLog();
        clsCommon oCommon = new clsCommon();
        JavaScriptSerializer js = new JavaScriptSerializer();
        List<result> lstResult = new List<result>();
        #endregion

        #region Web Methods

        [WebMethod]
        public void GetUserInfo(string sCompany)
        {
            DataTable dtResult = new DataTable();
            string sResult = string.Empty;
            try
            {
                oLog.WriteToDebugLogFile("Starting Function", sFuncName);
                //sSQL = string.Format("call \"AE_SP001_UFF_Generation\" ('{0}')", sFileId);
                sSQL = string.Format("call \"AE_SP001_GETALLUSERS\"");

                oLog.WriteToDebugLogFile("Execute SQL" + sSQL, sFuncName);
                OdbcParameter[] Param = new OdbcParameter[0]; 
                dtResult = oCommon.ExecuteSQLQuery(sSQL, sCompany, Param);
                List<Users> lstUsers = new List<Users>();
                if (dtResult.Rows.Count > 0)
                {
                    foreach (DataRow r in dtResult.Rows)
                    {
                        Users _company = new Users();
                        _company.USERCODE = r["USERCODE"].ToString();
                        _company.USERNAME = r["USERNAME"].ToString();
                        _company.DEFAULTENTITY = r["DEFAULTENTITY"].ToString();
                        _company.DEFAULTBRANCHCODE = r["DEFAULTBRANCHCODE"].ToString();
                        _company.DEFAULTDEPTCODE = r["DEFAULTDEPTCODE"].ToString();
                        _company.PASSWORD = r["PASSWORD"].ToString();
                        _company.LOCKED = r["LOCKED"].ToString();
                        _company.DEFAULTAPPROVALLEVEL = r["DEFAULTAPPROVALLEVEL"].ToString();
                        _company.APPROVALSCOPE = r["APPROVAL SCOPE"].ToString();
                        _company.LANGUAGE = r["LANGUAGE"].ToString();
                        lstUsers.Add(_company);
                    }
                    oLog.WriteToDebugLogFile("Before Serializing the Company List ", sFuncName);
                    Context.Response.Output.Write(js.Serialize(lstUsers));
                    oLog.WriteToDebugLogFile("After Serializing the Company List , the Serialized data is ' " + js.Serialize(lstUsers) + " '", sFuncName);
                }
                else
                {
                    Context.Response.Output.Write(js.Serialize(lstUsers));
                }

                oLog.WriteToDebugLogFile("Ending Function", sFuncName);
            }
            catch (Exception ex)
            {
                sErrDesc = ex.Message.ToString();
                oLog.WriteToErrorLogFile(sErrDesc, sFuncName);
                oLog.WriteToDebugLogFile("Completed With ERROR  ", sFuncName);
                result objResult = new result();
                objResult.Result = "Error";
                objResult.DisplayMessage = sErrDesc;
                lstResult.Add(objResult);
                Context.Response.Output.Write(js.Serialize(lstResult));
            }
        }

        ////Vikrant - 26/04/2017
        ////Created to check whether log in user is valid or not
        //[WebMethod]
        //public void CheckValidUser(string sUserName,string sPassword)
        //{
        //    string sResult = string.Empty;
        //    int iValid = 0;
        //    try
        //    {
        //        sFuncName = "CheckValidUser";
        //        oLog.WriteToDebugLogFile("Validate Login Function", sFuncName);

        //        sSQL = string.Format("call \"AE_SP002_VALIDUSER\"('" + sUserName + "','" + sPassword + "')");

        //        oLog.WriteToDebugLogFile("Execute SQL" + sSQL, sFuncName);
        //        OdbcParameter[] Param = new OdbcParameter[0];

        //        iValid = oCommon.ExecuteNonQuery_DR(sSQL, Param);
                
        //        if(iValid == 1)
        //        {
        //            sResult = "SUCCESS";
        //        }
        //        else if (iValid == 2)
        //        {
        //            sResult = "INCORRECT USERNAME";
        //        }
        //        else if (iValid == 3)
        //        {
        //            sResult = "INCORRECT PASSWORD";
        //        }
        //        else 
        //        {
        //            sResult = "FAILURE";
        //        }

        //        oLog.WriteToDebugLogFile("Ending Function", sFuncName);
        //    }
        //    catch (Exception ex)
        //    {
        //        sErrDesc = ex.Message.ToString();
        //        oLog.WriteToErrorLogFile(sErrDesc, sFuncName);
        //        oLog.WriteToDebugLogFile("Completed With ERROR  ", sFuncName);
        //        result objResult = new result();
        //        objResult.Result = "Error";
        //        objResult.DisplayMessage = sErrDesc;
        //        lstResult.Add(objResult);
        //        Context.Response.Output.Write(js.Serialize(lstResult));
        //    }

        //    return sResult;
        //}

        //Vivek - 15/05/2017
        //Created to check whether log in user is valid or not
        [WebMethod]
        public void CheckValidUser(string sJsonInput)
        {
            string sResult = string.Empty;
            result objResult = new result();
            int iValid = 0;
            try
            {
                string sUserName = string.Empty;
                string sPassword = string.Empty;

                sFuncName = "CheckValidUser";
                oLog.WriteToDebugLogFile("Validate Login Function", sFuncName);

                sJsonInput = "[" + sJsonInput + "]";
                oLog.WriteToDebugLogFile("Getting the Json Input 1 from web  '" + sJsonInput + "'", sFuncName);
                DataTable dtInput = JsonStringToDataTable(sJsonInput);
                if (dtInput != null && dtInput.Rows.Count > 0)
                {
                    sUserName = dtInput.Rows[0]["sUserName"].ToString();
                    sPassword = dtInput.Rows[0]["sPassword"].ToString();
                }

                sSQL = string.Format("call \"AE_SP002_VALIDUSER\"('" + sUserName + "','" + sPassword + "')");

                oLog.WriteToDebugLogFile("Execute SQL" + sSQL, sFuncName);
                OdbcParameter[] Param = new OdbcParameter[0];

                iValid = oCommon.ExecuteNonQuery_DR(sSQL, Param);

                if (iValid == 1)
                {
                    sResult = "SUCCESS";
                }
                else if (iValid == 2)
                {
                    sResult = "INCORRECT USERNAME";
                }
                else if (iValid == 3)
                {
                    sResult = "INCORRECT PASSWORD";
                }
                else
                {
                    sResult = "FAILURE";
                }

                oLog.WriteToDebugLogFile("Ending Function", sFuncName);
                oLog.WriteToDebugLogFile("Completed With SUCCESS ", sFuncName);
                objResult.Result = "SUCCESS";
                objResult.DisplayMessage = "Delivery Order Completed successfully";
                lstResult.Add(objResult);
                Context.Response.Output.Write(js.Serialize(lstResult));
            }
            catch (Exception ex)
            {
                sErrDesc = ex.Message.ToString();
                oLog.WriteToErrorLogFile(sErrDesc, sFuncName);
                oLog.WriteToDebugLogFile("Completed With ERROR  ", sFuncName);
                objResult.Result = "Error";
                objResult.DisplayMessage = sErrDesc;
                lstResult.Add(objResult);
                Context.Response.Output.Write(js.Serialize(lstResult));
            }
        }

        //Vikrant - 03/05/2017
        //Created to create new user
        [WebMethod]
        public void CreateUser(string sUserCode, string sUserName, string sDefaultEntity, string sDefaultBranchCode, string sDefaultDeptCode, string sPassword, string sLocked, string sDefaultApprovalLevel, string sApprovalScope, string sLanguage)
        {
            string sCompany = string.Empty;
            string sResult = string.Empty;
            sFuncName = "CreateUser";

            try
            {
                oLog.WriteToDebugLogFile("Create New User Function", sFuncName);

                sSQL = string.Format("call \"AE_SP003_CREATEUSER\"('" + sUserCode + "','" + sUserName + "','" + sDefaultEntity + "','" + sDefaultBranchCode + "','" + sDefaultDeptCode + "','" + sPassword + "','" + sLocked + "','" + sDefaultApprovalLevel + "','" + sApprovalScope + "','" + sLanguage + "')");
                
                oLog.WriteToDebugLogFile("Execute SQL" + sSQL, sFuncName);
                OdbcParameter[] Param = new OdbcParameter[0];
                oCommon.ExecuteQuery(sSQL, sCompany, Param);
                
                oLog.WriteToDebugLogFile("Ending Function", sFuncName);
            }
            catch (Exception ex)
            {
                sErrDesc = ex.Message.ToString();
                oLog.WriteToErrorLogFile(sErrDesc, sFuncName);
                oLog.WriteToDebugLogFile("Completed With ERROR  ", sFuncName);
                result objResult = new result();
                objResult.Result = "Error";
                objResult.DisplayMessage = sErrDesc;
                lstResult.Add(objResult);
                Context.Response.Output.Write(js.Serialize(lstResult));
            }
        }


        #endregion

        #region Class

        class result
        {
            public string Result { get; set; }
            public string DisplayMessage { get; set; }
        }

        class Users
        {
            public string USERCODE { get; set; }
            public string USERNAME { get; set; }
            public string DEFAULTENTITY { get; set; }
            public string DEFAULTBRANCHCODE { get; set; }
            public string DEFAULTDEPTCODE { get; set; }
            public string PASSWORD { get; set; }
            public string LOCKED { get; set; }
            public string DEFAULTAPPROVALLEVEL { get; set; }
            public string APPROVALSCOPE { get; set; }
            public string LANGUAGE { get; set; }
        }

        #endregion

        #region Public Methods
        public DataTable JsonStringToDataTable(string jsonString)
        {
            DataTable dt = new DataTable();
            string sFuncName = string.Empty;
            try
            {
                sFuncName = "JsonStringToDataTable()";
                oLog.WriteToDebugLogFile("Starting Function ", sFuncName);

                string[] jsonStringArray = Regex.Split(jsonString.Replace("[", "").Replace("]", ""), "},{");
                if (jsonStringArray[0].ToString() != string.Empty)
                {
                    List<string> ColumnsName = new List<string>();
                    foreach (string jSA in jsonStringArray)
                    {
                        string sjSA = jSA;
                        if (jSA.Contains("base64,"))
                        {
                            sjSA = jSA.Replace("base64,", "base64;");
                        }
                        string[] jsonStringData = Regex.Split(sjSA.Replace("{", "").Replace("}", ""), "\",");
                        foreach (string ColumnsNameData in jsonStringData)
                        {
                            try
                            {
                                int idx = ColumnsNameData.IndexOf(":");
                                string ColumnsNameString = ColumnsNameData.Substring(0, idx - 1).Replace("\"", "");
                                if (!ColumnsName.Contains(ColumnsNameString.Trim()))
                                {

                                    ColumnsName.Add(ColumnsNameString.Trim());
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception(string.Format("Error Parsing Column Name : {0}", ColumnsNameData));
                            }
                        }
                        break;
                    }
                    foreach (string AddColumnName in ColumnsName)
                    {
                        if (AddColumnName.Contains("Date"))
                        { dt.Columns.Add(AddColumnName, typeof(DateTime)); }
                        else
                        { dt.Columns.Add(AddColumnName); }

                    }
                    foreach (string jSA in jsonStringArray)
                    {
                        string sjSA = jSA;
                        if (jSA.Contains("base64,"))
                        {
                            sjSA = jSA.Replace("base64,", "base64;");
                        }
                        string[] RowData = Regex.Split(sjSA.Replace("{", "").Replace("}", ""), "\",");
                        DataRow nr = dt.NewRow();
                        foreach (string rowData in RowData)
                        {
                            try
                            {
                                string RowDataString = string.Empty;
                                int idx = rowData.Trim().IndexOf(":");
                                string RowColumns = rowData.Trim().Substring(0, idx - 1).Replace("\"", "");
                                if (rowData.Trim().Substring(idx + 1).Replace("\"", "").Contains("base64;"))
                                {
                                    RowDataString = rowData.Trim().Substring(idx + 1).Replace("\"", "").Replace("base64;", "base64,");
                                }
                                else
                                {
                                    RowDataString = rowData.Trim().Substring(idx + 1).Replace("\"", "");
                                }

                                nr[RowColumns] = RowDataString.Trim();

                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }
                        dt.Rows.Add(nr);
                    }
                }
            }
            catch (Exception ex)
            {
                sErrDesc = ex.Message.ToString();
                oLog.WriteToErrorLogFile(sErrDesc, sFuncName);
                oLog.WriteToDebugLogFile("Completed With ERROR  ", sFuncName);
            }
            return dt;
        }

        #endregion Public Methods
    }
}
