using System.Collections.Generic;
using System.Web.Services;
using System.Web.Script.Serialization;
using System.Web.Script.Services;
using System.Data;
using System;
using System.Text.RegularExpressions;
using AE_OutletManagerService_BLL;
using System.Data.Odbc;

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
    }
}
