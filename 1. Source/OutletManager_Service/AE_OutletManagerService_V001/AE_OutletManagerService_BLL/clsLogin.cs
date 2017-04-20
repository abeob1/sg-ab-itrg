using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using System.Configuration;

namespace AE_OutletManagerService_BLL
{
    public class clsLogin
    {
        #region Objects
        clsLog oLog = new clsLog();
        clsCommon oCommon = new clsCommon();
        public Int16 p_iDebugMode = DEBUG_ON;

        public const Int16 RTN_SUCCESS = 1;
        public const Int16 RTN_ERROR = 0;
        public const Int16 DEBUG_ON = 1;
        public const Int16 DEBUG_OFF = 0;
        public string sErrDesc = string.Empty;
        string sSql = string.Empty;

        #endregion

        #region Methods
        public string GenerateUFF(DataTable oDT_FinalResult, string sPAth, string sFileName, string sDate, string sOrgId, string sSenderName, string sErrDesc)
        {
            string sResult = string.Empty;
            string sFuncName = string.Empty;

            StreamWriter sw = null;
            int dCount = 0;
            double dAmount = 0.0;

            try
            {
                sFuncName = "GenerateUFF()";

                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Starting Function ", sFuncName);

                if (File.Exists(sPAth + sFileName + ".txt"))
                {
                    File.Delete(sPAth + sFileName + ".txt");
                }

                var utf8WithoutBom = new System.Text.UTF8Encoding(false);
                sw = new StreamWriter(sPAth + sFileName + ".txt", false, utf8WithoutBom);
                
                //----------------------- Header Portion -----------------------------------------------
               
                string sHeader = "HEADER," + sDate + "," + sOrgId + "," + sSenderName;
                sw.WriteLine(sHeader.Trim());
                
                //----------------------- Detail Portion -----------------------------------------------
                foreach (DataRow dr in oDT_FinalResult.Rows)
                {
                    dCount += 1;
                    dAmount += Convert.ToDouble(dr["DocTotal"].ToString().Trim());
                    sw.WriteLine(dr["Line"].ToString().Trim());
                }
                //----------------------- Detail Portion -----------------------------------------------
                sw.WriteLine("TRAILER," + dCount + "," + dAmount);
                sw.Close();
                sResult = "SUCCESS";
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Completed With SUCCESS ", sFuncName);
            }
            catch (Exception ex)
            {
                sErrDesc = ex.Message;
                sResult = sErrDesc;
                oLog.WriteToErrorLogFile(sErrDesc, sFuncName);
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Completed with ERROR", sFuncName);
                sw.Dispose();
                sw.Close();
            }
            return sResult;
        }
        #endregion
    }
}
