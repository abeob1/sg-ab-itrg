using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Configuration;
using System.Data.Common;
using System.Data.Odbc;

namespace AE_OutletManagerService_BLL
{
    public class clsCommon
    {
        #region Objects
        clsLog oLog = new clsLog();
        public Int16 p_iDebugMode = DEBUG_ON;

        public const Int16 RTN_SUCCESS = 1;
        public const Int16 RTN_ERROR = 0;
        public const Int16 DEBUG_ON = 1;
        public const Int16 DEBUG_OFF = 0;
        public string sErrDesc = string.Empty;

        #endregion

        #region Methods

        public DataTable ExecuteSQLQuery(string sQuery, string sCompanyCode, OdbcParameter[] param)
        {
            string sFuncName = "ExecuteSQLQuery()";
            string sConstr = ConfigurationManager.ConnectionStrings["DBConnection"].ToString();

            string[] sArray = sConstr.Split(';');
            string sSplitCompany = sConstr.Split(';').Last();
            string sSplit1 = sSplitCompany.Split('=').First();
            string sCompanyGenerate = sSplit1 + "=" + sCompanyCode;

            sConstr = sArray[0] + ";" + sArray[1] + ";" + sArray[2] + ";" + sArray[3] + ";" + sCompanyGenerate;

            System.Data.Odbc.OdbcConnection oCon = new System.Data.Odbc.OdbcConnection(sConstr);
            System.Data.Odbc.OdbcCommand oCmd = new System.Data.Odbc.OdbcCommand();
            DataSet oDs = new DataSet();

            try
            {
                oCon.Open();
                oCmd.CommandType = CommandType.Text;
                oCmd.CommandText = sQuery;
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("SQL Query : " + sQuery, sFuncName);
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Before adding Parameters", sFuncName);
                foreach (var item in param)
                {
                    oCmd.Parameters.Add(item);
                }
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("After adding parameters", sFuncName);
                oCmd.Connection = oCon;
                oCmd.CommandTimeout = 120;
                System.Data.Odbc.OdbcDataAdapter da = new System.Data.Odbc.OdbcDataAdapter(oCmd);
                da.Fill(oDs);
                oCon.Close();
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Completed with SUCCESS", sFuncName);

            }
            catch (Exception ex)
            {
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Completed with ERROR", sFuncName);
                oCon.Dispose();
                throw new Exception(ex.Message);
            }
            return oDs.Tables[0];
        }

        public DataTable ExecuteNonQuery(string sQuery, OdbcParameter[] param)
        {
            string sFuncName = "ExecuteNonQuery";

            if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Starting Function", sFuncName);
            string sConstr = ConfigurationManager.ConnectionStrings["DBConnection"].ToString();
            System.Data.Odbc.OdbcConnection oCon = new System.Data.Odbc.OdbcConnection(sConstr);
            System.Data.Odbc.OdbcCommand oCmd = new System.Data.Odbc.OdbcCommand();
            DataSet oDs = new DataSet();

            try
            {
                oCon.Open();
                oCmd.CommandType = CommandType.Text;
                oCmd.CommandText = sQuery;
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("SQL Query : " + sQuery, sFuncName);
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Before adding Parameters", sFuncName);
                foreach (var item in param)
                {
                    oCmd.Parameters.Add(item);
                }
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("After adding parameters", sFuncName);

                oCmd.Connection = oCon;
                oCmd.CommandTimeout = 120;
                System.Data.Odbc.OdbcDataAdapter da = new System.Data.Odbc.OdbcDataAdapter(oCmd);
                da.Fill(oDs);
                oCon.Close();
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Completed with SUCCESS", sFuncName);
            }
            catch (Exception ex)
            {
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Completed with ERROR", sFuncName);
                oCon.Dispose();
                throw new Exception(ex.Message);
            }
            return oDs.Tables[0];
        }

        public string ExecuteQuery(string sQuery, string sCompanyCode, OdbcParameter[] param)
        {
            string sFuncName = "ExecuteQuery()";

            string sConstr = ConfigurationManager.ConnectionStrings["DBConnection"].ToString();

            string[] sArray = sConstr.Split(';');
            string sSplitCompany = sConstr.Split(';').Last();
            string sSplit1 = sSplitCompany.Split('=').First();
            string sCompanyGenerate = sSplit1 + "=" + sCompanyCode;

            sConstr = sArray[0] + ";" + sArray[1] + ";" + sArray[2] + ";" + sArray[3] + ";" + sCompanyGenerate;
            if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Connection String : " + sConstr, sFuncName);

            System.Data.Odbc.OdbcConnection oCon = new System.Data.Odbc.OdbcConnection(sConstr);
            System.Data.Odbc.OdbcCommand oCmd = new System.Data.Odbc.OdbcCommand();

            try
            {
                oCon.Open();
                oCmd.CommandType = CommandType.Text;
                oCmd.CommandText = sQuery;
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("SQL Query : " + sQuery, sFuncName);
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Before adding Parameters", sFuncName);
                foreach (var item in param)
                {
                    oCmd.Parameters.Add(item);
                }
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("After adding parameters", sFuncName);
                oCmd.Connection = oCon;
                oCmd.CommandTimeout = 120;
                oCmd.ExecuteNonQuery();
                oCon.Close();
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Completed with SUCCESS", sFuncName);
            }
            catch (Exception ex)
            {
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Completed with ERROR", sFuncName);
                oCon.Dispose();
                throw new Exception(ex.Message);
            }
            return "SUCCESS";
        }

        public Int16 ExecuteNonQuery_DR(string sQuery, OdbcParameter[] param)
        {
            string sFuncName = "ExecuteNonQuery_DR";
            Int16 iResult = 0;

            if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Starting Function", sFuncName);
            string sConstr = ConfigurationManager.ConnectionStrings["DBConnection"].ToString();
            System.Data.Odbc.OdbcConnection oCon = new System.Data.Odbc.OdbcConnection(sConstr);
            System.Data.Odbc.OdbcCommand oCmd = new System.Data.Odbc.OdbcCommand();
            DbDataReader oDr;

            try
            {
                oCon.Open();
                oCmd.CommandType = CommandType.Text;
                oCmd.CommandText = sQuery;
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("SQL Query : " + sQuery, sFuncName);
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Before adding Parameters", sFuncName);
                foreach (var item in param)
                {
                    oCmd.Parameters.Add(item);
                }
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("After adding parameters", sFuncName);

                oCmd.Connection = oCon;
                oCmd.CommandTimeout = 120;
                oDr = oCmd.ExecuteReader();
                if (oDr.Read() && oDr.GetValue(0) != DBNull.Value)
                {
                    iResult = Convert.ToInt16(oDr[0]);
                }
                else
                {
                    iResult = 0;
                }
                oCon.Close();
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Completed with SUCCESS", sFuncName);
            }
            catch (Exception ex)
            {
                if (p_iDebugMode == DEBUG_ON) oLog.WriteToDebugLogFile("Completed with ERROR", sFuncName);
                oCon.Dispose();
                throw new Exception(ex.Message);
            }
            return iResult;
        }


        #endregion
    }
}
