using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;

namespace Jamiyah_Web_Integration.SAPServices
{
    public class ConstantClass
    {
        //--------------------------------------------SAP B1 SETTINGS---------------------------------------------------------------------//
        public static string SBOServer = System.Configuration.ConfigurationManager.AppSettings["SBOServer"].ToString();
        public static string ServerUN = System.Configuration.ConfigurationManager.AppSettings["ServerUN"].ToString();
        public static string ServerPW = System.Configuration.ConfigurationManager.AppSettings["ServerPW"].ToString();
        public static string ServerVersion = System.Configuration.ConfigurationManager.AppSettings["ServerVersion"].ToString();
        public static string SAPUser = System.Configuration.ConfigurationManager.AppSettings["SAPUsername"].ToString();
        public static string SAPPassword = System.Configuration.ConfigurationManager.AppSettings["SAPPassword"].ToString();
        public static string AppSetting = System.Configuration.ConfigurationManager.AppSettings["AppStart"].ToString();
        //--------------------------------------------SAP B1 SETTINGS---------------------------------------------------------------------//

        //--------------------------------------------TAIDII SETTING----------------------------------------------------------------------//
        public static string APIBaseURL = "";
        public static string APIKey = "";
        public static string APIClient = "";
        public static string APIFormat = System.Configuration.ConfigurationManager.AppSettings["APIFormat"].ToString();
        public static string APILastTimeStamp = System.Configuration.ConfigurationManager.AppSettings["APILastTimeStamp"].ToString();
        public static string APILastDate = System.Configuration.ConfigurationManager.AppSettings["APILastDate"].ToString();
        public static string ERRORPATH = System.Configuration.ConfigurationManager.AppSettings["ERRORPATH"].ToString();
        //--------------------------------------------TAIDII SETTING----------------------------------------------------------------------//


    }
    public class clsStart
    {
        public bool GetStarted()
        {
            bool retBool = true;
            string FileName = string.Empty;
            string GlobalErrorLogPath = ConstantClass.ERRORPATH + "\\";

            string ErrorMessage = string.Empty;
            string ConnString = string.Empty;
            string SelectQuery = string.Empty;
            //HanaConnection _HanaConnection = null;
            SqlConnection _SqlConnection = null;
            DataTable _DataTable = null;

            try
            {
                Console.WriteLine("API Connection is now Initializing!");
                //Call initialized constants      
                InitializedConstants();
                if (ConstantClass.AppSetting == "0")
                {
                    try
                    {
                        
                        //Connection String for MSSQL
                        ConnString = "Data Source=" + SBOConstantClass.SBOServer + ";Initial Catalog=TAIDII_SAP; User ID=" + SBOConstantClass.ServerUN + ";Password=" + SBOConstantClass.ServerPW + ";Integrated Security=false;";
                        _SqlConnection = new SqlConnection(ConnString);
                        _SqlConnection.Open();

                        SelectQuery = "select *,CONVERT(VARCHAR, GETDATE(), 23) \"current_date\" from " + " \"TAIDII_SAP\"" + ".." + "\"axxis_tb_IntegrationSetup\"";

                        SqlDataAdapter _SqlDataAdapter = new SqlDataAdapter(SelectQuery, _SqlConnection);
                        DataSet _DataSet = new DataSet();
                        _SqlDataAdapter.Fill(_DataSet);
                        _DataTable = _DataSet.Tables[0];

                        if (_DataTable.Rows.Count > 0)
                        {
                            for (int i = 0; i <= _DataTable.Rows.Count - 1; i++)
                            {
                                SBOConstantClass.Database = _DataTable.Rows[i]["companyDB"].ToString();
                                string ErrorLogPath = ConstantClass.ERRORPATH + "\\" + SBOConstantClass.Database + "\\";

                                if (!Directory.Exists(ErrorLogPath))
                                {
                                    Directory.CreateDirectory(ErrorLogPath);
                                }                                
                            }
                        }                     
                    }
                    catch (Exception ex)
                    {
                        string Message = ex.Message;
                        retBool = false;

                        ErrorMessage += "Exception Error : " + ex.Message + Environment.NewLine;
                        retBool = false;
                        //write to text;
                    }
                    finally
                    {
                        if (!string.IsNullOrEmpty(ErrorMessage))
                        {
                            WriteLog(GlobalErrorLogPath, ErrorMessage);
                        }
                    }
                }                
                return retBool;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        static void WriteLog(string filename, string msg)
        {
            try
            {
                string fname = filename + DateTime.Now.ToString("dd.MM.yyyy HHmm") + ".txt";
                using (StreamWriter writer = new StreamWriter(fname, true))
                {
                    if (msg == "----------------------------------------------------------------")
                    {
                        writer.WriteLine(msg);
                    }
                    else
                    {
                        writer.WriteLine(msg + " || timestamp: " + DateTime.Now.ToString("HH:mm:ss"));
                    }

                    writer.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        public void InitializedConstants()
        {
            SBOConstantClass.SBOServer = ConstantClass.SBOServer;
            SBOConstantClass.ServerVersion = ConstantClass.ServerVersion;
            SBOConstantClass.SAPUser = ConstantClass.SAPUser;
            SBOConstantClass.SAPPassword = ConstantClass.SAPPassword;
            SBOConstantClass.ServerUN = ConstantClass.ServerUN;
            SBOConstantClass.ServerPW = ConstantClass.ServerPW;
        }
    }
}