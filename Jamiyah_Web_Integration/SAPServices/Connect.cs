using Jamiyah_Web_Integration.SAPModels;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Jamiyah_Web_Integration.SAPServices
{
    public class SBOConnect
    {
        #region "Variables"
        public static SAPbobsCOM.Company sapCompany { get; set; }
        public static SAPbobsCOM.Recordset sapRecSet { get; set; }
        public static SAPbobsCOM.BusinessPartners oBusinessPartners { get; set; }
        public static SAPbobsCOM.Documents oInvoice { get; set; }
        public static SAPbobsCOM.Documents oCreditNote { get; set; }
        public static SAPbobsCOM.Payments oIncomingPayment { get; set; }
        public static SAPbobsCOM.Payments oOutgoingPayment { get; set; }
        public static string oItemDescription { get; set; }
        List<API_BusinessPartners> BusinessPartnersModel { get; set; }
        List<API_Invoice> InvoiceModelHeader { get; set; }
        List<API_InvoiceDetails> InvoiceModelDetails { get; set; }
        List<API_CreditNote> CreditNoteModelHeader { get; set; }
        List<API_CreditNoteDetails> CreditNoteModelDetails { get; set; }
        List<API_CreditRefund> CreditRefundModel { get; set; }
        List<API_Receipt> ReceiptModelHeader { get; set; }
        List<API_ReceiptDetails> ReceiptModelDetails { get; set; }
        List<API_FinanceItem> ItemModel { get; set; }
        public List<API_FinanceItem> ItemMasterModel { get; set; }
        List<ResponseResultSuccess> listResponseResultSuccess { get; set; }
        List<ResponseResultFailed> listResponseResultFailed { get; set; }
        string lastMessage { get; set; }
        string strQuery { get; set; }
        public string oCenterCode { get; set; }
        public string oCurrency { get; set; }
        public string oCountry { get; set; }
        public string oGroup { get; set; }
        public string oDivision { get; set; }
        public string oProduct { get; set; }
        public string api_key { get; set; }
        public string client { get; set; }
        public string base_url { get; set; }
        public int pricelistcode { get; set; }

        #endregion

        #region "Setting Values"
        public void GetIntegrationSetup()
        {
            try
            {
                string oQuery = "select * from \"@INTEGRATIONSETUP\" where \"U_CompanyDB\" = '" + SBOConstantClass.Database + "'";
                sapRecSet.DoQuery(oQuery);
                if (sapRecSet.RecordCount > 0)
                {
                    oCenterCode = sapRecSet.Fields.Item("Name").Value.ToString();
                    oCurrency = sapRecSet.Fields.Item("U_Curr").Value.ToString();
                    oCountry = sapRecSet.Fields.Item("U_Country").Value.ToString();
                    oGroup = sapRecSet.Fields.Item("U_Group").Value.ToString();
                    oDivision = sapRecSet.Fields.Item("U_Division").Value.ToString();
                    oProduct = sapRecSet.Fields.Item("U_Product").Value.ToString();
                    api_key = sapRecSet.Fields.Item("U_api_key").Value.ToString();
                    client = sapRecSet.Fields.Item("Name").Value.ToString();
                    base_url = sapRecSet.Fields.Item("U_base_url").Value.ToString();
                    pricelistcode = Convert.ToInt16(sapRecSet.Fields.Item("U_pricelist_code").Value.ToString());
                }
            }
            catch (Exception ex)
            {
                Console.Write("Cost Center Method" + ex.ToString());
            }
        }
        #endregion


        public bool SBOconnectToLoginCompany(string ServerName, string CompanyDB, string DBUserName, string DBPassword, string SBOUserName, string SBOPassword)
        {
            bool functionReturnValue = false;

            int lErrCode = 0;

            try
            {
                //// Initialize the Company Object.
                //// Create a new company object
                sapCompany = new SAPbobsCOM.Company();

                //// Set the mandatory properties for the connection to the database.
                //// To use a remote Db Server enter his name instead of the string "(local)"
                //// This string is used to work on a DB installed on your local machine

                sapCompany.Server = ServerName;
                sapCompany.CompanyDB = CompanyDB;
                sapCompany.UserName = SBOUserName;
                sapCompany.Password = SBOPassword;
                sapCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
                //// Use Windows authentication for database server.
                //// True for NT server authentication,
                //// False for database server authentication.

                sapCompany.UseTrusted = false;
                if (SBOConstantClass.ServerVersion == "dst_MSSQL2005")
                {
                    sapCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005;
                }
                else if (SBOConstantClass.ServerVersion == "dst_MSSQL2008")
                {
                    sapCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
                }
                else if (SBOConstantClass.ServerVersion == "dst_MSSQL2012")
                {
                    sapCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
                }
                else if (SBOConstantClass.ServerVersion == "dst_MSSQL2014")
                {
                    sapCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
                }
                //else if (SBOConstantClass.ServerVersion == "dst_MSSQL2016")
                //{
                //    sapCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
                //}
                else if (SBOConstantClass.ServerVersion == "dst_HANADB")
                {
                    sapCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
                }

                sapCompany.DbUserName = DBUserName;
                sapCompany.DbPassword = DBPassword;

                //// connection status
                lErrCode = sapCompany.Connect();

                //// Check for errors during connect
                if (lErrCode != 0)
                {
                    lastMessage = "SAP B1 DI API Connection Error : " + sapCompany.GetLastErrorDescription();
                    Console.Write(lastMessage);
                    functionReturnValue = false;
                }
                else
                {
                    sapRecSet = (Recordset)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SBOConstantClass.GetServerDate = sapCompany.GetDBServerDate().ToString("yyyy-MM-dd 00:00:00");
                    functionReturnValue = true;
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return functionReturnValue;
        }
    }

    public static class SBOstrManipulation
    {
        /// <summary>
        /// Get string value after [first] a.
        /// </summary>
        public static string BeforeCharacter(this string value, string a)
        {
            int posA = value.IndexOf(a);
            if (posA == -1)
            {
                return "";
            }
            return value.Substring(0, posA);
        }

        /// <summary>
        /// Get string value after [last] a.
        /// </summary>
        public static string AfterCharacter(this string value, string a)
        {
            int posA = value.LastIndexOf(a);
            if (posA == -1)
            {
                return "";
            }
            int adjustedPosA = posA + a.Length;
            if (adjustedPosA >= value.Length)
            {
                return "";
            }
            return value.Substring(adjustedPosA);
        }
    }
    public class SBOConstantClass
    {
        public static string SBOServer = "";
        public static string ServerUN = "";
        public static string ServerPW = "";
        public static string ServerVersion = "";
        public static string SAPUser = "";
        public static string SAPPassword = "";
        public static string Database = "";
        public static string GetServerDate = "";
    }
    public class SBOCompanyData
    {
        public string CompanyCode { get; set; }
        public string CompanyName { get; set; }
    }
}