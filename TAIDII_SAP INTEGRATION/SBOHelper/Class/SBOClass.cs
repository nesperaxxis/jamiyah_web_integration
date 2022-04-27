using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using SAPbobsCOM;
using System.Globalization;
using Newtonsoft.Json;
using System.Xml.Serialization;
using System.Xml.Linq;
using System.Web.Http;
using System.Net.Http;

namespace SBOHelper.Class
{
    public class SBOClass
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
        List<Models.API_BusinessPartners> BusinessPartnersModel { get; set; }
        List<Models.API_Invoice> InvoiceModelHeader { get; set; }
        List<Models.API_InvoiceDetails> InvoiceModelDetails { get; set; }
        List<Models.API_CreditNote> CreditNoteModelHeader { get; set; }
        List<Models.API_CreditNoteDetails> CreditNoteModelDetails { get; set; }
        List<Models.API_CreditRefund> CreditRefundModel { get; set; }
        List<Models.API_Receipt> ReceiptModelHeader { get; set; }
        List<Models.API_ReceiptDetails> ReceiptModelDetails { get; set; }
        List<Models.API_FinanceItem> ItemModel { get; set; }
        public List<Models.API_FinanceItem> ItemMasterModel { get; set; }
        List<Models.ResponseResultSuccess> listResponseResultSuccess { get; set; }
        List<Models.ResponseResultFailed> listResponseResultFailed { get; set; }
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

        #region "SAP B1 DI API Connection"
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
        #endregion

        #region "TAIDII - SAP B1 Integration"
        public bool SBOPostBusinessPartners(List<SBOHelper.Models.API_BusinessPartners> listBusinessParters, string APILastTimeStamp)
        {
            bool functionReturnValue = false;
            int lErrCode = 0;
            int RecordCount = 0;
            string oLogExist = string.Empty;
            string oCardCode = string.Empty;
            string oCountry = string.Empty;
            string oBPMaster = string.Empty;
            string GroupCode = string.Empty;
            SBOGetRecord clsSBOGetRecord = new SBOGetRecord();
            try
            {
                if (SBOconnectToLoginCompany(SBOConstantClass.SBOServer, SBOConstantClass.Database, SBOConstantClass.ServerUN, SBOConstantClass.ServerPW, SBOConstantClass.SAPUser, SBOConstantClass.SAPPassword))
                {
                    foreach (var iRowBP in listBusinessParters)
                    {
                        try
                        {
                            oBPMaster = iRowBP.BPMaster;

                            oBusinessPartners = (BusinessPartners)sapCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                            BusinessPartnersModel = new List<SBOHelper.Models.API_BusinessPartners>();
                            ////**** Create a list of Business Partners ****////
                            BusinessPartnersModel.Add(new SBOHelper.Models.API_BusinessPartners()
                            {
                                id = iRowBP.id,
                                BPMaster = iRowBP.BPMaster,
                                fullname = iRowBP.fullname,
                                nric = iRowBP.nric,
                                gender = iRowBP.gender,
                                dob = iRowBP.dob,
                                student_care_type = iRowBP.student_care_type,
                                program_type = iRowBP.program_type,
                                registration_no = iRowBP.registration_no,
                                subsidy = iRowBP.subsidy,
                                additional_subsidy = iRowBP.additional_subsidy,
                                financial_assistance = iRowBP.financial_assistance,
                                deposit = iRowBP.deposit,
                                nationality = iRowBP.nationality,
                                race = iRowBP.race,
                                address = iRowBP.address,
                                unit_no = iRowBP.unit_no,
                                postal_code = iRowBP.postal_code,
                                date_of_withdrawal = iRowBP.date_of_withdrawal,
                                country = iRowBP.country,
                                contact_name = iRowBP.contact_name,
                                contact_nric = iRowBP.contact_nric,
                                contact_relation = iRowBP.contact_relation,
                                contact_email = iRowBP.contact_email,
                                contact_telephone = iRowBP.contact_telephone,
                                contact_office_no = iRowBP.contact_office_no,
                                contact_home_phone = iRowBP.contact_home_phone,
                                bank_name = iRowBP.bank_name,
                                account_name = iRowBP.account_name,
                                cdac_bank_no = iRowBP.cdac_bank_no,
                                customer_ref_no = iRowBP.customer_ref_no,
                                admission_date = iRowBP.admission_date,
                                level = iRowBP.level
                            });

                            string strJSON = JsonConvert.SerializeObject(BusinessPartnersModel);

                            oLogExist = (String)clsSBOGetRecord.GetSingleValue("select * from " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " where \"companyDB\" = '" + sapCompany.CompanyDB + "' and \"module\" = 'Student' and \"uniqueId\" = '" + TrimData(iRowBP.BPMaster) + "'", sapCompany);
                            if (oLogExist == "" || oLogExist == "0")
                            {
                                Console.WriteLine("Adding Students:" + iRowBP.BPMaster + " in the integration log. Please wait...");
                                strQuery = "insert into " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " (\"lastTimeStamp\",\"companyDB\",\"module\",\"uniqueId\",\"docStatus\",\"status\",\"JSON\",\"statusCode\",\"successDesc\",\"failDesc\",\"logDate\",\"reference\",\"objType\") select '" + iif(APILastTimeStamp != "", APILastTimeStamp, sapCompany.GetDBServerDate().ToString("yyyy-MM-dd")) + "','" + TrimData(sapCompany.CompanyDB) + "','Student','" + TrimData(iRowBP.BPMaster) + "','Confirmed','','" + TrimData(strJSON) + "','','','',null,'" + iRowBP.id + "',2 " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "", "from dummy;") + "";

                                sapRecSet.DoQuery(strQuery);
                                RecordCount += 1;
                            }
                            ////**** Create a list of Business Partners ****////

                            if (BusinessPartnersModel.Count > 0)
                                Console.WriteLine("Processing Student:" + iRowBP.fullname + " in SAP B1. Please wait...");

                            oCardCode = (String)clsSBOGetRecord.GetSingleValue("select \"CardCode\" from \"OCRD\" where \"CardCode\" = '" + TrimData(iRowBP.BPMaster) + "'", sapCompany);
                            if (oCardCode == "" || oCardCode == "0")
                            {
                                ////**** Creation of Business Partners in SAP B1 ****/////
                                oBusinessPartners.CardType = BoCardTypes.cCustomer;

                                GroupCode = (String)clsSBOGetRecord.GetSingleValue("select \"GroupCode\" from \"OCRG\" where \"GroupName\" = 'STUDENT (Customer)'", sapCompany);

                                if (!string.IsNullOrEmpty(GroupCode) && GroupCode != "0")
                                    oBusinessPartners.GroupCode = Convert.ToInt16(GroupCode);

                                if (!string.IsNullOrEmpty(iRowBP.BPMaster))
                                    oBusinessPartners.CardCode = iRowBP.BPMaster;

                                if (!string.IsNullOrEmpty(iRowBP.fullname))
                                    oBusinessPartners.CardName = iRowBP.fullname;


                                oBusinessPartners.Addresses.AddressType = BoAddressType.bo_BillTo;
                                oBusinessPartners.Addresses.AddressName = "Home Address";

                                if (!string.IsNullOrEmpty(iRowBP.address))
                                    oBusinessPartners.Addresses.Street = iRowBP.address;

                                if (!string.IsNullOrEmpty(iRowBP.unit_no))
                                    oBusinessPartners.Addresses.Block = iRowBP.unit_no;

                                if (!string.IsNullOrEmpty(iRowBP.postal_code))
                                    oBusinessPartners.Addresses.ZipCode = iRowBP.postal_code;

                                //oBusinessPartners.Addresses.Street = "Sta. Monica Heights Subd. San Rafael Tarlac City";
                                //oBusinessPartners.Addresses.Block = "Center Road";
                                //oBusinessPartners.Addresses.ZipCode = "2301";

                                ////**** Contact Person ****////
                                if (!string.IsNullOrEmpty(iRowBP.contact_name))
                                    oBusinessPartners.ContactEmployees.Name = iRowBP.contact_name;

                                if (!string.IsNullOrEmpty(iRowBP.contact_relation))
                                    oBusinessPartners.ContactEmployees.Position = iRowBP.contact_relation;

                                if (!string.IsNullOrEmpty(iRowBP.contact_nric))
                                    oBusinessPartners.ContactEmployees.FirstName = iRowBP.contact_nric;

                                if (!string.IsNullOrEmpty(iRowBP.contact_telephone))
                                    oBusinessPartners.ContactEmployees.Phone1 = iRowBP.contact_telephone;

                                if (!string.IsNullOrEmpty(iRowBP.contact_email))
                                    oBusinessPartners.ContactEmployees.E_Mail = iRowBP.contact_email;
                                ////**** Contact Person ****////

                                ////**** User defined fields ****////
                                if (!string.IsNullOrEmpty(iRowBP.country) && iRowBP.country != "N.A")
                                    oBusinessPartners.UserFields.Fields.Item("U_Country").Value = iRowBP.country;

                                if (!string.IsNullOrEmpty(iRowBP.level))
                                    oBusinessPartners.UserFields.Fields.Item("U_Level").Value = iRowBP.level;

                                if (!string.IsNullOrEmpty(iRowBP.nric))
                                    oBusinessPartners.UserFields.Fields.Item("U_IC").Value = iRowBP.nric;

                                if (!string.IsNullOrEmpty(Convert.ToString(iRowBP.gender)))
                                    oBusinessPartners.UserFields.Fields.Item("U_Gender").Value = iRowBP.gender;

                                if (!string.IsNullOrEmpty(iRowBP.dob))
                                    oBusinessPartners.UserFields.Fields.Item("U_DOB").Value = Convert.ToDateTime(iRowBP.dob);

                                if (!string.IsNullOrEmpty(iRowBP.student_care_type))
                                    oBusinessPartners.UserFields.Fields.Item("U_STD_CARE_TYPE").Value = iRowBP.student_care_type;

                                if (!string.IsNullOrEmpty(iRowBP.program_type))
                                    oBusinessPartners.UserFields.Fields.Item("U_ProgramType").Value = iRowBP.program_type;

                                if (!string.IsNullOrEmpty(iRowBP.admission_date))
                                    oBusinessPartners.UserFields.Fields.Item("U_AdmissionDate").Value = Convert.ToDateTime(iRowBP.admission_date);

                                if (!string.IsNullOrEmpty(iRowBP.registration_no))
                                    oBusinessPartners.UserFields.Fields.Item("U_RegNo").Value = iRowBP.registration_no;

                                if (iRowBP.subsidy != 0)
                                    oBusinessPartners.UserFields.Fields.Item("U_Subsidy").Value = iRowBP.subsidy;

                                if (iRowBP.additional_subsidy != 0)
                                    oBusinessPartners.UserFields.Fields.Item("U_Add_Subsidy").Value = iRowBP.additional_subsidy;

                                if (iRowBP.financial_assistance != 0)
                                    oBusinessPartners.UserFields.Fields.Item("U_FinAssist").Value = iRowBP.financial_assistance;

                                if (iRowBP.deposit != 0)
                                    oBusinessPartners.UserFields.Fields.Item("U_Deposit").Value = iRowBP.deposit;

                                if (!string.IsNullOrEmpty(iRowBP.nationality))
                                    oBusinessPartners.UserFields.Fields.Item("U_Nationality").Value = iRowBP.nationality;

                                if (!string.IsNullOrEmpty(iRowBP.race))
                                    oBusinessPartners.UserFields.Fields.Item("U_Race").Value = iRowBP.race;

                                if (!string.IsNullOrEmpty(iRowBP.bank_name))
                                    oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_Bankname").Value = iRowBP.bank_name;

                                if (!string.IsNullOrEmpty(iRowBP.account_name))
                                    oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_AccName").Value = iRowBP.account_name;

                                if (!string.IsNullOrEmpty(iRowBP.cdac_bank_no))
                                    oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_BankAccNo").Value = iRowBP.cdac_bank_no;

                                if (!string.IsNullOrEmpty(iRowBP.customer_ref_no))
                                    oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_CusRefNo").Value = iRowBP.customer_ref_no;
                                ////**** User defined fields ****////

                                lErrCode = oBusinessPartners.Add();
                                if (lErrCode == 0)
                                {
                                    try
                                    {
                                        oCardCode = sapCompany.GetNewObjectKey();
                                        lastMessage = "Successfully created Customer Code: " + oCardCode + " in SAP B1.";
                                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Posted',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oCardCode + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Student' and \"uniqueId\" = '" + TrimData(iRowBP.BPMaster) + "'");

                                        functionReturnValue = false;
                                    }
                                    catch (Exception)
                                    {
                                        lastMessage = sapCompany.GetLastErrorDescription();
                                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'false',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Student' and \"uniqueId\" = '" + iRowBP.BPMaster + "'");

                                        functionReturnValue = true;
                                    }
                                }
                                else
                                {
                                    lastMessage = sapCompany.GetLastErrorDescription();
                                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'false',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Student' and \"uniqueId\" = '" + iRowBP.BPMaster + "'");

                                    functionReturnValue = true;

                                    goto isAddWithError;
                                }

                            isAddWithError: ;

                                ////**** Creation of Business Partners in SAP B1 ****/////
                            }
                            else
                            {
                                if (oBusinessPartners.GetByKey(oCardCode) == true)
                                {
                                    if (!string.IsNullOrEmpty(iRowBP.fullname))
                                        oBusinessPartners.CardName = iRowBP.fullname;

                                    oBusinessPartners.Addresses.AddressType = BoAddressType.bo_BillTo;
                                    oBusinessPartners.Addresses.AddressName = "Home Address";

                                    if (!string.IsNullOrEmpty(iRowBP.address))
                                        oBusinessPartners.Addresses.Street = iRowBP.address;

                                    if (!string.IsNullOrEmpty(iRowBP.unit_no))
                                        oBusinessPartners.Addresses.Block = iRowBP.unit_no;

                                    if (!string.IsNullOrEmpty(iRowBP.postal_code))
                                        oBusinessPartners.Addresses.ZipCode = iRowBP.postal_code;

                                    //oBusinessPartners.Addresses.Street = "Sta. Monica Heights Subd. San Rafael Tarlac City";
                                    //oBusinessPartners.Addresses.Block = "Center Road";
                                    //oBusinessPartners.Addresses.ZipCode = "2301";

                                    ////**** Contact Person ****////
                                    if (!string.IsNullOrEmpty(iRowBP.contact_name))
                                        oBusinessPartners.ContactEmployees.Name = iRowBP.contact_name;

                                    if (!string.IsNullOrEmpty(iRowBP.contact_relation))
                                        oBusinessPartners.ContactEmployees.Position = iRowBP.contact_relation;

                                    if (!string.IsNullOrEmpty(iRowBP.contact_nric))
                                        oBusinessPartners.ContactEmployees.FirstName = iRowBP.contact_nric;

                                    if (!string.IsNullOrEmpty(iRowBP.contact_telephone))
                                        oBusinessPartners.ContactEmployees.Phone1 = iRowBP.contact_telephone;

                                    if (!string.IsNullOrEmpty(iRowBP.contact_email))
                                        oBusinessPartners.ContactEmployees.E_Mail = iRowBP.contact_email;
                                    ////**** Contact Person ****////

                                    ////**** User defined fields ****////
                                    if (!string.IsNullOrEmpty(iRowBP.country) && iRowBP.country != "N.A")
                                        oBusinessPartners.UserFields.Fields.Item("U_Country").Value = iRowBP.country;

                                    if (!string.IsNullOrEmpty(iRowBP.level))
                                        oBusinessPartners.UserFields.Fields.Item("U_Level").Value = iRowBP.level;

                                    if (!string.IsNullOrEmpty(iRowBP.nric))
                                        oBusinessPartners.UserFields.Fields.Item("U_IC").Value = iRowBP.nric;

                                    if (!string.IsNullOrEmpty(Convert.ToString(iRowBP.gender)))
                                        oBusinessPartners.UserFields.Fields.Item("U_Gender").Value = iRowBP.gender;

                                    if (!string.IsNullOrEmpty(iRowBP.dob))
                                        oBusinessPartners.UserFields.Fields.Item("U_DOB").Value = Convert.ToDateTime(iRowBP.dob);

                                    if (!string.IsNullOrEmpty(iRowBP.student_care_type))
                                        oBusinessPartners.UserFields.Fields.Item("U_STD_CARE_TYPE").Value = iRowBP.student_care_type;

                                    if (!string.IsNullOrEmpty(iRowBP.program_type))
                                        oBusinessPartners.UserFields.Fields.Item("U_ProgramType").Value = iRowBP.program_type;

                                    if (!string.IsNullOrEmpty(iRowBP.admission_date))
                                        oBusinessPartners.UserFields.Fields.Item("U_AdmissionDate").Value = Convert.ToDateTime(iRowBP.admission_date);

                                    if (!string.IsNullOrEmpty(iRowBP.registration_no))
                                        oBusinessPartners.UserFields.Fields.Item("U_RegNo").Value = iRowBP.registration_no;

                                    if (iRowBP.subsidy != 0)
                                        oBusinessPartners.UserFields.Fields.Item("U_Subsidy").Value = iRowBP.subsidy;

                                    if (iRowBP.additional_subsidy != 0)
                                        oBusinessPartners.UserFields.Fields.Item("U_Add_Subsidy").Value = iRowBP.additional_subsidy;

                                    if (iRowBP.financial_assistance != 0)
                                        oBusinessPartners.UserFields.Fields.Item("U_FinAssist").Value = iRowBP.financial_assistance;

                                    if (iRowBP.deposit != 0)
                                        oBusinessPartners.UserFields.Fields.Item("U_Deposit").Value = iRowBP.deposit;

                                    if (!string.IsNullOrEmpty(iRowBP.nationality))
                                        oBusinessPartners.UserFields.Fields.Item("U_Nationality").Value = iRowBP.nationality;

                                    if (!string.IsNullOrEmpty(iRowBP.race))
                                        oBusinessPartners.UserFields.Fields.Item("U_Race").Value = iRowBP.race;

                                    if (!string.IsNullOrEmpty(iRowBP.bank_name))
                                        oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_Bankname").Value = iRowBP.bank_name;

                                    if (!string.IsNullOrEmpty(iRowBP.account_name))
                                        oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_AccName").Value = iRowBP.account_name;

                                    if (!string.IsNullOrEmpty(iRowBP.cdac_bank_no))
                                        oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_BankAccNo").Value = iRowBP.cdac_bank_no;

                                    if (!string.IsNullOrEmpty(iRowBP.customer_ref_no))
                                        oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_CusRefNo").Value = iRowBP.customer_ref_no;
                                    ////**** User defined fields ****////

                                    lErrCode = oBusinessPartners.Update();
                                    if (lErrCode == 0)
                                    {
                                        try
                                        {
                                            oCardCode = sapCompany.GetNewObjectKey();
                                            lastMessage = "Successfully updated Customer Code: " + oCardCode + " in SAP B1.";
                                            strQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Posted',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oCardCode + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Student' and \"uniqueId\" = '" + TrimData(iRowBP.BPMaster) + "'";
                                            sapRecSet.DoQuery(strQuery);

                                            functionReturnValue = false;
                                        }
                                        catch
                                        { }
                                    }
                                    else
                                    {
                                        lastMessage = sapCompany.GetLastErrorDescription();
                                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'false',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Student' and \"uniqueId\" = '" + TrimData(iRowBP.BPMaster) + "'");

                                        functionReturnValue = true;

                                        goto isUpdateWithError;
                                    }
                                }

                            isUpdateWithError: ;

                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBusinessPartners);
                        }
                        catch (Exception ex)
                        {
                            lastMessage = ex.ToString();
                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'false',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Student' and \"uniqueId\" = '" + TrimData(iRowBP.BPMaster) + "'");

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBusinessPartners);
                        }
                    }
                    Console.WriteLine("Done adding the List of " + string.Format("{0:#,##0}", RecordCount) + " Students in the integration log. Please wait...");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return functionReturnValue;
        }

        public bool SBOPostInvoice(List<SBOHelper.Models.API_Invoice> listInvoice, string APILastTimeStamp)
        {
            bool functionReturnValue = false;
            int lErrCode = 0;
            int RecordCount = 0;
            int oId = 0;
            int oStatus = 0;
            string oLogExist = string.Empty;
            string oTransId = string.Empty;
            string oCardCode = string.Empty;
            string oCardName = string.Empty;
            string oDocEntry = string.Empty;
            string oDescription = string.Empty;
            string oItemCode = string.Empty;
            string oDocType = string.Empty;
            SBOGetRecord clsSBOGetRecord = new SBOGetRecord();
            try
            {
                if (SBOconnectToLoginCompany(SBOConstantClass.SBOServer, SBOConstantClass.Database, SBOConstantClass.ServerUN, SBOConstantClass.ServerPW, SBOConstantClass.SAPUser, SBOConstantClass.SAPPassword))
                {
                    GetIntegrationSetup();
                    foreach (var iRowInv in listInvoice)
                    {
                        try
                        {
                            oId = iRowInv.id;
                            oStatus = iRowInv.status;

                            InvoiceModelHeader = new List<SBOHelper.Models.API_Invoice>();
                            InvoiceModelDetails = new List<SBOHelper.Models.API_InvoiceDetails>();

                            ////**** Create a list of Invoices ****////
                            foreach (var iRowInvDtl in iRowInv.items.ToList())
                            {
                                InvoiceModelDetails.Add(new SBOHelper.Models.API_InvoiceDetails()
                                {
                                    description = iRowInvDtl.description,
                                    item_code = iRowInvDtl.item_code,
                                    date_for = iRowInvDtl.date_for,
                                    unit_price = iRowInvDtl.unit_price,
                                    quantity = iRowInvDtl.quantity,
                                    total = iRowInvDtl.total
                                });
                            }

                            InvoiceModelHeader.Add(new SBOHelper.Models.API_Invoice()
                            {
                                id = iRowInv.id,
                                invoice_no = iRowInv.invoice_no,
                                date_created = iRowInv.date_created,
                                date_due = iRowInv.date_due,
                                status = iRowInv.status,
                                remarks = iRowInv.remarks,
                                void_remarks = iRowInv.void_remarks,
                                student = iRowInv.student,
                                level = iRowInv.level,
                                program_type = iRowInv.program_type,
                                items = InvoiceModelDetails.ToList()
                            });

                            string strJSON = JsonConvert.SerializeObject(InvoiceModelHeader);

                            oLogExist = (String)clsSBOGetRecord.GetSingleValue("select * from " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowInv.id + "'", sapCompany);

                            if (oLogExist == "" || oLogExist == "0")
                            {
                                Console.WriteLine("Adding Invoice with Transaction Id:" + iRowInv.id + " in the integration log. Please wait...");
                                strQuery = "insert into " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " (\"lastTimeStamp\",\"companyDB\",\"module\",\"uniqueId\",\"docStatus\",\"status\",\"JSON\",\"statusCode\",\"successDesc\",\"failDesc\",\"logDate\",\"reference\") select '" + iif(APILastTimeStamp != "", APILastTimeStamp, sapCompany.GetDBServerDate().ToString("yyyy-MM-dd")) + "','" + TrimData(sapCompany.CompanyDB) + "','Invoice','" + iRowInv.id + "','" + iif(iRowInv.status == 1, "Confirmed", "Void") + "','Draft','" + TrimData(strJSON) + "','For Process','','',null,'" + iRowInv.invoice_no + "' " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "", "from dummy;") + "";
                                sapRecSet.DoQuery(strQuery);
                                RecordCount += 1;
                            }
                            else
                            {
                                if (iRowInv.status == 2)
                                {
                                    Console.WriteLine("Updating Invoice with Transaction Id:" + iRowInv.id + " in the integration log. Please wait...");
                                    strQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"docStatus\" = '" + iif(iRowInv.status == 1, "Confirmed", "Void") + "',\"statusCode\" = 'For Process',\"JSON\" = '" + TrimData(strJSON) + "',\"logDate\" = '" + iif(APILastTimeStamp != "", APILastTimeStamp, sapCompany.GetDBServerDate().ToString("yyyy-MM-dd")) + "',\"objType\" = 14 where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowInv.id + "' and \"docStatus\" = 'Confirmed'";
                                    sapRecSet.DoQuery(strQuery);
                                }
                            }
                            ////**** Create a list of Invoices ****////

                            if (iRowInv.status == 1)
                            {
                                Console.WriteLine("Processing Invoice with Transaction Id:" + iRowInv.id + " in SAP B1 Draft. Please wait...");
                               
                                string Query="select \"U_TransId\" from \"ODRF\" where \"U_TransId\" = '" + iRowInv.id + "' and \"CANCELED\" = 'N' and \"NumAtCard\" = '" + iRowInv.invoice_no + "' and \"ObjType\" = 13 " + Environment.NewLine +
                                "union all " + Environment.NewLine +
                                "select \"U_TransId\" from \"OINV\" where \"U_TransId\" = '" + iRowInv.id + "' and \"CANCELED\" = 'N' and \"NumAtCard\" = '" + iRowInv.invoice_no + "'";
                                oTransId = (String)clsSBOGetRecord.GetSingleValue(Query, sapCompany);
                                if (oTransId == "" || oTransId == "0")
                                {
                                    oInvoice = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oDrafts);
                                    oInvoice.DocObjectCode = BoObjectTypes.oInvoices;

                                    oCardCode = (String)clsSBOGetRecord.GetSingleValue("select \"CardCode\" from \"OCRD\" where \"CardCode\" = '" + TrimData(iRowInv.student) + "'", sapCompany);
                                    if (oCardCode != "")
                                    {
                                        oInvoice.CardCode = oCardCode;
                                    }
                                    else
                                    {
                                        lastMessage = "Customer Code:" + iRowInv.student + " is not found in SAP B1";
                                        strQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowInv.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowInv.id + "'";
                                        sapRecSet.DoQuery(strQuery);

                                        functionReturnValue = true;

                                        goto isAddWithError;
                                    }

                                    oInvoice.DocDate = Convert.ToDateTime(iRowInv.date_created);
                                    oInvoice.NumAtCard = iRowInv.invoice_no;
                                    oInvoice.DocDueDate = Convert.ToDateTime(iRowInv.date_due);

                                    if (iRowInv.status == 1)
                                        oInvoice.Comments = iRowInv.remarks;
                                    else
                                        oInvoice.Comments = iRowInv.void_remarks;

                                    ////**** UDF *****/////
                                    if (iRowInv.id != 0)
                                        oInvoice.UserFields.Fields.Item("U_TransId").Value = iRowInv.id.ToString();

                                    if (iRowInv.level != "")
                                        oInvoice.UserFields.Fields.Item("U_Level").Value = iRowInv.level;

                                    if (iRowInv.program_type != "")
                                        oInvoice.UserFields.Fields.Item("U_ProgramType").Value = iRowInv.program_type;
                                    ////**** UDF *****/////

                                    foreach (var iRowInvDtls in iRowInv.items.ToList())
                                    {
                                        if (iRowInvDtls.item_code == "" || string.IsNullOrEmpty(iRowInvDtls.item_code))
                                        {
                                            oDocType = "dDocument_Service";
                                            string iReplaceDesc = " (" + TrimData(iRowInv.level) + " - " + TrimData(iRowInv.program_type) + ")";
                                            //oDescription = SBOstrManipulation.BeforeCharacter(iRowInvDtls.description, " (");
                                            oDescription = iRowInvDtls.description.Replace(iReplaceDesc, "");

                                            if (oDescription != "")
                                            {
                                                string description = oDescription;
                                                string iDescription = (String)clsSBOGetRecord.GetSingleValue("select \"U_Description\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(iRowInv.level) + "' and \"U_ProgramType\" = '" + TrimData(iRowInv.program_type) + "'", sapCompany);
                                                if (iDescription != "")
                                                {
                                                    string idate_created = string.Empty;
                                                    string idate_for = string.Empty;
                                                    string iGLAccount = string.Empty;
                                                    string oDateFor = string.Empty;

                                                    if (!string.IsNullOrEmpty(iRowInvDtls.date_for))
                                                    {
                                                        idate_for = iRowInvDtls.date_for;
                                                        oDateFor = Convert.ToDateTime(idate_for).ToString("MMM") + " " + Convert.ToDateTime(idate_for).Year.ToString();
                                                    }
                                                    else
                                                    {
                                                        idate_for = iRowInv.date_created;
                                                        oDateFor = Convert.ToDateTime(idate_for).ToString("MMM") + " " + Convert.ToDateTime(idate_for).Year.ToString();
                                                    }

                                                    oCardName = (String)clsSBOGetRecord.GetSingleValue("select \"CardName\" from \"OCRD\" where \"CardCode\" = '" + TrimData(iRowInv.student) + "'", sapCompany);

                                                    string oTaxCode = (String)clsSBOGetRecord.GetSingleValue("select \"U_TaxCode\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(iRowInv.level) + "' and \"U_ProgramType\" = '" + TrimData(iRowInv.program_type) + "'", sapCompany);

                                                    if (!string.IsNullOrEmpty(oTaxCode))
                                                        oInvoice.Lines.VatGroup = oTaxCode;

                                                    oItemDescription = oCardName + " - " + oDateFor + " - " + iRowInvDtls.description;
                                                    oInvoice.Lines.UserFields.Fields.Item("U_Dscription").Value = oItemDescription;

                                                    string Dscription = string.Empty;
                                                    if (oItemDescription.Length > 100)
                                                    {
                                                        Dscription = oItemDescription.Substring(0, 100);
                                                        oInvoice.Lines.ItemDescription = Dscription;
                                                    }
                                                    else
                                                    {
                                                        oInvoice.Lines.ItemDescription = oItemDescription;
                                                    }

                                                    oInvoice.Lines.LineTotal = iRowInvDtls.unit_price;

                                                    if (!string.IsNullOrEmpty(iRowInv.date_created))
                                                        idate_created = iRowInv.date_created;

                                                    if (string.IsNullOrEmpty(oCountry) || string.IsNullOrEmpty(oGroup) || string.IsNullOrEmpty(oDivision) || string.IsNullOrEmpty(oProduct))
                                                    {
                                                        lastMessage = "Cost Center is not defined in SAP B1. Please define in the integration setup.";
                                                        string oQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowInv.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowInv.id + "'";
                                                        sapRecSet.DoQuery(oQuery);

                                                        functionReturnValue = true;

                                                        goto isAddWithError;
                                                    }

                                                    if (!string.IsNullOrEmpty(oCountry))
                                                        oInvoice.Lines.CostingCode = oCountry;

                                                    if (!string.IsNullOrEmpty(oGroup))
                                                        oInvoice.Lines.CostingCode2 = oGroup;

                                                    if (!string.IsNullOrEmpty(oDivision))
                                                        oInvoice.Lines.CostingCode3 = oDivision;

                                                    if (!string.IsNullOrEmpty(oProduct))
                                                        oInvoice.Lines.CostingCode4 = oProduct;

                                                    if (!string.IsNullOrEmpty(idate_for))
                                                        oInvoice.Lines.UserFields.Fields.Item("U_date_for").Value = Convert.ToDateTime(idate_for);

                                                    if (CheckDate(idate_created) == true && CheckDate(idate_for) == true)
                                                    {
                                                        if (Convert.ToDateTime(idate_for) > Convert.ToDateTime(idate_created))
                                                        {
                                                            iGLAccount = (String)clsSBOGetRecord.GetSingleValue("select \"U_FuturePeriod\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(iRowInv.level) + "' and \"U_ProgramType\" = '" + TrimData(iRowInv.program_type) + "'", sapCompany);
                                                        }
                                                        else
                                                        {
                                                            iGLAccount = (String)clsSBOGetRecord.GetSingleValue("select \"U_CurrentPeriod\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(iRowInv.level) + "' and \"U_ProgramType\" = '" + TrimData(iRowInv.program_type) + "'", sapCompany);
                                                        }
                                                    }

                                                    if (!string.IsNullOrEmpty(iGLAccount))
                                                        oInvoice.Lines.AccountCode = iGLAccount;

                                                    oInvoice.Lines.Add();
                                                }
                                                else
                                                {
                                                    lastMessage = "Description:" + iRowInvDtls.description + ", Level: " + iRowInv.level + " or Program type:" + iRowInv.program_type + " is not defined in SAP B1. Please define in the table.";
                                                    string oQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowInv.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowInv.id + "'";
                                                    sapRecSet.DoQuery(oQuery);

                                                    functionReturnValue = true;

                                                    goto isAddWithError;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            oDocType = "dDocument_Items";
                                            oItemCode = string.Empty;
                                            oItemCode = (String)clsSBOGetRecord.GetSingleValue("select \"ItemCode\" from \"OITM\" where \"ItemCode\" = '" + TrimData(iRowInvDtls.item_code) + "'", sapCompany);
                                            if (oItemCode != "" || !string.IsNullOrEmpty(oItemCode))
                                            {
                                                oInvoice.Lines.ItemCode = iRowInvDtls.item_code;

                                                oInvoice.Lines.FreeText = iRowInvDtls.description;

                                                if (iRowInvDtls.quantity != 0)
                                                    oInvoice.Lines.Quantity = iRowInvDtls.quantity;

                                                if (iRowInvDtls.unit_price != 0)
                                                    oInvoice.Lines.UnitPrice = iRowInvDtls.unit_price;

                                                if (!string.IsNullOrEmpty(oCountry))
                                                    oInvoice.Lines.CostingCode = oCountry;

                                                if (!string.IsNullOrEmpty(oGroup))
                                                    oInvoice.Lines.CostingCode2 = oGroup;

                                                if (!string.IsNullOrEmpty(oDivision))
                                                    oInvoice.Lines.CostingCode3 = oDivision;

                                                if (!string.IsNullOrEmpty(oProduct))
                                                    oInvoice.Lines.CostingCode4 = oProduct;

                                                oInvoice.Lines.Add();
                                            }
                                            else
                                            {
                                                lastMessage = "ItemCode: " + iRowInvDtls.item_code + " does not exist in SAP B1.";
                                                sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowInv.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowInv.id + "'");

                                                functionReturnValue = true;

                                                goto isAddWithError;
                                            }
                                        }
                                    }

                                    if (oDocType == "dDocument_Items")
                                        oInvoice.DocType = BoDocumentTypes.dDocument_Items;
                                    else
                                        oInvoice.DocType = BoDocumentTypes.dDocument_Service;

                                    lErrCode = oInvoice.Add();
                                    if (lErrCode == 0)
                                    {
                                        try
                                        {
                                            oDocEntry = sapCompany.GetNewObjectKey();
                                            lastMessage = "Successfully created Invoice (Draft) with Transaction Id:" + iRowInv.id + " in SAP B1.";
                                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Draft',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "',\"objType\" = 112 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowInv.id + "'");

                                            functionReturnValue = false;
                                        }
                                        catch
                                        { }
                                    }
                                    else
                                    {
                                        lastMessage = sapCompany.GetLastErrorDescription();
                                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowInv.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowInv.id + "'");

                                        functionReturnValue = true;

                                        goto isAddWithError;
                                    }

                                isAddWithError: ;

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvoice);

                                }
                                else
                                {
                                    oDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ODRF\" where \"U_TransId\" = '" + iRowInv.id + "' and \"CANCELED\" = 'N' and \"NumAtCard\" = '" + iRowInv.invoice_no + "' and \"ObjType\" = 13", sapCompany);
                                    lastMessage = "Invoice with Transaction Id:" + iRowInv.id + " is already existing in SAP B1 Draft.";

                                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Draft',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "',\"objType\" = 112 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowInv.id + "' and \"sapCode\" is null");

                                    functionReturnValue = true;
                                }
                            }
                            else
                            {
                                oDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"OINV\" where \"U_TransId\" = '" + iRowInv.id + "' and \"CANCELED\" = 'N'", sapCompany);
                                if (oDocEntry != "" && oDocEntry != "0")
                                {
                                    oInvoice = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oInvoices);
                                    if (oInvoice.GetByKey(Convert.ToInt16(oDocEntry)) == true)
                                    {
                                        oCreditNote = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oDrafts);
                                        oCreditNote.DocObjectCode = BoObjectTypes.oCreditNotes;

                                        oCreditNote.CardCode = oInvoice.CardCode;
                                        oCreditNote.NumAtCard = oInvoice.NumAtCard;
                                        oCreditNote.DocDate = oInvoice.DocDate;
                                        oCreditNote.Comments = "Reference of Invoice with Transaction Id:" + oInvoice.UserFields.Fields.Item("U_TransId").Value;

                                        if (oInvoice.DocType == BoDocumentTypes.dDocument_Items)
                                        {
                                            oCreditNote.DocType = BoDocumentTypes.dDocument_Items;
                                        }
                                        else
                                        {
                                            oCreditNote.DocType = BoDocumentTypes.dDocument_Service;
                                        }

                                        ////**** UDF ****////
                                        oCreditNote.UserFields.Fields.Item("U_TransId").Value = iRowInv.id.ToString();
                                        oCreditNote.UserFields.Fields.Item("U_Level").Value = oInvoice.UserFields.Fields.Item("U_Level").Value;
                                        oCreditNote.UserFields.Fields.Item("U_ProgramType").Value = oInvoice.UserFields.Fields.Item("U_ProgramType").Value;
                                        ////**** UDF ****////

                                        for (int i = 0; i < oInvoice.Lines.Count; i++)
                                        {
                                            oCreditNote.Lines.BaseEntry = Convert.ToInt16(oDocEntry);
                                            oCreditNote.Lines.BaseType = (int)SAPbobsCOM.BoObjectTypes.oInvoices;
                                            oCreditNote.Lines.BaseLine = oInvoice.Lines.LineNum;

                                            oCreditNote.Lines.Add();
                                        }

                                        lErrCode = oCreditNote.Add();
                                        if (lErrCode == 0)
                                        {
                                            try
                                            {
                                                lastMessage = "Successfully created Credit Note (Draft) to void Invoice with Transaction Id: " + iRowInv.id + " in SAP B1. Subject for manual posting the Draft to cancel the Invoice.";
                                                sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Posted',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "',\"objType\" = 112 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowInv.id + "'");

                                                functionReturnValue = false;
                                            }
                                            catch
                                            { }
                                        }
                                        else
                                        {
                                            lastMessage = sapCompany.GetLastErrorDescription();
                                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowInv.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowInv.id + "'");

                                            functionReturnValue = true;
                                        }
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCreditNote);
                                    }
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvoice);
                                }
                                else
                                {
                                    Int16 iDocEntry = CreateInvoiceVoid(listInvoice);
                                    if (iDocEntry != 0)
                                    {
                                        functionReturnValue = false;
                                    }
                                    else
                                        functionReturnValue = true;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            lastMessage = ex.ToString();
                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowInv.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowInv.id + "' and \"sapCode\" is not null");
                        }
                    }
                    Console.WriteLine("Done adding the List of " + string.Format("{0:#,##0}", RecordCount) + " Invoice(s) in the integration log. Please wait...");
                }
            }
            catch (Exception ex)
            {
                lastMessage = ex.ToString();
                sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(oStatus == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + oId + "' and \"sapCode\" is not null");
            }

            return functionReturnValue;
        }

        public bool SBOPostCreditNote(List<SBOHelper.Models.API_CreditNote> listCreditNote, string APILastTimeStamp)
        {
            bool functionReturnValue = false;
            int lErrCode = 0;
            int RecordCount = 0;
            int oId = 0;
            int oStatus = 0;
            string oLogExist = string.Empty;
            string oTransId = string.Empty;
            string oCardCode = string.Empty;
            string oCardName = string.Empty;
            string oDocEntry = string.Empty;
            string oDescription = string.Empty;
            string oDocType = string.Empty;
            SBOGetRecord clsSBOGetRecord = new SBOGetRecord();
            try
            {
                if (SBOconnectToLoginCompany(SBOConstantClass.SBOServer, SBOConstantClass.Database, SBOConstantClass.ServerUN, SBOConstantClass.ServerPW, SBOConstantClass.SAPUser, SBOConstantClass.SAPPassword))
                {
                    GetIntegrationSetup();
                    foreach (var iRowCreditNote in listCreditNote)
                    {
                        try
                        {
                            oId = iRowCreditNote.id;
                            oStatus = iRowCreditNote.status;

                            CreditNoteModelHeader = new List<SBOHelper.Models.API_CreditNote>();
                            CreditNoteModelDetails = new List<SBOHelper.Models.API_CreditNoteDetails>();

                            ////**** Create a list of Credit Note ****////
                            foreach (var iRowCreditNoteDtl in iRowCreditNote.items.ToList())
                            {
                                CreditNoteModelDetails.Add(new SBOHelper.Models.API_CreditNoteDetails()
                                {
                                    description = iRowCreditNoteDtl.description,
                                    date_for = iRowCreditNoteDtl.date_for,
                                    amount = iRowCreditNoteDtl.amount,
                                    gst = iRowCreditNoteDtl.gst
                                });
                            }

                            CreditNoteModelHeader.Add(new SBOHelper.Models.API_CreditNote()
                            {
                                id = iRowCreditNote.id,
                                credit_no = iRowCreditNote.credit_no,
                                credit_type = iRowCreditNote.credit_type,
                                overpaid_receipt = iRowCreditNote.overpaid_receipt,
                                student = iRowCreditNote.student,
                                date_created = iRowCreditNote.date_created,
                                status = iRowCreditNote.status,
                                remarks = iRowCreditNote.remarks,
                                void_remarks = iRowCreditNote.void_remarks,
                                type = iRowCreditNote.type,
                                level = iRowCreditNote.level,
                                program_type = iRowCreditNote.program_type,
                                payment_method = iRowCreditNote.payment_method,
                                items = CreditNoteModelDetails.ToList()
                            });

                            string strJSON = JsonConvert.SerializeObject(CreditNoteModelHeader);

                            oLogExist = (String)clsSBOGetRecord.GetSingleValue("select * from " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " where \"companyDB\" = '" + sapCompany.CompanyDB + "' and \"module\" = 'Credit Note' and \"uniqueId\" = '" + iRowCreditNote.id + "' ", sapCompany);

                            if (oLogExist == "" || oLogExist == "0")
                            {
                                Console.WriteLine("Adding Credit Note with Transaction Id:" + iRowCreditNote.id + " in the integration log. Please wait...");
                                strQuery = "insert into " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " (\"lastTimeStamp\",\"companyDB\",\"module\",\"uniqueId\",\"docStatus\",\"status\",\"JSON\",\"statusCode\",\"successDesc\",\"failDesc\",\"logDate\",\"reference\") select '" + iif(APILastTimeStamp != "", APILastTimeStamp, sapCompany.GetDBServerDate().ToString("yyyy-MM-dd")) + "','" + TrimData(sapCompany.CompanyDB) + "','Credit Note','" + iRowCreditNote.id + "','" + iif(iRowCreditNote.status == 1, "Confirmed", "Void") + "','Draft','" + TrimData(strJSON) + "','For Process','','',null,'" + iRowCreditNote.credit_no + "' " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "", "from dummy;") + "";
                                sapRecSet.DoQuery(strQuery);
                                RecordCount += 1;
                            }
                            else
                            {
                                if (iRowCreditNote.status == 2)
                                {
                                    Console.WriteLine("Updating Credit Note with Transaction Id:" + iRowCreditNote.id + " in the integration log. Please wait...");
                                    strQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"docStatus\" = '" + iif(iRowCreditNote.status == 1, "Confirmed", "Void") + "',\"statusCode\" = 'For Process',\"JSON\" = '" + TrimData(strJSON) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Credit Note' and \"uniqueId\" = '" + iRowCreditNote.id + "' and \"docStatus\" = 'Confirmed'";
                                    sapRecSet.DoQuery(strQuery);
                                }
                            }

                            ////**** Create a list of Credit Note ****////

                            if (iRowCreditNote.status == 1)
                            {
                                Console.WriteLine("Processing Credit Note with Transaction Id:" + iRowCreditNote.id + " in SAP B1 Draft and A/R Credit Memo. Please wait...");
                               
                                string Query = "select \"U_TransId\" from \"ODRF\" where \"U_TransId\" = '" + iRowCreditNote.id + "' and \"CANCELED\" = 'N' and \"NumAtCard\" = '" + iRowCreditNote.credit_no + "' and \"ObjType\" = 14 and \"U_CreatedByVoucher\" = 0 " + Environment.NewLine +
                               "union all " + Environment.NewLine +
                               "select \"U_TransId\" from \"ORIN\" where \"U_TransId\" = '" + iRowCreditNote.id + "' and \"CANCELED\" = 'N' and \"NumAtCard\" = '" + iRowCreditNote.credit_no + "'";
                                oTransId = (String)clsSBOGetRecord.GetSingleValue(Query,sapCompany);
                                if (oTransId == "" || oTransId == "0")
                                {
                                    oCreditNote = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oDrafts);
                                    oCreditNote.DocObjectCode = BoObjectTypes.oCreditNotes;

                                    oCardCode = (String)clsSBOGetRecord.GetSingleValue("select \"CardCode\" from \"OCRD\" where \"CardCode\" = '" + iRowCreditNote.student + "'", sapCompany);
                                    if (oCardCode != "")
                                    {
                                        oCreditNote.CardCode = oCardCode;
                                    }
                                    else
                                    {
                                        lastMessage = "Customer Code:" + iRowCreditNote.student + " is not found in SAP B1";
                                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditNote.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + lastMessage + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + SBOConstantClass.Database + "' and \"module\" = 'Credit Note' and \"uniqueId\" = '" + iRowCreditNote.id + "'");

                                        functionReturnValue = true;

                                        goto isAddWithError;
                                    }

                                    oCreditNote.DocDate = Convert.ToDateTime(iRowCreditNote.date_created);
                                    oCreditNote.NumAtCard = iRowCreditNote.credit_no;

                                    if (iRowCreditNote.status == 1)
                                        oCreditNote.Comments = iRowCreditNote.remarks;
                                    else
                                        oCreditNote.Comments = iRowCreditNote.void_remarks;

                                    ////**** UDF ****////
                                    if (iRowCreditNote.id != 0)
                                        oCreditNote.UserFields.Fields.Item("U_TransId").Value = iRowCreditNote.id.ToString();

                                    if (iRowCreditNote.credit_type != -1)
                                        oCreditNote.UserFields.Fields.Item("U_CreditType").Value = iRowCreditNote.credit_type;

                                    if (iRowCreditNote.type != -1)
                                        oCreditNote.UserFields.Fields.Item("U_Type").Value = iRowCreditNote.type;

                                    if (iRowCreditNote.program_type != "")
                                        oCreditNote.UserFields.Fields.Item("U_ProgramType").Value = iRowCreditNote.program_type;

                                    if (iRowCreditNote.level != "")
                                        oCreditNote.UserFields.Fields.Item("U_Level").Value = iRowCreditNote.level;
                                    ////**** UDF ****/////

                                    foreach (var iRowCreditNoteDtls in iRowCreditNote.items.ToList())
                                    {
                                        if (iRowCreditNoteDtls.description != "")
                                        {
                                            oDocType = "dDocument_Service";

                                            string iReplaceDesc = " (" + TrimData(iRowCreditNote.level) + " - " + TrimData(iRowCreditNote.program_type) + ")";
                                            //oDescription = SBOstrManipulation.BeforeCharacter(iRowCreditNoteDtls.description, " (");
                                            oDescription = iRowCreditNoteDtls.description.Replace(iReplaceDesc, "");

                                            if (oDescription != "")
                                            {
                                                string description = oDescription;
                                                string iDescription = (String)clsSBOGetRecord.GetSingleValue("select \"U_Description\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + description + "' and \"U_Level\" = '" + iRowCreditNote.level + "' and \"U_ProgramType\" = '" + iRowCreditNote.program_type + "'", sapCompany);
                                                if (iDescription != "")
                                                {
                                                    string idate_created = string.Empty;
                                                    string idate_for = string.Empty;
                                                    string iGLAccount = string.Empty;
                                                    string oDateFor = string.Empty;

                                                    if (string.IsNullOrEmpty(iRowCreditNoteDtls.date_for))
                                                    {
                                                        idate_for = iRowCreditNote.date_created;
                                                        oDateFor = Convert.ToDateTime(idate_for).ToString("MMM") + " " + Convert.ToDateTime(idate_for).Year.ToString();
                                                    }
                                                    else
                                                    {
                                                        idate_for = iRowCreditNoteDtls.date_for;
                                                        oDateFor = Convert.ToDateTime(idate_for).ToString("MMM") + " " + Convert.ToDateTime(idate_for).Year.ToString();
                                                    }

                                                    oCardName = (String)clsSBOGetRecord.GetSingleValue("select \"CardName\" from \"OCRD\" where \"CardCode\" = '" + TrimData(iRowCreditNote.student) + "'", sapCompany);

                                                    string oTaxCode = (String)clsSBOGetRecord.GetSingleValue("select \"U_TaxCode\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(iRowCreditNote.level) + "' and \"U_ProgramType\" = '" + TrimData(iRowCreditNote.program_type) + "'", sapCompany);

                                                    if (!string.IsNullOrEmpty(oTaxCode))
                                                        oCreditNote.Lines.VatGroup = oTaxCode;

                                                    oItemDescription = oCardName + " - " + oDateFor + " - " + iRowCreditNoteDtls.description;
                                                    oCreditNote.Lines.UserFields.Fields.Item("U_Dscription").Value = oItemDescription;

                                                    string Dscription = string.Empty;
                                                    if (oItemDescription.Length > 100)
                                                    {
                                                        Dscription = oItemDescription.Substring(0, 100);
                                                        oCreditNote.Lines.ItemDescription = Dscription;
                                                    }
                                                    else
                                                    {
                                                        oCreditNote.Lines.ItemDescription = oItemDescription;
                                                    }

                                                    oCreditNote.Lines.LineTotal = iRowCreditNoteDtls.amount;

                                                    if (!string.IsNullOrEmpty(iRowCreditNote.date_created))
                                                        idate_created = iRowCreditNote.date_created;

                                                    if (string.IsNullOrEmpty(oCountry) || string.IsNullOrEmpty(oGroup) || string.IsNullOrEmpty(oDivision) || string.IsNullOrEmpty(oProduct))
                                                    {
                                                        lastMessage = "Cost Center is not defined in SAP B1. Please define in the integration setup.";
                                                        string oQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditNote.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Note' and \"uniqueId\" = '" + iRowCreditNote.id + "'";
                                                        sapRecSet.DoQuery(oQuery);

                                                        functionReturnValue = true;

                                                        goto isAddWithError;
                                                    }

                                                    if (!string.IsNullOrEmpty(oCountry))
                                                        oCreditNote.Lines.CostingCode = oCountry;

                                                    if (!string.IsNullOrEmpty(oGroup))
                                                        oCreditNote.Lines.CostingCode2 = oGroup;

                                                    if (!string.IsNullOrEmpty(oDivision))
                                                        oCreditNote.Lines.CostingCode3 = oDivision;

                                                    if (!string.IsNullOrEmpty(oProduct))
                                                        oCreditNote.Lines.CostingCode4 = oProduct;

                                                    if (CheckDate(idate_created) == true && CheckDate(idate_for) == true)
                                                    {
                                                        if (Convert.ToDateTime(idate_for) > Convert.ToDateTime(idate_created))
                                                        {
                                                            iGLAccount = (String)clsSBOGetRecord.GetSingleValue("select \"U_FuturePeriod\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + description + "' and \"U_Level\" = '" + iRowCreditNote.level + "' and \"U_ProgramType\" = '" + iRowCreditNote.program_type + "'", sapCompany);
                                                        }
                                                        else
                                                        {
                                                            iGLAccount = (String)clsSBOGetRecord.GetSingleValue("select \"U_CurrentPeriod\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + description + "' and \"U_Level\" = '" + iRowCreditNote.level + "' and \"U_ProgramType\" = '" + iRowCreditNote.program_type + "'", sapCompany);
                                                        }
                                                    }

                                                    if (!string.IsNullOrEmpty(iGLAccount))
                                                        oCreditNote.Lines.AccountCode = iGLAccount;

                                                    oCreditNote.Lines.Add();
                                                }
                                                else
                                                {
                                                    lastMessage = "Description:" + iRowCreditNoteDtls.description + ", Level: " + iRowCreditNote.level + " or Program type:" + iRowCreditNote.program_type + " is not defined in SAP B1. Please define in the table.";
                                                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditNote.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + lastMessage + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + SBOConstantClass.Database + "' and \"module\" = 'Credit Note' and \"uniqueId\" = '" + iRowCreditNote.id + "'");

                                                    functionReturnValue = true;

                                                    goto isAddWithError;
                                                }
                                            }
                                        }
                                    }

                                    if (oDocType == "dDocument_Items")
                                        oCreditNote.DocType = BoDocumentTypes.dDocument_Items;
                                    else
                                        oCreditNote.DocType = BoDocumentTypes.dDocument_Service;

                                    lErrCode = oCreditNote.Add();
                                    if (lErrCode == 0)
                                    {
                                        try
                                        {
                                            oDocEntry = sapCompany.GetNewObjectKey();
                                            lastMessage = "Successfully created Credit Note (Draft) with Transaction Id:" + iRowCreditNote.id + " in SAP B1.";
                                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Draft',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "',\"objType\" = 112 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Note' and \"uniqueId\" = '" + iRowCreditNote.id + "'");

                                            functionReturnValue = false;
                                        }
                                        catch
                                        { }
                                    }
                                    else
                                    {
                                        lastMessage = sapCompany.GetLastErrorDescription();
                                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditNote.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + lastMessage + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Note' and \"uniqueId\" = '" + iRowCreditNote.id + "'");

                                        functionReturnValue = true;

                                        goto isAddWithError;
                                    }

                                isAddWithError: ;

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCreditNote);

                                }
                                else
                                {
                                    oDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ODRF\" where \"U_TransId\" = '" + iRowCreditNote.id + "' and \"CANCELED\" = 'N' and \"NumAtCard\" = '" + iRowCreditNote.credit_no + "' and \"ObjType\" = 14 and \"U_CreatedByVoucher\" = 0", sapCompany);
                                    lastMessage = "Credit Note with Transaction Id:" + iRowCreditNote.id + " is already existing in SAP B1 Draft.";

                                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Draft',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "',\"objType\" = 112 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowCreditNote.id + "' and \"sapCode\" is null");

                                    functionReturnValue = true;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            lastMessage = ex.ToString();
                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditNote.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + lastMessage + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + sapCompany.CompanyDB + "' and \"module\" = 'Credit Note' and \"uniqueId\" = '" + iRowCreditNote.id + "' and \"sapCode\" is not null");
                        }
                    }
                    Console.WriteLine("Done adding the List of " + string.Format("{0:#,##0}", RecordCount) + " Credit Note(s) in the integration log. Please wait...");
                }
            }
            catch (Exception ex)
            {
                lastMessage = ex.ToString();
                sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(oStatus == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + lastMessage + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + sapCompany.CompanyDB + "' and \"module\" = 'Credit Note' and \"uniqueId\" = '" + oId + "' and \"sapCode\" is not null");
            }

            return functionReturnValue;
        }

        public bool SBOPostReceipt(List<SBOHelper.Models.API_Receipt> listReceipt, string APILastTimeStamp)
        {
            bool functionReturnValue = false;
            int RecordCount = 0;
            string oLogExist = string.Empty;
            string oTransId = string.Empty;
            string oCardCode = string.Empty;
            string oDocEntry = string.Empty;
            string oInvDocEntry = string.Empty;
            string oCreditNoteDocEntry = string.Empty;
            string oAcctCode = string.Empty;
            string oBankName = string.Empty;
            string oCheckBankName = string.Empty;
            string iReference = string.Empty;

            SBOGetRecord clsSBOGetRecord = new SBOGetRecord();
            try
            {
                if (SBOconnectToLoginCompany(SBOConstantClass.SBOServer, SBOConstantClass.Database, SBOConstantClass.ServerUN, SBOConstantClass.ServerPW, SBOConstantClass.SAPUser, SBOConstantClass.SAPPassword))
                {
                    foreach (var iRowReceipt in listReceipt)
                    {
                        try
                        {
                            ReceiptModelHeader = new List<SBOHelper.Models.API_Receipt>();
                            ReceiptModelDetails = new List<SBOHelper.Models.API_ReceiptDetails>();

                            ////**** Create a list of Receipt ****////
                            foreach (var iRowReceiptDtl in iRowReceipt.payment_methods.ToList())
                            {
                                ReceiptModelDetails.Add(new SBOHelper.Models.API_ReceiptDetails()
                                {
                                    method = iRowReceiptDtl.method,
                                    reference = iRowReceiptDtl.reference,
                                    reference_id = iRowReceiptDtl.reference_id,
                                    amount = iRowReceiptDtl.amount,
                                });
                            }

                            ReceiptModelHeader.Add(new SBOHelper.Models.API_Receipt()
                            {
                                id = iRowReceipt.id,
                                receipt_no = iRowReceipt.receipt_no,
                                student = iRowReceipt.student,
                                level = iRowReceipt.level,
                                program_type = iRowReceipt.program_type,
                                invoice_id = iRowReceipt.invoice_id.ToList(),
                                invoice_no = iRowReceipt.invoice_no.ToList(),
                                invoice_paid = iRowReceipt.invoice_paid.ToList(),
                                payment_type = iRowReceipt.payment_type,
                                date_created = iRowReceipt.date_created,
                                status = iRowReceipt.status,
                                remarks = iRowReceipt.remarks,
                                void_remarks = iRowReceipt.void_remarks,
                                offset_references = iRowReceipt.offset_references.ToList(),
                                payment_methods = ReceiptModelDetails.ToList()
                            });

                            string strJSON = JsonConvert.SerializeObject(ReceiptModelHeader);

                            oLogExist = (String)clsSBOGetRecord.GetSingleValue("select * from " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "' ", sapCompany);

                            if (oLogExist == "" || oLogExist == "0")
                            {
                                Console.WriteLine("Adding Receipt with Transaction Id:" + iRowReceipt.id + " in the integration log. Please wait...");
                                strQuery = "insert into " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " (\"lastTimeStamp\",\"companyDB\",\"module\",\"uniqueId\",\"docStatus\",\"status\",\"JSON\",\"statusCode\",\"successDesc\",\"failDesc\",\"logDate\",\"reference\",\"objType\") select '" + iif(APILastTimeStamp != "", APILastTimeStamp, sapCompany.GetDBServerDate().ToString("yyyy-MM-dd")) + "','" + sapCompany.CompanyDB + "','Receipt','" + iRowReceipt.id + "','" + iif(iRowReceipt.status == 0, "Confirmed", "Void") + "','Draft','" + TrimData(strJSON) + "','For Process','','',null,'" + iRowReceipt.receipt_no + "', 140 " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "", "from dummy;") + "";
                                sapRecSet.DoQuery(strQuery);
                                RecordCount += 1;
                            }
                            else
                            {
                                if (iRowReceipt.status == 1)
                                {
                                    Console.WriteLine("Updating Receipt with Transaction Id:" + iRowReceipt.id + " in the integration log. Please wait...");
                                    strQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"docStatus\" = '" + iif(iRowReceipt.status == 0, "Confirmed", "Void") + "',\"statusCode\" = 'For Process',\"JSON\" = '" + TrimData(strJSON) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "' and \"docStatus\" = 'Confirmed'";
                                    sapRecSet.DoQuery(strQuery);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            lastMessage = ex.ToString();
                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'false',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");
                        }
                    }
                    Console.WriteLine("Done adding the List of " + string.Format("{0:#,##0}", RecordCount) + " Receipt(s) in the integration log. Please wait...");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return functionReturnValue;
        }

        public bool SBOPostCreditRefund(List<SBOHelper.Models.API_CreditRefund> listCreditRefund, string APILastTimeStamp)
        {
            bool functionReturnValue = false;
            int RecordCount = 0;
            string oLogExist = string.Empty;
            string oTransId = string.Empty;
            string oCardCode = string.Empty;
            string oDocEntry = string.Empty;
            string oCreditNoteDocEntry = string.Empty;
            string oAcctCode = string.Empty;
            string oBankName = string.Empty;
            string oCheckBankName = string.Empty;
            string iReference = string.Empty;

            SBOGetRecord clsSBOGetRecord = new SBOGetRecord();
            try
            {
                if (SBOconnectToLoginCompany(SBOConstantClass.SBOServer, SBOConstantClass.Database, SBOConstantClass.ServerUN, SBOConstantClass.ServerPW, SBOConstantClass.SAPUser, SBOConstantClass.SAPPassword))
                {
                    foreach (var iRowCreditRefund in listCreditRefund)
                    {
                        try
                        {
                            ////**** Create a list of Receipt ****////
                            CreditRefundModel = new List<SBOHelper.Models.API_CreditRefund>();

                            CreditRefundModel.Add(new SBOHelper.Models.API_CreditRefund()
                            {
                                id = iRowCreditRefund.id,
                                credit_id = iRowCreditRefund.credit_id,
                                credit_type = iRowCreditRefund.credit_type,
                                student = iRowCreditRefund.student,
                                status = iRowCreditRefund.status,
                                date_created = iRowCreditRefund.date_created,
                                remarks = iRowCreditRefund.remarks,
                                void_remarks = iRowCreditRefund.void_remarks,
                                payment_method = iRowCreditRefund.payment_method,
                                payment_reference = iRowCreditRefund.payment_reference,
                                amount = iRowCreditRefund.amount,
                                overpaid_offsets = iRowCreditRefund.overpaid_offsets,
                                overpaid_offsets_receipt_id = iRowCreditRefund.overpaid_offsets_receipt_id,
                                overpaid_offsets_credit_notes = iRowCreditRefund.overpaid_offsets_credit_notes.ToList(),
                                overpaid_offsets_invoices = iRowCreditRefund.overpaid_offsets_invoices.ToList()
                            });

                            string strJSON = JsonConvert.SerializeObject(CreditRefundModel);

                            oLogExist = (String)clsSBOGetRecord.GetSingleValue("select * from " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Credit Refund' and \"uniqueId\" = '" + iRowCreditRefund.id + "' ", sapCompany);

                            if (oLogExist == "" || oLogExist == "0")
                            {
                                Console.WriteLine("Adding Credit Refund with Transaction Id:" + iRowCreditRefund.id + " in the integration log. Please wait...");
                                strQuery = "insert into " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " (\"lastTimeStamp\",\"companyDB\",\"module\",\"uniqueId\",\"docStatus\",\"status\",\"JSON\",\"statusCode\",\"successDesc\",\"failDesc\",\"logDate\",\"reference\",\"objType\") select '" + iif(APILastTimeStamp != "", APILastTimeStamp, sapCompany.GetDBServerDate().ToString("yyyy-MM-dd")) + "','" + TrimData(sapCompany.CompanyDB) + "','Credit Refund','" + iRowCreditRefund.id + "','" + iif(iRowCreditRefund.status == 0, "Confirmed", "Void") + "','Draft','" + TrimData(strJSON) + "','For Process','','',null,'" + iRowCreditRefund.credit_id + "', 140 " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "", "from dummy;") + "";
                                sapRecSet.DoQuery(strQuery);
                                RecordCount += 1;
                            }
                            else
                            {
                                if (iRowCreditRefund.status == 1)
                                {
                                    Console.WriteLine("Updating Credit Refund with Transaction Id:" + iRowCreditRefund.id + " in the integration log. Please wait...");
                                    strQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"docStatus\" = '" + iif(iRowCreditRefund.status == 0, "Confirmed", "Void") + "',\"statusCode\" = 'For Process',\"JSON\" = '" + TrimData(strJSON) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Credit Refund' and \"uniqueId\" = '" + iRowCreditRefund.id + "' and \"docStatus\" = 'Confirmed'";
                                    sapRecSet.DoQuery(strQuery);
                                }
                            }
                            ////**** Create a list of Receipt ****////
                        }
                        catch (Exception ex)
                        {
                            lastMessage = ex.ToString();
                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'false',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Credit Refund' and \"uniqueId\" = '" + iRowCreditRefund.id + "'");
                        }
                    }
                    Console.WriteLine("Done adding the List of " + string.Format("{0:#,##0}", RecordCount) + " Credit Refund(s) in the integration log. Please wait...");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return functionReturnValue;
        }

        public Int16 CreateInvoiceVoid(List<Models.API_Invoice> olistInvoice)
        {
            {
                string ifunctionReturnValue = "0";
                string ilastErrorMessage = string.Empty; ;
                int olErrCode;
                string iTransId = string.Empty;
                string iCardCode = string.Empty;
                string iCardName = string.Empty;
                string iDocEntry = string.Empty;
                string iDescription = string.Empty;
                string iItemCode = string.Empty;
                string iDocType = string.Empty;
                SBOGetRecord clsSBOGetRecord = new SBOGetRecord();
                foreach (var oRowInv in olistInvoice)
                {
                    try
                    {
                        if (oRowInv.status == 2)
                        {
                            iDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ODRF\" where \"U_TransId\" = '" + oRowInv.id + "' and \"CANCELED\" = 'N' and \"ObjType\" = 13", sapCompany);
                            if (iDocEntry == "" || iDocEntry == "0")
                            {
                                oInvoice = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oDrafts);
                                oInvoice.DocObjectCode = BoObjectTypes.oInvoices;

                                iCardCode = (String)clsSBOGetRecord.GetSingleValue("select \"CardCode\" from \"OCRD\" where \"CardCode\" = '" + TrimData(oRowInv.student) + "'", sapCompany);
                                if (iCardCode != "")
                                {
                                    oInvoice.CardCode = iCardCode;
                                }
                                else
                                {
                                    lastMessage = "Customer Code:" + oRowInv.student + " is not found in SAP B1";
                                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(oRowInv.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "' where \"module\" = 'Invoice' and \"uniqueId\" = '" + oRowInv.id + "' and \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "'");

                                    ifunctionReturnValue = "0";

                                    goto isAddWithError;
                                }

                                oInvoice.DocDate = Convert.ToDateTime(oRowInv.date_created);
                                oInvoice.NumAtCard = oRowInv.invoice_no;
                                oInvoice.DocDueDate = Convert.ToDateTime(oRowInv.date_due);

                                if (oRowInv.status == 1)
                                    oInvoice.Comments = oRowInv.remarks;
                                else
                                    oInvoice.Comments = oRowInv.void_remarks;

                                ////**** UDF *****/////
                                if (oRowInv.id != 0)
                                    oInvoice.UserFields.Fields.Item("U_TransId").Value = oRowInv.id.ToString();

                                if (oRowInv.level != "")
                                    oInvoice.UserFields.Fields.Item("U_Level").Value = oRowInv.level;

                                if (oRowInv.program_type != "")
                                    oInvoice.UserFields.Fields.Item("U_ProgramType").Value = oRowInv.program_type;
                                ////**** UDF *****/////

                                foreach (var oRowInvDtls in oRowInv.items.ToList())
                                {
                                    if (oRowInvDtls.item_code == "" || string.IsNullOrEmpty(oRowInvDtls.item_code))
                                    {
                                        iDocType = "dDocument_Service";
                                        string iReplaceDesc = " (" + TrimData(oRowInv.level) + " - " + TrimData(oRowInv.program_type) + ")";
                                        //oDescription = SBOstrManipulation.BeforeCharacter(oRowInvDtls.description, " (");
                                        iDescription = oRowInvDtls.description.Replace(iReplaceDesc, "");

                                        if (iDescription != "")
                                        {
                                            string description = iDescription;
                                            string oDescription = (String)clsSBOGetRecord.GetSingleValue("select \"U_Description\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(oRowInv.level) + "' and \"U_ProgramType\" = '" + TrimData(oRowInv.program_type) + "'", sapCompany);
                                            if (oDescription != "")
                                            {
                                                string idate_created = string.Empty;
                                                string idate_for = string.Empty;
                                                string iGLAccount = string.Empty;
                                                string oDateFor = string.Empty;

                                                if (!string.IsNullOrEmpty(oRowInvDtls.date_for))
                                                {
                                                    oDateFor = Convert.ToDateTime(oRowInvDtls.date_for).ToString("MMM") + " " + Convert.ToDateTime(oRowInvDtls.date_for).Year.ToString();
                                                    idate_for = oRowInvDtls.date_for;
                                                }
                                                else
                                                {
                                                    idate_for = oRowInv.date_created;
                                                    oDateFor = Convert.ToDateTime(idate_for).ToString("MMM") + " " + Convert.ToDateTime(idate_for).Year.ToString();
                                                }

                                                iCardName = (String)clsSBOGetRecord.GetSingleValue("select \"CardName\" from \"OCRD\" where \"CardCode\" = '" + TrimData(oRowInv.student) + "'", sapCompany);

                                                string oTaxCode = (String)clsSBOGetRecord.GetSingleValue("select \"U_TaxCode\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(oRowInv.level) + "' and \"U_ProgramType\" = '" + TrimData(oRowInv.program_type) + "'", sapCompany);

                                                if (oTaxCode != "")
                                                    oInvoice.Lines.VatGroup = oTaxCode;

                                                oItemDescription = iCardName + " - " + oDateFor + " - " + oRowInvDtls.description;
                                                oInvoice.Lines.UserFields.Fields.Item("U_Dscription").Value = oItemDescription;

                                                string Dscription = string.Empty;
                                                if (oItemDescription.Length > 100)
                                                {
                                                    Dscription = oItemDescription.Substring(0, 100);
                                                    oInvoice.Lines.ItemDescription = Dscription;
                                                }
                                                else
                                                {
                                                    oInvoice.Lines.ItemDescription = oItemDescription;
                                                }

                                                oInvoice.Lines.LineTotal = oRowInvDtls.unit_price;

                                                if (!string.IsNullOrEmpty(oRowInv.date_created))
                                                    idate_created = oRowInv.date_created;

                                                if (!string.IsNullOrEmpty(oRowInvDtls.date_for))
                                                    idate_for = oRowInvDtls.date_for;

                                                if (!string.IsNullOrEmpty(oCountry))
                                                    oInvoice.Lines.CostingCode = oCountry;

                                                if (!string.IsNullOrEmpty(oGroup))
                                                    oInvoice.Lines.CostingCode2 = oGroup;

                                                if (!string.IsNullOrEmpty(oDivision))
                                                    oInvoice.Lines.CostingCode3 = oDivision;

                                                if (!string.IsNullOrEmpty(oProduct))
                                                    oInvoice.Lines.CostingCode4 = oProduct;

                                                if (idate_for != "")
                                                    oInvoice.Lines.UserFields.Fields.Item("U_date_for").Value = Convert.ToDateTime(idate_for);

                                                if (CheckDate(idate_created) == true && CheckDate(idate_for) == true)
                                                {
                                                    if (Convert.ToDateTime(idate_for) > Convert.ToDateTime(idate_created))
                                                    {
                                                        iGLAccount = (String)clsSBOGetRecord.GetSingleValue("select \"U_FuturePeriod\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(oRowInv.level) + "' and \"U_ProgramType\" = '" + TrimData(oRowInv.program_type) + "'", sapCompany);
                                                    }
                                                    else
                                                    {
                                                        iGLAccount = (String)clsSBOGetRecord.GetSingleValue("select \"U_CurrentPeriod\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(oRowInv.level) + "' and \"U_ProgramType\" = '" + TrimData(oRowInv.program_type) + "'", sapCompany);
                                                    }
                                                }

                                                if (!string.IsNullOrEmpty(iGLAccount))
                                                    oInvoice.Lines.AccountCode = iGLAccount;

                                                oInvoice.Lines.Add();
                                            }
                                            else
                                            {
                                                lastMessage = "Description:" + oRowInvDtls.description + ", Level: " + oRowInv.level + " or Program type:" + oRowInv.program_type + " is not defined in SAP B1. Please define in the table.";
                                                string oQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(oRowInv.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + oRowInv.id + "'";
                                                sapRecSet.DoQuery(oQuery);

                                                ifunctionReturnValue = "0";

                                                goto isAddWithError;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        iDocType = "dDocument_Items";
                                        iItemCode = string.Empty;
                                        iItemCode = (String)clsSBOGetRecord.GetSingleValue("select \"ItemCode\" from \"OITM\" where \"ItemCode\" = '" + TrimData(oRowInvDtls.item_code) + "'", sapCompany);
                                        if (iItemCode != "")
                                        {
                                            oInvoice.Lines.ItemCode = oRowInvDtls.item_code;

                                            if (oRowInvDtls.quantity != 0)
                                                oInvoice.Lines.Quantity = oRowInvDtls.quantity;

                                            if (oRowInvDtls.unit_price != 0)
                                                oInvoice.Lines.UnitPrice = oRowInvDtls.unit_price;

                                            if (!string.IsNullOrEmpty(oCountry))
                                                oInvoice.Lines.CostingCode = oCountry;

                                            if (!string.IsNullOrEmpty(oGroup))
                                                oInvoice.Lines.CostingCode2 = oGroup;

                                            if (!string.IsNullOrEmpty(oDivision))
                                                oInvoice.Lines.CostingCode3 = oDivision;

                                            if (!string.IsNullOrEmpty(oProduct))
                                                oInvoice.Lines.CostingCode4 = oProduct;

                                            oInvoice.Lines.Add();
                                        }
                                        else
                                        {
                                            lastMessage = "ItemCode: " + oRowInvDtls.item_code + " does not exist in SAP B1.";
                                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(oRowInv.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + oRowInv.id + "'");

                                            ifunctionReturnValue = "0";

                                            goto isAddWithError;
                                        }
                                    }
                                }

                                if (iDocType == "dDocument_Items")
                                    oInvoice.DocType = BoDocumentTypes.dDocument_Items;
                                else
                                    oInvoice.DocType = BoDocumentTypes.dDocument_Service;

                                olErrCode = oInvoice.Add();
                                if (olErrCode == 0)
                                {
                                    try
                                    {
                                        iDocEntry = sapCompany.GetNewObjectKey();
                                        lastMessage = "Successfully created Invoice (Draft) with Transaction Id: " + oRowInv.id + " in SAP B1. Subject for manual posting and cancellation.";
                                        string oQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Draft',\"failDesc\" = '',\"successDesc\" = '" + lastMessage + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"objType\" = 112 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + oRowInv.id + "'";
                                        sapRecSet.DoQuery(oQuery);

                                        ifunctionReturnValue = iDocEntry;
                                    }
                                    catch
                                    { }
                                }
                                else
                                {
                                    lastMessage = sapCompany.GetLastErrorDescription();
                                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(oRowInv.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + oRowInv.id + "'");

                                    ifunctionReturnValue = "0";

                                    goto isAddWithError;
                                }

                            isAddWithError: ;

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvoice);

                            }
                            else
                                ifunctionReturnValue = iDocEntry;
                        }
                    }
                    catch (Exception ex)
                    {
                        lastMessage = ex.ToString();
                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(oRowInv.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + oRowInv.id + "'");

                        ifunctionReturnValue = "0";
                    }
                }
                return Convert.ToInt16(ifunctionReturnValue);
            }
        }

        public Int16 CreateCreditNoteVoid(List<Models.API_CreditNote> listCreditNote)
        {
            {
                string ifunctionReturnValue = "0";
                string ilastErrorMessage = string.Empty; ;
                int olErrCode;
                string iTransId = string.Empty;
                string iCardCode = string.Empty;
                string iCardName = string.Empty;
                string iDocEntry = string.Empty;
                string iDescription = string.Empty;
                string iItemCode = string.Empty;
                string iDocType = string.Empty;
                SBOGetRecord clsSBOGetRecord = new SBOGetRecord();
                foreach (var oRowCreditNote in listCreditNote)
                {
                    try
                    {
                        if (oRowCreditNote.status == 2)
                        {
                            iDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ORIN\" where \"U_TransId\" = '" + oRowCreditNote.id + "' and \"CANCELED\" = 'N' and \"NumAtCard\" = '" + oRowCreditNote.credit_no + "'", sapCompany);
                            if (iDocEntry == "" || iDocEntry == "0")
                            {
                                oCreditNote = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oCreditNotes);

                                iCardCode = (String)clsSBOGetRecord.GetSingleValue("select \"CardCode\" from \"OCRD\" where \"CardCode\" = '" + TrimData(oRowCreditNote.student) + "'", sapCompany);
                                if (iCardCode != "")
                                {
                                    oCreditNote.CardCode = iCardCode;
                                }
                                else
                                {
                                    lastMessage = "Customer Code:" + oRowCreditNote.student + " is not found in SAP B1";
                                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(oRowCreditNote.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "' where \"module\" = 'Credit Note' and \"uniqueId\" = '" + oRowCreditNote.id + "' and \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "'");

                                    ifunctionReturnValue = "0";

                                    goto isAddWithError;
                                }

                                oCreditNote.DocDate = Convert.ToDateTime(oRowCreditNote.date_created);
                                oCreditNote.NumAtCard = oRowCreditNote.credit_no;

                                if (oRowCreditNote.status == 1)
                                    oCreditNote.Comments = oRowCreditNote.remarks;
                                else
                                    oCreditNote.Comments = oRowCreditNote.void_remarks;

                                ////**** UDF *****/////
                                if (oRowCreditNote.id != 0)
                                    oCreditNote.UserFields.Fields.Item("U_TransId").Value = oRowCreditNote.id.ToString();

                                if (oRowCreditNote.level != "")
                                    oCreditNote.UserFields.Fields.Item("U_Level").Value = oRowCreditNote.level;

                                if (oRowCreditNote.program_type != "")
                                    oCreditNote.UserFields.Fields.Item("U_ProgramType").Value = oRowCreditNote.program_type;
                                ////**** UDF *****/////

                                foreach (var oRowCreditNoteDtls in oRowCreditNote.items.ToList())
                                {
                                    if (oRowCreditNoteDtls.description != "")
                                    {
                                        iDocType = "dDocument_Service";
                                        string iReplaceDesc = " (" + TrimData(oRowCreditNote.level) + " - " + TrimData(oRowCreditNote.program_type) + ")";
                                        //oDescription = SBOstrManipulation.BeforeCharacter(oRowCreditNoteDtls.description, " (");
                                        iDescription = oRowCreditNoteDtls.description.Replace(iReplaceDesc, "");

                                        if (iDescription != "")
                                        {
                                            string description = iDescription;
                                            string oDescription = (String)clsSBOGetRecord.GetSingleValue("select \"U_Description\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(oRowCreditNote.level) + "' and \"U_ProgramType\" = '" + TrimData(oRowCreditNote.program_type) + "'", sapCompany);
                                            if (oDescription != "")
                                            {
                                                string idate_created = string.Empty;
                                                string idate_for = string.Empty;
                                                string iGLAccount = string.Empty;
                                                string oDateFor = string.Empty;

                                                if (!string.IsNullOrEmpty(oRowCreditNoteDtls.date_for))
                                                {
                                                    oDateFor = Convert.ToDateTime(oRowCreditNoteDtls.date_for).ToString("MMM") + " " + Convert.ToDateTime(oRowCreditNoteDtls.date_for).Year.ToString();
                                                    idate_for = oRowCreditNoteDtls.date_for;
                                                }
                                                else
                                                {
                                                    idate_for = oRowCreditNote.date_created;
                                                    oDateFor = Convert.ToDateTime(idate_for).ToString("MMM") + " " + Convert.ToDateTime(idate_for).Year.ToString();
                                                }

                                                iCardName = (String)clsSBOGetRecord.GetSingleValue("select \"CardName\" from \"OCRD\" where \"CardCode\" = '" + TrimData(oRowCreditNote.student) + "'", sapCompany);

                                                string oTaxCode = (String)clsSBOGetRecord.GetSingleValue("select \"U_TaxCode\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(oRowCreditNote.level) + "' and \"U_ProgramType\" = '" + TrimData(oRowCreditNote.program_type) + "'", sapCompany);

                                                if (oTaxCode != "")
                                                    oCreditNote.Lines.VatGroup = oTaxCode;

                                                oItemDescription = iCardName + " - " + oDateFor + " - " + oRowCreditNoteDtls.description;
                                                oCreditNote.Lines.UserFields.Fields.Item("U_Dscription").Value = oItemDescription;

                                                string Dscription = string.Empty;
                                                if (oItemDescription.Length > 100)
                                                {
                                                    Dscription = oItemDescription.Substring(0, 100);
                                                    oCreditNote.Lines.ItemDescription = Dscription;
                                                }
                                                else
                                                {
                                                    oCreditNote.Lines.ItemDescription = oItemDescription;
                                                }

                                                oCreditNote.Lines.LineTotal = oRowCreditNoteDtls.amount;

                                                if (!string.IsNullOrEmpty(oRowCreditNote.date_created))
                                                    idate_created = oRowCreditNote.date_created;

                                                if (!string.IsNullOrEmpty(oRowCreditNoteDtls.date_for))
                                                    idate_for = oRowCreditNoteDtls.date_for;

                                                if (!string.IsNullOrEmpty(oCountry))
                                                    oCreditNote.Lines.CostingCode = oCountry;

                                                if (!string.IsNullOrEmpty(oGroup))
                                                    oCreditNote.Lines.CostingCode2 = oGroup;

                                                if (!string.IsNullOrEmpty(oDivision))
                                                    oCreditNote.Lines.CostingCode3 = oDivision;

                                                if (!string.IsNullOrEmpty(oProduct))
                                                    oCreditNote.Lines.CostingCode4 = oProduct;

                                                if (idate_for != "")
                                                    oCreditNote.Lines.UserFields.Fields.Item("U_date_for").Value = Convert.ToDateTime(idate_for);

                                                if (CheckDate(idate_created) == true && CheckDate(idate_for) == true)
                                                {
                                                    if (Convert.ToDateTime(idate_for) > Convert.ToDateTime(idate_created))
                                                    {
                                                        iGLAccount = (String)clsSBOGetRecord.GetSingleValue("select \"U_FuturePeriod\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(oRowCreditNote.level) + "' and \"U_ProgramType\" = '" + TrimData(oRowCreditNote.program_type) + "'", sapCompany);
                                                    }
                                                    else
                                                    {
                                                        iGLAccount = (String)clsSBOGetRecord.GetSingleValue("select \"U_CurrentPeriod\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(oRowCreditNote.level) + "' and \"U_ProgramType\" = '" + TrimData(oRowCreditNote.program_type) + "'", sapCompany);
                                                    }
                                                }

                                                if (!string.IsNullOrEmpty(iGLAccount))
                                                    oCreditNote.Lines.AccountCode = iGLAccount;

                                                oCreditNote.Lines.Add();
                                            }
                                            else
                                            {
                                                lastMessage = "Description:" + oRowCreditNoteDtls.description + ", Level: " + oRowCreditNote.level + " or Program type:" + oRowCreditNote.program_type + " is not defined in SAP B1. Please define in the table.";
                                                string oQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(oRowCreditNote.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Note' and \"uniqueId\" = '" + oRowCreditNote.id + "'";
                                                sapRecSet.DoQuery(oQuery);

                                                ifunctionReturnValue = "0";

                                                goto isAddWithError;
                                            }
                                        }
                                    }
                                }

                                if (iDocType == "dDocument_Items")
                                    oCreditNote.DocType = BoDocumentTypes.dDocument_Items;
                                else
                                    oCreditNote.DocType = BoDocumentTypes.dDocument_Service;

                                olErrCode = oCreditNote.Add();
                                if (olErrCode == 0)
                                {
                                    try
                                    {
                                        iDocEntry = sapCompany.GetNewObjectKey();
                                        ifunctionReturnValue = iDocEntry;
                                    }
                                    catch
                                    { }
                                }
                                else
                                {
                                    lastMessage = sapCompany.GetLastErrorDescription();
                                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(oRowCreditNote.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Credit Note' and \"uniqueId\" = '" + oRowCreditNote.id + "'");

                                    ifunctionReturnValue = "0";

                                    goto isAddWithError;
                                }

                            isAddWithError: ;

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCreditNote);

                            }
                            else
                                ifunctionReturnValue = iDocEntry;
                        }
                    }
                    catch (Exception ex)
                    {
                        lastMessage = ex.ToString();
                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(oRowCreditNote.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Note' and \"uniqueId\" = '" + oRowCreditNote.id + "'");

                        ifunctionReturnValue = "0";
                    }
                }
                return Convert.ToInt16(ifunctionReturnValue);
            }
        }

        public bool ItemMasterData(string oDate = "")
        {
            try
            {
                //Declarations
                bool functionReturnValue = false;
                string oLogExist = string.Empty;
                string lastMessage = string.Empty;
                SBOGetRecord clsSBOGetRecord = new SBOGetRecord();

                if (SBOconnectToLoginCompany(SBOConstantClass.SBOServer, SBOConstantClass.Database, SBOConstantClass.ServerUN, SBOConstantClass.ServerPW, SBOConstantClass.SAPUser, SBOConstantClass.SAPPassword))
                {
                    GetIntegrationSetup();

                    int oItmsGrpCod = Convert.ToInt16(clsSBOGetRecord.GetSingleValue("select \"ItmsGrpCod\" from \"OITB\" where \"ItmsGrpNam\" like '%MERCHANDISE%'", sapCompany));

                    ////** Declarations **//////
                    ItemModel = new List<SBOHelper.Models.API_FinanceItem>();

                    string oQuery = "select " + Environment.NewLine +
                    "\"a\".\"ItemCode\" \"item_code\", " + Environment.NewLine +
                    "\"a\".\"ItemName\" \"description\", " + Environment.NewLine +
                    "\"b\".\"ItmsGrpNam\" \"type\", " + Environment.NewLine +
                    "\"c\".\"Price\" \"unit_price\", " + Environment.NewLine +
                    "\"a\".\"FrgnName\" \"remarks\", " + Environment.NewLine +
                    "1 \"tax\" " + Environment.NewLine +
                    "from \"OITM\" \"a\" " + Environment.NewLine +
                    "left join \"OITB\" \"b\" on \"b\".\"ItmsGrpCod\" = \"a\".\"ItmsGrpCod\" " + Environment.NewLine +
                    "left join \"ITM1\" \"c\" on \"c\".\"ItemCode\" = \"a\".\"ItemCode\" " + Environment.NewLine +
                    "where \"c\".\"PriceList\" = " + pricelistcode + " and \"a\".\"CreateDate\" = '" + oDate + "' and \"a\".\"ItmsGrpCod\" = " + oItmsGrpCod + " " + Environment.NewLine +
                    "or \"c\".\"PriceList\" = " + pricelistcode + " and \"a\".\"UpdateDate\" = '" + oDate + "' and \"a\".\"ItmsGrpCod\" = " + oItmsGrpCod + "";
                    sapRecSet.DoQuery(oQuery);

                    ItemMasterModel = new List<SBOHelper.Models.API_FinanceItem>();

                    if (sapRecSet.RecordCount > 0)
                    {
                        ////** Parse Business Partners **//////
                        XDocument xItemMasterData = XDocument.Parse(sapRecSet.GetAsXML());
                        if (xItemMasterData.Root != null)
                        {
                            ItemModel = xItemMasterData.Descendants("row").Select(oItemMaster =>
                            new SBOHelper.Models.API_FinanceItem
                            {
                                item_code = oItemMaster.Element("item_code").Value,
                                description = oItemMaster.Element("description").Value,
                                type = oItemMaster.Element("type").Value,
                                unit_price = Convert.ToDouble(oItemMaster.Element("unit_price").Value),
                                remarks = oItemMaster.Element("remarks").Value,
                                tax = Convert.ToInt16(oItemMaster.Element("tax").Value)
                            }).ToList();
                        }

                        ////**** Create a list of Products ****////
                        foreach (var iRowItems in ItemModel)
                        {
                            ItemMasterModel = new List<SBOHelper.Models.API_FinanceItem>();
                            ItemMasterModel.Add(new SBOHelper.Models.API_FinanceItem()
                            {
                                item_code = iRowItems.item_code,
                                description = iRowItems.description,
                                type = iRowItems.type,
                                unit_price = Convert.ToDouble(iRowItems.unit_price),
                                remarks = iRowItems.remarks,
                                tax = iRowItems.tax
                            });

                            string strJSON = JsonConvert.SerializeObject(ItemMasterModel);
                            oLogExist = (String)clsSBOGetRecord.GetSingleValue("select * from " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Product' and \"uniqueId\" = '" + TrimData(iRowItems.item_code) + "' ", sapCompany);

                            if (oLogExist == "" || oLogExist == "0")
                            {
                                Console.WriteLine("Adding Product:" + iRowItems.item_code + " in the integration log. Please wait...");
                                strQuery = "insert into " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " (\"lastTimeStamp\",\"companyDB\",\"module\",\"uniqueId\",\"docStatus\",\"status\",\"JSON\",\"statusCode\",\"successDesc\",\"failDesc\",\"logDate\",\"objType\") select '" + oDate + "','" + TrimData(sapCompany.CompanyDB) + "','Product','" + TrimData(iRowItems.item_code) + "','Confirmed','','" + TrimData(strJSON) + "','','','',null,4" + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "", " from dummy;") + "";
                                sapRecSet.DoQuery(strQuery);
                            }
                            else
                            {
                                Console.WriteLine("Updating Product:" + iRowItems.item_code + " in the integration log. Please wait...");
                                strQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"JSON\" = '" + TrimData(strJSON) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Product' and \"uniqueId\" = '" + TrimData(iRowItems.item_code) + "'";
                                sapRecSet.DoQuery(strQuery);
                            }

                            if (ItemModel.Count > 0)
                                Console.WriteLine("Processing Product:" + iRowItems.item_code + " in TAIDII. Please wait...");

                            string BaseUrl = string.Empty;
                            string MethodUrl = string.Empty;
                            string JSONResult = string.Empty;
                            string oResponseResult = string.Empty;

                            //Set Base URL Address for API Call
                            BaseUrl = base_url;

                            //Set Method for the API Call
                            MethodUrl = "centeritem/create/";

                            HttpClient httpClient = new HttpClient();
                            httpClient.BaseAddress = new Uri(BaseUrl);

                            HttpContent content = new FormUrlEncodedContent(
                            new List<KeyValuePair<string, string>> { 
                        new KeyValuePair<string, string>("api_key", api_key),
                        new KeyValuePair<string,string>("client",client),
                        new KeyValuePair<string,string>("items",strJSON)
                    });

                            HttpResponseMessage Response = httpClient.PostAsync(MethodUrl, content).Result;
                            if (Response.IsSuccessStatusCode)
                            {
                                oResponseResult = Response.Content.ReadAsStringAsync().Result;
                                if (oResponseResult.Contains("id") == true)
                                {
                                    listResponseResultSuccess = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Models.ResponseResultSuccess>>(oResponseResult);

                                    foreach (var iRowSuccess in listResponseResultSuccess)
                                    {
                                        if (iRowSuccess.status == 1)
                                        {
                                            lastMessage = "Successfully " + iif(iRowSuccess.log == "new", "created new Item", "updated existing item") + " in TAIDII Portal.";
                                            strQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Posted',\"failDesc\" = '',\"successDesc\" = '" + lastMessage + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + iRowItems.item_code + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Product' and \"uniqueId\" = '" + iRowItems.item_code + "'";
                                            sapRecSet.DoQuery(strQuery);

                                            functionReturnValue = false;
                                        }
                                    }
                                }
                                else
                                {
                                    listResponseResultFailed = Newtonsoft.Json.JsonConvert.DeserializeObject<List<SBOHelper.Models.ResponseResultFailed>>(oResponseResult);
                                    foreach (var iRowFailed in listResponseResultFailed)
                                    {
                                        if (iRowFailed.status == 0)
                                        {
                                            int errCnt = iRowFailed.errors.Count;
                                            int counter = 0;
                                            lastMessage = string.Empty;
                                            foreach (var iRowFailedDtl in iRowFailed.errors.ToList())
                                            {
                                                if (counter == 0 && errCnt != counter)
                                                {
                                                    lastMessage += iRowFailedDtl.ToString() + ", ";
                                                }
                                                else
                                                    lastMessage += iRowFailedDtl.ToString() + ", ";

                                                counter += 1;
                                            }

                                            if (lastMessage.Length > 0)
                                                lastMessage = lastMessage.Substring(0, lastMessage.Length - 2);

                                            strQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'false',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + iRowItems.item_code + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Product' and \"uniqueId\" = '" + iRowItems.item_code + "'";
                                            sapRecSet.DoQuery(strQuery);

                                            functionReturnValue = false;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                oResponseResult = Response.Content.ReadAsStringAsync().Result;
                            }
                        }
                        ////**** Create a list of Products ****////
                        Console.WriteLine("Done adding the List of " + string.Format("{0:#,##0}", sapRecSet.RecordCount) + " Product(s) in the integration log. Please wait...");
                    }
                }
                return functionReturnValue;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        #region "CreateSettings"
        public string createUDF(String tableName, String fieldName, String fieldDescription, SAPbobsCOM.BoFieldTypes fieldType, int fieldSize, SAPbobsCOM.BoFldSubTypes subfieldType = SAPbobsCOM.BoFldSubTypes.st_None, String fieldValues = "", String defaultValue = "", String linkTable = null, string DBCompany = "")
        {
            //Declarations for SQLQuery

            try
            {
                string sqlScript = "select Top 1 \"fieldID\" from \"CUFD\" where \"TableID\" = '" + tableName + "' and \"AliasID\" = '" + fieldName + "'";
                SAPbobsCOM.Recordset oRecset = (SAPbobsCOM.Recordset)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecset.DoQuery(sqlScript);

                //Execute Selected Query
                if (oRecset.RecordCount != 0)
                {
                    try
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecset);
                    }
                    catch
                    {
                    }
                    return "UDF Already Exist!";
                }

                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecset);
                }
                catch
                {
                }

                GC.Collect();
                SAPbobsCOM.UserFieldsMD oUDF = default(SAPbobsCOM.UserFieldsMD);
                oUDF = null;

                oUDF = (SAPbobsCOM.UserFieldsMD)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                //Filling userdefinefields data.
                oUDF.Name = fieldName;
                oUDF.Type = fieldType;
                oUDF.Size = fieldSize;
                oUDF.Description = fieldDescription;
                oUDF.TableName = tableName;
                oUDF.EditSize = fieldSize;
                oUDF.SubType = subfieldType;
                if (fieldValues.Length > 0)
                {
                    foreach (String s1 in fieldValues.Split('|'))
                    {
                        if ((s1.Length > 0))
                        {
                            String[] s2 = s1.Split('-');
                            oUDF.ValidValues.Description = s2[1];
                            oUDF.ValidValues.Value = s2[1];
                            oUDF.ValidValues.Add();
                        }

                    }
                }
                oUDF.DefaultValue = defaultValue;
                oUDF.LinkedTable = linkTable;
                if (oUDF.Add() == 0)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDF);
                    oUDF = null;
                    GC.Collect();

                    return "Successfully Added UDF";
                }
                else
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDF);
                    oUDF = null;
                    GC.Collect();
                    string dat = sapCompany.GetLastErrorDescription();
                    lastMessage += Environment.NewLine + DBCompany + "- Error Adding UDF : " + sapCompany.GetLastErrorDescription();

                    return "Error";
                }

            }
            catch (Exception ex)
            {
                return ex.Message;
                throw ex;
            }

        }

        public bool createUDT(String tableName, String description, SAPbobsCOM.BoUTBTableType tableType)
        {

            try
            {
                int iRet = -1;
                SAPbobsCOM.UserTablesMD ouTables = default(SAPbobsCOM.UserTablesMD);

                ouTables = (SAPbobsCOM.UserTablesMD)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

                if (!ouTables.GetByKey(tableName))
                {
                    ouTables.TableName = tableName;
                    ouTables.TableDescription = description;
                    ouTables.TableType = tableType;
                    iRet = ouTables.Add();
                }

                if (iRet == 0)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ouTables);
                    ouTables = null;
                    GC.Collect();
                    return true;
                }
                else
                {
                    lastMessage += "Fail to Add UDT " + sapCompany.GetLastErrorDescription();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ouTables);
                    ouTables = null;
                    GC.Collect();
                    return false;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region "Help"
        public string DataString(string value, int length)
        {
            if (String.IsNullOrEmpty(value)) return string.Empty;

            return value.Length <= length ? value : value.Substring(value.Length - length);
        }

        public bool CheckDate(String date)
        {
            try
            {
                DateTime iDateTIme = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region "Properties"
        public string LastErrorMessage
        {
            get
            {
                return lastMessage;
            }
        }

        public object iif(bool expression, object truePart, object falsePart)
        { return expression ? truePart : falsePart; }

        public object TrimData(string oValue)
        { return oValue.Replace("'", "''"); }

        #endregion
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
