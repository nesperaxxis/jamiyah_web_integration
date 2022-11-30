using Jamiyah_Web_Integration.SAPModels;
using Newtonsoft.Json;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Jamiyah_Web_Integration.SAPServices
{
    public class Posting
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
        public string lastMessage { get; set; }
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
                    pricelistcode = 13; // Convert.ToInt16(sapRecSet.Fields.Item("U_pricelist_code").Value.ToString());
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

        public bool SBOPostBusinessPartners(List<API_BusinessPartners> listBusinessParters, string APILastTimeStamp)
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

                            BusinessPartnersModel = new List<API_BusinessPartners>();
                            ////**** Create a list of Business Partners ****////
                            BusinessPartnersModel.Add(new API_BusinessPartners()
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
                                oBusinessPartners.PriceListNum = 12;
                                oBusinessPartners.DebitorAccount = "3020-007";
                                //GENDER = GlblLocNum  - TODO

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
                                    oBusinessPartners.ContactEmployees.Phone1 = iRowBP.contact_telephone.Length >= 17 ? iRowBP.contact_telephone.Substring(0,17) : iRowBP.contact_telephone;

                                if (!string.IsNullOrEmpty(iRowBP.contact_email))
                                    oBusinessPartners.ContactEmployees.E_Mail = iRowBP.contact_email;
                                ////**** Contact Person ****////

                                ////**** User defined fields ****////
                                if (!string.IsNullOrEmpty(iRowBP.country) && iRowBP.country != "N.A")
                                    oBusinessPartners.UserFields.Fields.Item("U_Country").Value = iRowBP.country;

                                if (!string.IsNullOrEmpty(iRowBP.level))
                                    oBusinessPartners.UserFields.Fields.Item("U_Level").Value = iRowBP.level;

                                if (!string.IsNullOrEmpty(iRowBP.nric))
                                    oBusinessPartners.UserFields.Fields.Item("U_ic_no").Value = iRowBP.nric;

                                //if (!string.IsNullOrEmpty(iRowBP.nric))
                                //    oBusinessPartners.UserFields.Fields.Item("U_IDType").Value = "NRIC";

                                if (!string.IsNullOrEmpty(Convert.ToString(iRowBP.gender)))
                                    oBusinessPartners.UserFields.Fields.Item("U_Gender").Value = (object)iRowBP.gender;

                                //if (!string.IsNullOrEmpty(iRowBP.dob))
                                //    oBusinessPartners.UserFields.Fields.Item("U_DOB").Value = Convert.ToDateTime(iRowBP.dob);

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

                                //if (!string.IsNullOrEmpty(iRowBP.bank_name))
                                //    oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_Bankname").Value = iRowBP.bank_name;

                                //if (!string.IsNullOrEmpty(iRowBP.account_name))
                                //    oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_AccName").Value = iRowBP.account_name;

                                //if (!string.IsNullOrEmpty(iRowBP.cdac_bank_no))
                                //    oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_BankAccNo").Value = iRowBP.cdac_bank_no;

                                //if (!string.IsNullOrEmpty(iRowBP.customer_ref_no))
                                //    oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_CusRefNo").Value = iRowBP.customer_ref_no;
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

                            isAddWithError:;

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
                                    oBusinessPartners.PriceListNum = 12;
                                    oBusinessPartners.DebitorAccount = "3020-007";
                                    //GENDER = GlblLocNum  - TODO

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
                                        oBusinessPartners.ContactEmployees.Phone1 = iRowBP.contact_telephone.Length >= 17 ? iRowBP.contact_telephone.Substring(0, 17) : iRowBP.contact_telephone;

                                    if (!string.IsNullOrEmpty(iRowBP.contact_email))
                                        oBusinessPartners.ContactEmployees.E_Mail = iRowBP.contact_email;
                                    ////**** Contact Person ****////

                                    ////**** User defined fields ****////
                                    if (!string.IsNullOrEmpty(iRowBP.country) && iRowBP.country != "N.A")
                                        oBusinessPartners.UserFields.Fields.Item("U_Country").Value = iRowBP.country;

                                    if (!string.IsNullOrEmpty(iRowBP.level))
                                        oBusinessPartners.UserFields.Fields.Item("U_Level").Value = iRowBP.level;

                                    if (!string.IsNullOrEmpty(iRowBP.nric))
                                        oBusinessPartners.UserFields.Fields.Item("U_ic_no").Value = iRowBP.nric;

                                    //if (!string.IsNullOrEmpty(iRowBP.nric))
                                    //    oBusinessPartners.UserFields.Fields.Item("U_IDType").Value = "NRIC";

                                    if (!string.IsNullOrEmpty(Convert.ToString(iRowBP.gender)))
                                        oBusinessPartners.UserFields.Fields.Item("U_Gender").Value = (object)iRowBP.gender;

                                    //if (!string.IsNullOrEmpty(iRowBP.dob))
                                    //    oBusinessPartners.UserFields.Fields.Item("U_DOB").Value = Convert.ToDateTime(iRowBP.dob);

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

                                    //if (!string.IsNullOrEmpty(iRowBP.bank_name))
                                    //    oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_Bankname").Value = iRowBP.bank_name;

                                    //if (!string.IsNullOrEmpty(iRowBP.account_name))
                                    //    oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_AccName").Value = iRowBP.account_name;

                                    //if (!string.IsNullOrEmpty(iRowBP.cdac_bank_no))
                                    //    oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_BankAccNo").Value = iRowBP.cdac_bank_no;

                                    //if (!string.IsNullOrEmpty(iRowBP.customer_ref_no))
                                    //    oBusinessPartners.ContactEmployees.UserFields.Fields.Item("U_CusRefNo").Value = iRowBP.customer_ref_no;
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

                            isUpdateWithError:;

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

        public bool SBOPostDownpayment(List<API_CreditNote> lstdp, string APILastTimeStamp)
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
                    lstdp = lstdp.Where(x => x.status == 1).ToList();


                    lstdp = lstdp.Where(x => x.credit_no == "ADV-000069").ToList();

                    foreach (var iRowCreditNote in lstdp)
                    {
                        try
                        {
                            if (iRowCreditNote.credit_type != 2)
                            {
                                continue;
                            }
                            string _checkIfExists = "select \"U_TransId\" from \"ODPI\" where \"U_TransId\" = '" + iRowCreditNote.id + "' and \"CANCELED\" = 'N' and \"NumAtCard\" = '" + iRowCreditNote.credit_no + "'";
                            oTransId = (String)clsSBOGetRecord.GetSingleValue(_checkIfExists, sapCompany);
                            if (oTransId != "" && oTransId != "0")
                            {
                                continue;
                            }

                            oId = iRowCreditNote.id;
                            oStatus = iRowCreditNote.status;

                            CreditNoteModelHeader = new List<API_CreditNote>();
                            CreditNoteModelDetails = new List<API_CreditNoteDetails>();

                            ////**** Create a list of Credit Note ****////
                            foreach (var iRowCreditNoteDtl in iRowCreditNote.items.ToList())
                            {
                                CreditNoteModelDetails.Add(new API_CreditNoteDetails()
                                {
                                    description = iRowCreditNoteDtl.description,
                                    date_for = iRowCreditNoteDtl.date_for,
                                    amount = iRowCreditNoteDtl.amount,
                                    gst = iRowCreditNoteDtl.gst
                                });
                            }

                            CreditNoteModelHeader.Add(new API_CreditNote()
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

                                string Query = "select \"U_TransId\" from \"ODPI\" where \"U_TransId\" = '" + iRowCreditNote.id + "' and \"CANCELED\" = 'N' and \"NumAtCard\" = '" + iRowCreditNote.credit_no + "' and \"ObjType\" = 203  ";                           
                                oTransId = (String)clsSBOGetRecord.GetSingleValue(Query, sapCompany);
                                if (oTransId == "" || oTransId == "0")
                                {
                                    oCreditNote = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oDownPayments);
                                    oCreditNote.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptInvoice;
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
                                    if (DateTime.Parse(iRowCreditNote.date_created) >= DateTime.Parse("07/01/2022") && DateTime.Parse(iRowCreditNote.date_created) <= DateTime.Parse("07/31/2022"))
                                    {
                                        iRowCreditNote.date_created = "07/01/2022";
                                    }
                                    else if (DateTime.Parse(iRowCreditNote.date_created) >= DateTime.Parse("08/01/2022") && DateTime.Parse(iRowCreditNote.date_created) <= DateTime.Parse("08/31/2022"))
                                    {
                                        iRowCreditNote.date_created = "08/01/2022";
                                    }


                                    oCreditNote.BPL_IDAssignedToInvoice = 5; //"Jamiyah Education Centre (JEC)";
                                    oCreditNote.DocDate = Convert.ToDateTime(iRowCreditNote.date_created);
                                    oCreditNote.NumAtCard = iRowCreditNote.credit_no;

                                    if (iRowCreditNote.status == 1)
                                        oCreditNote.Comments = iRowCreditNote.remarks;
                                    else
                                        oCreditNote.Comments = iRowCreditNote.void_remarks;

                                    ////**** UDF *****/////
                                    if (iRowCreditNote.id != 0)
                                        oCreditNote.UserFields.Fields.Item("U_TransId").Value = iRowCreditNote.id.ToString();

                                    if (iRowCreditNote.level != "")
                                        oCreditNote.UserFields.Fields.Item("U_Level").Value = iRowCreditNote.level;

                                    if (iRowCreditNote.program_type != "")
                                        oCreditNote.UserFields.Fields.Item("U_ProgramType").Value = iRowCreditNote.program_type;


                                    oCreditNote.UserFields.Fields.Item("U_branch").Value = "Jamiyah Education Centre (JEC)";
                                  
                                    foreach (var iRowCreditNoteDtls in iRowCreditNote.items.ToList())
                                    {
                                        if (iRowCreditNoteDtls.description != "")
                                        {
                                            oDocType = "dDocument_Items";

                                            string iReplaceDesc = " (" + TrimData(iRowCreditNote.level) + " - " + TrimData(iRowCreditNote.program_type) + ")";
                                            //oDescription = SBOstrManipulation.BeforeCharacter(iRowCreditNoteDtls.description, " (");
                                            oDescription = iRowCreditNoteDtls.description.Replace(iReplaceDesc, "");

                                            if (oDescription != "")
                                            {
                                                string description = oDescription;
                                                //string iDescription = (String)clsSBOGetRecord.GetSingleValue("select \"U_Description\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + description + "' and \"U_Level\" = '" + iRowCreditNote.level + "' and \"U_ProgramType\" = '" + iRowCreditNote.program_type + "'", sapCompany);
                                                string itemCode = (String)clsSBOGetRecord.GetSingleValue("select \"U_ccode\" from Jamiyah_Live..\"@JEC\" where \"U_descript\" = '" + TrimData(iRowCreditNoteDtls.description) + "' and \"U_unitprice\" = '" + iRowCreditNoteDtls.amount + "'", sapCompany);
                                                itemCode = "JEC100102";
                                                if (itemCode != "")
                                                {
                                                    string idate_created = string.Empty;
                                                    string idate_for = string.Empty;
                                                    string iGLAccount = "5000-005";//string.Empty;
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

                                                    oCreditNote.Lines.ItemCode = itemCode;
                                                    oCreditNote.Lines.Quantity = 1;

                                                    oCardName = (String)clsSBOGetRecord.GetSingleValue("select \"CardName\" from \"OCRD\" where \"CardCode\" = '" + TrimData(iRowCreditNote.student) + "'", sapCompany);

                                                    string oTaxCode = "SR";
                                                    //(String)clsSBOGetRecord.GetSingleValue("select \"U_TaxCode\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(iRowCreditNote.level) + "' and \"U_ProgramType\" = '" + TrimData(iRowCreditNote.program_type) + "'", sapCompany);

                                                    if (!string.IsNullOrEmpty(oTaxCode))
                                                        oCreditNote.Lines.VatGroup = oTaxCode;
                                                    oCreditNote.Lines.PriceAfterVAT = iRowCreditNoteDtls.amount;

                                                    oItemDescription = oCardName + " - " + oDateFor + " - " + iRowCreditNoteDtls.description;
                                                    //oCreditNote.Lines.UserFields.Fields.Item("U_Dscription").Value = oItemDescription;
                                                    oCreditNote.Lines.ItemDescription = iRowCreditNoteDtls.description;
                                                    //oCreditNote.Lines.Price =   iRowCreditNoteDtls.amount;
                                                    //oCreditNote.Lines.LineTotal = iRowCreditNoteDtls.amount;
                                                    oCreditNote.Lines.UnitsOfMeasurment = 1;
                                                    //if (!string.IsNullOrEmpty(iRowCreditNote.date_created))
                                                    //    idate_created = iRowCreditNote.date_created;

                                                    //oCreditNote.Lines. = "1";
                                                    oCreditNote.Lines.ProjectCode = "00";
                                                    oCreditNote.Lines.CostingCode = "21";
                                                    oCreditNote.Lines.CostingCode2 = "N/A";
                                                    oCreditNote.Lines.CostingCode3 = "G_00";
                                                    oCreditNote.Lines.CostingCode4 = "Default";
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
                                            string seriesNum = (String)clsSBOGetRecord.GetSingleValue("select TOP 1 Series from \"NNM1\" \"e\"  where \"e\".SeriesName like '%JEC%' AND BPLId = 5 AND Indicator = YEAR(GETDATE()) AND ObjectCode = '24'", sapCompany);
                                            var InPay = (SAPbobsCOM.Payments)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                                            InPay.Series = int.Parse(seriesNum);
                                            InPay.CardCode = oCreditNote.CardCode;
                                            InPay.BPLID = oCreditNote.BPL_IDAssignedToInvoice;
                                            InPay.DocDate = oCreditNote.DocDate;
                                            InPay.CashAccount = "3040-001";
                                            InPay.CashSum = iRowCreditNote.items.Sum(x => x.amount);
                                            InPay.Invoices.DocEntry = Convert.ToInt32(oDocEntry);
                                            InPay.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_DownPayment;
                                            InPay.UserFields.Fields.Item("U_TransId").Value = iRowCreditNote.id.ToString();
                                            //oIncomingPayment.UserFields.Fields.Item("U_StatusTaidii").Value = iRowCreditNote.status.ToString();
                                            InPay.UserFields.Fields.Item("U_tax").Value = "N/A";
                                            InPay.UserFields.Fields.Item("U_ipc").Value = "NON-IPC";
                                            //oIncomingPayment.UserFields.Fields.Item("U_Level").Value = iRowReceipt.level;
                                            //oIncomingPayment.UserFields.Fields.Item("U_ProgramType").Value = iRowReceipt.program_type;
                                            InPay.UserFields.Fields.Item("U_ReceiptNo").Value = "N/A";
                                            InPay.UserFields.Fields.Item("U_branch").Value = "Jamiyah Education Centre (JEC)";
                                            var lRetCode = InPay.Add();

                                            if (lRetCode != 0)
                                            {
                                                lastMessage = sapCompany.GetLastErrorDescription();
                                                sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditNote.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + lastMessage + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Note' and \"uniqueId\" = '" + iRowCreditNote.id + "'");
                                            }
                                            else
                                            {
                                                lastMessage = "Successfully created Payment for Downpayment Invoice with Transaction Id:" + iRowCreditNote.id + " in SAP B1.";
                                                sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Draft',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "',\"objType\" = 112 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Note' and \"uniqueId\" = '" + iRowCreditNote.id + "'");
                                            }

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

                                isAddWithError:;

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

        public bool SBOPostInvoice(List<API_Invoice> listInvoice, string APILastTimeStamp)
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
            List<API_Invoice> notAdded = new List<API_Invoice>();
            SBOGetRecord clsSBOGetRecord = new SBOGetRecord();
            try
            {
                if (SBOconnectToLoginCompany(SBOConstantClass.SBOServer, SBOConstantClass.Database, SBOConstantClass.ServerUN, SBOConstantClass.ServerPW, SBOConstantClass.SAPUser, SBOConstantClass.SAPPassword))
                {
                    GetIntegrationSetup();

                    #region Cancel Invoices and Credit Notes
                    ////Cancel Invoices and Credit Notes
                    //               var forRevertInvoices = clsSBOGetRecord.TaidiiInvoicesDocEntries("01/01/2022", "05/31/2022", sapCompany);

                    //               foreach (var inv in forRevertInvoices)
                    //               {
                    //                   var Invoice = (Documents)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                    //                   Invoice.GetByKey(inv.DocEntry);
                    //                   var CancelInvoice = Invoice.CreateCancellationDocument();
                    //                   CancelInvoice.Comments = $"[REVERT]";
                    //                   CancelInvoice.Add();
                    //	System.Runtime.InteropServices.Marshal.ReleaseComObject(Invoice);
                    //	System.Runtime.InteropServices.Marshal.ReleaseComObject(CancelInvoice);
                    //}

                    //               var forRevertCNs = clsSBOGetRecord.TaidiiCreditNotesDocEntries("01/01/2022", "05/31/2022", sapCompany);

                    //               foreach (var inv in forRevertCNs)
                    //               {

                    //                   var CreditNote = (Documents)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
                    //                   CreditNote.GetByKey(inv.DocEntry);
                    //                   var CancelCreditNote = CreditNote.CreateCancellationDocument();
                    //                   CancelCreditNote.Comments = $"[REVERT]";
                    //                   CancelCreditNote.Add();
                    //	System.Runtime.InteropServices.Marshal.ReleaseComObject(CreditNote);
                    //	System.Runtime.InteropServices.Marshal.ReleaseComObject(CancelCreditNote);
                    //}
                    //               return true;
                    #endregion

                    var invoices = listInvoice.Where(x => x.status == 1).ToList();
                    //var invoices = listInvoice.Where(x => x.status == 1 && x.invoice_no == "INV-002786").ToList();
                    //return true;
                    //invoices = listInvoice.Where(x => DateTime.Parse(x.date_created) >= DateTime.Parse("02/01/2022") && DateTime.Parse(x.date_created) <= DateTime.Parse("02/28/2022")).ToList();
                    //invoices = invoices.Where(x => DateTime.Parse(x.date_created) >= DateTime.Parse("09/01/2022") && DateTime.Parse(x.date_created) <= DateTime.Parse("09/30/2022")).ToList();
                    //return true;

                    foreach (var iRowInv in invoices)
                    {
                        string _checkIfExists = "select \"U_TransId\" from \"OINV\" where \"U_TransId\" = '" + iRowInv.id + "' and \"CANCELED\" = 'N' and \"NumAtCard\" = '" + iRowInv.invoice_no + "'";
                        oTransId = (String)clsSBOGetRecord.GetSingleValue(_checkIfExists, sapCompany);
                        //if (oTransId != "" && oTransId != "0" && !iRowInv.downPaymentAmount.HasValue && !iRowInv.downPaymentDocEntry.HasValue)
                        if ((oTransId != "" && oTransId != "0" && !iRowInv.isAppliedDP.HasValue) || (iRowInv.isAppliedDP.HasValue && iRowInv.isAppliedDP.Value == true && iRowInv.OpenDPs?.Count < 1))
                        {
                            continue;
                        }                       

                        bool hasItemCode = true;
                        try
                        {
                            oId = iRowInv.id;
                            oStatus = iRowInv.status;

                            InvoiceModelHeader = new List<API_Invoice>();
                            InvoiceModelDetails = new List<API_InvoiceDetails>();

                            ////**** Create a list of Invoices ****////
                            foreach (var iRowInvDtl in iRowInv.items.ToList())
                            {
                                iRowInvDtl.item_code = (String)clsSBOGetRecord.GetSingleValue("select \"U_ccode\" from Jamiyah_Live..\"@JEC\" where \"U_descript\" = '" + TrimData(iRowInvDtl.description) + "' and \"U_unitprice\" = '" + iRowInvDtl.unit_price + "'", sapCompany);
                                InvoiceModelDetails.Add(new API_InvoiceDetails()
                                {
                                    description = iRowInvDtl.description,
                                    item_code = iRowInvDtl.item_code,
                                    date_for = iRowInvDtl.date_for,
                                    unit_price = iRowInvDtl.unit_price,
                                    quantity = iRowInvDtl.unit_price < 1 ? -1 : iRowInvDtl.quantity,
                                    total = iRowInvDtl.total
                                });


                                if (String.IsNullOrEmpty(iRowInvDtl.item_code))
                                {
                                    hasItemCode = false;
                                    notAdded.Add(iRowInv);
                                }
                            }
                            if (!hasItemCode)
                            {
                                continue;
                            }

                            //iRowInv.date_created = "07/02/2022";
                            //iRowInv.date_due = "07/02/2022";
                            //if (DateTime.Parse(iRowInv.date_created) >= DateTime.Parse("07/01/2022") && DateTime.Parse(iRowInv.date_created) <= DateTime.Parse("07/31/2022"))
                            //{
                            //    iRowInv.date_created = "07/07/2022";
                            //    iRowInv.date_due = "07/07/2022";
                            //}
                            //else if (DateTime.Parse(iRowInv.date_created) >= DateTime.Parse("08/01/2022") && DateTime.Parse(iRowInv.date_created) <= DateTime.Parse("08/31/2022"))
                            //{
                            //    iRowInv.date_created = "08/01/2022";
                            //    iRowInv.date_due = "08/01/2022";
                            //}

                            ////else if (DateTime.Parse(iRowInv.date_created) >= DateTime.Parse("06/01/2022") && DateTime.Parse(iRowInv.date_created) <= DateTime.Parse("06/30/2022"))
                            ////{
                            ////    iRowInv.date_created = "07/06/2022";
                            ////    iRowInv.date_due = "07/06/2022";
                            ////}
                            //else
                            //{
                            //    continue;
                            //}


                            InvoiceModelHeader.Add(new API_Invoice()
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



                            if (iRowInv.status == 1)
                            {
                                Console.WriteLine("Processing Invoice with Transaction Id:" + iRowInv.id + " in SAP B1 Draft. Please wait...");

                                string Query = "select \"U_TransId\" from \"ODRF\" where \"U_TransId\" = '" + iRowInv.id + "' and \"CANCELED\" = 'N' and \"NumAtCard\" = '" + iRowInv.invoice_no + "' and \"ObjType\" = 13 " + Environment.NewLine +
                                "union all " + Environment.NewLine +
                                "select \"U_TransId\" from \"OINV\" where \"U_TransId\" = '" + iRowInv.id + "' and \"CANCELED\" = 'N' and \"NumAtCard\" = '" + iRowInv.invoice_no + "'";
                                oTransId = (String)clsSBOGetRecord.GetSingleValue(Query, sapCompany);
                                if (oTransId == "" || oTransId == "0" || (iRowInv.isAppliedDP.HasValue && iRowInv.isAppliedDP.Value == true && iRowInv.OpenDPs.Count > 0))
                                {
                                    oInvoice = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oInvoices);
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

                                    string seriesNum = (String)clsSBOGetRecord.GetSingleValue("select TOP 1 Series from \"NNM1\" \"e\"  where \"e\".SeriesName like 'JEC%' AND BPLId = 5 AND Indicator = YEAR(GETDATE()) AND ObjectCode = '13' ORDER BY LastNum DESC", sapCompany);
                                    oInvoice.BPL_IDAssignedToInvoice = 5; //"Jamiyah Education Centre (JEC)";
                                    oInvoice.DocDate = Convert.ToDateTime(iRowInv.date_created);
                                    oInvoice.NumAtCard = iRowInv.invoice_no;
                                    oInvoice.DocDueDate = Convert.ToDateTime(iRowInv.date_due);
                                    oInvoice.Series = int.Parse(seriesNum);
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


                                    oInvoice.UserFields.Fields.Item("U_branch").Value = "Jamiyah Education Centre (JEC)";
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
                                                //string iDescription = (String)clsSBOGetRecord.GetSingleValue("select \"U_Description\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(iRowInv.level) + "' and \"U_ProgramType\" = '" + TrimData(iRowInv.program_type) + "'", sapCompany);
                                                //if (iDescription != "")
                                                //{
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

                                                string oTaxCode = ""; //(String)clsSBOGetRecord.GetSingleValue("select \"U_TaxCode\" from \"@GLACCTMAPPING\" where \"U_Description\" = '" + TrimData(description) + "' and \"U_Level\" = '" + TrimData(iRowInv.level) + "' and \"U_ProgramType\" = '" + TrimData(iRowInv.program_type) + "'", sapCompany);

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


                                                oInvoice.Lines.Add();

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

                                                oInvoice.Lines.WarehouseCode = "JEC";
                                                //oInvoice.Lines.AccountCode = "1065-002";
                                                oInvoice.Lines.ProjectCode = "00";
                                                oInvoice.Lines.CostingCode = "21";
                                                oInvoice.Lines.CostingCode2 = "N/A";
                                                oInvoice.Lines.CostingCode3 = "G_00";
                                                oInvoice.Lines.CostingCode4 = "Default";

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

                                    oInvoice.DocType = BoDocumentTypes.dDocument_Items;

                                    if (iRowInv.OpenDPs?.Count > 0)
                                    {

                                        CancelInvoiceForDP(0, iRowInv.oldInvDocEntry.Value);

                                        var docTotal = (double)iRowInv.items.Sum(x => x.unit_price);
                                        double appliedAmt = 0;
                                        int i = 0;
                                        while (appliedAmt < docTotal && iRowInv.OpenDPs?.Count > i)
                                        {
                                            var dpDraw = oInvoice.DownPaymentsToDraw;
                                            var dpAmount = double.Parse(iRowInv.OpenDPs[i].Ammount.ToString());

                                            if (dpAmount > docTotal)
                                            {
                                                dpDraw.AmountToDraw = docTotal;
                                            }
                                            else
                                            {
                                                dpDraw.AmountToDraw = dpAmount;
                                            }
                                            dpDraw.DocEntry = int.Parse(iRowInv.OpenDPs[i].DocEntry);
                                            dpDraw.Add();
                                            appliedAmt = appliedAmt + dpAmount;
                                            docTotal = docTotal - dpAmount;
                                            i++;
                                        }
                                        //var dpDraw = oInvoice.DownPaymentsToDraw;
                                        //var dpAmount = double.Parse(iRowInv.downPaymentAmount.ToString());
                                        //var docTotal = iRowInv.items.Sum(x => x.unit_price);
                                        //if (dpAmount > docTotal)
                                        //{
                                        //    dpDraw.AmountToDraw = docTotal;
                                        //}
                                        //else
                                        //{
                                        //    dpDraw.AmountToDraw = dpAmount;
                                        //}
                                       
                                        //dpDraw.DocEntry = iRowInv.downPaymentDocEntry.Value;
                                    }                                    

                                    lErrCode = oInvoice.Add();
                                    if (lErrCode == 0)
                                    {
                                        try
                                        {
                                            oDocEntry = sapCompany.GetNewObjectKey();
                                            lastMessage = "Successfully created Invoice (Draft) with Transaction Id:" + iRowInv.id + " in SAP B1.";
                                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Draft',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "',\"objType\" = 112 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowInv.id + "'");

                                            if (iRowInv.OpenDPs?.Count > 0)
                                            {
                                                UpdateCancelledInvoiceForDP(int.Parse(oDocEntry), iRowInv.oldInvDocEntry.Value);
                                            }                                          

                                            functionReturnValue = false;
                                        }
                                        catch (Exception ex)
                                        {
                                            lastMessage = sapCompany.GetLastErrorDescription();
                                        }
                                    }
                                    else
                                    {
                                        lastMessage = sapCompany.GetLastErrorDescription();
                                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowInv.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowInv.id + "'");

                                        functionReturnValue = true;

                                        goto isAddWithError;
                                    }

                                isAddWithError:;

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvoice);

                                }

                            }
                        }
                        catch (Exception ex)
                        {
                            lastMessage = ex.ToString();
                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowInv.status == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + iRowInv.id + "' and \"sapCode\" is not null");
                        }
                    }

                    var voidInvoices = listInvoice.Where(x => x.status == 2).ToList();
                    //voidInvoices = voidInvoices.Where(x => DateTime.Parse(x.date_created) >= DateTime.Parse("03/01/2022") && DateTime.Parse(x.date_created) <= DateTime.Parse("03/31/2022")).ToList();
                    voidInvoices = voidInvoices.Where(x => DateTime.Parse(x.date_created) >= DateTime.Parse("04/01/2022") && DateTime.Parse(x.date_created) <= DateTime.Parse("06/30/2022")).ToList();
                    if (voidInvoices.Count > 0)
                    {
                        Int32 iDocEntry = CreateInvoiceVoid(voidInvoices);
                        if (iDocEntry != 0)
                        {
                            functionReturnValue = false;
                        }
                        else
                            functionReturnValue = true;
                    }
                    Console.WriteLine("Done adding the List of " + string.Format("{0:#,##0}", RecordCount) + " Invoice(s) in the integration log. Please wait...");
                }
            }
            catch (Exception ex)
            {
                lastMessage = ex.ToString();
                sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(oStatus == 1, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Invoice' and \"uniqueId\" = '" + oId + "' and \"sapCode\" is not null");
            }
            notAdded = notAdded;
            var _s = String.Join("\r\n", notAdded.Select(x => x.items).Select(y => String.Join("\r\n", y.Select(z => "Item: " + z.description + " \r\n" + "price: " + z.unit_price.ToString()).Distinct().ToList())).Select(v => v).Distinct().ToList().Select(x => x).Distinct().ToList());
            return functionReturnValue;
        }   

        public bool SBOPostReceipt(List<API_Receipt> listReceipt)
        {
            //var receiptNums = new string[] { "RCP-000353", "RCP-000347", "RCP-000285", "RCP-000256", "RCP-000037", "RCP-000669", "RCP-000668", "RCP-000586", "RCP-001491", "RCP-000604", "RCP-001114","RCP-001115", " RCP-001116", "RCP-001128" };
            //var receiptNums = new string[] { "RCP-001491", "RCP-000604", "RCP-001114", "RCP-001115", "RCP-001128" };
            //var receiptNums = new string[] { "RCP-000668", "RCP-000669", "RCP-000037", "RCP-000285", "RCP-000347", "RCP-000353", "RCP-000586", "RCP-000256" };
            //var receiptNums = new string[] { "RCP-003139" };
            bool functionReturnValue = false;
            int lErrCode = 0;
            string oLogExist = string.Empty;
            string oTransId = string.Empty;
            string oCardCode = string.Empty;
            string oCardName = string.Empty;
            string oDocEntry = string.Empty;
            string oInvDocEntry = string.Empty;
            string oCreditNoteDocEntry = string.Empty;
            string oModeOfPayment = string.Empty;
            string oAcctCode = string.Empty;
            string oBankName = string.Empty;
            string oCheckBankName = string.Empty;
            string iReference = string.Empty;

            SBOGetRecord clsSBOGetRecord = new SBOGetRecord();
            try
            {
                if (SBOconnectToLoginCompany(SBOConstantClass.SBOServer, SBOConstantClass.Database, SBOConstantClass.ServerUN, SBOConstantClass.ServerPW, SBOConstantClass.SAPUser, SBOConstantClass.SAPPassword))
                {
                    GetIntegrationSetup();

                    //listReceipt = listReceipt.Where(x => DateTime.Parse(x.date_created) >= DateTime.Parse("04/01/2022") && DateTime.Parse(x.date_created) <= DateTime.Parse("06/30/2022")).ToList();

                    //CreateReceiptVoid(listReceipt);
                    //return functionReturnValue;
                    //listReceipt = listReceipt.Where(x => receiptNums.Contains(x.receipt_no)).ToList();

                    var newReceipt = listReceipt.Where(x => x.status == 0).ToList();

                    //newReceipt = newReceipt.Where(x => x.receipt_no == "RCP-003554").ToList();

                    foreach (var iRowReceipt in newReceipt)
                    {
                        try
                        {
                            //0 = no offset
                            //1 = has both payment and offset
                            //2 = only offset type
                            if (iRowReceipt.payment_type == 0 || iRowReceipt.payment_type == 1)
                            {
                                oIncomingPayment = (Payments)sapCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments);
                                oDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ORCT\" where \"U_TransId\" = '" + iRowReceipt.id + "' and \"Canceled\" = 'N'", sapCompany);
                                if (iRowReceipt.status == 0 && (oDocEntry == "" || oDocEntry == "0"))
                                {
                                    oTransId = (String)clsSBOGetRecord.GetSingleValue("select \"U_TransId\" from \"ORCT\" where \"U_TransId\" = '" + iRowReceipt.id + "' and \"Canceled\" = 'N'", sapCompany);
                                    oIncomingPayment.DocObjectCode = BoPaymentsObjectType.bopot_IncomingPayments;
                                    if (oTransId == "" || oTransId == "0")
                                    {
                                        oCardCode = (String)clsSBOGetRecord.GetSingleValue("select \"CardCode\" from \"OCRD\" where \"CardCode\" = '" + TrimData(iRowReceipt.student) + "'", sapCompany);
                                        if (oCardCode != "")
                                        {
                                            oIncomingPayment.CardCode = oCardCode;

                                            oCardName = (String)clsSBOGetRecord.GetSingleValue("select \"CardName\" from \"OCRD\" where \"CardCode\" = '" + TrimData(iRowReceipt.student) + "'", sapCompany);
                                        }
                                        else
                                        {
                                            lastMessage = "Customer Code:" + iRowReceipt.student + " is not found in SAP B1";
                                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                            functionReturnValue = true;

                                            goto isAddWithError;
                                        }

                                        var _invoiceIds = String.Join(",", iRowReceipt.invoice_id);
                                        string _earliestInvDate = (String)clsSBOGetRecord.GetSingleValue("SELECT TOP 1 DocDate FROM OINV WHERE \"U_TransId\" IN ('" + _invoiceIds + "') ORDER BY DocDate ASC", sapCompany);
                                        string _latestInvDate = (String)clsSBOGetRecord.GetSingleValue("SELECT TOP 1 DocDate FROM OINV WHERE \"U_TransId\" IN ('" + _invoiceIds + "') ORDER BY DocDate DESC", sapCompany);

                                        oIncomingPayment.BPLID = 5; //"Jamiyah Education Centre (JEC)";
                                        oIncomingPayment.DocTypte = BoRcptTypes.rCustomer;
                                        //if (DateTime.Parse(iRowReceipt.date_created) >= DateTime.Parse("04/01/2022") && DateTime.Parse(iRowReceipt.date_created) <= DateTime.Parse("04/30/2022"))
                                        //{
                                        //    oIncomingPayment.DocDate = Convert.ToDateTime("07/04/2022");
                                        //}
                                        //else if (DateTime.Parse(iRowReceipt.date_created) >= DateTime.Parse("05/01/2022") && DateTime.Parse(iRowReceipt.date_created) <= DateTime.Parse("05/31/2022"))
                                        //{
                                        //    oIncomingPayment.DocDate = Convert.ToDateTime("07/05/2022");
                                        //}

                                        //else if (DateTime.Parse(iRowReceipt.date_created) >= DateTime.Parse("06/01/2022") && DateTime.Parse(iRowReceipt.date_created) <= DateTime.Parse("06/30/2022"))
                                        //{
                                        //    oIncomingPayment.DocDate = Convert.ToDateTime("07/06/2022");
                                        //}
                                        //else
                                        //{
                                        //    continue;
                                        //}
                                        //oIncomingPayment.DocDate = Convert.ToDateTime("07/06/2022");

                                        oIncomingPayment.DocDate = (!String.IsNullOrEmpty(_earliestInvDate) && Convert.ToDateTime(_earliestInvDate) > Convert.ToDateTime(iRowReceipt.date_created)
                                                            ? Convert.ToDateTime(_earliestInvDate) : Convert.ToDateTime(iRowReceipt.date_created));
                                        string seriesNum = (String)clsSBOGetRecord.GetSingleValue("select TOP 1 Series from \"NNM1\" \"e\"  where \"e\".SeriesName like '%JEC%' AND BPLId = 5 AND Indicator = YEAR(GETDATE()) AND ObjectCode = '24'", sapCompany);
                                        oIncomingPayment.Series = int.Parse(seriesNum);
                                        ////**** UDF ****\\\\     
                                        oIncomingPayment.UserFields.Fields.Item("U_TransId").Value = iRowReceipt.id.ToString();
                                        oIncomingPayment.UserFields.Fields.Item("U_StatusTaidii").Value = iRowReceipt.status.ToString();
                                        oIncomingPayment.UserFields.Fields.Item("U_tax").Value = "N/A";
                                        oIncomingPayment.UserFields.Fields.Item("U_ipc").Value = "NON-IPC";
                                        //oIncomingPayment.UserFields.Fields.Item("U_Level").Value = iRowReceipt.level;
                                        //oIncomingPayment.UserFields.Fields.Item("U_ProgramType").Value = iRowReceipt.program_type;
                                        oIncomingPayment.UserFields.Fields.Item("U_ReceiptNo").Value = iRowReceipt.receipt_no;
                                        oIncomingPayment.UserFields.Fields.Item("U_branch").Value = "Jamiyah Education Centre (JEC)";
                                        ////**** UDF ****\\\\

                                        if (iRowReceipt.status == 0)

                                            if (iRowReceipt.remarks.Length >= 200)
                                            {
                                                oIncomingPayment.Remarks = oCardName.Substring(0, oCardName.Length > 50 ? 50 : oCardName.Length - 1) + " " + iRowReceipt.remarks.Substring(0, iRowReceipt.remarks.Length > 200 ? 200 : iRowReceipt.remarks.Length - 1);
                                            }
                                            else
                                                oIncomingPayment.Remarks = oCardName + " " + iRowReceipt.remarks;
                                        else
                                        {
                                            if (iRowReceipt.void_remarks.Length >= 200)
                                            {
                                                oIncomingPayment.Remarks = oCardName.Substring(0, oCardName.Length > 50 ? 50 : oCardName.Length - 1) + " " + iRowReceipt.void_remarks.Substring(0, iRowReceipt.void_remarks.Length > 200 ? 200 : iRowReceipt.void_remarks.Length - 1);
                                                //oIncomingPayment.Remarks = oCardName.Substring(0, oCardName.Length > 50 ? 50 : oCardName.Length - 1) + " " + iRowReceipt.void_remarks;
                                            }
                                            else
                                                oIncomingPayment.Remarks = oCardName + " " + iRowReceipt.void_remarks;
                                        }

                                        ////**** Adding of List of Invoice to Incoming Payment ****\\\\
                                        int invoiceCount = 0;
                                        int invPaidCount;
                                        foreach (var iRowReceiptInvDtl in iRowReceipt.invoice_id.ToList())
                                        {
                                            invoiceCount += 1;
                                            oInvDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"OINV\" " + Environment.NewLine +
                                            "where \"U_TransId\" = '" + iRowReceiptInvDtl.ToString() + "' and \"CANCELED\" = 'N'", sapCompany);
                                            if (oInvDocEntry != "" && oInvDocEntry != "0")
                                            {
                                                invPaidCount = 0;
                                                foreach (var iRowReceiptInvPaidDtl in iRowReceipt.invoice_paid.ToList())
                                                {
                                                    invPaidCount += 1;
                                                    if (invoiceCount == invPaidCount)
                                                    {
                                                        oIncomingPayment.Invoices.DocEntry = Convert.ToInt32(oInvDocEntry);
                                                        oIncomingPayment.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                                                        oIncomingPayment.Invoices.SumApplied = Convert.ToDouble(iRowReceiptInvPaidDtl.ToString());
                                                        oIncomingPayment.Invoices.Add();
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                lastMessage = "Invoice with Transaction id:" + iRowReceiptInvDtl.ToString() + " does not exist in SAP B1.";
                                                sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                                functionReturnValue = true;

                                                goto isAddWithError;
                                            }
                                        }
                                        ////**** Adding of List of Invoice to Incoming Payment ****\\\\

                                        ////**** Adding of List of Credit Note to Incoming Payment ****\\\\
                                        iReference = string.Empty;
                                        foreach (var iRowReceiptInvDtls in iRowReceipt.payment_methods.ToList())
                                        {
                                            if (iRowReceiptInvDtls.method == 3 || iRowReceiptInvDtls.method == 8 || iRowReceiptInvDtls.method == 10) //**OFFSET_DEPOSIT = 3**\\
                                            {
                                                if (!string.IsNullOrEmpty(iRowReceiptInvDtls.reference_id) && iRowReceiptInvDtls.reference_id != "N.A")
                                                {
                                                    oCreditNoteDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ORIN\" where \"U_TransId\" = '" + TrimData(iRowReceiptInvDtls.reference_id) + "' and \"CANCELED\" = 'N' and \"U_CreatedByVoucher\" = 0", sapCompany);
                                                    if (oCreditNoteDocEntry != "" && oCreditNoteDocEntry != "0")
                                                    {
                                                        oIncomingPayment.Invoices.DocEntry = Convert.ToInt16(oCreditNoteDocEntry);
                                                        oIncomingPayment.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;
                                                        oIncomingPayment.Invoices.Add();
                                                    }
                                                    else
                                                    {
                                                        string oDraftDocEntry = string.Empty;
                                                        oDraftDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ODRF\" where \"U_TransId\" = '" + iRowReceiptInvDtls.reference_id + "' and \"CANCELED\" = 'N' and \"ObjType\" = 14", sapCompany);
                                                        if (oDraftDocEntry != "" && oDraftDocEntry != "0")
                                                        {
                                                            SAPbobsCOM.Documents oDraft = (Documents)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                                            if (oDraft.GetByKey(Convert.ToInt16(oDraftDocEntry)))
                                                            {
                                                                int ErrCode = oDraft.SaveDraftToDocument();
                                                                if (ErrCode == 0)
                                                                {
                                                                    oCreditNoteDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ORIN\" where \"U_TransId\" = '" + iRowReceiptInvDtls.reference_id + "' and \"CANCELED\" = 'N' and \"U_CreatedByVoucher\" = 0", sapCompany);
                                                                    if (oCreditNoteDocEntry != "" && oCreditNoteDocEntry != "0")
                                                                    {
                                                                        oIncomingPayment.Invoices.DocEntry = Convert.ToInt16(oCreditNoteDocEntry);
                                                                        oIncomingPayment.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;
                                                                        oIncomingPayment.Invoices.Add();
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    lastMessage = sapCompany.GetLastErrorDescription();
                                                                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                                                    functionReturnValue = true;

                                                                    goto isAddWithError;
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            lastMessage = "Credit Note with Reference Id:" + iRowReceiptInvDtls.reference_id + " does not exist in SAP B1 Drafts";
                                                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                                            functionReturnValue = true;

                                                            goto isAddWithError;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (iRowReceiptInvDtls.reference != "N.A")
                                                        iReference += iRowReceiptInvDtls.reference + ", ";
                                                }
                                            }
                                            else
                                            {
                                                if (iRowReceiptInvDtls.reference != "N.A")
                                                    iReference += iRowReceiptInvDtls.reference + ", ";
                                            }
                                        }
                                        ////**** Adding of List of Credit Note to Incoming Payment ****\\\\

                                        string oJournalRemarks = string.Empty;
                                        if (!string.IsNullOrEmpty(iReference))
                                        {
                                            oJournalRemarks = iReference.Substring(0, iReference.Length - 2);
                                        }

                                        if (!string.IsNullOrEmpty(oJournalRemarks))
                                            oIncomingPayment.JournalRemarks = oJournalRemarks;

                                        ////**** Payment Means for the List of Invoices ****\\\\
                                        foreach (var iRowReceiptDtls in iRowReceipt.payment_methods.ToList())
                                        {
                                            if (string.IsNullOrEmpty(iRowReceiptDtls.reference_id) || iRowReceiptDtls.reference_id == "N.A")
                                            {
                                                oAcctCode = (String)clsSBOGetRecord.GetSingleValue("select \"U_GLAccount\" from \"@PAYMENTCODES\" where \"U_PaymentCodeMethod\" = " + iRowReceiptDtls.method + "", sapCompany);

                                                oModeOfPayment = (String)clsSBOGetRecord.GetSingleValue("select \"U_ModePayment\" from \"@PAYMENTCODES\" where \"U_PaymentCodeMethod\" = " + iRowReceiptDtls.method + "", sapCompany);

                                                oIncomingPayment.UserFields.Fields.Item("U_type").Value = "Local";
                                                oIncomingPayment.UserFields.Fields.Item("U_giro").Value = "N/A";

                                                if (oModeOfPayment == "Cash" || oModeOfPayment == "NETS")
                                                {
                                                    if (!string.IsNullOrEmpty(oAcctCode))
                                                    {
                                                        oIncomingPayment.UserFields.Fields.Item("U_cash").Value = oModeOfPayment;
                                                        oIncomingPayment.CashAccount = oAcctCode;
                                                    }
                                                    if (iRowReceiptDtls.amount != 0)
                                                        oIncomingPayment.CashSum = iRowReceiptDtls.amount;
                                                }
                                                else if (oModeOfPayment == "Check")
                                                {
                                                    if (!string.IsNullOrEmpty(oAcctCode))
                                                    {
                                                        oIncomingPayment.UserFields.Fields.Item("U_cash").Value = "CHQ";
                                                        oIncomingPayment.CheckAccount = oAcctCode;
                                                    }
                                                    if (iRowReceiptDtls.amount != 0)
                                                        oIncomingPayment.Checks.CheckSum = iRowReceiptDtls.amount;

                                                    oIncomingPayment.Checks.Add();
                                                }
                                                else if (oModeOfPayment == "Bank Transfer" || oModeOfPayment == "GIRO" || oModeOfPayment == "Paynow")
                                                {
                                                    oIncomingPayment.TransferReference = iRowReceiptDtls.reference;

                                                    if (!string.IsNullOrEmpty(oAcctCode))
                                                        oIncomingPayment.TransferAccount = oAcctCode;

                                                    if (oModeOfPayment == "GIRO" || oModeOfPayment == "Paynow")
                                                    {
                                                        oIncomingPayment.UserFields.Fields.Item("U_cash").Value = "GIRO";
                                                        oIncomingPayment.UserFields.Fields.Item("U_giro").Value = "Yes";
                                                    }

                                                    if (iRowReceiptDtls.amount != 0)
                                                        oIncomingPayment.TransferSum = iRowReceiptDtls.amount;

                                                }
                                                else if (oModeOfPayment == "CC")
                                                {
                                                    //string creditCardName = cls.GetSingleValue("SELECT \"CreditCard\" FROM OCRC WHERE \"CardName\" = '" + oIncomingPaymentLines.creditCardName + "'", company);
                                                    //if (creditCardName != "")
                                                    //{
                                                    //    oIncomingPayment.CreditCards.CreditCard = Convert.ToInt16(creditCardName);
                                                    //    oIncomingPayment.CreditCards.CardValidUntil = Convert.ToDateTime(oIncomingPaymentLines.creditCardValidDate);
                                                    //    oIncomingPayment.CreditCards.CreditCardNumber = oIncomingPaymentLines.creditCardNumber;

                                                    //    if (oIncomingPaymentLines.creditCardAmount != 0)
                                                    //        oIncomingPayment.CreditCards.CreditSum = oIncomingPaymentLines.creditCardAmount;

                                                    //    oIncomingPayment.CreditCards.VoucherNum = oIncomingPaymentLines.creditCardApproval;
                                                    //    oIncomingPayment.CreditCards.Add();
                                                    //}
                                                }
                                                else if (oModeOfPayment == "CN")
                                                {
                                                    string oDocDate = string.Empty;
                                                    string CNDesc = string.Empty;
                                                    if (!string.IsNullOrEmpty(iRowReceiptDtls.reference_id))
                                                    { }
                                                    else
                                                    {
                                                        string oVoucherTaxCode = (String)clsSBOGetRecord.GetSingleValue("select \"U_TaxCode\" from \"@PAYMENTCODES\" where \"U_PaymentCodeMethod\" = " + iRowReceiptDtls.method + "", sapCompany);

                                                        oCardName = (String)clsSBOGetRecord.GetSingleValue("select \"CardName\" from \"OCRD\" where \"CardCode\" = '" + TrimData(iRowReceipt.student) + "'", sapCompany);

                                                        CNDesc = oCardName + " Voucher " + Convert.ToDateTime(iRowReceipt.date_created).ToString("MMM") + " " + Convert.ToDateTime(iRowReceipt.date_created).Year + " " + iRowReceipt.level + " " + iRowReceipt.program_type;

                                                        Int16 CNDocEntry = CreateCreditNoteVoucher(oCardCode, iRowReceipt.receipt_no, iRowReceipt.date_created, CNDesc, oAcctCode, iRowReceiptDtls.amount, oVoucherTaxCode, iRowReceipt.invoice_no[0].ToString());
                                                        if (CNDocEntry != 0)
                                                        {
                                                            oIncomingPayment.Invoices.DocEntry = CNDocEntry;
                                                            oIncomingPayment.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;
                                                            oIncomingPayment.Invoices.Add();
                                                        }
                                                        else
                                                        {
                                                            lastMessage = "Credit Note (Voucher) with Transaction id:" + iRowReceipt.id + " and Receipt No:" + iRowReceipt.receipt_no + " does not exist in SAP B1.";
                                                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                                            functionReturnValue = true;

                                                            goto isAddWithError;
                                                        }
                                                    }
                                                }
                                                else if (oModeOfPayment == "NA")
                                                { }
                                                else
                                                { }
                                            }
                                        }
                                        ////**** Payment Means for the List of Invoices and Credit Note ****\\\\

                                        lErrCode = oIncomingPayment.Add();
                                        if (lErrCode == 0)
                                        {
                                            try
                                            {
                                                oDocEntry = sapCompany.GetNewObjectKey();
                                                lastMessage = "Successfully created Incoming Payment with Transaction Id:" + iRowReceipt.id + " in SAP B1.";
                                                sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Draft',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "',\"objType\" = 140 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                                functionReturnValue = false;
                                            }
                                            catch
                                            { }

                                            //try
                                            //{
                                            //	if (iRowReceipt.status == 1 && oIncomingPayment.GetByKey(Convert.ToInt32(oDocEntry)) == true)
                                            //	{
                                            //		lErrCode = oIncomingPayment.Cancel();
                                            //		if (lErrCode == 0)
                                            //		{
                                            //			try
                                            //			{
                                            //				lastMessage = "Successfully canceled Incoming Payment with Transaction Id:" + iRowReceipt.id + " in SAP B1.";
                                            //				sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Posted',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                            //				functionReturnValue = false;
                                            //			}
                                            //			catch (Exception ex)
                                            //			{ }
                                            //		}
                                            //		else
                                            //		{
                                            //			lastMessage = sapCompany.GetLastErrorDescription();
                                            //			sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                            //			functionReturnValue = true;
                                            //		}
                                            //		System.Runtime.InteropServices.Marshal.ReleaseComObject(oIncomingPayment);
                                            //	}
                                            //}
                                            //catch (Exception ex) 
                                            //{ }
                                        }
                                        else
                                        {
                                            lastMessage = sapCompany.GetLastErrorDescription();
                                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                            functionReturnValue = true;

                                            goto isAddWithError;
                                        }

                                    isAddWithError:;

                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oIncomingPayment);

                                    }
                                    else
                                    {
                                        oDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"OPDF\" where \"U_TransId\" = '" + iRowReceipt.id + "' and \"ObjType\" = 24", sapCompany);

                                        lastMessage = "Incoming Payment with Transaction Id:" + iRowReceipt.id + " is already existing in SAP B1.";
                                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'Draft',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "',\"objType\" = 140 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                        functionReturnValue = true;
                                    }
                                }
                                #region unused code
                                //else
                                //{                                  

                                //    oDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ORCT\" where \"U_TransId\" = '" + iRowReceipt.id + "' and \"Canceled\" = 'N'", sapCompany);
                                //    if (oDocEntry != "" && oDocEntry != "0") //**** Voiding of Incoming Payment when it is already existing SAP B1. ****\\
                                //    {
                                //        oIncomingPayment = (Payments)sapCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments);
                                //        if (oIncomingPayment.GetByKey(Convert.ToInt32(oDocEntry)) == true)
                                //        {
                                //            lErrCode = oIncomingPayment.Cancel();
                                //            if (lErrCode == 0)
                                //            {
                                //                try
                                //                {
                                //                    lastMessage = "Successfully canceled Incoming Payment with Transaction Id:" + iRowReceipt.id + " in SAP B1.";
                                //                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Posted',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                //                    functionReturnValue = false;
                                //                }
                                //                catch (Exception ex)
                                //                { }
                                //            }
                                //            else
                                //            {
                                //                lastMessage = sapCompany.GetLastErrorDescription();
                                //                sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                //                functionReturnValue = true;
                                //            }
                                //            System.Runtime.InteropServices.Marshal.ReleaseComObject(oIncomingPayment);
                                //        }
                                //    }
                                //    //else
                                //    //{
                                //    //    CreateReceiptVoid(listReceipt.Where(x => x.id == iRowReceipt.id).ToList());
                                //    //}

                                //    //else //**** Creation of Incoming Payment in SAP B1 before voiding the Incoming Payment. ****\\
                                //    //{
                                //    //	Int32 oDocEntryORCT = 0;
                                //    //	string oDocEntryOPDF = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ORCT\" where \"U_TransId\" = '" + iRowReceipt.id + "' and \"Canceled\" = 'N'", sapCompany);
                                //    //	if (oDocEntryOPDF == "" || oDocEntryOPDF == "0")
                                //    //	{
                                //    //		oDocEntryORCT = 
                                //    //		SAPbobsCOM.Payments oIncomingDraft = (Payments)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                                //    //		if (oDocEntryORCT != 0)
                                //    //		{
                                //    //			if (oIncomingDraft.GetByKey(oDocEntryORCT))
                                //    //			{
                                //    //				int ErrCode = oIncomingDraft.Update();
                                //    //				if (ErrCode == 0)
                                //    //				{
                                //    //					oDocEntry = sapCompany.GetNewObjectKey();
                                //    //					oIncomingPayment = (Payments)sapCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments);
                                //    //					if (oIncomingPayment.GetByKey(Convert.ToInt32(oDocEntry)) == true)
                                //    //					{
                                //    //						lErrCode = oIncomingPayment.Cancel();
                                //    //						if (lErrCode == 0)
                                //    //						{
                                //    //							try
                                //    //							{
                                //    //								oDocEntry = sapCompany.GetNewObjectKey();
                                //    //								lastMessage = "Successfully canceled Incoming Payment with Transaction Id:" + iRowReceipt.id + " in SAP B1.";
                                //    //								sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Posted',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "',\"objType\" = 24 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                //    //								functionReturnValue = false;
                                //    //							}
                                //    //							catch (Exception ex)
                                //    //							{ }
                                //    //						}
                                //    //						else
                                //    //						{
                                //    //							lastMessage = sapCompany.GetLastErrorDescription();
                                //    //							sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                //    //							functionReturnValue = true;
                                //    //						}
                                //    //						System.Runtime.InteropServices.Marshal.ReleaseComObject(oIncomingPayment);
                                //    //					}
                                //    //				}
                                //    //			}
                                //    //			System.Runtime.InteropServices.Marshal.ReleaseComObject(oIncomingDraft);
                                //    //		}
                                //    //	}
                                //    //}

                                //}
                                #endregion
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oIncomingPayment);
                            }
                            
                            else
                            {
                                //0 = no offset
                                //1 = has both payment and offset
                                //2 = only offset type
                                foreach (var iRowReceiptOffSetDtls in iRowReceipt.offset_references.ToList())
                                {
                                    string oDocEntryODPI = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ODPI\" where \"U_TransId\" = '" + iRowReceiptOffSetDtls.ToString() + "' and \"CANCELED\" = 'N' and \"ObjType\" = 203", sapCompany);

                                    if (String.IsNullOrEmpty(oDocEntryODPI) || oDocEntryODPI == "0") 
                                    { 
                                        continue; 
                                    }

                                    //if (iRowReceiptOffSetDtls.ToString() != "225948" && iRowReceiptOffSetDtls.ToString() != "222444")
                                    //{
                                    //    continue;
                                    //}

                                    var openDPs = clsSBOGetRecord.OpenDPEntries(iRowReceipt.student, sapCompany);
                                    if (openDPs.Count < 1)
                                    {
                                        continue;
                                    }

                                    string DpmAppl = (String)clsSBOGetRecord.GetSingleValue("select CAST(DpmAppl as nvarchar(max)) [DpmAppl] from \"ODPI\" where \"U_TransId\" = '" + iRowReceiptOffSetDtls.ToString() + "' and \"CANCELED\" = 'N' and \"ObjType\" = 203", sapCompany);
                                    string DpmAmnt = (String)clsSBOGetRecord.GetSingleValue("select CAST(DpmAmnt as nvarchar(max)) [DpmAmnt] from \"ODPI\" where \"U_TransId\" = '" + iRowReceiptOffSetDtls.ToString() + "' and \"CANCELED\" = 'N' and \"ObjType\" = 203", sapCompany);

                                    var remainingDPAmount = (Double.Parse(DpmAmnt) - Double.Parse(DpmAppl));
                                    if (remainingDPAmount <= 0) 
                                        continue;

                                    foreach (var offsetInvoiceNo in iRowReceipt.invoice_id.ToList())
                                    {
                                        //SAPbobsCOM.Documents _dpDoc = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oDownPayments);
                                        //_dpDoc.GetByKey(int.Parse(oDocEntryODPI));

                                        string oDocEntryOINV = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"OINV\" where \"U_TransId\" = '" + offsetInvoiceNo.ToString() + "' and \"DocStatus\" = 'O' and \"CANCELED\" = 'N' and \"ObjType\" = 13", sapCompany);
                                        if (!String.IsNullOrEmpty(oDocEntryOINV) && oDocEntryOINV != "0" && !String.IsNullOrEmpty(oDocEntryODPI) && oDocEntryODPI != "0")
                                        {

                                            SAPbobsCOM.Documents _invDoc = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oInvoices);
                                            _invDoc.GetByKey(int.Parse(oDocEntryOINV));

                                            var newInv = new API_Invoice()
                                            {
                                                id = int.Parse(_invDoc.UserFields.Fields.Item("U_TransId").Value?.ToString() ?? "0"),
                                                invoice_no = _invDoc.NumAtCard,
                                                date_created = _invDoc.DocDate.ToString(),
                                                date_due = _invDoc.DocDueDate.ToString(),
                                                status = 1,
                                                remarks = _invDoc.Comments,
                                                void_remarks = _invDoc.Comments,
                                                student = _invDoc.CardCode,
                                                level = _invDoc.UserFields.Fields.Item("U_Level").Value?.ToString(),
                                                program_type = _invDoc.UserFields.Fields.Item("U_ProgramType").Value?.ToString(),
                                                //downPaymentAmount = float.Parse(remainingDPAmount.ToString()),
                                                //downPaymentDocEntry = _dpDoc.DocEntry,
                                                OpenDPs = openDPs,
                                                isAppliedDP = true,
                                                oldInvDocEntry = int.Parse(oDocEntryOINV),
                                            };
                                      
                                            newInv.items = new List<API_InvoiceDetails>();
                                            for(int i = 0; i < _invDoc.Lines.Count; i++)
                                            {
                                                _invDoc.Lines.SetCurrentLine(i);
                                                newInv.items.Add(new API_InvoiceDetails()
                                                {
                                                    description = _invDoc.Lines.ItemDescription,
                                                    item_code = _invDoc.Lines.ItemCode,
                                                    date_for =  _invDoc.DocDate.ToString(),
                                                    unit_price = float.Parse(_invDoc.Lines.LineTotal.ToString()),
                                                    quantity = float.Parse(_invDoc.Lines.Quantity.ToString()),
                                                    total = float.Parse(_invDoc.Lines.Quantity.ToString())
                                                });
                                            }
                                            var invoices = new List<API_Invoice>();
                                            invoices.Add(newInv);
                                            SBOPostInvoice(invoices, DateTime.Now.Date.ToString());


                                            
                                        }
                                    }
                                }
                               
                            }
                        }
                        catch (Exception ex)
                        {
                            lastMessage = ex.ToString();
                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");
                            functionReturnValue = true;
                        }

                    }

                    CreateReceiptVoid(listReceipt.Where(x => x.status != 0).ToList());
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return functionReturnValue;
        }

        public Int16 CreateCreditNoteVoucher(string student, string receipt_no, string create_date, string description, string account, float amount, string vatGroup, string invoice_no)
        {
            string functionReturnValue = string.Empty;
            string lastErrorMessage = string.Empty;
            string oDocEntry = string.Empty;
            int lErrCode;
            SBOGetRecord clsSBOGetValue = new SBOGetRecord();
            try
            {
                oDocEntry = (String)clsSBOGetValue.GetSingleValue("select \"DocEntry\" from \"ORIN\" where \"U_ReceiptNo\" = '" + receipt_no + "' and \"CANCELED\" = 'N' and \"U_CreatedByVoucher\" = 1", sapCompany);
                if (oDocEntry == "" || oDocEntry == "0")
                {
                    oCreditNote = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oCreditNotes);

                    oCreditNote.CardCode = student;
                    oCreditNote.NumAtCard = receipt_no;
                    oCreditNote.DocDate = Convert.ToDateTime(create_date);
                    oCreditNote.DocType = BoDocumentTypes.dDocument_Service;
                    oCreditNote.Comments = "Invoice No:" + invoice_no + ", Receipt No:" + receipt_no;

                    oCreditNote.UserFields.Fields.Item("U_ReceiptNo").Value = receipt_no;
                    oCreditNote.UserFields.Fields.Item("U_CreatedByVoucher").Value = 1;

                    if (!string.IsNullOrEmpty(oCountry))
                        oCreditNote.Lines.CostingCode = oCountry;

                    if (!string.IsNullOrEmpty(oGroup))
                        oCreditNote.Lines.CostingCode2 = oGroup;

                    if (!string.IsNullOrEmpty(oDivision))
                        oCreditNote.Lines.CostingCode3 = oDivision;

                    if (!string.IsNullOrEmpty(oProduct))
                        oCreditNote.Lines.CostingCode4 = oProduct;

                    oCreditNote.Lines.AccountCode = account;
                    oCreditNote.Lines.UserFields.Fields.Item("U_Dscription").Value = description;

                    string Dscription = string.Empty;
                    if (description.Length > 100)
                    {
                        Dscription = description.Substring(0, 100);
                        oCreditNote.Lines.ItemDescription = Dscription;
                    }
                    else
                    {
                        oCreditNote.Lines.ItemDescription = description;
                    }

                    oCreditNote.Lines.VatGroup = vatGroup;
                    oCreditNote.Lines.LineTotal = amount / 1.07;

                    lErrCode = oCreditNote.Add();
                    if (lErrCode == 0)
                    {
                        try
                        {
                            functionReturnValue = sapCompany.GetNewObjectKey();
                        }
                        catch
                        { }
                    }
                    else
                    {
                        lastErrorMessage = sapCompany.GetLastErrorDescription();
                        functionReturnValue = "0";
                    }
                }
                else
                    functionReturnValue = oDocEntry;

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return Convert.ToInt16(functionReturnValue);
        }

        public Int16 CreateCreditRefundVoid(List<API_CreditRefund> listCreditRefund)
        {
            string functionReturnValue = "";
            int lErrCode = 0;
            string oLogExist = string.Empty;
            string oTransId = string.Empty;
            string oCardCode = string.Empty;
            string oDocEntry = string.Empty;
            string oCreditNoteDocEntry = string.Empty;
            string oInvoiceDocEntry = string.Empty;
            string oCountryCod = string.Empty;
            string oModeOfPayment = string.Empty;
            string oAcctCode = string.Empty;
            string oBankCode = string.Empty;
            string oBankName = string.Empty;
            string oBranch = string.Empty;
            string oCheckAccount = string.Empty;
            string oCheckNumber = string.Empty;
            string iReference = string.Empty;
            SBOGetRecord clsSBOGetRecord = new SBOGetRecord();
            try
            {
                foreach (var iRowCreditRefund in listCreditRefund)
                {
                    try
                    {
                        oOutgoingPayment = (Payments)sapCompany.GetBusinessObject(BoObjectTypes.oPaymentsDrafts);
                        if (iRowCreditRefund.status == 1)
                        {
                            oTransId = (String)clsSBOGetRecord.GetSingleValue("select \"U_TransId\" from \"OPDF\" where \"U_TransId\" = '" + iRowCreditRefund.id + "' and \"ObjType\" = 46", sapCompany);
                            oOutgoingPayment.DocObjectCode = BoPaymentsObjectType.bopot_OutgoingPayments;
                            if (oTransId == "" || oTransId == "0")
                            {
                                oCardCode = (String)clsSBOGetRecord.GetSingleValue("select \"CardCode\" from \"OCRD\" where \"CardCode\" = '" + TrimData(iRowCreditRefund.student) + "'", sapCompany);
                                if (oCardCode != "")
                                {
                                    oOutgoingPayment.CardCode = oCardCode;
                                }
                                else
                                {
                                    lastMessage = "Customer Code:" + iRowCreditRefund.student + " is not found in SAP B1";
                                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditRefund.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Refund' and \"uniqueId\" = '" + iRowCreditRefund.id + "'");

                                    functionReturnValue = "0";

                                    goto isAddWithError;
                                }

                                oOutgoingPayment.DocType = BoRcptTypes.rCustomer;
                                oOutgoingPayment.DocDate = Convert.ToDateTime(iRowCreditRefund.date_created);

                                if (iRowCreditRefund.payment_reference != "N.A" && !string.IsNullOrEmpty(iRowCreditRefund.payment_reference))
                                    oOutgoingPayment.JournalRemarks = iRowCreditRefund.payment_reference;

                                ////**** UDF ****\\\\
                                oOutgoingPayment.UserFields.Fields.Item("U_TransId").Value = iRowCreditRefund.id.ToString();
                                oOutgoingPayment.UserFields.Fields.Item("U_Status").Value = iRowCreditRefund.status.ToString();
                                oOutgoingPayment.UserFields.Fields.Item("U_CreditType").Value = iRowCreditRefund.credit_type;
                                ////**** UDF ****\\\\

                                if (iRowCreditRefund.status == 0)
                                    oOutgoingPayment.Remarks = iRowCreditRefund.remarks;
                                else
                                    oOutgoingPayment.Remarks = iRowCreditRefund.void_remarks;


                                if (iRowCreditRefund.overpaid_offsets == 0)
                                {
                                    //** credit id **//
                                    oCreditNoteDocEntry = string.Empty;
                                    oCreditNoteDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ORIN\" where \"U_TransId\" = '" + iRowCreditRefund.credit_id + "' and \"CANCELED\" = 'N'", sapCompany);
                                    if (oCreditNoteDocEntry != "" && oCreditNoteDocEntry != "0")
                                    {
                                        oOutgoingPayment.Invoices.DocEntry = Convert.ToInt16(oCreditNoteDocEntry);
                                        oOutgoingPayment.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;
                                        oOutgoingPayment.Invoices.Add();
                                    }
                                    else
                                    {
                                        string oDraftDocEntry = string.Empty;
                                        oDraftDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ODRF\" where \"U_TransId\" = '" + iRowCreditRefund.credit_id + "' and \"CANCELED\" = 'N'", sapCompany);
                                        if (oDraftDocEntry != "" && oDraftDocEntry != "0")
                                        {
                                            SAPbobsCOM.Documents oDraft = (Documents)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                            if (oDraft.GetByKey(Convert.ToInt16(oDraftDocEntry)))
                                            {
                                                int ErrCode = oDraft.SaveDraftToDocument();
                                                if (ErrCode == 0)
                                                {
                                                    oCreditNoteDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ORIN\" where \"U_TransId\" = '" + iRowCreditRefund.credit_id + "' and \"CANCELED\" = 'N'", sapCompany);
                                                    if (oCreditNoteDocEntry != "" && oCreditNoteDocEntry != "0")
                                                    {
                                                        oOutgoingPayment.Invoices.DocEntry = Convert.ToInt16(oCreditNoteDocEntry);
                                                        oOutgoingPayment.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;
                                                        oOutgoingPayment.Invoices.Add();
                                                    }
                                                }
                                                else
                                                {
                                                    lastMessage = sapCompany.GetLastErrorDescription();
                                                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditRefund.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Refund' and \"uniqueId\" = '" + iRowCreditRefund.id + "'");

                                                    functionReturnValue = "0";

                                                    goto isAddWithError;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            lastMessage = "Credit Note (Draft) with Reference Id:" + iRowCreditRefund.credit_id + " does not exist in SAP B1.";
                                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditRefund.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Refund' and \"uniqueId\" = '" + iRowCreditRefund.id + "'");

                                            functionReturnValue = "0";

                                            goto isAddWithError;

                                        }
                                    }
                                    //** credit id **//
                                }
                                else
                                {
                                    //**overpaid_offsets_credit_notes**//
                                    foreach (var iRowCreditRefundCN in iRowCreditRefund.overpaid_offsets_credit_notes.ToList())
                                    {
                                        oCreditNoteDocEntry = string.Empty;
                                        oCreditNoteDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ORIN\" where \"U_TransId\" = '" + iRowCreditRefundCN.ToString() + "' and \"CANCELED\" = 'N'", sapCompany);
                                        if (oCreditNoteDocEntry != "" && oCreditNoteDocEntry != "0")
                                        {
                                            oOutgoingPayment.Invoices.DocEntry = Convert.ToInt16(oCreditNoteDocEntry);
                                            oOutgoingPayment.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;
                                            oOutgoingPayment.Invoices.Add();
                                        }
                                        else
                                        {
                                            string oDraftDocEntry = string.Empty;
                                            oDraftDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ODRF\" where \"U_TransId\" = '" + iRowCreditRefundCN.ToString() + "' and \"CANCELED\" = 'N' and \"ObjType\" = 14", sapCompany);
                                            if (oDraftDocEntry != "" && oDraftDocEntry != "0")
                                            {
                                                SAPbobsCOM.Documents oDraft = (Documents)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                                if (oDraft.GetByKey(Convert.ToInt16(oDraftDocEntry)))
                                                {
                                                    int ErrCode = oDraft.SaveDraftToDocument();
                                                    if (ErrCode == 0)
                                                    {
                                                        oCreditNoteDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ORIN\" where \"U_TransId\" = '" + iRowCreditRefundCN.ToString() + "' and \"CANCELED\" = 'N'", sapCompany);
                                                        if (oCreditNoteDocEntry != "" && oCreditNoteDocEntry != "0")
                                                        {
                                                            oOutgoingPayment.Invoices.DocEntry = Convert.ToInt16(oCreditNoteDocEntry);
                                                            oOutgoingPayment.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;
                                                            oOutgoingPayment.Invoices.Add();
                                                        }
                                                    }
                                                    else
                                                    {
                                                        lastMessage = sapCompany.GetLastErrorDescription();
                                                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditRefund.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Refund' and \"uniqueId\" = '" + iRowCreditRefund.id + "'");

                                                        functionReturnValue = "0";

                                                        goto isAddWithError;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                lastMessage = "Credit Note (Draft) with Reference Id:" + iRowCreditRefundCN.ToString() + " does not exist in SAP B1.";
                                                sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditRefund.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Refund' and \"uniqueId\" = '" + iRowCreditRefund.id + "'");

                                                functionReturnValue = "0";

                                                goto isAddWithError;

                                            }
                                        }
                                    }
                                    //**overpaid_offsets_credit_notes**//
                                }

                                //**overpaid_offsets_invoices**//
                                foreach (var iRowCreditRefundInv in iRowCreditRefund.overpaid_offsets_invoices.ToList())
                                {
                                    oInvoiceDocEntry = string.Empty;
                                    oInvoiceDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"OINV\" where \"U_TransId\" = '" + iRowCreditRefundInv.ToString() + "' and \"CANCELED\" = 'N'", sapCompany);
                                    if (oInvoiceDocEntry != "" && oInvoiceDocEntry != "0")
                                    {
                                        oOutgoingPayment.Invoices.DocEntry = Convert.ToInt16(oInvoiceDocEntry);
                                        oOutgoingPayment.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                                        oOutgoingPayment.Invoices.Add();
                                    }
                                    else
                                    {
                                        string oDraftDocEntry = string.Empty;
                                        oDraftDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ODRF\" where \"U_TransId\" = '" + iRowCreditRefundInv.ToString() + "' and \"CANCELED\" = 'N' and \"ObjType\" = 13", sapCompany);
                                        if (oDraftDocEntry != "" && oDraftDocEntry != "0")
                                        {
                                            SAPbobsCOM.Documents oDraft = (Documents)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                            if (oDraft.GetByKey(Convert.ToInt16(oDraftDocEntry)))
                                            {
                                                int ErrCode = oDraft.SaveDraftToDocument();
                                                if (ErrCode == 0)
                                                {
                                                    oInvoiceDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"OINV\" where \"U_TransId\" = '" + iRowCreditRefundInv.ToString() + "' and \"CANCELED\" = 'N'", sapCompany);
                                                    if (oInvoiceDocEntry != "" && oInvoiceDocEntry != "0")
                                                    {
                                                        oOutgoingPayment.Invoices.DocEntry = Convert.ToInt16(oCreditNoteDocEntry);
                                                        oOutgoingPayment.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                                                        oOutgoingPayment.Invoices.Add();
                                                    }
                                                }
                                                else
                                                {
                                                    lastMessage = sapCompany.GetLastErrorDescription();
                                                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditRefund.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Refund' and \"uniqueId\" = '" + iRowCreditRefund.id + "'");

                                                    functionReturnValue = "0";

                                                    goto isAddWithError;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            lastMessage = "Invoice (Draft) with Reference Id:" + iRowCreditRefundInv.ToString() + " does not exist in SAP B1.";
                                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditRefund.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Refund' and \"uniqueId\" = '" + iRowCreditRefund.id + "'");

                                            functionReturnValue = "0";

                                            goto isAddWithError;

                                        }
                                    }
                                }
                                //**overpaid_offsets_invoices**//


                                ////**** Payment Means for the List of Credit Note ****\\\\
                                oAcctCode = (String)clsSBOGetRecord.GetSingleValue("select \"U_GLAccount\" from \"@PAYMENTCODES\" where \"U_PaymentCodeMethod\" = '" + iRowCreditRefund.payment_method + "'", sapCompany);
                                oModeOfPayment = (String)clsSBOGetRecord.GetSingleValue("select \"U_ModePayment\" from \"@PAYMENTCODES\" where \"U_PaymentCodeMethod\" = '" + iRowCreditRefund.payment_method + "'", sapCompany);

                                if (oModeOfPayment == "CA")
                                {
                                    if (!string.IsNullOrEmpty(oAcctCode))
                                        oOutgoingPayment.CashAccount = oAcctCode;

                                    if (iRowCreditRefund.amount != 0)
                                        oOutgoingPayment.CashSum = iRowCreditRefund.amount;
                                }
                                else if (oModeOfPayment == "CK")
                                {
                                    if (!string.IsNullOrEmpty(iRowCreditRefund.payment_reference))
                                    {
                                        oBankName = iRowCreditRefund.payment_reference.Substring(0, iRowCreditRefund.payment_reference.IndexOf(' '));

                                        oCheckNumber = iRowCreditRefund.payment_reference.Replace(oBankName, "");

                                        oBankCode = (String)clsSBOGetRecord.GetSingleValue("select \"BankCode\" from \"ODSC\" where \"BankName\" = '" + TrimData(oBankName) + "'", sapCompany);

                                        if (string.IsNullOrEmpty(oBankCode))
                                        {
                                            lastMessage = "Bank:" + oBankName + " is not found in SAP B1";
                                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditRefund.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Refund' and \"uniqueId\" = '" + iRowCreditRefund.id + "'");

                                            functionReturnValue = "0";

                                            goto isAddWithError;
                                        }

                                        oCountryCod = (String)clsSBOGetRecord.GetSingleValue("select \"Country\" from \"DSC1\" where \"BankCode\" = '" + TrimData(oBankCode) + "'", sapCompany);

                                        oAcctCode = (String)clsSBOGetRecord.GetSingleValue("select \"GLAccount\" from \"DSC1\" where \"BankCode\" = '" + TrimData(oBankCode) + "'", sapCompany);

                                        oCheckAccount = (String)clsSBOGetRecord.GetSingleValue("select \"Account\" from \"DSC1\" where \"BankCode\" = '" + TrimData(oBankCode) + "'", sapCompany);

                                        oBranch = (String)clsSBOGetRecord.GetSingleValue("select \"Branch\" from \"DSC1\" where \"BankCode\" = '" + TrimData(oBankCode) + "'", sapCompany);
                                    }

                                    if (!string.IsNullOrEmpty(oCountryCod))
                                        oOutgoingPayment.Checks.CountryCode = oCountryCod;

                                    if (!string.IsNullOrEmpty(oBankCode))
                                        oOutgoingPayment.Checks.BankCode = oBankCode;

                                    if (!string.IsNullOrEmpty(oAcctCode))
                                        oOutgoingPayment.Checks.CheckAccount = oAcctCode;

                                    if (!string.IsNullOrEmpty(oCheckAccount))
                                        oOutgoingPayment.Checks.AccounttNum = oCheckAccount;

                                    if (!string.IsNullOrEmpty(oBranch))
                                        oOutgoingPayment.Checks.Branch = oBranch;

                                    if (!string.IsNullOrEmpty(oCheckNumber))
                                        oOutgoingPayment.Checks.CheckNumber = Convert.ToInt32(TrimData(oCheckNumber));

                                    if (iRowCreditRefund.amount != 0)
                                        oOutgoingPayment.Checks.CheckSum = iRowCreditRefund.amount;

                                    oOutgoingPayment.Checks.Add();
                                }
                                else if (oModeOfPayment == "BT")
                                {
                                    if (iReference != "N.A" && iReference != "")
                                        oOutgoingPayment.TransferReference = iReference;

                                    if (string.IsNullOrEmpty(oAcctCode))
                                        oOutgoingPayment.TransferAccount = oAcctCode;

                                    if (iRowCreditRefund.amount != 0)
                                        oOutgoingPayment.TransferSum = iRowCreditRefund.amount;
                                }
                                else if (oModeOfPayment == "CC")
                                {
                                    //string creditCardName = cls.GetSingleValue("SELECT \"CreditCard\" FROM OCRC WHERE \"CardName\" = '" + oIncomingPaymentLines.creditCardName + "'", company);
                                    //if (creditCardName != "")
                                    //{
                                    //    oIncomingPayment.CreditCards.CreditCard = Convert.ToInt16(creditCardName);
                                    //    oIncomingPayment.CreditCards.CardValidUntil = Convert.ToDateTime(oIncomingPaymentLines.creditCardValidDate);
                                    //    oIncomingPayment.CreditCards.CreditCardNumber = oIncomingPaymentLines.creditCardNumber;

                                    //    if (oIncomingPaymentLines.creditCardAmount != 0)
                                    //        oIncomingPayment.CreditCards.CreditSum = oIncomingPaymentLines.creditCardAmount;

                                    //    oIncomingPayment.CreditCards.VoucherNum = oIncomingPaymentLines.creditCardApproval;
                                    //    oIncomingPayment.CreditCards.Add();
                                    //}
                                }
                                else if (oModeOfPayment == "CN")
                                { }
                                else if (oModeOfPayment == "NA")
                                { }
                                else
                                { }

                                ////**** Payment Means for the List of Invoices ****\\\\

                                lErrCode = oOutgoingPayment.Add();
                                if (lErrCode == 0)
                                {
                                    try
                                    {
                                        oDocEntry = sapCompany.GetNewObjectKey();
                                        lastMessage = "Successfully created Outgoing Payment (Draft) with Transaction Id:" + iRowCreditRefund.id + " in SAP B1.";
                                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Draft',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "',\"objType\" = 46 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Refund' and \"uniqueId\" = '" + iRowCreditRefund.id + "'");

                                        functionReturnValue = oDocEntry;
                                    }
                                    catch
                                    { }
                                }
                                else
                                {
                                    lastMessage = sapCompany.GetLastErrorDescription();
                                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditRefund.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Refund' and \"uniqueId\" = '" + iRowCreditRefund.id + "'");

                                    functionReturnValue = "0";

                                    goto isAddWithError;
                                }

                            isAddWithError:;

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oOutgoingPayment);

                            }
                            else
                            {
                                oDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"OPDF\" where \"U_TransId\" = '" + iRowCreditRefund.id + "' and \"ObjType\" = 46", sapCompany);

                                lastMessage = "Outgoing Payment with Transaction Id:" + iRowCreditRefund.id + " is already existing in SAP B1.";
                                sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Posted',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "',\"objType\" = 46 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Refund' and \"uniqueId\" = '" + iRowCreditRefund.id + "'");

                                functionReturnValue = "0";
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        lastMessage = ex.ToString();
                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowCreditRefund.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Credit Refund' and \"uniqueId\" = '" + iRowCreditRefund.id + "'");
                        functionReturnValue = "0";
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return Convert.ToInt16(functionReturnValue);
        }

        public Int32 CreateInvoiceVoid(List<API_Invoice> olistInvoice)
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

                //olistInvoice = olistInvoice.Where(x => DateTime.Parse(x.date_created) < DateTime.Parse("06/01/2022")).ToList();
                olistInvoice = olistInvoice.Where(x => DateTime.Parse(x.date_created) >= DateTime.Parse("09/01/2022") && DateTime.Parse(x.date_created) <= DateTime.Parse("09/30/2022")).ToList();
                foreach (var oRowInv in olistInvoice)
                {
                    try
                    {
                        if (oRowInv.status == 2)
                        {
                            bool hasItemCode = true;

                            string _checkIfExists = "select \"DocEntry\" from \"OINV\" where \"U_TransId\" = '" + oRowInv.id + "' and \"CANCELED\" = 'N' and \"NumAtCard\" = '" + oRowInv.invoice_no + "'";
                            var invoiceID = (String)clsSBOGetRecord.GetSingleValue(_checkIfExists, sapCompany);

                            if (String.IsNullOrEmpty(invoiceID) || invoiceID == "0")
                            {
                                continue;
                            }

                            //iDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ODRF\" where \"U_TransId\" = '" + oRowInv.id + "' and \"CANCELED\" = 'N' and \"ObjType\" = 13", sapCompany);
                            string Query = "select \"U_TransId\" from \"ORIN\" where \"U_TransId\" = '" + oRowInv.id + "' and \"CANCELED\" = 'N' and \"NumAtCard\" = '" + oRowInv.invoice_no + "'";
                            iDocEntry = (String)clsSBOGetRecord.GetSingleValue(Query, sapCompany);
                            if (iDocEntry == "" || iDocEntry == "0")
                            {

                                var Invoice = (Documents)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                                Invoice.GetByKey(int.Parse(invoiceID));
                                var CancelInvoice = Invoice.CreateCancellationDocument();
                                CancelInvoice.Comments = oRowInv.void_remarks;
                                if (CancelInvoice.Add() == 0) continue;
                               

                                //if (DateTime.Parse(oRowInv.date_created) >= DateTime.Parse("07/01/2022") && DateTime.Parse(oRowInv.date_created) <= DateTime.Parse("07/31/2022"))
                                //{
                                //    oRowInv.date_created = "07/07/2022";
                                //    oRowInv.date_due = "07/07/2022";
                                //}
                                //else if (DateTime.Parse(oRowInv.date_created) >= DateTime.Parse("08/01/2022") && DateTime.Parse(oRowInv.date_created) <= DateTime.Parse("08/31/2022"))
                                //{
                                //    oRowInv.date_created = "08/01/2022";
                                //    oRowInv.date_due = "08/01/2022";
                                //}
                                ////else if (DateTime.Parse(oRowInv.date_created) >= DateTime.Parse("06/01/2022") && DateTime.Parse(oRowInv.date_created) <= DateTime.Parse("06/30/2022"))
                                ////{
                                ////    oRowInv.date_created = "07/06/2022";
                                ////    oRowInv.date_due = "07/06/2022";
                                ////}
                                //else
                                //{
                                //    continue;
                                //}

                                oInvoice = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oCreditNotes);
                                oInvoice.DocObjectCode = BoObjectTypes.oCreditNotes;

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
                                string seriesNum = (String)clsSBOGetRecord.GetSingleValue("select TOP 1 Series from \"NNM1\" \"e\"  where \"e\".SeriesName like 'JEC%' AND BPLId = 5 AND Indicator = YEAR(GETDATE()) AND ObjectCode = '14'", sapCompany);

                                var _invoiceIds = String.Join(",", oRowInv.invoice_no);
                                string _earliestInvDate = (String)clsSBOGetRecord.GetSingleValue("SELECT TOP 1 DocDate FROM OINV WHERE \"U_TransId\" IN ('" + _invoiceIds + "') ORDER BY DocDate ASC", sapCompany);
                                string _latestInvDate = (String)clsSBOGetRecord.GetSingleValue("SELECT TOP 1 DocDate FROM OINV WHERE \"U_TransId\" IN ('" + _invoiceIds + "') ORDER BY DocDate DESC", sapCompany);
                                //oRowInv.date_created = "07/02/2022";
                                //oRowInv.date_due = "07/02/2022";
                                

                                oInvoice.BPL_IDAssignedToInvoice = 5; //"Jamiyah Education Centre (JEC)";
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

                                oInvoice.UserFields.Fields.Item("U_branch").Value = "Jamiyah Education Centre (JEC)";
                                ////**** UDF *****/////

                                foreach (var oRowInvDtls in oRowInv.items.ToList())
                                {
                                    oRowInvDtls.item_code = (String)clsSBOGetRecord.GetSingleValue("select \"U_ccode\" from \"@JEC\" where \"U_descript\" = '" + TrimData(oRowInvDtls.description) + "' and \"U_unitprice\" = '" + oRowInvDtls.unit_price + "'", sapCompany);
                                    if (String.IsNullOrEmpty(oRowInvDtls.item_code))
                                    {
                                        hasItemCode = false;
                                    }

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
                                            var _glAcct = (String)clsSBOGetRecord.GetSingleValue("SELECT  \"ARCMAct\" FROM \"OITB\" T1 JOIN \"OITM\" T2 ON T2.\"ItmsGrpCod\" = T1.\"ItmsGrpCod\" WHERE T2.\"ItemCode\" = '" + TrimData(oRowInvDtls.item_code) + "'", sapCompany);

                                            oInvoice.Lines.ItemCode = oRowInvDtls.item_code;
                                            oInvoice.Lines.UnitsOfMeasurment = 1;
                                            oInvoice.Lines.VatGroup = "SR";
                                            oInvoice.Lines.AccountCode = _glAcct;

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

                                if (!hasItemCode)
                                {
                                    continue;
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

                            isAddWithError:;

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
                return Convert.ToInt32(ifunctionReturnValue);
            }
        }

        public Int32 CreateReceiptVoid(List<API_Receipt> listReceipt)
        {
            bool functionReturnValue = false;
            int lErrCode = 0;
            string oLogExist = string.Empty;
            string oTransId = string.Empty;
            string oCardCode = string.Empty;
            string oCardName = string.Empty;
            string oDocEntry = string.Empty;
            string oInvDocEntry = string.Empty;
            string oCreditNoteDocEntry = string.Empty;
            string oModeOfPayment = string.Empty;
            string oAcctCode = string.Empty;
            string oBankName = string.Empty;
            string oCheckBankName = string.Empty;
            string iReference = string.Empty;

            SBOGetRecord clsSBOGetRecord = new SBOGetRecord();
            try
            {
                //listReceipt = listReceipt.Where(x => DateTime.Parse(x.date_created) < DateTime.Parse("06/01/2022")).ToList();
                foreach (var iRowReceipt in listReceipt)
                {
                    try
                    {
                        oIncomingPayment = (Payments)sapCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments);
                        oDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ORCT\" where \"U_TransId\" = '" + iRowReceipt.id + "' and \"Canceled\" = 'N'", sapCompany);
                        if (oIncomingPayment.GetByKey(Convert.ToInt32(oDocEntry)) == true)
                        {
                            oIncomingPayment.Remarks = "[Revert] - " + (iRowReceipt.remarks.Length >= 200 ? oCardName.Substring(0, 50) + " " + iRowReceipt.remarks : oCardName + " " + iRowReceipt.remarks);                        
                            lErrCode = oIncomingPayment.Cancel();
                            if (lErrCode == 0)
                            {
                                try
                                {
                                    lastMessage = "Successfully canceled Incoming Payment with Transaction Id:" + iRowReceipt.id + " in SAP B1.";
                                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Posted',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                    functionReturnValue = false;
                                }
                                catch (Exception ex)
                                { }
                            }
                            else
                            {
                                lastMessage = sapCompany.GetLastErrorDescription();
                                sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                                functionReturnValue = true;
                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oIncomingPayment);
                        }
                    }
                    catch (Exception ex)
                    { }
                }
                return Convert.ToInt16(functionReturnValue);
                #region Old Code for Create Void Receipt
                //foreach (var iRowReceipt in listReceipt)
                //{
                //    try
                //    {
                //        //0 = no offset
                //        //1 = has both payment and offset
                //        //2 = only offset type

                //        if (iRowReceipt.payment_type == 0 || iRowReceipt.payment_type == 1)
                //        {
                //            oIncomingPayment = (Payments)sapCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments);
                //            if (iRowReceipt.status == 1)
                //            {
                //                oTransId = (String)clsSBOGetRecord.GetSingleValue("select \"U_TransId\" from \"ORCT\" where \"U_TransId\" = '" + iRowReceipt.id + "' and \"Canceled\" = 'N'", sapCompany);
                //                oIncomingPayment.DocObjectCode = BoPaymentsObjectType.bopot_IncomingPayments;
                //                if (oTransId == "" || oTransId == "0")
                //                {
                //                    oCardCode = (String)clsSBOGetRecord.GetSingleValue("select \"CardCode\" from \"OCRD\" where \"CardCode\" = '" + TrimData(iRowReceipt.student) + "'", sapCompany);
                //                    if (oCardCode != "")
                //                    {
                //                        oIncomingPayment.CardCode = oCardCode;

                //                        oCardName = (String)clsSBOGetRecord.GetSingleValue("select \"CardName\" from \"OCRD\" where \"CardCode\" = '" + TrimData(iRowReceipt.student) + "'", sapCompany);
                //                    }
                //                    else
                //                    {
                //                        lastMessage = "Customer Code:" + iRowReceipt.student + " is not found in SAP B1";
                //                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                //                        functionReturnValue = true;

                //                        goto isAddWithError;
                //                    }
                //                    var _invoiceIds = String.Join(",", iRowReceipt.invoice_id);
                //                    string _earliestInvDate = (String)clsSBOGetRecord.GetSingleValue("SELECT TOP 1 DocDate FROM OINV WHERE \"U_TransId\" IN ('" + _invoiceIds + "') ORDER BY DocDate ASC", sapCompany);
                //                    string _latestInvDate = (String)clsSBOGetRecord.GetSingleValue("SELECT TOP 1 DocDate FROM OINV WHERE \"U_TransId\" IN ('" + _invoiceIds + "') ORDER BY DocDate DESC", sapCompany);

                //                    oIncomingPayment.BPLID = 5; //"Jamiyah Education Centre (JEC)";
                //                    oIncomingPayment.DocTypte = BoRcptTypes.rCustomer;
                //                    //oIncomingPayment.DocDate = Convert.ToDateTime(iRowReceipt.date_created);
                //                    oIncomingPayment.DocDate = Convert.ToDateTime("07/03/2022");
                //                    //oIncomingPayment.DocDate = (!String.IsNullOrEmpty(_earliestInvDate) && Convert.ToDateTime(_earliestInvDate) > Convert.ToDateTime(iRowReceipt.date_created)
                //                    //					? Convert.ToDateTime(_earliestInvDate) : Convert.ToDateTime(iRowReceipt.date_created));
                //                    string seriesNum = (String)clsSBOGetRecord.GetSingleValue("select TOP 1 Series from \"NNM1\" \"e\"  where \"e\".SeriesName like '%JEC%' AND BPLId = 5 AND Indicator = YEAR(GETDATE()) AND ObjectCode = '24'", sapCompany);
                //                    oIncomingPayment.Series = int.Parse(seriesNum);
                //                    ////**** UDF ****\\\\     
                //                    oIncomingPayment.UserFields.Fields.Item("U_TransId").Value = iRowReceipt.id.ToString();
                //                    oIncomingPayment.UserFields.Fields.Item("U_StatusTaidii").Value = iRowReceipt.status.ToString();
                //                    oIncomingPayment.UserFields.Fields.Item("U_tax").Value = "N/A";
                //                    oIncomingPayment.UserFields.Fields.Item("U_ipc").Value = "NON-IPC";
                //                    //oIncomingPayment.UserFields.Fields.Item("U_Level").Value = iRowReceipt.level;
                //                    //oIncomingPayment.UserFields.Fields.Item("U_ProgramType").Value = iRowReceipt.program_type;
                //                    oIncomingPayment.UserFields.Fields.Item("U_ReceiptNo").Value = iRowReceipt.receipt_no;
                //                    oIncomingPayment.UserFields.Fields.Item("U_branch").Value = "Jamiyah Education Centre (JEC)";
                //                    ////**** UDF ****\\\\

                //                    if (iRowReceipt.status == 0)

                //                        if (iRowReceipt.remarks.Length >= 200)
                //                        {
                //                            oIncomingPayment.Remarks = oCardName.Substring(0, 50) + " " + iRowReceipt.remarks;
                //                        }
                //                        else
                //                            oIncomingPayment.Remarks = oCardName + " " + iRowReceipt.remarks;
                //                    else
                //                    {
                //                        if (iRowReceipt.void_remarks.Length >= 200)
                //                        {
                //                            oIncomingPayment.Remarks = "[REVERT] - " + oCardName.Substring(0, 50) + " " + iRowReceipt.void_remarks;
                //                        }
                //                        else
                //                            oIncomingPayment.Remarks = "[REVERT] - " + oCardName + " " + iRowReceipt.void_remarks;
                //                    }

                //                    ////**** Adding of List of Invoice to Incoming Payment ****\\\\
                //                    int invoiceCount = 0;
                //                    int invPaidCount;
                //                    foreach (var iRowReceiptInvDtl in iRowReceipt.invoice_id.ToList())
                //                    {
                //                        invoiceCount += 1;
                //                        oInvDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"OINV\" " + Environment.NewLine +
                //                        "where \"U_TransId\" = '" + iRowReceiptInvDtl.ToString() + "' and \"CANCELED\" = 'N'", sapCompany);
                //                        if (oInvDocEntry != "" && oInvDocEntry != "0")
                //                        {
                //                            invPaidCount = 0;
                //                            foreach (var iRowReceiptInvPaidDtl in iRowReceipt.invoice_paid.ToList())
                //                            {
                //                                invPaidCount += 1;
                //                                if (invoiceCount == invPaidCount)
                //                                {
                //                                    oIncomingPayment.Invoices.DocEntry = Convert.ToInt32(oInvDocEntry);
                //                                    oIncomingPayment.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                //                                    oIncomingPayment.Invoices.SumApplied = Convert.ToDouble(iRowReceiptInvPaidDtl.ToString());
                //                                    oIncomingPayment.Invoices.Add();
                //                                }
                //                            }
                //                        }
                //                        else
                //                        {
                //                            lastMessage = "Invoice with Transaction id:" + iRowReceiptInvDtl.ToString() + " does not exist in SAP B1.";
                //                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                //                            functionReturnValue = true;

                //                            goto isAddWithError;
                //                        }
                //                    }
                //                    ////**** Adding of List of Invoice to Incoming Payment ****\\\\

                //                    ////**** Adding of List of Credit Note to Incoming Payment ****\\\\
                //                    iReference = string.Empty;
                //                    foreach (var iRowReceiptInvDtls in iRowReceipt.payment_methods.ToList())
                //                    {
                //                        if (iRowReceiptInvDtls.method == 3 || iRowReceiptInvDtls.method == 8 || iRowReceiptInvDtls.method == 10) //**OFFSET_DEPOSIT = 3**\\
                //                        {
                //                            if (!string.IsNullOrEmpty(iRowReceiptInvDtls.reference_id) && iRowReceiptInvDtls.reference_id != "N.A")
                //                            {
                //                                oCreditNoteDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ORIN\" where \"U_TransId\" = '" + TrimData(iRowReceiptInvDtls.reference_id) + "' and \"CANCELED\" = 'N' and \"U_CreatedByVoucher\" = 0", sapCompany);
                //                                if (oCreditNoteDocEntry != "" && oCreditNoteDocEntry != "0")
                //                                {
                //                                    oIncomingPayment.Invoices.DocEntry = Convert.ToInt16(oCreditNoteDocEntry);
                //                                    oIncomingPayment.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;
                //                                    oIncomingPayment.Invoices.Add();
                //                                }
                //                                else
                //                                {
                //                                    string oDraftDocEntry = string.Empty;
                //                                    oDraftDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ODRF\" where \"U_TransId\" = '" + iRowReceiptInvDtls.reference_id + "' and \"CANCELED\" = 'N' and \"ObjType\" = 14", sapCompany);
                //                                    if (oDraftDocEntry != "" && oDraftDocEntry != "0")
                //                                    {
                //                                        SAPbobsCOM.Documents oDraft = (Documents)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                //                                        if (oDraft.GetByKey(Convert.ToInt16(oDraftDocEntry)))
                //                                        {
                //                                            int ErrCode = oDraft.SaveDraftToDocument();
                //                                            if (ErrCode == 0)
                //                                            {
                //                                                oCreditNoteDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"ORIN\" where \"U_TransId\" = '" + iRowReceiptInvDtls.reference_id + "' and \"CANCELED\" = 'N' and \"U_CreatedByVoucher\" = 0", sapCompany);
                //                                                if (oCreditNoteDocEntry != "" && oCreditNoteDocEntry != "0")
                //                                                {
                //                                                    oIncomingPayment.Invoices.DocEntry = Convert.ToInt16(oCreditNoteDocEntry);
                //                                                    oIncomingPayment.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;
                //                                                    oIncomingPayment.Invoices.Add();
                //                                                }
                //                                            }
                //                                            else
                //                                            {
                //                                                lastMessage = sapCompany.GetLastErrorDescription();
                //                                                sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                //                                                functionReturnValue = true;

                //                                                goto isAddWithError;
                //                                            }
                //                                        }
                //                                    }
                //                                    else
                //                                    {
                //                                        lastMessage = "Credit Note with Reference Id:" + iRowReceiptInvDtls.reference_id + " does not exist in SAP B1 Drafts";
                //                                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                //                                        functionReturnValue = true;

                //                                        goto isAddWithError;
                //                                    }
                //                                }
                //                            }
                //                            else
                //                            {
                //                                if (iRowReceiptInvDtls.reference != "N.A")
                //                                    iReference += iRowReceiptInvDtls.reference + ", ";
                //                            }
                //                        }
                //                        else
                //                        {
                //                            if (iRowReceiptInvDtls.reference != "N.A")
                //                                iReference += iRowReceiptInvDtls.reference + ", ";
                //                        }
                //                    }
                //                    ////**** Adding of List of Credit Note to Incoming Payment ****\\\\

                //                    string oJournalRemarks = string.Empty;
                //                    if (!string.IsNullOrEmpty(iReference))
                //                    {
                //                        oJournalRemarks = iReference.Substring(0, iReference.Length - 2);
                //                    }

                //                    if (!string.IsNullOrEmpty(oJournalRemarks))
                //                        oIncomingPayment.JournalRemarks = oJournalRemarks;

                //                    ////**** Payment Means for the List of Invoices ****\\\\
                //                    foreach (var iRowReceiptDtls in iRowReceipt.payment_methods.ToList())
                //                    {
                //                        if (string.IsNullOrEmpty(iRowReceiptDtls.reference_id) || iRowReceiptDtls.reference_id == "N.A")
                //                        {
                //                            oAcctCode = (String)clsSBOGetRecord.GetSingleValue("select \"U_GLAccount\" from \"@PAYMENTCODES\" where \"U_PaymentCodeMethod\" = " + iRowReceiptDtls.method + "", sapCompany);

                //                            oModeOfPayment = (String)clsSBOGetRecord.GetSingleValue("select \"U_ModePayment\" from \"@PAYMENTCODES\" where \"U_PaymentCodeMethod\" = " + iRowReceiptDtls.method + "", sapCompany);

                //                            oIncomingPayment.UserFields.Fields.Item("U_type").Value = "Local";
                //                            oIncomingPayment.UserFields.Fields.Item("U_giro").Value = "N/A";

                //                            if (oModeOfPayment == "Cash" || oModeOfPayment == "NETS")
                //                            {
                //                                if (!string.IsNullOrEmpty(oAcctCode))
                //                                {
                //                                    oIncomingPayment.UserFields.Fields.Item("U_cash").Value = oModeOfPayment;
                //                                    oIncomingPayment.CashAccount = oAcctCode;
                //                                }
                //                                if (iRowReceiptDtls.amount != 0)
                //                                    oIncomingPayment.CashSum = iRowReceiptDtls.amount;
                //                            }
                //                            else if (oModeOfPayment == "Check")
                //                            {
                //                                if (!string.IsNullOrEmpty(oAcctCode))
                //                                {
                //                                    oIncomingPayment.UserFields.Fields.Item("U_cash").Value = "CHQ";
                //                                    oIncomingPayment.CheckAccount = oAcctCode;
                //                                }
                //                                if (iRowReceiptDtls.amount != 0)
                //                                    oIncomingPayment.Checks.CheckSum = iRowReceiptDtls.amount;

                //                                oIncomingPayment.Checks.Add();
                //                            }
                //                            else if (oModeOfPayment == "Bank Transfer" || oModeOfPayment == "GIRO" || oModeOfPayment == "Paynow")
                //                            {
                //                                oIncomingPayment.TransferReference = iRowReceiptDtls.reference;

                //                                if (!string.IsNullOrEmpty(oAcctCode))
                //                                    oIncomingPayment.TransferAccount = oAcctCode;

                //                                if (oModeOfPayment == "GIRO" || oModeOfPayment == "Paynow")
                //                                {
                //                                    oIncomingPayment.UserFields.Fields.Item("U_cash").Value = "GIRO";
                //                                    oIncomingPayment.UserFields.Fields.Item("U_giro").Value = "Yes";
                //                                }

                //                                if (iRowReceiptDtls.amount != 0)
                //                                    oIncomingPayment.TransferSum = iRowReceiptDtls.amount;

                //                            }
                //                            else if (oModeOfPayment == "CC")
                //                            {
                //                                //string creditCardName = cls.GetSingleValue("SELECT \"CreditCard\" FROM OCRC WHERE \"CardName\" = '" + oIncomingPaymentLines.creditCardName + "'", company);
                //                                //if (creditCardName != "")
                //                                //{
                //                                //    oIncomingPayment.CreditCards.CreditCard = Convert.ToInt16(creditCardName);
                //                                //    oIncomingPayment.CreditCards.CardValidUntil = Convert.ToDateTime(oIncomingPaymentLines.creditCardValidDate);
                //                                //    oIncomingPayment.CreditCards.CreditCardNumber = oIncomingPaymentLines.creditCardNumber;

                //                                //    if (oIncomingPaymentLines.creditCardAmount != 0)
                //                                //        oIncomingPayment.CreditCards.CreditSum = oIncomingPaymentLines.creditCardAmount;

                //                                //    oIncomingPayment.CreditCards.VoucherNum = oIncomingPaymentLines.creditCardApproval;
                //                                //    oIncomingPayment.CreditCards.Add();
                //                                //}
                //                            }
                //                            else if (oModeOfPayment == "CN")
                //                            {
                //                                string oDocDate = string.Empty;
                //                                string CNDesc = string.Empty;
                //                                if (!string.IsNullOrEmpty(iRowReceiptDtls.reference_id))
                //                                { }
                //                                else
                //                                {
                //                                    string oVoucherTaxCode = (String)clsSBOGetRecord.GetSingleValue("select \"U_TaxCode\" from \"@PAYMENTCODES\" where \"U_PaymentCodeMethod\" = " + iRowReceiptDtls.method + "", sapCompany);

                //                                    oCardName = (String)clsSBOGetRecord.GetSingleValue("select \"CardName\" from \"OCRD\" where \"CardCode\" = '" + TrimData(iRowReceipt.student) + "'", sapCompany);

                //                                    CNDesc = oCardName + " Voucher " + Convert.ToDateTime(iRowReceipt.date_created).ToString("MMM") + " " + Convert.ToDateTime(iRowReceipt.date_created).Year + " " + iRowReceipt.level + " " + iRowReceipt.program_type;

                //                                    Int16 CNDocEntry = CreateCreditNoteVoucher(oCardCode, iRowReceipt.receipt_no, iRowReceipt.date_created, CNDesc, oAcctCode, iRowReceiptDtls.amount, oVoucherTaxCode, iRowReceipt.invoice_no[0].ToString());
                //                                    if (CNDocEntry != 0)
                //                                    {
                //                                        oIncomingPayment.Invoices.DocEntry = CNDocEntry;
                //                                        oIncomingPayment.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;
                //                                        oIncomingPayment.Invoices.Add();
                //                                    }
                //                                    else
                //                                    {
                //                                        lastMessage = "Credit Note (Voucher) with Transaction id:" + iRowReceipt.id + " and Receipt No:" + iRowReceipt.receipt_no + " does not exist in SAP B1.";
                //                                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                //                                        functionReturnValue = true;

                //                                        goto isAddWithError;
                //                                    }
                //                                }
                //                            }
                //                            else if (oModeOfPayment == "NA")
                //                            { }
                //                            else
                //                            { }
                //                        }
                //                    }
                //                    ////**** Payment Means for the List of Invoices and Credit Note ****\\\\

                //                    lErrCode = oIncomingPayment.Add();
                //                    if (lErrCode == 0)
                //                    {
                //                        try
                //                        {
                //                            oDocEntry = sapCompany.GetNewObjectKey();
                //                            lastMessage = "Successfully created Incoming Payment with Transaction Id:" + iRowReceipt.id + " in SAP B1.";
                //                            sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Draft',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "',\"objType\" = 140 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                //                            functionReturnValue = false;
                //                        }
                //                        catch
                //                        { }

                //                        try
                //                        {
                //                            if (iRowReceipt.status == 1 && oIncomingPayment.GetByKey(Convert.ToInt32(oDocEntry)) == true)
                //                            {
                //                                lErrCode = oIncomingPayment.Cancel();
                //                                if (lErrCode == 0)
                //                                {
                //                                    try
                //                                    {
                //                                        lastMessage = "Successfully canceled Incoming Payment with Transaction Id:" + iRowReceipt.id + " in SAP B1.";
                //                                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Posted',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                //                                        functionReturnValue = false;
                //                                    }
                //                                    catch (Exception ex)
                //                                    { }
                //                                }
                //                                else
                //                                {
                //                                    lastMessage = sapCompany.GetLastErrorDescription();
                //                                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                //                                    functionReturnValue = true;
                //                                }
                //                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oIncomingPayment);
                //                            }
                //                        }
                //                        catch (Exception ex)
                //                        { }
                //                    }
                //                    else
                //                    {
                //                        lastMessage = sapCompany.GetLastErrorDescription();
                //                        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                //                        functionReturnValue = true;

                //                        goto isAddWithError;
                //                    }

                //                isAddWithError:;

                //                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oIncomingPayment);

                //                }
                //                else
                //                {
                //                    oDocEntry = (String)clsSBOGetRecord.GetSingleValue("select \"DocEntry\" from \"OPDF\" where \"U_TransId\" = '" + iRowReceipt.id + "' and \"ObjType\" = 24", sapCompany);

                //                    lastMessage = "Incoming Payment with Transaction Id:" + iRowReceipt.id + " is already existing in SAP B1.";
                //                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oDocEntry + "',\"objType\" = 24 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");

                //                    functionReturnValue = false;
                //                }
                //            }
                //        }
                //    }
                //    catch (Exception ex)
                //    {
                //        lastMessage = ex.ToString();
                //        sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = '" + iif(iRowReceipt.status == 0, "Draft", "Void") + "',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + iRowReceipt.id + "'");
                //        functionReturnValue = false;
                //    }
                //}
                #endregion
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return Convert.ToInt16(functionReturnValue);
        }

        public void UpdateCancelledInvoiceForDP(int newInvDocEntry, int oldInvDocEntry)
        {
            sapRecSet.DoQuery($"Update {SBOConstantClass.Database}..OINV SET U_NewDocEntry={newInvDocEntry} WHERE DocEntry={oldInvDocEntry}");
        }
        public bool CancelInvoiceForDP(int newInvDocEntry, int oldInvDocEntry)
        {
            try
            {
                SBOGetRecord clsSBOGetRecord = new SBOGetRecord();
                SAPbobsCOM.Documents _invDoc = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oInvoices);
                _invDoc.GetByKey(oldInvDocEntry);

                var receiptNo = clsSBOGetRecord.GetSingleValue("select DocEntry from OINV WHERE DocEntry = '" + oldInvDocEntry + "'", sapCompany);



                var CancelInvoice = _invDoc.CreateCancellationDocument();
                if (CancelInvoice == null) return true;
                CancelInvoice.Comments = $"[REVERT] - This invoice has linked with downpayment.";
                int ErrCode = CancelInvoice.Add();
                if (ErrCode == 0)
                {
                    lastMessage = "Successfully updated the invoice with Transaction Id:" + oldInvDocEntry + " in SAP B1.";
                    sapRecSet.DoQuery("update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Posted',\"failDesc\" = '',\"successDesc\" = '" + TrimData(lastMessage) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + oldInvDocEntry + "',\"objType\" = 14 where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Receipt' and \"uniqueId\" = '" + oldInvDocEntry + "'");
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {

            }
            return false;
        }

        public string ItemMasterData(string oDate = "")
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
                    ItemModel = new List<API_FinanceItem>();
                    pricelistcode = 13;

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
                    "left join \"NNM1\" \"d\" on \"d\".\"Series\" = \"a\".\"Series\" " + Environment.NewLine +
                    "where \"c\".\"PriceList\" = " + pricelistcode + Environment.NewLine +
                    "and \"d\".\"SeriesName\" like 'JEC%'" + Environment.NewLine +
                    "and (YEAR(\"a\".\"UpdateDate\") = YEAR(GETDATE())" + Environment.NewLine +
                    "OR YEAR(\"a\".\"CreateDate\") = YEAR(GETDATE()))" + Environment.NewLine;
                    //+ "and MONTH(\"a\".\"UpdateDate\") = MONTH(GETDATE())" + Environment.NewLine  
                    //"or \"c\".\"PriceList\" = " + pricelistcode + " and \"a\".\"UpdateDate\" = '" + oDate + "' and \"a\".\"ItmsGrpCod\" = " + oItmsGrpCod + "";
                    sapRecSet.DoQuery(oQuery);

                    ItemMasterModel = new List<API_FinanceItem>();

                    if (sapRecSet.RecordCount > 0)
                    {
                        ////** Parse Business Partners **//////
                        XDocument xItemMasterData = XDocument.Parse(sapRecSet.GetAsXML());
                        if (xItemMasterData.Root != null)
                        {
                            ItemModel = xItemMasterData.Descendants("row").Select(oItemMaster =>
                            new API_FinanceItem
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

                            ItemMasterModel.Add(new API_FinanceItem()
                            {
                                item_code = iRowItems.item_code,
                                description = iRowItems.description,
                                type = iRowItems.type,
                                unit_price = Convert.ToDouble(iRowItems.unit_price),
                                remarks = iRowItems.remarks,
                                tax = iRowItems.tax
                            });

                            var strJson = JsonConvert.SerializeObject(ItemMasterModel);


                            oLogExist = (String)clsSBOGetRecord.GetSingleValue("select * from " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " where \"companyDB\" = '" + TrimData(sapCompany.CompanyDB) + "' and \"module\" = 'Product' and \"uniqueId\" = '" + TrimData(iRowItems.item_code) + "' ", sapCompany);

                            if (oLogExist == "" || oLogExist == "0")
                            {
                                Console.WriteLine("Adding Product:" + iRowItems.item_code + " in the integration log. Please wait...");
                                strQuery = "insert into " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " (\"lastTimeStamp\",\"companyDB\",\"module\",\"uniqueId\",\"docStatus\",\"status\",\"JSON\",\"statusCode\",\"successDesc\",\"failDesc\",\"logDate\",\"objType\") select '" + oDate + "','" + TrimData(sapCompany.CompanyDB) + "','Product','" + TrimData(iRowItems.item_code) + "','Confirmed','','" + TrimData(strJson) + "','','','',null,4" + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "", " from dummy;") + "";
                                sapRecSet.DoQuery(strQuery);
                            }
                            else
                            {
                                Console.WriteLine("Updating Product:" + iRowItems.item_code + " in the integration log. Please wait...");
                                strQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"JSON\" = '" + TrimData(strJson) + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Product' and \"uniqueId\" = '" + TrimData(iRowItems.item_code) + "'";
                                sapRecSet.DoQuery(strQuery);
                            }

                            if (ItemModel.Count > 0)
                                Console.WriteLine("Processing Product:" + iRowItems.item_code + " in TAIDII. Please wait...");

                            //string BaseUrl = string.Empty;
                            //string MethodUrl = string.Empty;
                            //string JSONResult = string.Empty;
                            //string oResponseResult = string.Empty;

                            ////Set Base URL Address for API Call
                            //BaseUrl = "https://dev-new.taidii.com/api/open/sap/"; // base_url;

                            ////Set Method for the API Call
                            //MethodUrl = "centeritem/create/";

                            //HttpClient httpClient = new HttpClient();
                            //httpClient.BaseAddress = new Uri(BaseUrl);

                            //HttpContent content = new FormUrlEncodedContent(
                            //    new List<KeyValuePair<string, string>> {
                            //    new KeyValuePair<string, string>("api_key", "piRUbJ7d4AoXlH1TADBO"),
                            //    new KeyValuePair<string,string>("client","jmdev"),
                            //    new KeyValuePair<string,string>("items",strJSON)
                            //});

                            //HttpResponseMessage Response = httpClient.PostAsync(MethodUrl, content).Result;
                            //if (Response.IsSuccessStatusCode)
                            //{
                            //    oResponseResult = Response.Content.ReadAsStringAsync().Result;
                            //    if (oResponseResult.Contains("id") == true)
                            //    {
                            //        listResponseResultSuccess = Newtonsoft.Json.JsonConvert.DeserializeObject<List<ResponseResultSuccess>>(oResponseResult);

                            //        foreach (var iRowSuccess in listResponseResultSuccess)
                            //        {
                            //            if (iRowSuccess.status == 1)
                            //            {
                            //                lastMessage = "Successfully " + iif(iRowSuccess.log == "new", "created new Item", "updated existing item") + " in TAIDII Portal.";
                            //                strQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'true',\"statusCode\" = 'Posted',\"failDesc\" = '',\"successDesc\" = '" + lastMessage + "',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + iRowItems.item_code + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Product' and \"uniqueId\" = '" + iRowItems.item_code + "'";
                            //                sapRecSet.DoQuery(strQuery);

                            //                functionReturnValue = false;
                            //            }
                            //        }
                            //    }
                            //    else
                            //    {
                            //        listResponseResultFailed = Newtonsoft.Json.JsonConvert.DeserializeObject<List<ResponseResultFailed>>(oResponseResult);
                            //        foreach (var iRowFailed in listResponseResultFailed)
                            //        {
                            //            if (iRowFailed.status == 0)
                            //            {
                            //                int errCnt = iRowFailed.errors.Count;
                            //                int counter = 0;
                            //                lastMessage = string.Empty;
                            //                foreach (var iRowFailedDtl in iRowFailed.errors.ToList())
                            //                {
                            //                    if (counter == 0 && errCnt != counter)
                            //                    {
                            //                        lastMessage += iRowFailedDtl.ToString() + ", ";
                            //                    }
                            //                    else
                            //                        lastMessage += iRowFailedDtl.ToString() + ", ";

                            //                    counter += 1;
                            //                }

                            //                if (lastMessage.Length > 0)
                            //                    lastMessage = lastMessage.Substring(0, lastMessage.Length - 2);

                            //                strQuery = "update " + iif(SBOConstantClass.ServerVersion != "dst_HANADB", "\"TAIDII_SAP\"..\"axxis_tb_IntegrationLog\"", "\"TAIDII_SAP\".\"axxis_tb_IntegrationLog\"") + " set \"status\" = 'false',\"statusCode\" = 'For Process',\"failDesc\" = '" + TrimData(lastMessage) + "',\"successDesc\" = '',\"logDate\" = '" + sapCompany.GetDBServerDate().ToString("yyyy-MM-dd") + "',\"sapCode\" = '" + iRowItems.item_code + "' where \"companyDB\" = '" + TrimData(SBOConstantClass.Database) + "' and \"module\" = 'Product' and \"uniqueId\" = '" + iRowItems.item_code + "'";
                            //                sapRecSet.DoQuery(strQuery);

                            //                functionReturnValue = false;
                            //            }
                            //        }
                            //    }
                            //}
                            //else
                            //{
                            //    oResponseResult = Response.Content.ReadAsStringAsync().Result;
                            //}
                        }
                        ////**** Create a list of Products ****////
                        //Console.WriteLine("Done adding the List of " + string.Format("{0:#,##0}", sapRecSet.RecordCount) + " Product(s) in the integration log. Please wait...");
                    }
                }

                return JsonConvert.SerializeObject(ItemMasterModel);
            }
            catch (Exception ex)
            {
                return "";
            }
        }


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

        public object iif(bool expression, object truePart, object falsePart)
        { return expression ? truePart : falsePart; }

        public object TrimData(string oValue)
        { return oValue.Replace("'", "''"); }

        public bool ReversePostedInvoice()
        {




            return true;
        }
    }
}