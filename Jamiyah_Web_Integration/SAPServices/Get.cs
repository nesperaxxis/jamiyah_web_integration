using Jamiyah_Web_Integration.SAPModels;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Jamiyah_Web_Integration.SAPServices
{
    public class SBOGetRecord
    {
        public string GetSingleValue(string StrQuery, Company SAPCompany)
        {
            try
            {
                Company company = SAPCompany;
                Recordset oRecSet = default(Recordset);
                oRecSet = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecSet.DoQuery(StrQuery);
                return Convert.ToString(oRecSet.Fields.Item(0).Value);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<TaidiiInvoicesDocEntriesResult> TaidiiInvoicesDocEntries(string _from, string _to, Company SAPCompany)
        {
            try
            {
                List<TaidiiInvoicesDocEntriesResult> _ids = new List<TaidiiInvoicesDocEntriesResult>();
                Company company = SAPCompany;
                Recordset oRecSet = default(Recordset);
                oRecSet = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                //oRecSet.DoQuery("SELECT U_TransId, DocEntry FROM OINV WHERE U_TransId IS NOT NULL AND DocStatus = 'O' AND DocDate BETWEEN '" + _from + "' AND '" + _to + "'");
                oRecSet.DoQuery("SELECT U_TransId, DocEntry FROM OINV WHERE U_TransId IS NOT NULL AND DocStatus = 'O' AND DocDate >= '" + _from + "'");
                while (!oRecSet.EoF)
                {
                    _ids.Add(new TaidiiInvoicesDocEntriesResult 
                    { 
                        TransId = int.Parse(Convert.ToString(oRecSet.Fields.Item(0).Value)), 
                        DocEntry = int.Parse(Convert.ToString(oRecSet.Fields.Item(1).Value))
                    });
                    oRecSet.MoveNext();
                }
                
                return _ids;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<TaidiiInvoicesDocEntriesResult> TaidiiCreditNotesDocEntries(string _from, string _to, Company SAPCompany)
        {
            try
            {
                List<TaidiiInvoicesDocEntriesResult> _ids = new List<TaidiiInvoicesDocEntriesResult>();
                Company company = SAPCompany;
                Recordset oRecSet = default(Recordset);
                oRecSet = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                //oRecSet.DoQuery("SELECT U_TransId, DocEntry FROM OINV WHERE U_TransId IS NOT NULL AND DocStatus = 'O' AND DocDate BETWEEN '" + _from + "' AND '" + _to + "'");
                oRecSet.DoQuery("SELECT U_TransId, DocEntry FROM ORIN WHERE U_TransId IS NOT NULL AND DocStatus = 'O' AND DocDate >= '" + _from + "'");
                while (!oRecSet.EoF)
                {
                    _ids.Add(new TaidiiInvoicesDocEntriesResult
                    {
                        TransId = int.Parse(Convert.ToString(oRecSet.Fields.Item(0).Value)),
                        DocEntry = int.Parse(Convert.ToString(oRecSet.Fields.Item(1).Value))
                    });
                    oRecSet.MoveNext();
                }

                return _ids;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<OpenDPEntries> OpenDPEntries(string cardCode, Company SAPCompany)
        {
            try
            {
                List<OpenDPEntries> entries  = new List<OpenDPEntries>();
                Company company = SAPCompany;
                Recordset oRecSet = default(Recordset);
                oRecSet = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                //oRecSet.DoQuery("SELECT U_TransId, DocEntry FROM OINV WHERE U_TransId IS NOT NULL AND DocStatus = 'O' AND DocDate BETWEEN '" + _from + "' AND '" + _to + "'");
                oRecSet.DoQuery("SELECT DocEntry, CAST(DpmAmnt  - DpmAppl as nvarchar(max)) [DpmAppl] from ODPI where CardCode = '" + cardCode + "' AND U_TransId IS NOT NULL AND (DpmAmnt  - DpmAppl) > 0 ORDER BY DocDate ASC");
                while (!oRecSet.EoF)
                {
                    entries.Add(new OpenDPEntries
                    {
                        DocEntry = Convert.ToString(oRecSet.Fields.Item(0).Value),
                        Ammount = Convert.ToString(oRecSet.Fields.Item(1).Value)
                    });
                    oRecSet.MoveNext();
                }

                return entries;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
    public class TaidiiInvoicesDocEntriesResult
    {
        public int DocEntry { get; set; }
        public int TransId { get; set; }
    }
     
    public class OpenDPEntries
    {
        public string DocEntry { get; set; }
        public string Ammount { get; set; }

    }
}