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
    }
}