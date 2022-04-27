using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using SAPbobsCOM;
namespace SBOHelper.Class
{
    public class SBOGetRecord
    {
        public string GetSingleValue(string StrQuery,SAPbobsCOM.Company SAPCompany)
        {
            try
            {
                SAPbobsCOM.Company company = SAPCompany;
                SAPbobsCOM.Recordset oRecSet = default(SAPbobsCOM.Recordset);
                oRecSet = (Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
