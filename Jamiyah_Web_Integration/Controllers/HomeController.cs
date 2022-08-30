using Jamiyah_Web_Integration.SAPModels;
using Jamiyah_Web_Integration.SAPServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace Jamiyah_Web_Integration.Controllers
{
    //[Authorize]
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public bool SyncStudent(string jsonData)
        {
            bool isSuccess = false;
            List<API_BusinessPartners> lstBusinessPartners = Newtonsoft.Json.JsonConvert.DeserializeObject<List<API_BusinessPartners>>(jsonData);
            if (lstBusinessPartners.Count > 0)
            {
                clsStart clsStart = new clsStart();
                clsStart.GetStarted();

                Posting clsSBOClass = new Posting();
                isSuccess = clsSBOClass.SBOPostBusinessPartners(lstBusinessPartners, DateTime.Now.Date.ToString());
            }
            return isSuccess;
        }
        
        public bool SyncInvoices(string jsonData)
        {
            bool isSuccess = false;
            List<API_Invoice> lstInvoices = Newtonsoft.Json.JsonConvert.DeserializeObject<List<API_Invoice>>(jsonData);
            if (lstInvoices.Count > 0)
            {
                clsStart clsStart = new clsStart();
                clsStart.GetStarted();

                Posting clsSBOClass = new Posting();
                isSuccess = clsSBOClass.SBOPostInvoice(lstInvoices, DateTime.Now.Date.ToString());
            }
            return isSuccess;
        }
        public bool SyncDownpayment(string jsonData)
        {
            bool isSuccess = false;
            List<API_CreditNote> lstdp = Newtonsoft.Json.JsonConvert.DeserializeObject<List<API_CreditNote>>(jsonData);
            if (lstdp.Count > 0)
            {
                clsStart clsStart = new clsStart();
                clsStart.GetStarted();

                Posting clsSBOClass = new Posting();
                isSuccess = clsSBOClass.SBOPostDownpayment(lstdp, DateTime.Now.Date.ToString());
            }
            return isSuccess;
        }

        [HttpGet]
        public string SyncSAPItems()
        {
            string returnData = "";
            clsStart clsStart = new clsStart();
            clsStart.GetStarted();

            Posting clsSBOClass = new Posting();
            returnData = clsSBOClass.ItemMasterData(DateTime.Now.Date.ToString());
            return returnData;
        }
        public bool SyncPayments(string jsonData)
        {
            bool isSuccess = false;
            List<API_Receipt> lstInvoices = Newtonsoft.Json.JsonConvert.DeserializeObject<List<API_Receipt>>(jsonData);
            if (lstInvoices.Count > 0)
            {
                clsStart clsStart = new clsStart();
                clsStart.GetStarted();

                Posting clsSBOClass = new Posting();
                isSuccess = clsSBOClass.SBOPostReceipt(lstInvoices);
            }
            return isSuccess;
        }
    }
}
