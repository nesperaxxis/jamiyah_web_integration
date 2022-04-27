using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Jamiyah_Web_Integration.SAPModels
{
    public class API_BusinessPartners
    {
        public int id { get; set; }
        public string BPMaster { get; set; }
        public string fullname { get; set; }
        public string nric { get; set; }
        public int gender { get; set; }
        public string dob { get; set; }
        public string student_care_type { get; set; }
        public string program_type { get; set; }
        public string registration_no { get; set; }
        public float subsidy { get; set; }
        public float additional_subsidy { get; set; }
        public float financial_assistance { get; set; }
        public float deposit { get; set; }
        public string nationality { get; set; }
        public string race { get; set; }
        public string address { get; set; }
        public string unit_no { get; set; }
        public string postal_code { get; set; }
        public string date_of_withdrawal { get; set; }
        public string country { get; set; }
        public string contact_name { get; set; }
        public string contact_nric { get; set; }
        public string contact_relation { get; set; }
        public string contact_email { get; set; }
        public string contact_telephone { get; set; }
        public string contact_office_no { get; set; }
        public string contact_home_phone { get; set; }
        public string bank_name { get; set; }
        public string account_name { get; set; }
        public string cdac_bank_no { get; set; }
        public string customer_ref_no { get; set; }
        public string admission_date { get; set; }
        public string level { get; set; }
    }

    //**Invoice List**//
    public class API_Invoice
    {
        public int id { get; set; }
        public string invoice_no { get; set; }
        public string date_created { get; set; }
        public string date_due { get; set; }
        public int status { get; set; }
        public string remarks { get; set; }
        public string void_remarks { get; set; }
        public string student { get; set; }
        public string level { get; set; }
        public string program_type { get; set; }
        public virtual List<API_InvoiceDetails> items { get; set; }
    }

    public class API_InvoiceDetails
    {
        public string description { get; set; }
        public string item_code { get; set; }
        public string date_for { get; set; }
        public float unit_price { get; set; }
        public float quantity { get; set; }
        public float total { get; set; }
    }

    //**Credit Note List**//
    public class API_CreditNote
    {
        public int id { get; set; }
        public string credit_no { get; set; }
        public int credit_type { get; set; }
        public string overpaid_receipt { get; set; }
        public string student { get; set; }
        public string date_created { get; set; }
        public int status { get; set; }
        public string remarks { get; set; }
        public string void_remarks { get; set; }
        public int type { get; set; }
        public string level { get; set; }
        public string program_type { get; set; }
        public int payment_method { get; set; }
        public virtual List<API_CreditNoteDetails> items { get; set; }
    }

    public class API_CreditNoteDetails
    {
        public string description { get; set; }
        public string date_for { get; set; }
        public float amount { get; set; }
        public float gst { get; set; }
    }

    //**Receipt List**//
    public class API_Receipt
    {
        public int id { get; set; }
        public string receipt_no { get; set; }
        public string student { get; set; }
        public string level { get; set; }
        public string program_type { get; set; }
        public virtual List<int> invoice_id { get; set; }
        public virtual List<string> invoice_no { get; set; }
        public virtual List<float> invoice_paid { get; set; }
        public int payment_type { get; set; }
        public string date_created { get; set; }
        public int status { get; set; }
        public string remarks { get; set; }
        public string void_remarks { get; set; }
        public virtual List<int> offset_references { get; set; }
        public virtual List<API_ReceiptDetails> payment_methods { get; set; }
    }

    public class API_ReceiptDetails
    {
        public int method { get; set; }
        public string reference { get; set; }
        public string reference_id { get; set; }
        public float amount { get; set; }
    }

    //**Credit List**//
    public class API_CreditRefund
    {
        public int id { get; set; }
        public int credit_id { get; set; }
        public int credit_type { get; set; }
        public string student { get; set; }
        public int status { get; set; }
        public string date_created { get; set; }
        public string remarks { get; set; }
        public string void_remarks { get; set; }
        public int payment_method { get; set; }
        public string payment_reference { get; set; }
        public float amount { get; set; }
        public int overpaid_offsets { get; set; }
        public int overpaid_offsets_receipt_id { get; set; }
        public virtual List<int> overpaid_offsets_credit_notes { get; set; }
        public virtual List<int> overpaid_offsets_invoices { get; set; }
    }

    //**Create Finance Item List**//
    public class API_FinanceItem
    {
        public string item_code { get; set; }
        public string description { get; set; }
        public string type { get; set; }
        public double unit_price { get; set; }
        public string remarks { get; set; }
        public int tax { get; set; }
    }

    //**Response**//
    public class SAPResponse
    {
        public string RecordStatus { get; set; }
        public string Message { get; set; }
    }
    public class ErrorResponse
    {
        public string Message { get; set; }
    }

    public class ResponseResultSuccess
    {
        public int status { get; set; }
        public int id { get; set; }
        public string item_code { get; set; }
        public string log { get; set; }

        //"status": 1,
        //"id": 5326,
        //"item_code": "135933",
        //"log": "updated"
    }

    public class ResponseResultFailed
    {
        public int status { get; set; }
        public string item_code { get; set; }
        public string description { get; set; }
        public float unit_price { get; set; }
        public string remarks { get; set; }
        public string type { get; set; }
        public int tax { get; set; }
        public virtual List<string> errors { get; set; }

        //"status": 0,
        //"errors": [
        //    "Missing item_code"
        //],
        //"description": "Database Configuration",
        //"tax": 0,
        //"unit_price": 670,
        //"item_code": "",
        //"remarks": "",
        //"type": "Others"
    }

    //**Response**//
}