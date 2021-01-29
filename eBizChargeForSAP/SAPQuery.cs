using System;
using System.Windows;
using System.Configuration;
using System.Threading;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Web;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography;
using System.Text;
using eBizChargeForSAP;
using System.Collections.Generic;
using System.Linq;
using eBizChargeForSAP.ServiceReference1;
using System.Windows.Forms;

partial class SAP
{

    DateTime oDT = DateTime.Parse("1/1/1900");
    static string ebizVersion = "";
    public string Truncate(string value, int maxLength)
    {
        return string.IsNullOrEmpty(value) ? value : value.Substring(0, Math.Min(value.Length, maxLength));
    }
    public string getAddOnVersion()
    {
        ebizVersion = GetConfig("AName");
        string sql = string.Format(" select \"AddOnVer\" from \"SBOCOMMON\".\"SARI\" where \"NameSpace\"='cbs' and \"AName\"='{0}' and \"AddOnVer\" <> ''", GetConfig("AName"));
        try
        {

            ebizVersion = GetConfig("AName") + "/" + getSQLString(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return ebizVersion;

    }
    public void insert(CCCUST cust)
    {
        int id = getNextTableID("@CCCUST");

        string sql = string.Format("insert into \"@CCCUST\" (\"DocEntry\", \"U_active\",\"U_cardCode\",\"U_CardName\",\"U_cardType\",\"U_CCAccountID\",\"U_checkingAccount\"" +
        ",\"U_city\",\"U_CustNum\",\"U_CustomerID\",\"U_Declined\",\"U_email\",\"U_expDate\",\"U_firstName\",\"U_GroupName\",\"U_lastName\",\"U_methodDescription\",\"U_MethodID\",\"U_routingNumber\",\"U_state\",\"U_street\",\"U_zip\",\"U_cardNum\", \"U_MethodName\", \"U_default\")" +
            "values({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}')"
            , id
            , cust.active
            , cust.cardCode
            , Truncate(cust.CardName == null ? "" : cust.CardName.Replace("'", "''"), 80)
            , cust.cardType
            , cust.CCAccountID
            , cust.checkingAccount
            , Truncate(cust.city == null ? "" : cust.city.Replace("'", "''"), 50)
            , cust.CustNum
            , cust.CustomerID == null ? "" : cust.CustomerID.Replace("'", "''")
            , cust.Declined
            , Truncate(cust.email == null ? "" : cust.email.Replace("'", "''"), 80)
            , cust.expDate
            , Truncate(cust.firstName == null ? "" : cust.firstName.Replace("'", "''"), 50)
            , Truncate(cust.GroupName == null ? "" : cust.GroupName.Replace("'", "''"), 80)
            , Truncate(cust.lastName == null ? "" : cust.lastName.Replace("'", "''"), 50)
            , Truncate(cust.methodDescription == null ? "" : cust.methodDescription.Replace("'", "''"), 250)
            , cust.MethodID
            , cust.routingNumber
            , cust.state
            , Truncate(cust.street == null ? "" : cust.street.Replace("'", "''"), 80)
            , cust.zip
            , cust.cardNum
            , Truncate(cust.methodName == null ? "" : cust.methodName.Replace("'", "''"), 80)
            , cust.@default
           );
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void insert(CCCPayment pmt)
    {
        int id = getNextTableID("@CCCPAYMENT");

        string sql = "";
        try
        {
            sql = string.Format("insert into \"@CCCPAYMENT\" (\"DocEntry\", \"U_CustomerID\", \r\n" +
"\"U_InvoiceID\", \r\n" +
"\"U_PaymentID\", \r\n" +
"\"U_Amount\", \r\n" +
"\"U_DateImported\")" +
            "values({0},'{1}','{2}','{3}','{4}',{5})"
            , id
            , pmt.CustomerID
            , pmt.InvoiceID
            , pmt.PaymentID
            , pmt.Amount
            , toDateString(pmt.DateImported)
            );

            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void insert(CCPayFormPayment pmt)
    {
        int id = getNextTableID("@CCPAYFORMPAYMENT");

        string sql = "";
        try
        {
            sql = string.Format("insert into \"@CCPAYFORMPAYMENT\" (\"DocEntry\", \"U_CustomerID\", \r\n" +
"\"U_InvoiceID\", \r\n" +
"\"U_PaymentID\", \r\n" +
"\"U_Amount\", \r\n" +
"\"U_DateImported\")" +
            "values({0},'{1}','{2}','{3}','{4}',{5})"
            , id
            , pmt.CustomerID
            , pmt.InvoiceID
            , pmt.PaymentID
            , pmt.Amount
            , toDateString(pmt.DateImported)
            );

            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void update(CCCPayment pmt)
    {
        string sql = "";
        try
        {
            sql = string.Format("Update \"@CCCPAYMENT\" (\"DocEntry\", \"U_CustomerID\" = '{1}', \r\n" +
"\"U_InvoiceID\" = '{2}', \r\n" +
"\"U_PaymentID\" = '{3}', \r\n" +
"\"U_Amount\" = '{4}', \r\n" +
"\"U_DateImported\" = TO_DATE('{5}','MM/DD/YYYY')) \r\n" +
             " where \"DocEntry\" = {0}"
            , pmt.recID
            , pmt.CustomerID
            , pmt.InvoiceID
            , pmt.PaymentID
            , pmt.Amount
            , pmt.DateImported == null ? "" : ((DateTime)pmt.DateImported).ToString("MM/dd/yyyy")
           );

            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void update(CCPayFormPayment pmt)
    {
        string sql = "";
        try
        {
            sql = string.Format("Update \"@CCPAYFORMPAYMENT\" (\"DocEntry\", \"U_CustomerID\" = '{1}', \r\n" +
"\"U_InvoiceID\" = '{2}', \r\n" +
"\"U_PaymentID\" = '{3}', \r\n" +
"\"U_Amount\" = '{4}', \r\n" +
"\"U_DateImported\" = TO_DATE('{5}','MM/DD/YYYY')) \r\n" +
             " where \"DocEntry\" = {0}"
            , pmt.recID
            , pmt.CustomerID
            , pmt.InvoiceID
            , pmt.PaymentID
            , pmt.Amount
            , pmt.DateImported == null ? "" : ((DateTime)pmt.DateImported).ToString("MM/dd/yyyy")
           );

            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void update(CCCUST cust)
    {
        string sql = "";
        try
        {
            sql = string.Format("update \"@CCCUST\"  set \"U_cardCode\" = '{1}' , \r\n" +
                "\"U_CardName\" = '{2}', \r\n" +
                "\"U_cardType\"='{3}', \r\n" +
                "\"U_CCAccountID\" = '{4}', \r\n" +
                "\"U_checkingAccount\" = '{5}', \r\n" +
        "\"U_city\" = '{6}', \r\n" +
        "\"U_CustNum\"='{7}', \r\n" +
        "\"U_CustomerID\"='{8}',\r\n" +
        "\"U_email\"='{9}',\r\n" +
        "\"U_expDate\"='{10}', \r\n" +
           "\"U_firstName\"='{11}',\r\n" +
           "\"U_GroupName\"='{12}', \r\n" +
           "\"U_lastName\"='{13}', \r\n" +
        "\"U_methodDescription\"='{14}', \r\n" +
        "\"U_MethodID\"='{15}', \r\n" +
        "\"U_routingNumber\"='{16}', \r\n" +
        "\"U_state\"='{17}', \r\n" +
        "\"U_street\"='{18}', \r\n" +
        "\"U_zip\"='{19}', \r\n" +
        "\"U_cardNum\"='{20}', \r\n" +
        "\"U_MethodName\"='{21}' \r\n" +
           " where \"DocEntry\" = {0}"
            , cust.recID
            , cust.cardCode
            , Truncate(cust.CardName == null ? "" : cust.CardName.Replace("'", "''"), 80)
            , cust.cardType
            , cust.CCAccountID
            , cust.checkingAccount
            , Truncate(cust.city == null ? "" : cust.city.Replace("'", "''"), 50)
            , cust.CustNum
            , cust.CustomerID
            , Truncate(cust.email == null ? "" : cust.email.Replace("'", "''"), 80)
            , cust.expDate
           , Truncate(cust.firstName == null ? "" : cust.firstName.Replace("'", "''"), 50)
            , Truncate(cust.GroupName == null ? "" : cust.GroupName.Replace("'", "''"), 80)
            , Truncate(cust.lastName == null ? "" : cust.lastName.Replace("'", "''"), 50)
            , Truncate(cust.methodDescription == null ? "" : cust.methodDescription.Replace("'", "''"), 255)
            , cust.MethodID
            , cust.routingNumber
            , cust.state
            , Truncate(cust.street == null ? "" : cust.street.Replace("'", "''"), 80)
            , cust.zip
            , cust.cardNum
            , Truncate(cust.methodName == null ? "" : cust.methodName.Replace("'", "''"), 80));
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void insert(CCJOB j)
    {
        int id = getNextTableID("@CCJOB");

        string sql = "";
        try
        {
            sql = string.Format("insert into \"@CCJOB\" (\"DocEntry\", \"U_Amount\",\"U_Cancelled\",\"U_CancelledBy\",\"U_CancelledDate\",\"U_CustomerID\",\"U_Description\"" +
    ",\"U_EndDate\",\"U_Frequency\",\"U_InvoiceID\", \"U_LastRunDate\",\"U_NextRunDate\",\"U_OrderID\",\"U_PaymentID\" ,\"U_Result\",\"U_StartDate\")" +
        "values({0},'{1}','{2}','{3}',{4},'{5}','{6}',{7},'{8}','{9}',{10},{11},'{12}','{13}','{14}',{15})"
        , id
        , j.Amount
        , j.Cancelled
        , j.CancelledBy
        , toDateString(j.CancelledDate)
        , j.CustomerID
        , j.Description
        , toDateString(j.EndDate)
        , j.Frequency
        , j.InvoiceID
        , toDateString(j.LastRunDate)
        , toDateString(j.NextRunDate)
        , j.OrderID
        , j.PaymentID
        , j.Result
        , toDateString(j.StartDate)
        );

            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void insert(CCConnectInvoice j)
    {
        int id = getNextTableID("@CCCONNECTINVOICE");

        string sql = "";
        try
        {

            sql = string.Format("insert into \"@CCCONNECTINVOICE\" (\"DocEntry\", \"U_InvoiceID\", \r\n" +
     "\"U_CustomerID\", \r\n" +
     "\"U_InvoiceGUID\", \r\n" +
     "\"U_PaidAmount\", \r\n" +
     "\"U_UploadedBalance\", \r\n" +
     "\"U_Status\", \r\n" +
     "\"U_UploadDate\", \r\n" +
     "\"U_PaidDate\") \r\n" +
                 "values({0},'{1}','{2}','{3}','{4}','{5}','{6}',{7},{8})"
                 , id
                 , j.InvoiceID
                 , j.CustomerID
                 , j.InvoiceGUID
                 , j.PaidAmount
                 , j.UploadedBalance
                 , j.Status
                 , toDateString(j.UploadDate)
                 , toDateString(j.PaidDate)
                 );
            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void insert(CCPayFormInvoice j)
    {
        int id = getNextTableID("@CCPAYFORMINVOICE");

        string sql = "";
        try
        {

            sql = string.Format("insert into \"@CCPAYFORMINVOICE\" (\"DocEntry\", \"U_InvoiceID\", \r\n" +
     "\"U_CustomerID\", \r\n" +
     "\"U_InvoiceGUID\", \r\n" +
     "\"U_PaidAmount\", \r\n" +
     "\"U_UploadedBalance\", \r\n" +
     "\"U_Status\", \r\n" +
     "\"U_UploadDate\", \r\n" +
     "\"U_PaidDate\") \r\n" +
                 "values({0},'{1}','{2}','{3}','{4}','{5}','{6}',{7},{8})"
                 , id
                 , j.InvoiceID
                 , j.CustomerID
                 , j.InvoiceGUID
                 , j.PaidAmount
                 , j.UploadedBalance
                 , j.Status
                 , toDateString(j.UploadDate)
                 , toDateString(j.PaidDate)
                 );
            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void update(CCJOB j)
    {
        string sql = "";
        try
        {
            sql = string.Format("update \"@CCJOB\" set \"U_Amount\" = '{1}' ,\"U_Cancelled\" = '{2}', \"U_CancelledBy\"='{3}', \"U_CancelledDate\" = {4}, \r\n" +
       "\"U_CustomerID\"='{5}', \"U_Description\"='{6}', \"U_EndDate\"={7}, \"U_Frequency\"='{8}', \"U_InvoiceID\"='{9}', \"U_LastRunDate\"={10}, \r\n" +
       "\"U_NextRunDate\"={11},\"U_OrderID\"='{12}', \"U_PaymentID\"='{13}',\"U_Result\"='{14}',\"U_StartDate\"={15} \r\n" +
           "where \"DocEntry\" = {0}"
           , j.jobID
           , j.Amount
           , j.Cancelled
            , j.CancelledBy
            , toDateString(j.CancelledDate)
            , j.CustomerID
            , Truncate(j.Description.Replace("'", "''"), 254)
            , toDateString(j.EndDate)
            , j.Frequency
            , j.InvoiceID
            , toDateString(j.LastRunDate)
            , toDateString(j.NextRunDate)
            , j.OrderID
            , j.PaymentID
            , j.Result
            , toDateString(j.StartDate)
            );

            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void update(CCConnectInvoice j)
    {
        string sql = "";
        try
        {
            sql = string.Format("update \"@CCCONNECTINVOICE\" set \"U_InvoiceID\" = '{1}', \r\n" +
    "\"U_CustomerID\" = '{2}', \r\n" +
    "\"U_InvoiceGUID\" = '{3}', \r\n" +
    "\"U_PaidAmount\" = '{4}', \r\n" +
    "\"U_UploadedBalance\" = '{5}', \r\n" +
    "\"U_Status\" = '{6}', \r\n" +
    "\"U_UploadDate\" = {7}, \r\n" +
    "\"U_PaidDate\" = {8} \r\n" +
               "where \"DocEntry\" = {0}"
               , j.recID
               , j.InvoiceID
               , j.CustomerID
               , j.InvoiceGUID
                , j.PaidAmount
                , j.UploadedBalance
                , j.Status
                , toDateString(j.UploadDate)
                , toDateString(j.PaidDate)
                );

            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void update(CCPayFormInvoice j)
    {
        string sql = "";
        try
        {
            sql = string.Format("update \"@CCPAYFORMINVOICE\" set \"U_InvoiceID\" = '{1}', \r\n" +
    "\"U_CustomerID\" = '{2}', \r\n" +
    "\"U_InvoiceGUID\" = '{3}', \r\n" +
    "\"U_PaidAmount\" = '{4}', \r\n" +
    "\"U_UploadedBalance\" = '{5}', \r\n" +
    "\"U_Status\" = '{6}', \r\n" +
    "\"U_UploadDate\" = {7}, \r\n" +
    "\"U_PaidDate\" = {8} \r\n" +
               "where \"DocEntry\" = {0}"
               , j.recID
               , j.InvoiceID
               , j.CustomerID
               , j.InvoiceGUID
                , j.PaidAmount
                , j.UploadedBalance
                , j.Status
                , toDateString(j.UploadDate)
                , toDateString(j.PaidDate)
                );

            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void insert(CCTRAN t)
    {
        int id = getNextTableID("@CCTRANS");

        string sql = string.Format("insert into \"@CCTRANS\" (\"DocEntry\", \"U_acsUrl\",\"U_amount\",\"U_authAmount\",\"U_authCode\",\"U_avsResult\",\"U_avsResultCode\"\r\n" +
        ",\"U_batchNum\",\"U_batchRef\",\"U_cardCodeResult\",\"U_cardCodeRCode\",\"U_CardHolder\",\"U_cardLevelResult\",\"U_cardLevelRCode\",\"U_CCAccountID\",\"U_CCORDERID\"\r\n" +
        ",\"U_command\",\"U_conversionRate\",\"U_convertedAmount\",\"U_convertedCurrency\",\"U_crCardNum\",\"U_CreditMemoID\",\"U_creditRef\",\"U_custNum\",\"U_customerID\"\r\n" +
        ",\"U_customerName\",\"U_Description\",\"U_DownPaymentInvID\",\"U_error\",\"U_errorCode\",\"U_InvoiceID\", \"U_jobID\",\"U_MethodID\",\"U_OrderID\", \"U_PaymentID\"\r\n" +
        ", \"U_profilerReason\", \"U_profilerResponse\", \"U_profilerScore\", \"U_recDate\", \"U_recID\", \"U_refNum\", \"U_remainingBalance\", \"U_result\", \"U_resultCode\"\r\n" +
        ", \"U_status\", \"U_statusCode\", \"U_vpasResultCode\")\r\n" +
        "values('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}'\r\n" +
        ",'{21}','{22}','{23}','{24}','{25}','{26}','{27}','{28}','{29}','{30}'\r\n" +
        ",'{31}','{32}','{33}','{44}','{35}','{36}','{37}',CURRENT_DATE,'{38}','{39}'\r\n" +
        ",'{40}','{41}','{42}','{43}','{44}','{45}')"
            , id    //0
            , t.acsUrl == null ? "" : t.acsUrl.Replace("'", "''")
            , t.amount
            , t.authAmount
            , t.authCode
            , t.avsResult
            , t.avsResultCode
            , t.batchNum
            , t.batchRef
            , t.cardCodeResult
            , t.cardCodeResultCode  //10
            , t.CardHolder
            , t.cardLevelResult
            , t.cardLevelResultCode
            , t.CCAccountID
            , t.CCORDERID
            , t.command
            , t.conversionRate
            , t.convertedAmount
            , t.convertedAmountCurrency
            , t.crCardNum //20
            , t.CreditMemoID
            , t.creditRef
            , t.custNum
            , t.customerID
            , t.customerName
            , Truncate(t.Description == null ? "" : t.Description.Replace("'", "''"), 50)
            , t.DownPaymentInvoiceID
            , t.error == null ? "" : t.error.Replace("'", "''")
            , t.errorCode
            , t.InvoiceID //30
            , t.jobID
            , t.MethodID
            , t.OrderID
            , t.PaymentID
            , t.profilerReason
            , t.profilerResponse
            , t.profilerScore
            , t.recID
            , t.refNum
            , t.remainingBalance //40
            , t.result == null ? "" : t.result.Replace("'", "''")
            , t.resultCode
            , t.status == null ? "" : t.status.Replace("'", "''")
            , t.statusCode
            , t.vpasResultCode
        );
        try
        {
            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }


    public int getJELineID(string transId, string cardcode)
    {
        string sql = string.Format("select \"Line_ID\" from JDT1 where \"TransId\" = '{0}' and \"ShortName\" = '{1}'", transId, cardcode);
        int rtn = 1;
        try
        {
            int i = getSQLInt(sql);
            if (i != 0)
                return i;
            sql = string.Format("select ExpnsCode from OEXD");
            return getSQLInt(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }

    public int getExpenseCode()
    {
        string sql = string.Format("select \"ExpnsCode\" from OEXD where \"ExpnsName\"='Freight'");
        int rtn = 1;
        try
        {
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);
            int i = getSQLInt(sql);
            if (i != 0)
                return i;
            sql = string.Format("select \"ExpnsCode\", \"ExpnsName\" from OEXD");
            return getSQLInt(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public int getPaymentTypeCode(int creditCardCode)
    {
        string sql = string.Format("select top 1 \"CrTypeCode\" from OCRP where \"CreditCard\" in (0, {0}) order by \"CreditCard\" desc", creditCardCode);
        //errorLog(sql);
        int rtn = 1;
        try
        {
            return getSQLInt(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public double getInvoicePaidToDate(int id, String table = "OINV")
    {
        string sql = string.Format("select \"PaidToDate\" from {1} where \"DocEntry\" = {0}", id, table);
        double rtn = 0;
        try
        {
            return getSQLDouble(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public string getEmptTaxCode()
    {
        string sql = string.Format("Select \"Code\" from OSTC where \"Rate\" = 0");
        string rtn = "";
        try
        {
            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public double getTaxRate(ref string code)
    {
        string sql = string.Format("Select \"Rate\" from OSTC where \"Code\" ='{0}'", code);
        double rtn = 0;
        try
        {
            double rate = getSQLDouble(sql);
            if (rate != 0)
            {
                return rate;
            }
            else
            {
                code = getEmptTaxCode();
                return getTaxRate(ref code);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;

    }
    public double getCashDiscountRate(string cardcode)
    {
        double rtn = 0;
        string sql = string.Format("select c.\"Discount\" from OCRD a, OCTG b, CDC1 c where a.\"GroupNum\" = b.\"GroupNum\" and b.\"DiscCode\" = c.\"CdcCode\" and a.\"CardCode\" = '{0}'", cardcode);
        return getSQLDouble(sql);

    }
    public double getSQLDouble(string sql)
    {
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {

            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                return (double)oRS.Fields.Item(0).Value;
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return 0;

    }

    public bool updateCustState(int id, string active)
    {
        string sql = string.Format("UPDATE \"@CCCUST\" set \"U_active\" = '{0}' where \"DocEntry\"={1}", active, id);
        try
        {

            execute(sql);


        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            return false;
        }
        return true;
    }
    public bool updateCustDefaultState(int id, string d, string customerID)
    {
        string sql = "";
        if (d == "Y")
        {
            sql = string.Format("UPDATE \"@CCCUST\" set \"U_default\" = 'N' where \"U_CustomerID\" ='{0}'", customerID);
            execute(sql);
            trace(sql);
        }
        execute(string.Format("UPDATE \"@CCCUST\" set \"U_default\" = '{0}' where \"DocEntry\"={1}", d, id));
        return true;
    }
    public bool updateCustDeclined(int id, string declined)
    {
        string sql = string.Format("UPDATE \"@CCCUST\" set \"U_Declined\" = '{0}' where \"DocEntry\"={1}", declined, id);
        try
        {

            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            return false;
        }
        return true;
    }
    public void removeImportError(DateTime created)
    {
        string sql = string.Format("delete from CCITEM where OrderID in (select ORderID from CCORDER where imported = 'E' and created > '{0}')\r\ndelete from CCORDER where imported = 'E' and created > '{0}'", created.ToString("MM/dd/yyyy"));
        try
        {

            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void setImportError(int id, string err)
    {
        string sql = string.Format("update CCORDER set imported = 'E', Error= '{0}' where OrderID={1}", err, id);
        try
        {

            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }

    public string getItemCode(string sku)
    {
        string rtn = "";
        string sql = string.Format("select \"ItemCode\" from OITM where \"FrgnName\" = '{0}' or \"ItemCode\"='{0}'", sku);
        try
        {

            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public int getDocEntryByDelivery(string customerid, int deliveryid)
    {
        int rtn = 0;
        string sql = string.Format("select \"DocEntry\" from INV1 where \"BaseCard\" = '{0}' and \"BaseEntry\" = {1} and \"BaseType\" = 15", customerid, deliveryid);
        try
        {

            return getSQLInt(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public int getInvDocEntryByOrder(string customerid, int orderid)
    {
        int rtn = 0;
        string sql = string.Format("select \"DocEntry\" from INV1 where \"BaseCard\" = '{0}' and \"BaseEntry\" = {1} and \"BaseType\" = 17", customerid, orderid);
        try
        {

            return getSQLInt(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public int getNextTableID(string tn)
    {
        int rtn = 0;
        string sql = string.Format("select max(\"DocEntry\") from \"{0}\" ", tn);
        try
        {
            return getSQLInt(sql) + 1;
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public string getCurrencyString()
    {
        string sql = string.Format("select \"CurrCode\" from OCRN");
        string rtn = "";
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            oRS.DoQuery(sql);
            rtn = (string)oRS.Fields.Item(0).Value;
            oRS.MoveNext();
            while (!oRS.EoF)
            {
                rtn = rtn + "," + (string)oRS.Fields.Item(0).Value;
                oRS.MoveNext();
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return rtn;
    }

    public bool isOrderPreauthed(string id)
    {
        string sql = string.Format("select 1 from \"@CCTRANS\" where \"U_OrderID\" = '{0}' and \"U_command\" = 'cc:authonly'", id);

        try
        {
            return !IsRSEOF(sql);


        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return false;
    }

    public string getCurrency(string tbname, string id)
    {
        if (tbname == "" || id == "")
            return "";
        string sql = string.Format("select \"DocCur\" from {0} where \"DocEntry\" = {1}", tbname, id);
        string rtn = "";
        try
        {
            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public int getNextTableNum(string tn)
    {
        int rtn = 0;
        string sql = string.Format("select max(\"DocNum\") from \"{0}\" where \"Handwrtten\" = 'N'", tn);
        try
        {

            return getSQLInt(sql) + 1;
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }



    public List<string> getDownpaymentToDraw(string soID)
    {
        List<string> list = new List<string>();
        string sql = string.Format("select \"U_DownPaymentInvID\" from \"@CCTRANS\" where \"U_OrderID\"='{0}' and \"U_DownPaymentInvID\"<>''", soID);

        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                list.Add((string)oRS.Fields.Item(0).Value);
                oRS.MoveNext();
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }

        return list;
    }
    public int getNextJEID()
    {
        int rtn = 0;
        string sql = "select max(\"TransId\") + 1 from OJDT";
        try
        {

            return getSQLInt(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public string getCCAccountNameByID(string ID)
    {
        string rtn = null;
        string sql = string.Format("Select \"CardName\" from OCRC where \"CreditCard\"='{0}'", ID);
        try
        {

            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public string getMethodIDByOrderID(string ID)
    {
        string rtn = null;
        string sql = string.Format("select \"U_MethodID\" from \"@CCTRANS\" where \"U_errorCode\"='0' and \"U_OrderID\"='{0}'", ID);
        try
        {

            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public int getDocEntry(string type, string DocNum, string cardcode)
    {
        //DocDate = string.Format("{0}/{1}/{2}", DocDate.Substring(4,2), DocDate.Substring(6,2), DocDate, 4));
        int rtn = 0;
        string sql = "";
        string tn = "";
        switch (type)
        {
            case "13":
                tn = "OINV";
                break;
            case "14":
                tn = "ORIN";
                break;
            case "203":
                tn = "ODPI";
                break;
            case "17":
                tn = "ORDR";
                break;
        }
        sql = string.Format("select \"DocEntry\" from {2} where \"DocNum\"={0} and \"CardCode\" = '{1}'", DocNum, cardcode, tn);
        if (tn == "")
            return int.Parse(DocNum);
        try
        {

            int i = getSQLInt(sql);
            if (i != 0)
                return i;

            sql = string.Format("select \"DocEntry\" from {1} where \"DocNum\"={0}", DocNum, tn);
            return getSQLInt(sql);


        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public int getDocNum(string tn, string DocEntry)
    {
        if (DocEntry == "")
        {
            return 0;
        }
        //DocDate = string.Format("{0}/{1}/{2}", DocDate.Substring(4,2), DocDate.Substring(6,2), DocDate, 4));
        int rtn = 0;
        string sql = "";
        sql = string.Format("select \"DocNum\" from {1} where \"DocEntry\"={0} ", DocEntry, tn);
        try
        {

            return getSQLInt(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public int getDocEntry(string DocNum, string tb)
    {
        if (DocNum == "")
        {
            return 0;
        }
        //DocDate = string.Format("{0}/{1}/{2}", DocDate.Substring(4,2), DocDate.Substring(6,2), DocDate, 4));
        int rtn = 0;
        string sql = "";
        sql = string.Format("select \"DocEntry\" from {1} where \"DocNum\"={0} ", DocNum, tb);
        try
        {

            return getSQLInt(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }

    public string getMethodIDByCustomerID(string ID)
    {
        string rtn = null;
        string sql = string.Format("select \"U_MethodID\" from \"@CCTRANS\" where \"U_errorCode\"='0' and \"U_CustomerID\"='{0}'  Order By \"DocEntry\" desc ", ID);
        try
        {

            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public string getJobCustomerIDFromTrans(int jobID)
    {
        string rtn = null;
        string sql = string.Format("select \"U_CustomerID\" from \"@CCTRANS\" where \"U_jobID\"='{0}'", jobID);
        try
        {

            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public string getMethodID(string ID, string fn)
    {
        string rtn = null;
        string sql = string.Format("select \"U_MethodID\" from \"@CCTRANS\" where \"U_errorCode\"='0' and \"{1}\"='{0}'", ID, fn);
        try
        {

            return getSQLString(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }

    public void CCTRANAdjustSOWithInvId(int InvID, int OrderID)
    {
        string sql = string.Format("update \"@CCTRANS\" set \"U_InvoiceID\" = '{1}' where \"U_OrderID\" = '{0}' and \"U_InvoiceID\"=''", OrderID, InvID);
        string sql2 = string.Format("update \"@CCTRANS\" set \"U_OrderID\" = '{0}' where \"U_InvoiceID\" = '{1}' and \"U_OrderID\"='' ", OrderID, InvID);
        try
        {
            execute(sql);
            execute(sql2);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            errorLog(sql2);
        }
    }
    public void CCTRANAdjustSOWithDPId(int DPID, int OrderID)
    {
        string sql = string.Format("update \"@CCTRANS\" set \"U_DownPaymentInvID\" = '{1}' where \"U_OrderID\" = '{0}' and \"U_DownPaymentInvID\" = ''", OrderID, DPID);
        string sql2 = string.Format("update \"@CCTRANS\" set  \"U_OrderID\" = '{0}' where \"U_DownPaymentInvID\" = '{1}' and \"U_OrderID\"='' ", OrderID, DPID);
        try
        {
            execute(sql);
            execute(sql2);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            errorLog(sql2);
        }
    }
    public void CCTRANAdjustInvWithCMId(int CMID, int InvID)
    {
        string sql = string.Format("update \"@CCTRANS\"  set \"U_InvoiceID\" = '{1}' where \"U_CreditMemoID\" = '{0}' and \"U_InvoiceID\"=''", CMID);
        string sql2 = string.Format("update \"@CCTRANS\" set \"U_CreditMemoID\" = '{0}' where \"U_InvoiceID\" = '{1}' and \"U_CreditMemoID\"=''", CMID, InvID);
        try
        {
            execute(sql);
            execute(sql2);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            errorLog(sql2);
        }
    }
    public void CCTRANAdjustWithSOId(string refNum, string OrderID)
    {
        string sql = string.Format("update \"@CCTRANS\" set \"U_OrderID\" = '{1}' where \"U_refNum\" = '{0}'", refNum, OrderID);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void CCTRANAdjustWithDPId(string refNum, string dpID)
    {
        string sql = string.Format("update \"@CCTRANS\" set \"U_DownPaymentInvID\" = '{0}' where \"U_refNum\" = '{1}'", dpID, refNum);
        try
        {
            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public string GetPreauthRefNum(string OrderID, ref string amount, ref string DeliveryID, ref string customerID)
    {
        string sRtn = "";
        string sql = string.Format("select \"U_refNum\",\"U_amount\", \"U_DeliveryID\", \"U_customerID\" from  \"@CCTRANS\" where \"U_OrderID\"= '{0}' and \"U_command\" = 'cc:authonly' order by \"U_recDate\" desc", OrderID);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                DeliveryID = (string)oRS.Fields.Item(2).Value;

                amount = (string)oRS.Fields.Item(1).Value;
                customerID = (string)oRS.Fields.Item(3).Value;
                return (string)oRS.Fields.Item(0).Value;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return sRtn;
    }
    public void CCTRANUpdateIncomingPaymentID(string refNum)
    {
        string sql = string.Format("Update \"@CCTRANS\" set \"U_SAPPaymentID\" = CAST(b.\"DocNum\" AS VARCHAR) from \"@CCTRANS\" a, RCT2 b where a.\"U_InvoiceID\" = CAST(b.\"DocEntry\" AS VARCHAR) and a.\"U_refNum\" = '{0}'", refNum);
        try
        {
            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void CCTRANUpdateOutgoingPaymentID(string refNum)
    {
        string sql = string.Format("update \"@CCTRANS\" set \"U_SAPpaymentID\" = CAST((select max(\"DocEntry\") from OVPM) AS VARCHAR) where \"U_refNum\" = '{0}'", refNum);
        try
        {
            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void CCTRANAdjustWithINVId(string refNum, string invID)
    {
        string sql = string.Format("update \"@CCTRANS\" set \"U_InvoiceID\" = '{0}' where \"U_refNum\" = '{1}'", invID, refNum);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void CCTRANAdjustWithCMId(string refNum, string cmID)
    {
        string sql = string.Format("update \"@CCTRANS\" set \"U_CreditMemoID\" = '{1}' where \"U_refNum\" = '{0}'", refNum, cmID);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void CCTRANAdjustJobId(string refNum, string jobID)
    {
        string sql = string.Format("update \"@CCTRANS\" set \"U_jobID\" = '{1}' where \"U_refNum\" = '{0}'", refNum, jobID);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void CancelJob(string jobID)
    {
        string sql = string.Format("update \"@CCJOB\" set \"U_Cancelled\"='Y', \"U_CancelledDate\" = CURRENT_DATE where \"DocEntry\" = {0}", jobID);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        populateRBillingMatrix();
    }
    public void ActivateJob(string jobID)
    {
        string sql = string.Format("update \"@CCJOB\" set \"U_Cancelled\"= null, \"U_CancelledDate\" = null where \"DocEntry\" = {0}", jobID);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        populateRBillingMatrix();
    }
    public void updateOrderWithSOId(int soID, int OrderID)
    {
        string sql = string.Format("Update CCORDER set SAPOrderID='{0}' where OrderID={1}", soID, OrderID);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void updateOrderWithInvId(int invID, int OrderID)
    {
        string sql = string.Format("Update CCORDER set SAPInvoiceID='{0}' where OrderID={1}", invID, OrderID);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void updateCCTRANSWithPaymentId(string refNum, string CCAccountID)
    {
        string sql = string.Format("Update \"@CCTRANS\" set \"U_PaymentID\"=CAST((select max(\"TransId\") from OJDT where \"BaseRef\" = \"U_SAPPaymentID\") AS VARCHAR) ,\"U_CCAccountID\"='{1}'  where \"U_refNum\"='{0}'", refNum, CCAccountID);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void BIupdateCCTRANS(string refNum, string orderID, string invID)
    {
        string sql = string.Format("Update \"@CCTRANS\" set \"InvoiceID\"='{2}', \"OrderID\" = '{1}'  where \"refNum\"='{0}'", refNum, orderID, invID);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void voidCCTRANS(string refNum)
    {
        string sql = string.Format("Update \"@CCTRANS\" set  \"U_command\"='cc:void' where \"U_refNum\"='{0}' ", refNum);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void voidErrorCCTRANS(string refNum)
    {
        string sql = string.Format("Update \"@CCTRANS\" set \"U_InvoiceID\" = (select top 1 \"U_InvoiceID\" from \"@CCTRANS\" where \"U_refNum\" = '{0}' and \"U_command\"<>'cc:void')   " +
", \"U_DownPaymentInvID\" = (select top 1 \"U_DownPaymentInvID\" from \"@CCTRANS\" where \"U_refNum\" = '{0}' and \"U_command\"<>'cc:void') " +
", \"U_OrderID\" = (select top 1 \"U_OrderID\" from \"@CCTRANS\" where \"U_refNum\" = '{0}' and \"U_command\"<>'cc:void')   " +
",\"U_customerID\" = (select top 1 \"U_customerID\" from \"@CCTRANS\" where \"U_refNum\" = '{0}' and \"U_command\"<>'cc:void')   " +
", \"U_Description\"='Rollback transaction on error.'  " +
",\"U_CreditMemoID\" = (select top 1 \"U_CreditMemoID\" from \"@CCTRANS\" where \"U_refNum\" = '{0}' and \"U_command\"<>'cc:void') " +
"where \"U_refNum\"='{0}' and \"U_command\" = 'cc:void'", refNum);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }

    public bool isCardCodeExists(string code)
    {
        string sql = "";
        try
        {
            sql = string.Format("Select CreditCard from OCRC where CreditCard = '{0}'", code);
            return !IsRSEOF(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return true;
    }
    public bool CanCapure(string key)
    {
        string sql = "";
        try
        {
            sql = string.Format("select \"DocEntry\"  from OINV where \"DocEntry\" = '{0}' and CANCELED = 'N' and \"DocStatus\" <> 'C'", key);
            return !IsRSEOF(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return true;
    }
    public bool isCardNameExists(string Name)
    {
        string sql = "";

        try
        {
            sql = string.Format("Select CreditCard from OCRC where CardName = '{0}'", Name);
            return !IsRSEOF(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return true;
    }
    public string getCardAcctCode(int cardcode)
    {
        string str = null;
        string sql = "";
        try
        {
            sql = string.Format("Select \"CreditCard\" from OCRC where \"CreditCard\" = {0}", cardcode);
            int i = getSQLInt(sql);
            if (i == 0)
            {
                sql = string.Format("Select \"CreditCard\" from OCRC");
                i = getSQLInt(sql);
            }
            return i.ToString();

        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return str;
    }
    public List<CCCUST> getCCUST(string customerID)
    {
        List<CCCUST> list = new List<CCCUST>();
        string sql = string.Format("SELECT \"DocEntry\" , \"U_CustomerID\" ,\"U_CustNum\" ,\"U_MethodID\",\"U_active\" ,\"U_Declined\" ,\"U_methodDescription\" ,\"U_cardNum\",\"U_expDate\", " +
    "\"U_cardCode\",\"U_routingNumber\",\"U_checkingAccount\",\"U_firstName\",\"U_lastName\",\"U_email\",\"U_street\",\"U_city\",\"U_state\",\"U_zip\" ,\"U_cardType\",\"U_GroupName\",\"U_CardName\",\"U_CCAccountID\",\"U_MethodName\",\"U_default\" FROM \"@CCCUST\" where \"U_CustomerID\" = '{0}' order by \"DocEntry\" desc", customerID);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                while (!oRS.EoF)
                {

                    CCCUST c = new CCCUST();
                    c.recID = (int)oRS.Fields.Item(0).Value;
                    c.CustomerID = (string)oRS.Fields.Item(1).Value;
                    c.CustNum = (string)oRS.Fields.Item(2).Value;
                    c.MethodID = (string)oRS.Fields.Item(3).Value;

                    c.active = 'Y';
                    try
                    {
                        if ((string)oRS.Fields.Item(4).Value == "Y")
                            c.active = 'Y';
                        else
                            c.active = 'N';
                    }
                    catch (Exception ex)
                    {
                        errorLog(ex.Message + ",active");
                    }
                    try
                    {
                        if ((string)oRS.Fields.Item(5).Value == "Y")
                            c.Declined = 'Y';
                        else
                            c.Declined = 'N';
                    }
                    catch (Exception ex)
                    {

                        errorLog(ex.Message + ",Declined");
                    }

                    c.methodDescription = (string)oRS.Fields.Item(6).Value;
                    c.cardNum = (string)oRS.Fields.Item(7).Value;

                    c.expDate = (string)oRS.Fields.Item(8).Value;
                    c.cardCode = (string)oRS.Fields.Item(9).Value;
                    c.routingNumber = (string)oRS.Fields.Item(10).Value;
                    if (c.routingNumber == "")
                        c.routingNumber = null;
                    c.checkingAccount = (string)oRS.Fields.Item(11).Value;
                    if (c.checkingAccount == "")
                        c.checkingAccount = null;
                    c.firstName = (string)oRS.Fields.Item(12).Value;
                    c.lastName = (string)oRS.Fields.Item(13).Value;
                    c.email = (string)oRS.Fields.Item(14).Value;
                    c.street = (string)oRS.Fields.Item(15).Value;
                    c.city = (string)oRS.Fields.Item(16).Value;
                    c.state = (string)oRS.Fields.Item(17).Value;
                    c.zip = (string)oRS.Fields.Item(18).Value;
                    c.cardType = (string)oRS.Fields.Item(19).Value;
                    c.GroupName = (string)oRS.Fields.Item(20).Value;
                    c.CardName = (string)oRS.Fields.Item(21).Value;
                    c.CCAccountID = (string)oRS.Fields.Item(22).Value;
                    c.methodName = (string)oRS.Fields.Item(23).Value;
                    c.@default = 'N';
                    try
                    {
                        if ((string)oRS.Fields.Item(24).Value == "Y")
                            c.@default = 'Y';
                        else
                            c.@default = 'N';
                    }
                    catch (Exception ex)
                    {
                        errorLog(ex.Message + ",active");
                    }
                    list.Add(c);
                    oRS.MoveNext();
                }
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return list;
    }
    public List<CCCUST> getCCUST()
    {
        List<CCCUST> list = new List<CCCUST>();
        string sql = string.Format("SELECT \"DocEntry\" , \"U_CustomerID\" ,\"U_CustNum\" ,\"U_MethodID\",\"U_active\" ,\"U_Declined\" ,\"U_methodDescription\" ,\"U_cardNum\",\"U_expDate\", " +
    "\"U_cardCode\",\"U_routingNumber\",\"U_checkingAccount\",\"U_firstName\",\"U_lastName\",\"U_email\",\"U_street\",\"U_city\",\"U_state\",\"U_zip\" ,\"U_cardType\",\"U_GroupName\",\"U_CardName\",\"U_CCAccountID\",\"U_MethodName\" FROM \"@CCCUST\" order by \"DocEntry\" desc");
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                while (!oRS.EoF)
                {

                    CCCUST c = new CCCUST();
                    c.recID = (int)oRS.Fields.Item(0).Value;
                    c.CustomerID = (string)oRS.Fields.Item(1).Value;
                    c.CustNum = (string)oRS.Fields.Item(2).Value;
                    c.MethodID = (string)oRS.Fields.Item(3).Value;

                    c.active = 'Y';
                    try
                    {
                        if ((string)oRS.Fields.Item(4).Value == "Y")
                            c.active = 'Y';
                        else
                            c.active = 'N';
                    }
                    catch (Exception ex)
                    {
                        errorLog(ex.Message + ",active");
                    }
                    try
                    {
                        if ((string)oRS.Fields.Item(5).Value == "Y")
                            c.Declined = 'Y';
                        else
                            c.Declined = 'N';
                    }
                    catch (Exception ex)
                    {

                        errorLog(ex.Message + ",Declined");
                    }

                    c.methodDescription = (string)oRS.Fields.Item(6).Value;
                    c.cardNum = (string)oRS.Fields.Item(7).Value;

                    c.expDate = (string)oRS.Fields.Item(8).Value;
                    c.cardCode = (string)oRS.Fields.Item(9).Value;
                    c.routingNumber = (string)oRS.Fields.Item(10).Value;
                    if (c.routingNumber == "")
                        c.routingNumber = null;
                    c.checkingAccount = (string)oRS.Fields.Item(11).Value;
                    if (c.checkingAccount == "")
                        c.checkingAccount = null;
                    c.firstName = (string)oRS.Fields.Item(12).Value;
                    c.lastName = (string)oRS.Fields.Item(13).Value;
                    c.email = (string)oRS.Fields.Item(14).Value;
                    c.street = (string)oRS.Fields.Item(15).Value;
                    c.city = (string)oRS.Fields.Item(16).Value;
                    c.state = (string)oRS.Fields.Item(17).Value;
                    c.zip = (string)oRS.Fields.Item(18).Value;
                    c.cardType = (string)oRS.Fields.Item(19).Value;
                    c.GroupName = (string)oRS.Fields.Item(20).Value;
                    c.CardName = (string)oRS.Fields.Item(21).Value;
                    c.CCAccountID = (string)oRS.Fields.Item(22).Value;
                    c.methodName = (string)oRS.Fields.Item(23).Value;
                    list.Add(c);
                    oRS.MoveNext();
                }
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return list;
    }
    public string getPaymentKey(string methodID, string customerID)
    {
        string str = null;
        string sql = "";
        try
        {
            sql = string.Format("select \"U_methodDescription\" from \"@CCCUST\" where \"U_MethodID\"='{0}' and \"U_CustomerID\"='{1}'", methodID, customerID);
            return getSQLString(sql);


        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return str;
    }

    public string getUserEmail(string userName)
    {
        string str = "";
        string sql = "";
        try
        {
            sql = string.Format("select \"email\" from OHEM a, OUSR b where a.\"userId\"=b.\"USERID\" and b.\"USER_CODE\" = '{0}' ", userName);
            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return str;
    }
    public int getBranchIdFromINV(string docentry)
    {

        string sql = "";
        try
        {
            sql = string.Format("select \"BPLId\" from \"OINV\" where \"DocEntry\"  = '{0}'", docentry);
            return getSQLInt(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return 0;
    }
    public int getBranchIdByCustomer(string customerid)
    {

        string sql = "";
        try
        {
            sql = string.Format("select \"BPLId\" from \"OINV\" where \"CardCode\"  = '{0}'", customerid);
            return getSQLInt(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return 0;
    }
    public int getBranchIdFromQUT(string docentry)
    {

        string sql = "";
        try
        {
            sql = string.Format("select \"BPLId\" from \"OQUT\" where \"DocEntry\"  = '{0}'", docentry);
            return getSQLInt(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return 0;
    }
    public int getBranchIdFromCM(string docentry)
    {

        string sql = "";
        try
        {
            sql = string.Format("select \"BPLId\" from \"ORIN\" where \"DocEntry\"  = '{0}'", docentry);
            return getSQLInt(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return 0;
    }
    public int getBranchIdFromORDR(string docentry)
    {

        string sql = "";
        try
        {
            sql = string.Format("select \"BPLId\" from \"ORDR\" where \"DocEntry\"  = '{0}'", docentry);
            return getSQLInt(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return 0;
    }
    public int getBranchIdFromOPDI(string docentry)
    {

        string sql = "";
        try
        {
            sql = string.Format("select \"BPLId\" from \"ODPI\" where \"DocEntry\"  = '{0}'", docentry);
            return getSQLInt(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return 0;
    }
    public bool isPaymentMethodExists(string custid, string methodid)
    {

        string sql = string.Format("select \"U_CustomerID\" from \"@CCCUST\" where \"U_CustomerID\" = '{0}' and \"U_MethodID\" = '{1}'", custid, methodid);
        return !IsRSEOF(sql);
    }
    public int getUserBranch(string userName)
    {

        string sql = "";
        try
        {
            sql = string.Format("select \"BPLId\" from USR6 where \"UserCode\"  = '{0}' Order By \"BPLId\"", userName);
            return getSQLInt(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return 0;
    }
    public void clearInvoiceIDForPreAuth()
    {

        string sql = "";
        try
        {

            sql = string.Format("Update \"@CCTRANS\" set \"U_InvoiceID\" = '' where \"U_command\"='cc:authonly'");

            execute(sql);
            //trace(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }


    }
    public bool CreditCardPaymentMethodCheck()
    {
        try
        {
            return !IsRSEOF("select 1 from OCRP");

        }
        catch (Exception ex)
        {

        }
        return false;
    }
    public bool IsRSEOF(string sql)
    {
        bool Flg = true;
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (oRS.EoF)
                return true;
            else
                return false;
        } 
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            Flg = false;
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return Flg;
    }
    public bool CheckTable(string sql)
    {
        bool CheckTable = true;
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
        }
        catch (Exception ex)
        {
            CheckTable = false;
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return CheckTable;
    }
    public int getSQLInt(string sql)
    {
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (oRS.EoF)
                return 0;
            else
            {
                return (int)oRS.Fields.Item(0).Value;
            }
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return 0;
    }
    public void updateDownPaymentTranLogInvId(int invID, int DownPaymentID)
    {
        if (DownPaymentID == 0)
            return;
        string sql = string.Format("Update \"@CCTRANS\" set \"U_InvoiceID\"={0} where \"U_DownPaymentInvID\"='{1}'", invID, DownPaymentID);
        try
        {
            execute(sql);

            sql = string.Format("Update \"@CCTRANS\" set \"U_DownPaymentInvID\"='{1}' where \"U_InvoiceID\"='{0}'", invID, DownPaymentID);
            execute(sql);
            //refresh();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public string getGroupName(string CardCode)
    {
        string sql = string.Format("select \"GroupName\" from OCRG a, OCRD b where a.\"GroupCode\" = b.\"GroupCode\" and b.\"CardCode\" = '{0}'", CardCode);
        string str = "";
        try
        {
            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return str;
    }
    public string getCCAcountIDByRefNum(string RefNum)
    {
        string sql = string.Format("select \"U_CCAccountID\" from \"@CCTRANS\" where \"U_refNum\" = '{0}'", RefNum);
        string str = "";
        try
        {
            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return str;
    }

    public void cfgInit()
    {
        string sql = "";
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            maskCardOnTrans();

            sql = "select \"U_Pin\", \"U_SourceKey\", \"U_eBizChargeUrl\", 'N', 'N', 'N',\"U_NegPayment\", 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', " +
                "'N', 'N', 'N', 'N', 'N','N','N',\"U_paymentGroup\", " +
                "\"U_PMAutoSelect\", 	\"U_importPin\", \"U_importSourceKey\", \"U_ProcessPreAuth\", 'N', 'N', \"U_sendCustomerEmail\",\"U_urlSandbox\", \"U_urlProduction\", \"U_merchantEmail\", " +
                "\"U_MerchantReceipt\", \"U_CustomerReceipt\", 'N', 'N', 'N',\"U_addressRequired\", \"U_merchantID\", \"U_useUserEmail\", 'N', \"U_BatchAutoMode\", 'N', \"U_recurringBilling\", " +
                " \"U_BtnPos\", \"U_DefaultTaxRate\", \"U_NoTransOnInv\", \"U_importCust\", \"U_showStartup\", \"U_EmailSubject\", \"U_SaveReceipt\", \"U_BatchPreauth\", \"U_DeclineUnapprove\", \"U_CaptureOnPrint\", \"U_ApprUDF\",\"U_CaptureOnly\",\"U_AmountField\", \"U_FieldSize\", \"U_HasTerminal\" ,\"AutoSync\" ,\"ReceiptEmailUDF\" ,\"DefaultContactID\",\"ConnectUseDP\"  " +
                " from \"@CCCFG\"";
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                cfgPin = (string)oRS.Fields.Item(0).Value;
                cfgSourceKey = (string)oRS.Fields.Item(1).Value;
                cfgUrl = (string)oRS.Fields.Item(2).Value;
                cfgOutgoingFilter = (string)oRS.Fields.Item(3).Value;
                // cfgBINoDelivery = (string)oRS.Fields.Item(4).Value; //cfgBIInvoiceOnly
                // ebiz.Url =cfgUrl; //https://secure.ebizcharge.com/soap/gate/CCEBDC0A;
                cfgAccountForOutgoingPayment = (string)oRS.Fields.Item(5).Value;
                cfgNegativeIncomingPayment = (string)oRS.Fields.Item(6).Value;
                cfgAutoSelectAccount = (string)oRS.Fields.Item(7).Value;
                cfgVisaIncomingAccount = (string)oRS.Fields.Item(8).Value;
                cfgMasterIncomingAccount = (string)oRS.Fields.Item(9).Value;
                cfgAMXIncomingAccount = (string)oRS.Fields.Item(10).Value;
                cfgOtherIncomingAccount = (string)oRS.Fields.Item(11).Value;
                cfgVisaOutgoingAccount = (string)oRS.Fields.Item(12).Value;
                cfgMasterOutgoingAccount = (string)oRS.Fields.Item(13).Value;
                cfgAMXOutgoingAccount = (string)oRS.Fields.Item(14).Value;
                cfgOtherOutgoingAccount = (string)oRS.Fields.Item(15).Value;
                cfgDiscIncomingAccount = (string)oRS.Fields.Item(16).Value;
                cfgDiscOutgoingAccount = (string)oRS.Fields.Item(17).Value;
                cfgDinnIncomingAccount = (string)oRS.Fields.Item(18).Value;
                cfgDinnOutgoingAccount = (string)oRS.Fields.Item(19).Value;
                cfgJCBIncomingAccount = (string)oRS.Fields.Item(20).Value;
                cfgJCBOutgoingAccount = (string)oRS.Fields.Item(21).Value;
                cfgeCheckIncomingAccount = (string)oRS.Fields.Item(22).Value;
                cfgCardEMTemplate = (string)oRS.Fields.Item(23).Value; //U_paymentGroup
                cfgAutoPreauth = (string)oRS.Fields.Item(24).Value;  //U_AutoSelect
                cfgBINoDelivery = (string)oRS.Fields.Item(25).Value;  //\"U_importPin\",
                cfgPreAuthEmail = (string)oRS.Fields.Item(26).Value;  //\"U_importSourceKey\"
                cfgProcessPreAuth = (string)oRS.Fields.Item(27).Value;
                cfgimportBPNextCode = (string)oRS.Fields.Item(28).Value;
                cfgimportBPHeaderCode = (string)oRS.Fields.Item(29).Value;
                cfgsendCustomerEmail = (string)oRS.Fields.Item(30).Value;
                cfgurlSandbox = (string)oRS.Fields.Item(31).Value;
                cfgurlProduction = (string)oRS.Fields.Item(32).Value;
                cfgmerchantEmail = (string)oRS.Fields.Item(33).Value;
                cfgMerchantReceipt = (string)oRS.Fields.Item(34).Value;
                cfgCustomerReceipt = (string)oRS.Fields.Item(35).Value;
                cfgRecurringBillingReceipt = (string)oRS.Fields.Item(36).Value;

                cfgDeclineReceipt = (string)oRS.Fields.Item(37).Value;
                cfgDeclineMerchantReceipt = (string)oRS.Fields.Item(38).Value;
                cfgaddressRequired = (string)oRS.Fields.Item(39).Value;
                cfgMerchantID = (string)oRS.Fields.Item(40).Value;
                cfgUseUserEmail = (string)oRS.Fields.Item(41).Value;
                cfgBatchQuery = (string)oRS.Fields.Item(42).Value;
                cfgBatchAutoMode = (string)oRS.Fields.Item(43).Value;
                cfgimportBPGroupName = (string)oRS.Fields.Item(44).Value;
                cfgrecurringBilling = (string)oRS.Fields.Item(45).Value;
                cfgBtnPos = (string)oRS.Fields.Item(46).Value;
                cfgDefaultTaxRate = (string)oRS.Fields.Item(47).Value;
                cfgNoTransOnInv = (string)oRS.Fields.Item(48).Value;
                cfgImportCust = (string)oRS.Fields.Item(49).Value;
                cfgshowStartup = (string)oRS.Fields.Item(50).Value;
                cfgEmailSubject = (string)oRS.Fields.Item(51).Value;
                cfgSaveReceipt = (string)oRS.Fields.Item(52).Value;
                cfgBatchPreauth = (string)oRS.Fields.Item(53).Value;
                cfgDeclineUnapprove = (string)oRS.Fields.Item(54).Value;
                cfgCaptureOnPrint = (string)oRS.Fields.Item(55).Value;
                cfgApprUDF = (string)oRS.Fields.Item(56).Value;
                cfgCaptureOnly = (string)oRS.Fields.Item(57).Value;
                cfgAutoSync = (string)oRS.Fields.Item("AutoSync").Value;
                cfgDefaultContactID = (string)oRS.Fields.Item("DefaultContactID").Value;
                cfgReceiptEmailUDF = (string)oRS.Fields.Item("ReceiptEmailUDF").Value;
                cfgAmountField = (string)oRS.Fields.Item(58).Value;
                cfgFieldSize = (string)oRS.Fields.Item(59).Value;
                cfgHasTerminal = (string)oRS.Fields.Item(60).Value;
                cfgConnectUseDP = (string)oRS.Fields.Item("ConnectUseDP").Value;


                if (cfgurlSandbox == null || cfgurlSandbox == "")
                    cfgurlSandbox = "https://sandbox.ebizcharge.com/soap/gate/CCEBDC0A";
                if (cfgurlProduction == null || cfgurlProduction == "")
                    cfgurlProduction = "https://secure.ebizcharge.com/soap/gate/CCEBDC0A";
                if (cfgUrl == cfgurlSandbox)
                    cfgIsSandbox = "Y";
                else
                    cfgIsSandbox = "N";
                if (cfgMerchantReceipt == null || cfgMerchantReceipt == "")
                    cfgMerchantReceipt = "vterm_merchant";
                if (cfgCustomerReceipt == null || cfgCustomerReceipt == "")
                    cfgCustomerReceipt = "vterm_customer";
                if (cfgMerchantReceipt == null || cfgRecurringBillingReceipt == "")
                    cfgRecurringBillingReceipt = "vterm_customer";
                if (cfgBtnPos == null || cfgBtnPos == "")
                    cfgBtnPos = "68";
                if (cfgDefaultTaxRate == null || cfgDefaultTaxRate == "")
                    cfgDefaultTaxRate = "8";
                if (cfgNoTransOnInv == null || cfgNoTransOnInv == "")
                    cfgNoTransOnInv = "Y";
                if (cfgImportCust == null || cfgImportCust == "")
                    cfgImportCust = "N";
                if (cfgshowStartup == null || cfgshowStartup == "")
                    cfgshowStartup = "Y";
                if (cfgBatchAutoMode == null || cfgBatchAutoMode == "")
                    cfgBatchAutoMode = "Y";
                if (cfgBatchPreauth == null || cfgBatchPreauth == "")
                    cfgBatchPreauth = "N";
                if (cfgDeclineUnapprove == null || cfgDeclineUnapprove == "")
                    cfgDeclineUnapprove = "N";
                if (cfgCaptureOnPrint == null || cfgCaptureOnPrint == "")
                    cfgCaptureOnPrint = "N";
                if (cfgCaptureOnPrint == null || cfgCaptureOnPrint == "")
                    cfgCaptureOnPrint = "N";
                if (cfgFieldSize == null || cfgFieldSize == "")
                    cfgFieldSize = "102,109,14";
                if (cfgHasTerminal == null || cfgHasTerminal == "")
                    cfgHasTerminal = "N";
                if (cfgBINoDelivery == null || cfgBINoDelivery == "")
                    cfgBINoDelivery = "N";
                if (cfgPreAuthEmail == null || cfgPreAuthEmail == "")
                    cfgPreAuthEmail = "N";
                if (cfgAutoPreauth == null || cfgAutoPreauth == "")
                    cfgAutoPreauth = "N";
                //ebiz.Url = cfgUrl;
                validDefaultTax();
            }

        }
        catch (Exception ex)
        {
            try
            {
                errorLog(ex);
                errorLog(oCompany.CompanyDB + "\r\n" + sql);
                showMessage("Please run setup on database: " + oCompany.CompanyDB);
            }
            catch (Exception) { };
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }

    }
    public double getInvoiceBalance(string id)
    {
        string sql = string.Format("select \"DocTotalFC\" , a.\"DocTotal\", a.\"PaidToDate\",a.\"PaidFC\" from OINV a where \"DocEntry\" = {0}", id);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        double rtn = 0;
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                double dtFC = (double)oRS.Fields.Item(0).Value;
                double dt = (double)oRS.Fields.Item(1).Value;
                double ptd = (double)oRS.Fields.Item(2).Value;
                double ptdFC = (double)oRS.Fields.Item(3).Value;
                if (dtFC > 0)
                    return dtFC - ptdFC;
                else
                    return dt - ptd;
            }


        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return rtn;

    }
    public bool IsTransactionExists(string refNum, string amt)
    {
        try
        {
            string sql = string.Format("select * from \"@CCTRANS\" where \"U_refNum\" = '{0}' and \"U_amount\" = '{1}'", refNum, amt);
            return !IsRSEOF(sql);

        }
        catch (Exception ex)
        {

        }
        return false;
    }
    public void maskCardOnTrans()
    {

        string sql = "update \"@CCTRANS\" set \"U_crCardNum\" = 'XXXXXXXXXX' || RIGHT(\"U_crCardNum\", 4) where Substring(\"U_crCardNum\", 5, 1) <> 'X'";
        try
        {
            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }

    public void initSetupCardAcct()
    {

        string sql = "delete from \"@CCCARDACCT\" where \"U_CardCode\" Not in (select \"CreditCard\" from OCRC) or \"U_CardCode\" is null";
        string sql2 = "insert into \"@CCCARDACCT\"(\"DocEntry\", \"U_CardCode\", \"U_CardName\")\r\n" +
"Select \"CreditCard\", \"CreditCard\", \"CardName\" from OCRC a where a.\"CreditCard\" not in (select \"U_CardCode\" from \"@CCCARDACCT\")";
        string sql3 = "Update \"@CCCARDACCT\" set \"U_CardName\" = a.\"CardName\" from OCRC a, \"@CCCARDACCT\" b where a.\"CreditCard\" = b.\"U_CardCode\"";
        try
        {
            execute(sql);
            execute(sql2);
            execute(sql3);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            errorLog(sql2);
            errorLog(sql3);

        }

    }
    public bool IsBPActive(string customerID)
    {


        string sql = string.Format("select \"frozenFor\" from OCRD where \"CardCode\" = '{0}'", customerID);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                string y = (string)oRS.Fields.Item(0).Value;
                if (y.ToUpper() == "Y")
                    return false;
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return true;
    }

    public void setJobCancelledToNULL()
    {

        string sql = string.Format("update \"@CCJOB\" set \"U_Cancelled\" = null where \"U_Cancelled\" <> 'Y'");
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }

    public void updateCardAcct(string id, string group, string cardtype, string pin, string sourceKey, string currency, string BranchID)
    {

        string sql = string.Format("update \"@CCCARDACCT\" set \"U_GroupName\"='{0}', \"U_CardType\"='{1}', \"U_Pin\"='{3}', \"U_SourceKey\"='{4}', \"U_Currency\"='{5}',\"U_BranchID\"='{6}'  where \"DocEntry\" = {2}", group.Replace("'", "''"), cardtype, id, pin, sourceKey, currency, BranchID);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void updateCardName(string id, string group, string cardname, string CCAccountID)
    {

        string sql = string.Format("Update \"@CCCUST\" set \"U_CardName\"='{0}', \"U_GroupName\"='{1}', \"U_CCAccountID\"='{2}' where \"DocEntry\" = {3}", cardname, group.Replace("'", "''"), CCAccountID, id);
        try
        {
            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public bool IsWalkIn(string id)
    {

        string sql = string.Format("select 1 from \"@CCWALKIN\" where \"U_CustomerID\"='{0}'", id);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (oRS.EoF)
                return false;
            else
                return true;

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return false;
    }
    public int getSODocEntryByInv(string id)
    {
        int nRet = 0;
        string sql = string.Format("select \"BaseEntry\", \"BaseType\" from INV1 where \"DocEntry\" = {0} ", id);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                int baseEntry = (int)oRS.Fields.Item(0).Value;
                int baseType = (int)oRS.Fields.Item(1).Value;
                if (baseType == 17)
                    return baseEntry;
                if (baseType == 15)
                {
                    sql = string.Format("select \"BaseEntry\", \"BaseType\" from DLN1 where \"DocEntry\" = {0} ", baseEntry);
                    oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRS.DoQuery(sql);
                    if (!oRS.EoF)
                    {
                        baseEntry = (int)oRS.Fields.Item(0).Value;
                        baseType = (int)oRS.Fields.Item(1).Value;
                        if (baseType == 17)
                            return baseEntry;

                    }
                }
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return nRet;
    }


    public double getBalance(string docEntry, String table = "OINV")
    {
        double ret = 0;
        String sql = "";
        if (table.Equals("OINV"))
        {
            sql = string.Format("select (\"DocTotal\" - \"PaidToDate\") as \"balance\" from {1} where \"DocEntry\" = {0} ", docEntry, table);
        }
        else if (table.Equals("ORDR"))
        {
            sql = String.Format("select Sum([DocTotal[) from (SELECT Min(T0.[DocTotal[) as [DocTotal[ FROM ODPI T0 INNER JOIN DPI1 T1 ON T0.[DocEntry[ = T1.[DocEntry[ inner JOIN ORDR T2 ON T1.[BaseEntry[= T2.[DocEntry[ inner join RCT2 T3 ON T3.[baseAbs[=T0.[DocEntry[ inner join ORCT T4 on T4.[DocEntry[ = T3.[DocNum[ WHERE T1.[BaseType[ = '17' and T0.[ObjType[ = '203' and  T0.[CANCELED[ = 'N' and T3.[InvType[='203' and T4.[Canceled[='N' and  T1.[BaseEntry[ = '{0}' group by T1.[DocEntry[ )T1", docEntry).Replace("[", "\"");
        }
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = (double)oRS.Fields.Item(0).Value;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return ret;
    }
    public string getCardAcctByCustNum(string custNum, ref string CustomerId)
    {
        string str = "";
        string sql = string.Format("select \"U_CCAccountID\", \"U_CustomerID\" from \"@CCCUST\" where \"U_CustNum\" = '{0}' ", custNum);
        try
        {
            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return str;
    }
    public string getCardAcctByCustomerID(string CustomerID)
    {
        string str = "";
        string sql = string.Format("select U_CCAccountID from \"@CCCUST\" where U_CustomerID='{0}' ", CustomerID);
        try
        {

            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return "";
    }
    public void CaptureUpdateCCTrans(CCTRAN t, string refNum)
    {
        string sql = string.Format("update \"@CCTRANS\" set \"U_command\" = 'cc:authonly-captured', \"U_InvoiceID\" = '{2}' where \"U_refNum\" = '{0}' and \"U_command\"= 'cc:authonly' and \"U_OrderID\" = '{1}' ", refNum, t.OrderID, t.InvoiceID);
        try
        {
            execute(sql);
            sql = string.Format("update \"@CCTRANS\" set \"U_OrderID\" = '{1}', \"U_InvoiceID\" = '{2}' where \"U_refNum\" = '{0}' and \"U_command\"= 'cc:capture' ", refNum, t.OrderID, t.InvoiceID);

            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public double getTransAmount(string SOID)
    {
        double amt = 0;
        if (SOID == "")
            return 0;
        string sql = string.Format("select sum(CAST(\"U_amount\" AS DECIMAL(10,2))) from \"@CCTRANS\" where \"U_OrderID\" = '{0}' and \"U_command\" not in  ('cc:void', 'cc:authonly-captured', 'cc:authonly-error') ", SOID);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {

            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                amt = (double)oRS.Fields.Item(0).Value;

            }
            else
            {

            }
            // refresh();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return amt;
    }
    public void initUpdateCFGWithOCRC(string field, string like)
    {

        string sql = string.Format("Update \"@CCCFG\" set {0} = (select \"CardName\" from OCRC where \"cardname\" like '%{1}%') where \"{0}\" is null", field, like);
        try
        {
            execute(sql);
            //refresh();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void initUpdateCardAcctWithOCRC(string field, string like)
    {

        string sql = string.Format("Update \"@CCCARDACCT\" set \"U_CardType\" = '{0}' from \"@CCCARDACCT\" a, OCRC b where  a.\"U_CardCode\" = b.\"CreditCard\" and b.\"CardName\" like '%{1}%' and a.\"U_CardType\" is null", field, like);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void updateDefaultEmailTemplate(string template)
    {
        string sql = "";
        try
        {

            sql = string.Format("Update \"@CCCFG\" set \"U_impWebServiceUrl\" = '{0}'", template);
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public string getDefaultEmailTemplate()
    {
        string sql = "";
        try
        {

            sql = string.Format("select \"U_impWebServiceUrl\" from \"@CCCFG\"");
            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return "";
    }
    public void updateConfig()
    {

        string sql = "";
        try
        {

            sql = string.Format("Update \"@CCCFG\" set \"U_Pin\"='{0}', \"U_SourceKey\"='{1}' ", cfgPin, cfgSourceKey);
            sql = sql + string.Format(", \"U_NegPayment\" = '{0}'", cfgNegativeIncomingPayment);

            sql = sql + string.Format(", \"U_importPin\" = '{0}'", cfgBINoDelivery);
            sql = sql + string.Format(", \"U_importSourceKey\" = '{0}'", cfgPreAuthEmail);
            sql = sql + string.Format(", \"U_ProcessPreAuth\" = '{0}'", cfgProcessPreAuth);
            sql = sql + string.Format(", \"U_sendCustomerEmail\" = '{0}'", cfgsendCustomerEmail);
            sql = sql + string.Format(", \"U_merchantEmail\" = '{0}'", cfgmerchantEmail.Replace("'", "''"));
            sql = sql + string.Format(", \"U_EmailSubject\" = '{0}'", cfgEmailSubject.Replace("'", "''"));
            sql = sql + string.Format(", \"U_MerchantReceipt\" = '{0}'", cfgMerchantReceipt.Replace("'", "''"));
            sql = sql + string.Format(", \"U_CustomerReceipt\" = '{0}'", cfgCustomerReceipt.Replace("'", "''"));
            sql = sql + string.Format(", \"U_urlSandbox\" = '{0}'", cfgurlSandbox);
            sql = sql + string.Format(", \"U_urlProduction\" = '{0}'", cfgurlProduction);
            sql = sql + string.Format(", \"U_addressRequired\" = '{0}'", cfgaddressRequired);
            sql = sql + string.Format(", \"U_merchantID\" = '{0}'", cfgMerchantID);
            sql = sql + string.Format(", \"U_BatchAutoMode\" = '{0}'", cfgBatchAutoMode);
            sql = sql + string.Format(", \"U_useUserEmail\" = '{0}'", cfgUseUserEmail);
            sql = sql + string.Format(", \"U_SaveReceipt\" = '{0}'", cfgSaveReceipt);
            sql = sql + string.Format(", \"U_recurringBilling\" = '{0}'", cfgrecurringBilling);
            sql = sql + string.Format(", \"U_DefaultTaxRate\" = '{0}'", cfgDefaultTaxRate);
            sql = sql + string.Format(", \"U_NoTransOnInv\" = '{0}'", cfgNoTransOnInv);
            sql = sql + string.Format(", \"U_BtnPos\" = '{0}'", cfgBtnPos);
            sql = sql + string.Format(", \"U_importCust\" = '{0}'", cfgImportCust);
            sql = sql + string.Format(", \"U_showStartup\" = '{0}'", cfgshowStartup);
            sql = sql + string.Format(", \"U_BatchPreauth\" = '{0}'", cfgBatchPreauth);
            sql = sql + string.Format(", \"U_DeclineUnapprove\" = '{0}'", cfgDeclineUnapprove);
            sql = sql + string.Format(", \"U_CaptureOnPrint\" = '{0}'", cfgCaptureOnPrint);
            sql = sql + string.Format(", \"U_ApprUDF\" = '{0}'", cfgApprUDF);
            sql = sql + string.Format(", \"U_CaptureOnly\" = '{0}'", cfgCaptureOnly);
            sql = sql + string.Format(", \"U_AmountField\" = '{0}'", cfgAmountField.Replace("'", "''"));
            sql = sql + string.Format(", \"U_FieldSize\" = '{0}'", cfgFieldSize);
            sql = sql + string.Format(", \"U_HasTerminal\" = '{0}'", cfgHasTerminal);
            sql = sql + string.Format(", \"U_PMAutoSelect\" = '{0}'", cfgAutoPreauth);
            sql = sql + string.Format(", \"U_paymentGroup\" = '{0}'", cfgCardEMTemplate);

            if (cfgIsSandbox == "Y")
                sql = sql + string.Format(", \"U_eBizChargeUrl\" = '{0}'", cfgurlSandbox);
            else
                sql = sql + string.Format(", \"U_eBizChargeUrl\" = '{0}'", cfgurlProduction);

            execute(sql);
            //ebiz.Url = cfgUrl;

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public string FindCustomerByID(string customerID)
    {
        string sql = string.Format("select \"CardCode\" from OCRD where \"CardCode\" = '{0}'", customerID);
        string str = "";
        try
        {
            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return str;
    }
    public string FindCustomerNameByID(string customerID)
    {
        string sql = string.Format("select \"CardName\" from OCRD where \"CardCode\" = '{0}'", customerID);
        string str = "";
        try
        {
            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return str;
    }
    public string FindItemByID(string ItemCode)
    {
        string sql = string.Format("select \"ItemCode\" from OITM where \"ItemCode\" = '{0}'", ItemCode);
        string str = "";
        try
        {
            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return str;
    }
    public string getCustEmail(string cardcode)
    {

        string sql = "";
        string str = "";
        try
        {
            sql = string.Format("select \"E_Mail\" from OCRD where \"CardCode\" = '{0}' ", cardcode);
            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return str;
    }
    public string getSQLString(string sql)
    {
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string str = "";
        try
        {
            oRS.DoQuery(sql);
            if (oRS.EoF)
                return "";
            else
            {
                str = (string)oRS.Fields.Item(0).Value;
            }
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return str;
    }
    public bool isPaymentMethodExists(string cardcode, string cardnum, string expdt)
    {

        string sql = "";
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            sql = string.Format("select 1 from \"@CCCUST\" where \"U_CustomerID\" = '{0}' and \"U_cardNum\" = '{1}' and \"U_expDate\" = '{2}' ", cardcode, cardnum, expdt);
            oRS.DoQuery(sql);
            if (oRS.EoF)
                return false;
            else
                return true;
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return false;
    }


    public string getAcctCode(string acctId)
    {
        string sRtn = "";
        acctId = acctId.Replace("-", "");

        string sql = string.Format("select \"AcctCode\" from OACT where \"FormatCode\" = '{0}'", acctId);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                return (string)oRS.Fields.Item(0).Value;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return sRtn;
    }

    public void deleteCCCust(int id)
    {

        string sql = string.Format("Delete from \"@CCCUST\" where \"DocEntry\"={0}", id);
        execute(sql);
    }

    public void voidUpdateTrans(string id)
    {

        string sql = string.Format("Update \"@CCTRANS\" set \"U_command\"= 'void-' || \"U_command\" where \"DocEntry\" = " + id);
        execute(sql);
    }

    public void execute(string sql)
    {
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        var CompDBName = oCompany.CompanyDB;
        var CompanyName = oCompany.CompanyName;

        try
        {
            oRS.DoQuery(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            if (bRunJob)
                throw (ex);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
    }
    public void setBatchResult(string refNum, string result, SAPbouiCOM.Matrix oMatrix, string Col, int Row)
    {
        setMatrixItem(oMatrix, Col, Row, result);
        string sql = string.Format("Update \"@CCTRANS\" set \"U_BatchResult\" = '{0}', \"U_BatchDate\"=CURRENT_DATE where \"U_command\"='cc:authonly' and \"U_refNum\"='{1}'", result, refNum);
        execute(sql);
    }
    public void setCaptureResult(string refNum, string result)
    {

        string sql = string.Format("Update \"@CCTRANS\" set \"U_command\" = '{0}' where \"U_command\"in('cc:authonly','cc:processing') and \"U_refNum\"='{1}'", result, refNum);
        execute(sql);
    }
    public void setErrorCheckResult(string refNum, string result)
    {

        string sql = string.Format("Update \"@CCTRANS\" set \"U_command\" = '{0}' where \"U_command\"in('cc:authonly-error') and \"U_refNum\"='{1}'", result, refNum);
        execute(sql);
    }
    public void BatchSetTransInvoiceID(string refNum, string InvoiceID, string OrderID, string accountID, string customerid)
    {

        string sql = string.Format("Update \"@CCTRANS\" set \"U_PaymentID\"=CAST((select max(\"TransId\") from OJDT where CAST(\"BaseRef\" AS VARCHAR) = \"U_SAPPaymentID\") AS VARCHAR), \"U_InvoiceID\" = '{0}', \"U_OrderID\"='{2}' , \"U_CCAccountID\"='{3}' , \"U_BatchResult\" = 'Success' " +
" , \"U_BatchDate\"=CURRENT_DATE, \"U_customerID\"='{4}'  where \"U_refNum\"='{1}'", InvoiceID, refNum, OrderID, accountID, customerid);
        execute(sql);
    }
    public int BatchGetDevliveryDocEntryByOrder(string id)
    {
        int nRet = 0;
        string sql = string.Format("select \"DocEntry\" from DLN1 where \"BaseEntry\" = {0} and \"BaseType\" = 17 ", id);

        try
        {
            return getSQLInt(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }

        return nRet;
    }
    /*
    public void formatConnectNum()
    {
        string sql = "Update \"@CCCONNECTINVOICE\" set \"U_PaidAmount\" = CAST(CAST(\"U_PaidAmount\" AS Decimal(10,2)) AS VARCHAR(50)), \"U_UploadedBalance\" = CAST(CAST(\"U_UploadedBalance\" AS DECIMAL(10,2)) AS VARCHAR(50))";

        try
        {
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        
    }
     * */
    public List<String> FindCustomer(string customerID)
    {
        List<String> list = new List<String>();
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        string sql = string.Format("select \"CardCode\" from OCRD where \"CardCode\" like '%{0}%' and \"CardType\"='C' order By \"CardCode\"", customerID);

        try
        {
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                string c = (string)oRS.Fields.Item(0).Value;
                list.Add(c);
                oRS.MoveNext();
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return list;
    }
    public List<CCTRAN> findTrans(string where)
    {
        string sql = "";
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        List<CCTRAN> list = new List<CCTRAN>();
        try
        {
            sql = "select \r\n" +
    "\"DocEntry\", \r\n" +
    "\"U_CCORDERID\", \r\n" +
    "\"U_CCAccountID\", \r\n" +
    "\"U_InvoiceID\", \r\n" +
    "\"U_PaymentID\", \r\n" +
    "\"U_DownPaymentInvID\", \r\n" +
    "\"U_OrderID\", \r\n" +
    "\"U_CreditMemoID\", \r\n" +
    "\"U_MethodID\", \r\n" +
    "\"U_customerID\", \r\n" +
    "\"U_customerName\", \r\n" +
    "\"U_CardHolder\", \r\n" +
    "\"U_creditRef\", \r\n" +
    "\"U_crCardNum\", \r\n" +
    "\"U_Description\", \r\n" +
    "\"U_recID\", \r\n" +
    "\"U_acsUrl\", \r\n" +
    "\"U_authAmount\", \r\n" +
    "\"U_authCode\", \r\n" +
    "\"U_avsResult\", \r\n" +
    "\"U_avsResultCode\", \r\n" +
    "\"U_batchNum\", \r\n" +
    "\"U_batchRef\", \r\n" +
    "\"U_cardCodeResult\", \r\n" +
    "\"U_cardCodeRCode\", \r\n" +
    "\"U_cardLevelResult\", \r\n" +
    "\"U_cardLevelRCode\", \r\n" +
    "\"U_conversionRate\", \r\n" +
    "\"U_convertedAmount\", \r\n" +
    "\"U_convertedCurrency\", \r\n" +
    "\"U_custNum\", \r\n" +
    "\"U_error\", \r\n" +
    "\"U_errorCode\", \r\n" +
    "\"U_isDuplicate\", \r\n" +
    "\"U_payload\", \r\n" +
    "\"U_profilerScore\", \r\n" +
    "\"U_profilerResponse\", \r\n" +
    "\"U_profilerReason\", \r\n" +
    "\"U_refNum\", \r\n" +
    "\"U_remainingBalance\", \r\n" +
    "\"U_result\", \r\n" +
    "\"U_resultCode\", \r\n" +
    "\"U_status\", \r\n" +
    "\"U_statusCode\", \r\n" +
    "\"U_vpasResultCode\", \r\n" +
    "\"U_recDate\", \r\n" +
    "\"U_command\", \r\n" +
    "\"U_amount\", \r\n" +
    "\"U_jobID\", \r\n" +
    "\"U_BatchResult\", \r\n" +
    "\"U_BatchDate\" \r\n from \"@CCTRANS\" ";
            sql = sql + where;
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                CCTRAN t = new CCTRAN();

                t.ccTRANSID = (int)oRS.Fields.Item(0).Value;
                t.CCORDERID = toInt(oRS.Fields.Item(1).Value);
                t.CCAccountID = toInt(oRS.Fields.Item(2).Value);
                t.InvoiceID = (string)oRS.Fields.Item(3).Value;
                t.PaymentID = (string)oRS.Fields.Item(4).Value;
                t.DownPaymentInvoiceID = (string)oRS.Fields.Item(5).Value;
                t.OrderID = (string)oRS.Fields.Item(6).Value;
                t.CreditMemoID = (string)oRS.Fields.Item(7).Value;
                t.MethodID = (string)oRS.Fields.Item(8).Value;
                t.customerID = (string)oRS.Fields.Item(9).Value;
                t.customerName = (string)oRS.Fields.Item(10).Value;
                t.CardHolder = (string)oRS.Fields.Item(11).Value;
                t.creditRef = (string)oRS.Fields.Item(12).Value;
                t.crCardNum = (string)oRS.Fields.Item(13).Value;
                t.Description = (string)oRS.Fields.Item(14).Value;
                t.recID = (string)oRS.Fields.Item(15).Value;
                t.acsUrl = (string)oRS.Fields.Item(16).Value;
                t.authAmount = (string)oRS.Fields.Item(17).Value;
                t.authCode = (string)oRS.Fields.Item(18).Value;
                t.avsResult = (string)oRS.Fields.Item(19).Value;
                t.avsResultCode = (string)oRS.Fields.Item(20).Value;
                t.batchNum = (string)oRS.Fields.Item(21).Value;
                t.batchRef = (string)oRS.Fields.Item(22).Value;
                t.cardCodeResult = (string)oRS.Fields.Item(23).Value;
                t.cardCodeResultCode = (string)oRS.Fields.Item(24).Value;
                t.cardLevelResult = (string)oRS.Fields.Item(25).Value;
                t.cardLevelResultCode = (string)oRS.Fields.Item(26).Value;
                t.conversionRate = (string)oRS.Fields.Item(27).Value;
                t.convertedAmount = (string)oRS.Fields.Item(28).Value;
                t.convertedAmountCurrency = (string)oRS.Fields.Item(29).Value;
                t.custNum = (string)oRS.Fields.Item(30).Value;
                t.error = (string)oRS.Fields.Item(31).Value;
                t.errorCode = (string)oRS.Fields.Item(32).Value;
                t.isDuplicate = (string)oRS.Fields.Item(33).Value;
                t.payload = (string)oRS.Fields.Item(34).Value;
                t.profilerScore = (string)oRS.Fields.Item(35).Value;
                t.profilerResponse = (string)oRS.Fields.Item(36).Value;
                t.profilerReason = (string)oRS.Fields.Item(37).Value;
                t.refNum = (string)oRS.Fields.Item(38).Value;
                t.remainingBalance = (string)oRS.Fields.Item(39).Value;
                t.result = (string)oRS.Fields.Item(40).Value;
                t.resultCode = (string)oRS.Fields.Item(41).Value;
                t.status = (string)oRS.Fields.Item(42).Value;
                t.statusCode = (string)oRS.Fields.Item(43).Value;
                t.vpasResultCode = (string)oRS.Fields.Item(44).Value;
                t.recDate = (DateTime)oRS.Fields.Item(45).Value;
                t.command = (string)oRS.Fields.Item(46).Value;
                t.amount = (string)oRS.Fields.Item(47).Value;
                t.jobID = toInt(oRS.Fields.Item(48).Value);
                t.BatchResult = (string)oRS.Fields.Item(49).Value;
                t.BatchDate = (DateTime)oRS.Fields.Item(50).Value;
                list.Add(t);
                oRS.MoveNext();
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return list;
    }
    public List<CCCUST> findCCCust(string where)
    {
        string sql = "";
        List<CCCUST> list = new List<CCCUST>();
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            sql = "select \r\n" +
    "\"DocEntry\", \r\n" +
"\"U_CustomerID\", \r\n" +
"\"U_CustNum\", \r\n" +
"\"U_MethodID\", \r\n" +
"\"U_active\", \r\n" +
"\"U_Declined\", \r\n" +
"\"U_methodDescription\", \r\n" +
"\"U_cardNum\", \r\n" +
"\"U_expDate\", \r\n" +
"\"U_cardCode\", \r\n" +
"\"U_routingNumber\", \r\n" +
"\"U_checkingAccount\", \r\n" +
"\"U_firstName\", \r\n" +
"\"U_lastName\", \r\n" +
"\"U_email\", \r\n" +
"\"U_street\", \r\n" +
"\"U_city\", \r\n" +
"\"U_state\", \r\n" +
"\"U_zip\", \r\n" +
"\"U_cardType\", \r\n" +
"\"U_GroupName\", \r\n" +
"\"U_CardName\", \r\n" +
"\"U_CCAccountID\"  \r\n" +
" from \"@CCCUST\" ";
            sql = sql + where;
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                CCCUST t = new CCCUST();

                t.recID = (int)oRS.Fields.Item(0).Value;
                t.CustomerID = (string)oRS.Fields.Item(1).Value;
                t.CustNum = (string)oRS.Fields.Item(2).Value;
                t.MethodID = (string)oRS.Fields.Item(3).Value;
                t.active = toChar(oRS.Fields.Item(4).Value);
                t.Declined = toChar(oRS.Fields.Item(5).Value);
                t.methodDescription = (string)oRS.Fields.Item(6).Value;
                t.cardNum = (string)oRS.Fields.Item(7).Value;
                t.expDate = (string)oRS.Fields.Item(8).Value;
                t.cardCode = (string)oRS.Fields.Item(9).Value;
                t.routingNumber = (string)oRS.Fields.Item(10).Value;
                t.checkingAccount = (string)oRS.Fields.Item(11).Value;
                t.firstName = (string)oRS.Fields.Item(12).Value;
                t.lastName = (string)oRS.Fields.Item(13).Value;
                t.email = (string)oRS.Fields.Item(14).Value;
                t.street = (string)oRS.Fields.Item(15).Value;
                t.city = (string)oRS.Fields.Item(16).Value;
                t.state = (string)oRS.Fields.Item(17).Value;
                t.zip = (string)oRS.Fields.Item(18).Value;
                t.cardType = (string)oRS.Fields.Item(19).Value;
                t.GroupName = (string)oRS.Fields.Item(20).Value;
                t.CardName = (string)oRS.Fields.Item(21).Value;
                t.CCAccountID = (string)oRS.Fields.Item(22).Value;
                list.Add(t);
                oRS.MoveNext();
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return list;
    }
    public List<CCCPayment> findCCPayment(string where)
    {
        string sql = "";
        List<CCCPayment> list = new List<CCCPayment>();
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            sql = "select \"DocEntry\", \r\n" +
"\"U_CustomerID\", \r\n" +
"\"U_InvoiceID\",  \r\n" +
"\"U_PaymentID\", \r\n" +
"\"U_Amount\", \r\n" +
"\"U_DateImported\" \r\n" +
 " from \"@CCCPAYMENT\" \r\n";
            sql = sql + where;
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                CCCPayment t = new CCCPayment();

                t.recID = (int)oRS.Fields.Item(0).Value;
                t.CustomerID = (string)oRS.Fields.Item(1).Value;
                t.PaymentID = (string)oRS.Fields.Item(2).Value;
                t.Amount = (string)oRS.Fields.Item(3).Value;
                t.DateImported = (DateTime)oRS.Fields.Item(4).Value;

                list.Add(t);
                oRS.MoveNext();
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return list;
    }
    public List<CCPayFormPayment> findCCPayFormPayment(string where)
    {
        string sql = "";
        List<CCPayFormPayment> list = new List<CCPayFormPayment>();
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            sql = "select \"DocEntry\", \r\n" +
"\"U_CustomerID\", \r\n" +
"\"U_InvoiceID\",  \r\n" +
"\"U_PaymentID\", \r\n" +
"\"U_Amount\", \r\n" +
"\"U_DateImported\" \r\n" +
 " from \"@CCPAYFORMPAYMENT\" \r\n";
            sql = sql + where;
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                CCPayFormPayment t = new CCPayFormPayment();

                t.recID = (int)oRS.Fields.Item(0).Value;
                t.CustomerID = (string)oRS.Fields.Item(1).Value;
                t.PaymentID = (string)oRS.Fields.Item(2).Value;
                t.Amount = (string)oRS.Fields.Item(3).Value;
                t.DateImported = (DateTime)oRS.Fields.Item(4).Value;

                list.Add(t);
                oRS.MoveNext();
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return list;
    }
    public List<CCJOB> findCCJOB(string where)
    {
        string sql = "";
        List<CCJOB> list = new List<CCJOB>();
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            sql = "select \"DocEntry\", \r\n" +
                "\"U_InvoiceID\",\r\n" +
"\"U_OrderID\",\r\n" +
"\"U_Description\",\r\n" +
"\"U_CustomerID\",\r\n" +
"\"U_Frequency\",\r\n" +
"\"U_StartDate\",\r\n" +
"\"U_EndDate\",\r\n" +
"\"U_NextRunDate\",\r\n" +
"\"U_LastRunDate\",\r\n" +
"\"U_Cancelled\",\r\n" +
"\"U_CancelledDate\", \r\n" +
"\"U_CancelledBy\",\r\n" +
"\"U_Amount\", \r\n" +
"\"U_PaymentID\", \r\n" +
"\"U_Result\" \r\n" +
" from \"@CCJOB\" \r\n";
            sql = sql + where;
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                CCJOB t = new CCJOB();

                t.jobID = (int)oRS.Fields.Item(0).Value;
                t.InvoiceID = toInt(oRS.Fields.Item(1).Value);
                t.OrderID = toInt(oRS.Fields.Item(2).Value);
                t.Description = (string)oRS.Fields.Item(3).Value;
                t.CustomerID = (string)oRS.Fields.Item(4).Value;
                t.Frequency = (string)oRS.Fields.Item(5).Value;
                t.StartDate = (DateTime)oRS.Fields.Item(6).Value;
                t.EndDate = (DateTime)oRS.Fields.Item(7).Value;
                t.NextRunDate = (DateTime)oRS.Fields.Item(8).Value;
                t.LastRunDate = (DateTime)oRS.Fields.Item(9).Value;
                t.Cancelled = toChar(oRS.Fields.Item(10).Value);
                t.CancelledDate = (DateTime)oRS.Fields.Item(11).Value;
                t.CancelledBy = (string)oRS.Fields.Item(12).Value;
                t.Amount = (string)oRS.Fields.Item(13).Value;
                t.PaymentID = (string)oRS.Fields.Item(14).Value;
                t.Result = (string)oRS.Fields.Item(15).Value;

                if (t.CancelledDate < oDT)
                    t.CancelledDate = null;
                if (t.EndDate < oDT)
                    t.EndDate = null;
                if (t.LastRunDate < oDT)
                    t.LastRunDate = null;

                list.Add(t);
                oRS.MoveNext();
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return list;
    }
    public List<CCConnectInvoice> findCCConnectInvoice(string where)
    {
        string sql = "";
        List<CCConnectInvoice> list = new List<CCConnectInvoice>();
        try
        {
            sql = "select \"DocEntry\", \r\n" +
               "\"U_InvoiceID\", \r\n" +
"\"U_CustomerID\", \r\n" +
"\"U_InvoiceGUID\", \r\n" +
"\"U_PaidAmount\", \r\n" +
"\"U_UploadedBalance\", \r\n" +
"\"U_Status\", \r\n" +
"\"U_UploadDate\", \r\n" +
"\"U_PaidDate\" \r\n" +
" from \"@CCCONNECTINVOICE\" \r\n";
            sql = sql + where;
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                CCConnectInvoice t = new CCConnectInvoice();

                t.recID = (int)oRS.Fields.Item(0).Value;
                t.InvoiceID = (string)oRS.Fields.Item(1).Value;
                t.CustomerID = (string)oRS.Fields.Item(2).Value;
                t.InvoiceGUID = (string)oRS.Fields.Item(3).Value;
                t.PaidAmount = (string)oRS.Fields.Item(4).Value;
                t.UploadedBalance = (string)oRS.Fields.Item(5).Value;
                t.Status = (string)oRS.Fields.Item(6).Value;
                t.UploadDate = (DateTime)oRS.Fields.Item(7).Value;
                t.PaidDate = (DateTime)oRS.Fields.Item(8).Value;
                if (t.PaidDate < oDT)
                    t.PaidDate = null;
                if (t.UploadDate < oDT)
                    t.UploadDate = null;


                list.Add(t);
                oRS.MoveNext();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return list;
    }
    public List<CCPayFormInvoice> findCCPayFormInvoice(string where)
    {
        string sql = "";
        List<CCPayFormInvoice> list = new List<CCPayFormInvoice>();
        try
        {
            sql = "select \"DocEntry\", \r\n" +
               "\"U_InvoiceID\", \r\n" +
"\"U_CustomerID\", \r\n" +
"\"U_InvoiceGUID\", \r\n" +
"\"U_PaidAmount\", \r\n" +
"\"U_UploadedBalance\", \r\n" +
"\"U_Status\", \r\n" +
"\"U_UploadDate\", \r\n" +
"\"U_PaidDate\" \r\n" +
" from \"@CCPAYFORMINVOICE\" \r\n";
            sql = sql + where;
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                CCPayFormInvoice t = new CCPayFormInvoice();

                t.recID = (int)oRS.Fields.Item(0).Value;
                t.InvoiceID = (string)oRS.Fields.Item(1).Value;
                t.CustomerID = (string)oRS.Fields.Item(2).Value;
                t.InvoiceGUID = (string)oRS.Fields.Item(3).Value;
                t.PaidAmount = (string)oRS.Fields.Item(4).Value;
                t.UploadedBalance = (string)oRS.Fields.Item(5).Value;
                t.Status = (string)oRS.Fields.Item(6).Value;
                t.UploadDate = (DateTime)oRS.Fields.Item(7).Value;
                t.PaidDate = (DateTime)oRS.Fields.Item(8).Value;
                if (t.PaidDate < oDT)
                    t.PaidDate = null;
                if (t.UploadDate < oDT)
                    t.UploadDate = null;


                list.Add(t);
                oRS.MoveNext();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return list;
    }
    public List<String> GetGroupNames()
    {
        List<String> list = new List<String>();
        string sql = string.Format("select \"GroupName\" from OCRG where \"GroupType\"='C'");
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                string c = (string)oRS.Fields.Item(0).Value;
                list.Add(c);
                oRS.MoveNext();
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return list;
    }
    public List<String> GetPaymentTerms()
    {
        List<String> list = new List<String>();
        string sql = string.Format("select \"PymntGroup\" from OCTG ");
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                string c = (string)oRS.Fields.Item(0).Value;
                list.Add(c);
                oRS.MoveNext();
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return list;
    }
    #region Export So Module
    public List<ExportBpModel> GetBusinessPartner(string where = "%")
    {
        string sql = "";
        List<ExportBpModel> list = new List<ExportBpModel>();
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            sql = String.Format("SELECT T0.[CardCode], T0.[CardName],T0.[Balance],T0.[E_Mail], IFNULL(T0.[U_EbizChargeID],'') as [EBiz],CASE WHEN IFNULL(T0.[U_EbizChargeID], '') = '' THEN 'False' ELSE 'True'END as [Sync] FROM OCRD T0 WHERE T0.[CardType] = 'C' and T0.[CardCode]='{0}'", where).Replace("]", "\"").Replace("[", "\"");

            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                ExportBpModel t = new ExportBpModel();

                t.CardCode = (string)oRS.Fields.Item("CardCode").Value;
                t.CardName = (string)oRS.Fields.Item("CardName").Value;
                t.Balance = 0;// (decimal)oRS.Fields.Item("Balance").Value;
                t.Email = (string)oRS.Fields.Item("E_Mail").Value;
                t.Sync = (string)oRS.Fields.Item("Sync").Value;
                t.EbizChargeKey = (string)oRS.Fields.Item("EBiz").Value;

                list.Add(t);
                oRS.MoveNext();
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return list;
    }
    public ExportBpModel GetSAPBP(string CardCode)
    {
        string sql = "";
        ExportBpModel oBP = null;
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            sql = String.Format("SELECT T0.[CardCode], T0.[CardName],T0.[Balance],T0.[E_Mail], IFNULL(T0.[U_EbizChargeID],'') as [EBiz],CASE WHEN IFNULL(T0.[U_EbizChargeID], '') = '' THEN 'False' ELSE 'True'END as [Sync] FROM OCRD T0 WHERE T0.[CardType] = 'C' and  [validFor]='Y' and T0.[CardCode]='{0}' ", CardCode).Replace("]", "\"").Replace("[", "\"");

            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {

                oBP = new ExportBpModel();
                oBP.CardCode = (string)oRS.Fields.Item("CardCode").Value;
                oBP.CardName = (string)oRS.Fields.Item("CardName").Value;
                oBP.Balance = ConvertDecimal(oRS.Fields.Item("Balance").Value);
                oBP.Email = (string)oRS.Fields.Item("E_Mail").Value;
                oBP.Sync = (string)oRS.Fields.Item("Sync").Value;
                oBP.EbizChargeKey = (string)oRS.Fields.Item("EBiz").Value;

                oRS.MoveNext();
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return oBP;
    }
    public ItemMasterSAP GetSAPItem(string ItemCode)
    {
        string sql = "";
        ItemMasterSAP oItem = null;
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            sql = String.Format("SELECT T0.[ItemCode],T0.[ItemName], T0.[OnHand], T1.[Price] FROM OITM T0  INNER JOIN ITM1 T1 ON T0.[ItemCode] = T1.[ItemCode] WHERE T1.[ItemCode] ='{0}' and [validFor]='Y' and  T1.[PriceList] =1", ItemCode).Replace("]", "\"").Replace("[", "\"");
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                oItem = new ItemMasterSAP();
                oItem.ItemCode = (string)oRS.Fields.Item("ItemCode").Value;
                oItem.ItemName = (string)oRS.Fields.Item("ItemName").Value;
                oItem.onHand = ConvertDecimal(oRS.Fields.Item("OnHand").Value);
                oItem.Price = ConvertDecimal(oRS.Fields.Item("Price").Value);

                oRS.MoveNext();
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return oItem;
    }
    public List<String> GetAllKeyList(string Table)
    {
        string sql = "";
        List<String> list = new List<String>();
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            sql = String.Format("SELECT \"Code\" from \"@{0}\"", Table).Replace("]", "\"").Replace("[", "\"");

            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                list.Add((string)oRS.Fields.Item("Code").Value);
                oRS.MoveNext();
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return list;
    }
    public string getDataFromLOG(String Table, String Select, String Where)
    {
        String ret = "";

        var sql = String.Format("select MIN({0}) from \"@{1}\" where {2}", Select, Table, Where);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = ConvertString(oRS.Fields.Item(0).Value);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return ret;
    }
    public string getMaxCCSOLOG()
    {
        String ret = "";

        var sql = "select IFNULL(max(cast(\"Code\" as int) ),0) + 1 from \"@CCSOLOG\"";
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = ConvertString(oRS.Fields.Item(0).Value);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return ret;
    }
    public string getMaxCCInvLOG()
    {
        String ret = "";

        var sql = "select IFNULL(max(cast(\"Code\" as int) ),0) + 1 from \"@CCINVLOG\"";
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = ConvertString(oRS.Fields.Item(0).Value);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return ret;
    }
    public string getMaxCCLOG(String Table)
    {
        String ret = "";

        var sql = String.Format("select IFNULL(max(cast(\"Code\" as int) ),0) + 1 from \"@{0}\"", Table);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = ConvertString(oRS.Fields.Item(0).Value);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return ret;
    }
    public string getOrderEbizUrl()
    {
        String ret = "";

        var sql = "select IFNULL(max(\"Code\"),0)+1 from \"@CCSOLOG\"";
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = ConvertString(oRS.Fields.Item(0).Value);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return ret;
    }
    public string getPreAuthStatus()
    {
        String ret = "";

        var sql = "select top 1 [U_ProcessPreAuth[ as Proceess from [@CCCFG[".Replace("[", "\"");
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = ConvertString(oRS.Fields.Item(0).Value);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return ret;
    }
    public string getSODocNum(int docEntry)
    {
        String ret = "";
        string sql = string.Format("select \"DocNum\" from ORDR where \"DocEntry\" = {0} ", docEntry);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = ConvertString(oRS.Fields.Item(0).Value);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return ret;
    }
    public string GetSOOfConnect(string ConnectDocNum)
    {
        String ret = "";
        string sql = string.Format("select \"DocNum\" from ORDR where \"U_ConnectDocNum\" = {0} ", ConnectDocNum);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = ConvertString(oRS.Fields.Item(0).Value);
            }
            else
            {
                ret = "NoResult";
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return ret;
    }
    public string getInvoiceDocNum(int docEntry)
    {
        String ret = "";
        string sql = string.Format("select \"DocNum\" from OINV where \"DocEntry\" = {0} ", docEntry);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = ConvertString(oRS.Fields.Item(0).Value);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return ret;
    }
    public string getSOpaymentType(string docNum)
    {
        String ret = "";
        string sql = string.Format("select \"U_EbizChargeMarkPayment\" from ORDR where \"DocNum\" = '{0}' ", docNum);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = ConvertString(oRS.Fields.Item(0).Value);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return ret;
    }
    public string getSODownPaymentStatus(string docNum)
    {
        String ret = "";
        string sql = string.Format("select count(*) from DPI1 where \"BaseRef\"='{0}' ", docNum);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = ConvertString(oRS.Fields.Item(0).Value);

            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return ret;
    }
    public string getSOAuthonlyAlreadyPostedStatus(string docNum)
    {
        String ret = "";
        string sql = string.Format("select count(*) from \"@CCTRANS\" where \"U_OrderID\"='{0}' ", docNum);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = ConvertString(oRS.Fields.Item(0).Value);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return ret;
    }
    public string getSODocEntry(string docNum)
    {
        String ret = "";
        string sql = string.Format("select \"DocEntry\" from ORDR where \"DocNum\" = '{0}' ", docNum);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = ConvertString(oRS.Fields.Item(0).Value);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return ret;
    }

    public decimal getSOAmountDue(string docNum)
    {
        var DocEntry = getSODocEntry(docNum);
        decimal ret = 0;
        var sql = String.Format("SELECT  sum(T0.\"DpmAmnt\") FROM ODPI T0  INNER JOIN DPI1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" inner join RDR1 on RDR1.\"DocEntry\" = T1.\"BaseEntry\" where RDR1.\"DocEntry\" ={0} and T0.\"DocStatus\" = 'O' and T0.\"CANCELED\" = 'N'", DocEntry);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = ConvertDecimal(oRS.Fields.Item(0).Value);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return ret;
    }
    public String getSOPaymentTypeConnect(string docNum)
    {
        var DocEntry = getSODocEntry(docNum);
        String ret = "";
        var sql = String.Format("select CASE  WHEN IFNULL(R.[U_EbizChargeMarkPayment[, '') = 'sales' THEN 'Deposit'  WHEN IFNULL(R.[U_EbizChargeMarkPayment[, '') = 'AuthOnly' THEN 'Pre-auth' ELSE '' END as [Type[ from ORDR R where R.[DocEntry[='{0}'".Replace("[", "\""), DocEntry);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                ret = ConvertString(oRS.Fields.Item(0).Value);
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
        }
        return ret;
    }
    public SalesOrder GetSO(string DocNum)
    {
        string sql = "";
        SalesOrder oSalesOrder = new SalesOrder();
        oSalesOrder.Software = Software;
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        List<Item> ListItem = new List<Item>();
        try
        {
            sql = String.Format("SELECT T0.[PaidToDate[,T0.[DocCur[,T0.[DocNum[,T0.[DocEntry[,TO_NVARCHAR(T0.[DocDate[, 'YYYY/MM/DD')[DocDate[,TO_NVARCHAR(T0.[DocDueDate[, 'YYYY/MM/DD')[DeliveryDate[, T0.[CardCode[, T0.[CardName[,OCRD.[E_Mail] ,T0.[Address[, T0.[DocTotal[, T1.[ItemCode[,T1.[LineNum], T1.[Quantity[, T1.[Price[ as [UnitPrice[ , T1.[LineTotal[ , T0.[{1}] FROM ORDR T0 INNER JOIN RDR1 T1 ON T0.[DocEntry[ = T1.[DocEntry] inner join OCRD on OCRD.[CardCode]=T0.[CardCode] WHERE T0.[DocNum[ = '{0}'", DocNum, "U_" + U_EbizChargeKey).Replace("]", "\"").Replace("[", "\"");
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {

                oSalesOrder.CustomerId = (string)oRS.Fields.Item("CardCode").Value;
                //oSalesOrder.BillingAddress = new Address();
                //oSalesOrder.BillingAddress.Address1 = (string)oRS.Fields.Item("Address").Value;
                //                oSalesOrder. = ConvertString(oRS.Fields.Item("DocNum").Value);
                oSalesOrder.SalesOrderNumber = ConvertString(oRS.Fields.Item("DocNum").Value);
                oSalesOrder.PoNum = ConvertString(oRS.Fields.Item("DocEntry").Value);
                oSalesOrder.Currency = ConvertString(oRS.Fields.Item("DocCur").Value);
                oSalesOrder.Date = ConvertString(oRS.Fields.Item("DocDate").Value);
                oSalesOrder.DueDate = ConvertString(oRS.Fields.Item("DeliveryDate").Value);
                oSalesOrder.Amount = ConvertDecimal(oRS.Fields.Item("DocTotal").Value);
                oSalesOrder.AmountDue = ConvertDecimal(oRS.Fields.Item("DocTotal").Value)- ConvertDecimal(getBalance(oSalesOrder.PoNum, "ORDR"));
                oSalesOrder.EmailTemplateID = ConvertString(oRS.Fields.Item("E_Mail").Value);
                try
                {
                    oSalesOrder.SalesOrderInternalId = ConvertString(oRS.Fields.Item("U_" + U_EbizChargeKey).Value);
                }
                catch (Exception ex)
                {
                    errorLog(ex);
                }
                Item item = new Item();
                item.ItemId = ConvertString(oRS.Fields.Item("ItemCode").Value);
                item.TotalLineAmount = ConvertDecimal(oRS.Fields.Item("LineTotal").Value);
                item.Qty = ConvertDecimal(oRS.Fields.Item("Quantity").Value);
                item.UnitPrice = ConvertDecimal(oRS.Fields.Item("UnitPrice").Value);

                //item.ItemLineNumber = Convert.ToInt32(oRS.Fields.Item("LineNum").Value)+"";

                ListItem.Add(item);
                oRS.MoveNext();

            }
            Item[] Lineitems = new Item[ListItem.Count];
            for (int i = 0; i < ListItem.Count; i++)
            {
                Lineitems[i] = ListItem[i];
            }
            oSalesOrder.Items = Lineitems;
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return oSalesOrder;
    }

    public Invoice GetInvoice(string DocNum)
    {
        string sql = "";
        Invoice oInvoice = new Invoice();
        oInvoice.Software = Software;
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        List<Item> ListItem = new List<Item>();
        try
        {

            sql = String.Format("SELECT cast(T0.[DocEntry[ as nvarchar(50)) as [DocEntry[,T0.[DocCur[,T0.[DocNum[,TO_NVARCHAR(T0.[DocDate[, 'YYYY/MM/DD')[DocDate[,TO_NVARCHAR(T0.[DocDueDate[, 'YYYY/MM/DD')[DeliveryDate[, T0.[CardCode[, T0.[CardName[,OCRD.[E_Mail] ,T0.[Address[, T0.[DocTotal[, T1.[ItemCode[,T1.[LineNum], T1.[Quantity[, T1.[Price[ as [UnitPrice[ , T1.[LineTotal[ , T0.[{1}] FROM OINV T0 INNER JOIN INV1 T1 ON T0.[DocEntry[ = T1.[DocEntry] inner join OCRD on OCRD.[CardCode]=T0.[CardCode] WHERE T0.[DocNum[ = '{0}'", DocNum, "U_" + U_EbizChargeKey).Replace("]", "\"").Replace("[", "\"");
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {

                oInvoice.CustomerId = (string)oRS.Fields.Item("CardCode").Value;
                oInvoice.PoNum = (string)oRS.Fields.Item("DocEntry").Value;
                oInvoice.InvoiceNumber = ConvertString(oRS.Fields.Item("DocNum").Value);
                oInvoice.InvoiceDate = ConvertString(oRS.Fields.Item("DocDate").Value);
                oInvoice.InvoiceDueDate = ConvertString(oRS.Fields.Item("DeliveryDate").Value);
                oInvoice.InvoiceAmount = ConvertDecimal(oRS.Fields.Item("DocTotal").Value);//
                oInvoice.AmountDue = ConvertDecimal(oRS.Fields.Item("DocTotal").Value);//
                oInvoice.Currency = ConvertString(oRS.Fields.Item("DocCur").Value);
                oInvoice.EmailTemplateID = ConvertString(oRS.Fields.Item("E_Mail").Value);
                try
                {
                    oInvoice.InvoiceInternalId = ConvertString(oRS.Fields.Item("U_" + U_EbizChargeKey).Value);
                }
                catch (Exception ex)
                {
                    errorLog(ex);
                }
                Item item = new Item();
                item.ItemId = ConvertString(oRS.Fields.Item("ItemCode").Value);
                item.TotalLineAmount = ConvertDecimal(oRS.Fields.Item("LineTotal").Value);
                item.Qty = ConvertDecimal(oRS.Fields.Item("Quantity").Value);
                item.UnitPrice = ConvertDecimal(oRS.Fields.Item("UnitPrice").Value);

                //item.ItemLineNumber = Convert.ToInt32(oRS.Fields.Item("LineNum").Value)+"";

                ListItem.Add(item);
                oRS.MoveNext();

            }
            Item[] Lineitems = new Item[ListItem.Count];
            for (int i = 0; i < ListItem.Count; i++)
            {
                Lineitems[i] = ListItem[i];
            }
            oInvoice.Items = Lineitems;
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return oInvoice;
    }
    public string getSOInternalID(string DoNum)
    {
        string sRtn = "";


        string sql = string.Format("select top 1 \"U_EbizChargeID\" from ORDR where \"DocNum\" = '{0}'", DoNum);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                return (string)oRS.Fields.Item(0).Value;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return sRtn;
    }

    public string getEbizURL(string DocNum)
    {
        string sRtn = "";
        string sql = string.Format("select top 1 \"U_EbizChargeURL\" from ORDR where \"DocNum\" = '{0}'", DocNum);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                return (string)oRS.Fields.Item(0).Value;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return sRtn;
    }
    public void updateSOPayment(string DocNum)
    {
        string sql = string.Format(" update ORDR set \"U_EbizChargeMarkPayment\"='Paid' where \"DocNum\"='{0}'", DocNum);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void updateSOPayment(string DocNum, String Status)
    {
        string sql = string.Format(" update ORDR set \"U_EbizChargeMarkPayment\"='{1}' where \"DocNum\"='{0}'", DocNum, Status);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void AddSOLog(LogTable oExportSOLogTable)
    {
        try
        {
            var Id = getMaxCCSOLOG();
            string insertQury = "";
            if (oExportSOLogTable.CreateDt != DateTime.MinValue)//[U_RefNum[,[U_CustomerID[,[U_DocEntry[,[U_Balance[,[U_PaidAmt[
            {
                insertQury = String.Format("insert into [@CCSOLOG[ ([Code[,[Name[,[U_DocNum[,[U_DocType[,[U_Status[,[U_PaymentStatus[,[U_CreateDt[,[U_UpdateDt[,[U_RefNum[,[U_CustomerID[,[U_DocEntry[,[U_Balance[,[U_PaidAmt[,[U_DocState[) values ({0},{0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}') ", Id, oExportSOLogTable.DocNum, oExportSOLogTable.DocType, oExportSOLogTable.Status, oExportSOLogTable.PaymentStatus, oExportSOLogTable.CreateDt.ToString("yyyy/MM/dd"), oExportSOLogTable.UpdateDt.ToString("yyyy/MM/dd"), oExportSOLogTable.RefNum, oExportSOLogTable.CardCode, oExportSOLogTable.DocEntry, oExportSOLogTable.Amount, 0, oExportSOLogTable.DocState).Replace("[", "\"");
            }
            else
            {
                insertQury = String.Format("insert into [@CCSOLOG[ ([Code[,[Name[,[U_DocNum[,[U_DocType[,[U_Status[,[U_PaymentStatus[,[U_UpdateDt[,[U_RefNum[,[U_CustomerID[,[U_DocEntry[,[U_Balance[,[U_PaidAmt[,[U_DocState[) values ({0},{0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}') ", Id, oExportSOLogTable.DocNum, oExportSOLogTable.DocType, oExportSOLogTable.Status, oExportSOLogTable.PaymentStatus, oExportSOLogTable.UpdateDt.ToString("yyyy/MM/dd"), oExportSOLogTable.RefNum, oExportSOLogTable.CardCode, oExportSOLogTable.DocEntry, oExportSOLogTable.Amount, 0, oExportSOLogTable.DocState).Replace("[", "\"");
            }
            try
            {
                execute(insertQury);
            }
            catch (Exception ex)
            {
                errorLog(ex);
                errorLog(insertQury);
            }

        }
        catch (Exception ex)
        {
            var error = ex.Message;
        }

    }

    public void AddInvoiceLog(LogTable oExportLogTable)
    {
        try
        {
            var Id = getMaxCCInvLOG();
            string insertQury = "";
            if (oExportLogTable.CreateDt != DateTime.MinValue)//DocEntry,CustomerID,Balance
            {
                insertQury = String.Format("insert into [@CCINVLOG[ ([Code[,[Name[,[U_DocNum[,[U_DocType[,[U_Status[,[U_PaymentStatus[,[U_CreateDt[,[U_UpdateDt[,[U_DocEntry[,[U_CustomerID[,[U_Balance[,[U_RefNum[) values ({0},{0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}') ", Id, oExportLogTable.DocNum, oExportLogTable.DocType, oExportLogTable.Status, oExportLogTable.PaymentStatus, oExportLogTable.CreateDt.ToString("yyyy/MM/dd"), oExportLogTable.UpdateDt.ToString("yyyy/MM/dd"), oExportLogTable.DocEntry, oExportLogTable.CardCode, oExportLogTable.Balance, oExportLogTable.RefNum).Replace("[", "\"");
            }
            else
            {
                insertQury = String.Format("insert into [@CCINVLOG[ ([Code[,[Name[,[U_DocNum[,[U_DocType[,[U_Status[,[U_PaymentStatus[,[U_UpdateDt[,[U_DocEntry[,[U_CustomerID[,[U_Balance[,[U_RefNum[) values ({0},{0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}') ", Id, oExportLogTable.DocNum, oExportLogTable.DocType, oExportLogTable.Status, oExportLogTable.PaymentStatus, oExportLogTable.UpdateDt.ToString("yyyy/MM/dd"), oExportLogTable.DocEntry, oExportLogTable.CardCode, oExportLogTable.Balance, oExportLogTable.RefNum).Replace("[", "\"");
            }
            try
            {
                execute(insertQury);
            }
            catch (Exception ex)
            {
                errorLog(ex);
                errorLog(insertQury);
            }

        }
        catch (Exception ex)
        {
            var error = ex.Message;
        }

    }
    #region Datamember

    public static SAPbobsCOM.GeneralService oGeneralService = null;
    public static SAPbobsCOM.GeneralData oGeneralData = null;
    public static SAPbobsCOM.GeneralDataParams oGeneralParams = null;
    public static SAPbobsCOM.CompanyService sCmp = null;
    public static SAPbobsCOM.GeneralData oChild = null;
    public static SAPbobsCOM.GeneralDataCollection oChildren = null;
    #endregion
    public void AddUploadDataLog(LogTable oLogTable, string Table)
    {
        try
        {
            var Id = getMaxCCLOG(Table);

            if (Table.Equals("CCITEMLOGTAB"))
            {
                SAPbobsCOM.UserTable oUsrTbl = oCompany.UserTables.Item("CCITEMLOGTAB");
                oUsrTbl.Code = Id;
                oUsrTbl.Name = Id;
                oUsrTbl.UserFields.Fields.Item("U_ItemCode").Value = oLogTable.ItemCode;
                oUsrTbl.UserFields.Fields.Item("U_Status").Value = oLogTable.Status;
                oUsrTbl.UserFields.Fields.Item("U_EbizKey").Value = oLogTable.Ebiz;
                oUsrTbl.UserFields.Fields.Item("U_CreateDt").Value = DateTime.Now;
                oUsrTbl.UserFields.Fields.Item("U_UpdateDt").Value = DateTime.Now;
                int Res = oUsrTbl.Add();
                int erroCode = 0;
                string errDescr = "";
                if (Res != 0)
                {
                    oCompany.GetLastError(out erroCode, out errDescr);
                    trace(errDescr);
                }
            }

            else if (Table.Equals("CCBPLOGTAB"))
            {
                SAPbobsCOM.UserTable oUsrTbl = oCompany.UserTables.Item(Table);
                oUsrTbl.Code = Id;
                oUsrTbl.Name = Id;
                oUsrTbl.UserFields.Fields.Item("U_CardCode").Value = oLogTable.CardCode;
                oUsrTbl.UserFields.Fields.Item("U_Status").Value = oLogTable.Status;
                oUsrTbl.UserFields.Fields.Item("U_EbizKey").Value = oLogTable.Ebiz;
                oUsrTbl.UserFields.Fields.Item("U_CreateDt").Value = DateTime.Now;
                oUsrTbl.UserFields.Fields.Item("U_UpdateDt").Value = DateTime.Now;
                int Res = oUsrTbl.Add();
                int erroCode = 0;
                string errDescr = "";
                if (Res != 0)
                {
                    oCompany.GetLastError(out erroCode, out errDescr);
                    trace(errDescr);
                }
            }
            else if (Table.Equals("CCSO"))
            {
                SAPbobsCOM.UserTable oUsrTbl = oCompany.UserTables.Item(Table);
                oUsrTbl.Code = Id;
                oUsrTbl.Name = Id;
                //DocNum,DocType,Status,PaymentStatus,DocSent,Email,MethodDigit,CardCode,AmountPaid
                oUsrTbl.UserFields.Fields.Item("U_CardCode").Value = oLogTable.CardCode;
                oUsrTbl.UserFields.Fields.Item("U_Status").Value = oLogTable.Status;
                oUsrTbl.UserFields.Fields.Item("U_DocNum").Value = oLogTable.DocNum;
                oUsrTbl.UserFields.Fields.Item("U_Type").Value = oLogTable.Type;
                oUsrTbl.UserFields.Fields.Item("U_DocType").Value = "SO";
                //  oUsrTbl.UserFields.Fields.Item("U_PaymentStatus").Value = oExportSOLogTable.PaymentStatus;
                //oUsrTbl.UserFields.Fields.Item("U_DocSent").Value = oExportSOLogTable.DocSent;
                oUsrTbl.UserFields.Fields.Item("U_DocPaidDate").Value = oLogTable.DatePaid;
                oUsrTbl.UserFields.Fields.Item("U_Email").Value = oLogTable.Email;
                oUsrTbl.UserFields.Fields.Item("U_MethodDigit").Value = oLogTable.Method;
                oUsrTbl.UserFields.Fields.Item("U_AmountPaid").Value = oLogTable.AmountPaid.ToString();
                oUsrTbl.UserFields.Fields.Item("U_CreateDt").Value = DateTime.Now;
                oUsrTbl.UserFields.Fields.Item("U_UpdateDt").Value = DateTime.Now;//
                int Res = oUsrTbl.Add();
                int erroCode = 0;
                string errDescr = "";
                if (Res != 0)
                {
                    oCompany.GetLastError(out erroCode, out errDescr);
                    trace(errDescr);
                }
            }
            else if (Table.Equals("CCBPIMP"))
            {
                SAPbobsCOM.UserTable oUsrTbl = oCompany.UserTables.Item(Table);
                oUsrTbl.Code = Id;
                oUsrTbl.Name = Id;
                //CardCode,CardName,Email,Balance,Software,Status,CreateDt,UpdateDt
                oUsrTbl.UserFields.Fields.Item("U_CardCode").Value = oLogTable.CardCode;
                oUsrTbl.UserFields.Fields.Item("U_CardName").Value = oLogTable.CardName;
                oUsrTbl.UserFields.Fields.Item("U_Email").Value = oLogTable.Email;
                oUsrTbl.UserFields.Fields.Item("U_Balance").Value = oLogTable.Balance.ToString();
                oUsrTbl.UserFields.Fields.Item("U_Software").Value = oLogTable.Software.ToString();
                oUsrTbl.UserFields.Fields.Item("U_Status").Value = oLogTable.Status;
                oUsrTbl.UserFields.Fields.Item("U_CreateDt").Value = DateTime.Now;
                oUsrTbl.UserFields.Fields.Item("U_UpdateDt").Value = DateTime.Now;
                int Res = oUsrTbl.Add();
                int erroCode = 0;
                string errDescr = "";
                if (Res != 0)
                {
                    oCompany.GetLastError(out erroCode, out errDescr);
                    trace(errDescr);
                }
            }
            else if (Table.Equals("CCITMIMP"))
            {
                SAPbobsCOM.UserTable oUsrTbl = oCompany.UserTables.Item(Table);
                oUsrTbl.Code = Id;
                oUsrTbl.Name = Id;
                //CardCode,CardName,Email,Balance,Software,Status,CreateDt,UpdateDt
                oUsrTbl.UserFields.Fields.Item("U_ItemCode").Value = oLogTable.ItemCode;
                oUsrTbl.UserFields.Fields.Item("U_ItemName").Value = oLogTable.ItemName;
                oUsrTbl.UserFields.Fields.Item("U_Quantity").Value = oLogTable.Quantity.ToString();
                oUsrTbl.UserFields.Fields.Item("U_UnitPrice").Value = oLogTable.UnitPrice.ToString();
                oUsrTbl.UserFields.Fields.Item("U_Software").Value = oLogTable.Software.ToString();
                oUsrTbl.UserFields.Fields.Item("U_Status").Value = oLogTable.Status;
                oUsrTbl.UserFields.Fields.Item("U_CreateDt").Value = DateTime.Now;
                oUsrTbl.UserFields.Fields.Item("U_UpdateDt").Value = DateTime.Now;
                int Res = oUsrTbl.Add();
                int erroCode = 0;
                string errDescr = "";
                if (Res != 0)
                {
                    oCompany.GetLastError(out erroCode, out errDescr);
                    trace(errDescr);
                }
            }
            else if (Table.Equals("CCSOIMP"))
            {
                SAPbobsCOM.UserTable oUsrTbl = oCompany.UserTables.Item(Table);
                oUsrTbl.Code = Id;
                oUsrTbl.Name = Id;
                //[U_DocNum],[U_CardCode],[U_Customer],[U_DatePaid],[U_Software],[U_Status],[U_CreateDt]
                oUsrTbl.UserFields.Fields.Item("U_DocNum").Value = oLogTable.DocNum;
                oUsrTbl.UserFields.Fields.Item("U_DocEntry").Value = oLogTable.DocEntry;
                oUsrTbl.UserFields.Fields.Item("U_CardCode").Value = oLogTable.CardCode;
                oUsrTbl.UserFields.Fields.Item("U_Customer").Value = oLogTable.CardName.ToString();
                oUsrTbl.UserFields.Fields.Item("U_DatePaid").Value = oLogTable.DatePaid.ToString();
                oUsrTbl.UserFields.Fields.Item("U_Software").Value = oLogTable.Software.ToString();
                oUsrTbl.UserFields.Fields.Item("U_Status").Value = oLogTable.Status;
                oUsrTbl.UserFields.Fields.Item("U_CreateDt").Value = DateTime.Now;
                oUsrTbl.UserFields.Fields.Item("U_UpdateDt").Value = DateTime.Now;
                int Res = oUsrTbl.Add();
                int erroCode = 0;
                string errDescr = "";
                if (Res != 0)
                {
                    oCompany.GetLastError(out erroCode, out errDescr);
                    trace(errDescr);
                }
            }//CCSOPIMP
            else if (Table.Equals("CCSOPIMP"))
            {
                SAPbobsCOM.UserTable oUsrTbl = oCompany.UserTables.Item(Table);
                oUsrTbl.Code = Id;
                oUsrTbl.Name = Id;
                //[U_DocNum],[U_CardCode],[U_Customer],[U_DatePaid],[U_Software],[U_Status],[U_CreateDt]
                oUsrTbl.UserFields.Fields.Item("U_DocNum").Value = oLogTable.DocNum;
                oUsrTbl.UserFields.Fields.Item("U_DocEntry").Value = oLogTable.DocEntry;
                oUsrTbl.UserFields.Fields.Item("U_CardCode").Value = oLogTable.CardCode;
                oUsrTbl.UserFields.Fields.Item("U_Customer").Value = oLogTable.CardName.ToString();
                oUsrTbl.UserFields.Fields.Item("U_PaymentType").Value = oLogTable.Method.ToString();
                oUsrTbl.UserFields.Fields.Item("U_DatePaid").Value = oLogTable.DatePaid.ToString();
                oUsrTbl.UserFields.Fields.Item("U_AmtPaid").Value = oLogTable.AmountPaid.ToString();
                oUsrTbl.UserFields.Fields.Item("U_Software").Value = oLogTable.Software.ToString();
                oUsrTbl.UserFields.Fields.Item("U_Status").Value = oLogTable.Status;
                oUsrTbl.UserFields.Fields.Item("U_CreateDt").Value = DateTime.Now;
                oUsrTbl.UserFields.Fields.Item("U_UpdateDt").Value = DateTime.Now;
                int Res = oUsrTbl.Add();
                int erroCode = 0;
                string errDescr = "";
                if (Res != 0)
                {
                    oCompany.GetLastError(out erroCode, out errDescr);
                    trace(errDescr);
                }
            }
        }
        catch (Exception ex)
        {
            var error = ex.Message;
            trace(ex.Message);
            errorLog(ex);
        }

    }
    #endregion
    public string getCardCodeByGroup(string grp)
    {
        string rtn = "";
        string sql = string.Format("select \"CardCode\" from OCRG a , OCRD b where a.\"GroupCode\" = b.\"GroupCode\" and a.\"GroupType\"='C' and a.\"GroupName\"= '{0}'", grp.Replace("'", "''"));
        if (grp == "All")
            sql = string.Format("select \"CardCode\" from OCRG a , OCRD b where a.\"GroupCode\" = b.\"GroupCode\" and a.\"GroupType\"='C'");
        sql = sql + " and exists(select 1 from OINV i where i.\"CardCode\" = b.\"CardCode\" and (i.\"DocTotal\" - i.\"PaidToDate\") > 0) ";
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {

            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                rtn = rtn + string.Format("'{0}'", (string)oRS.Fields.Item(0).Value);
                oRS.MoveNext();
                if (!oRS.EoF)
                    rtn = rtn + ",";
            }


        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return rtn;
    }
    public void UpdateCCTRANSInvoiceID()
    {
        string sql = "UPDATE \"@CCTRANS\" set \"U_InvoiceID\" = i.\"DocEntry\" \r\n" +
    " from INV1 i, DLN1 d, \"@CCTRANS\" t \r\n" +
    " where i.\"BaseType\" = 15 and i.\"BaseEntry\" = d.\"DocEntry\" \r\n" +
    " and CAST(d.\"BaseEntry\" as VARCHAR) = t.\"U_OrderID\" \r\n" +
    " and t.\"U_InvoiceID\" = '' ";

        try
        {

            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);
            sql = "UPDATE \"@CCTRANS\" set \"U_InvoiceID\" = i.\"DocEntry\" \r\n" +
                    "from INV1 i, \"@CCTRANS\" t \r\n" +
                    "where i.\"BaseType\" = 17 and CAST(i.\"BaseEntry\" AS VARCHAR) = t.\"U_OrderID\" \r\n" +
                    "and t.\"U_InvoiceID\" = '' ";
            execute(sql);


        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public bool IsConnectCustomerExist(string customerID)
    {


        string sql = string.Format("select 1 from \"@CCCONNECTCUSTOMER\" where \"U_CustomerID\" = '{0}'", customerID);
        try
        {
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                return true;
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return false;
    }

    public bool IsSODocNumExist(string DocNum)
    {
        string sql = string.Format("select * from ORDR where \"DocNum\"='{0}' ", DocNum);
        try
        {
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                return true;
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return false;
    }
    public void delete(CCConnectInvoice t)
    {
        string sql = "";
        try
        {
            sql = string.Format("delete from  \"@CCCONNECTINVOICE\" where \"DocEntry\" = {0}", t.recID);
            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }

    }

    public string toDateString(DateTime? dt)
    {
        if (dt == null)
            return "NULL";
        return string.Format("TO_DATE('{0}','MM-DD-YYYY')", ((DateTime)dt).ToString("MM-dd-yyyy"));
    }
    public int toInt(object value)
    {
        try
        {
            return int.Parse((string)value == "" ? "0" : (string)value);
        }
        catch (Exception) { }
        return 0;
    }
    public char toChar(object value)
    {
        try
        {
            return ((string)value).Length > 0 ? ((string)value)[0] : ' ';
        }
        catch (Exception) { }
        return ' ';

    }
    public void setDeliveryID()
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();

        string sql = string.Format("Update \"@CCTRANS\" set \"U_DeliveryID\" = b.\"TrgetEntry\" " +
    " from \"@CCTRANS\" a, \"RDR1\" b " +
    " where a.\"U_OrderID\" = CAST(b.\"DocEntry\" as VARCHAR) and b.\"TargetType\"=15 " +
    " and a.\"U_DeliveryID\" is NULL " +
    " and a.\"U_command\"='cc:authonly' " +
    " and not exists(select 1 from \"@CCTRANS\" cc2 where a.\"U_OrderID\" = cc2.\"U_OrderID\" and cc2.\"U_command\"='cc:authonly-remain')");
        try
        {
            execute(sql);


            sql = string.Format("Update \"@CCTRANS\" set \"U_DeliveryID\" = b.\"TrgetEntry\" " +
    " from \"@CCTRANS\" a, \"RDR1\" b " +
    " where a.\"U_OrderID\" = CAST(b.\"DocEntry\" as VARCHAR) and b.\"TargetType\"=15 " +
    " and a.\"U_DeliveryID\" is NULL " +
    " and a.\"U_command\"='cc:authonly-remain' " +
    " and not exists(select 1 from \"@CCTRANS\" cc2 where a.\"U_OrderID\" = cc2.\"U_OrderID\" and cc2.\"U_DeliveryID\"=b.\"TrgetEntry\")");
            execute(sql);

            sql = string.Format("Update \"@CCTRANS\" set \"U_InvoiceID\" = b.\"DocEntry\" from \"@CCTRANS\" a, INV1 b where a.\"U_DeliveryID\" = b.\"BaseEntry\" and a.\"U_InvoiceID\" = '' and b.\"BaseType\" = 15 and not exists ( select 1 from OINV c where cast(c.\"DocEntry\" as varchar) = a.\"U_InvoiceID\" and c.\"CardCode\" = a.\"U_customerID\")");

            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            if (bRunJob)
            {
                trace("Encounter db error. Program Existing...");
                System.Environment.Exit(0);
            }
        }
        finally
        {
            try
            {
                watch.Stop();
                //trace("Batch Invoicing Set Delivery ID completed. execution time: " + watch.ElapsedMilliseconds);

            }
            catch (Exception) { }
        }

    }
    public void setInvoiceID()
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();

        string sql = string.Format("Update \"@CCTRANS\" set \"U_InvoiceID\" = b.\"TrgetEntry\" " +
    " from \"@CCTRANS\" a, \"RDR1\" b " +
    " where a.\"U_OrderID\" = CAST(b.\"DocEntry\" as VARCHAR) and b.\"TargetType\"=13 " +
    " and a.\"U_InvoiceID\" in ('', NULL) " +
    " and a.\"U_command\"='cc:authonly' " +
    " and not exists(select 1 from \"@CCTRANS\" cc2 where a.\"U_OrderID\" = cc2.\"U_OrderID\" and cc2.\"U_command\"='cc:authonly-remain')");
        try
        {
            //  trace(sql);
            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            if (bRunJob)
            {
                trace("Encounter db error. Program Existing...");
                System.Environment.Exit(0);
            }
        }
        finally
        {
            try
            {
                watch.Stop();
                // trace("Batch Invoicing Set Invoice ID completed. execution time: " + watch.ElapsedMilliseconds);

            }
            catch (Exception) { }
        }

    }
    public void BatchSetTransInvoiceIDByDeliveryID(string InvoiceID, string refNum, string DeliveryID)
    {

        string sql = string.Format("Update \"@CCTRANS\" set \"U_PaymentID\"=(select max(\"TransId\") from \"OJDT\" where \"BaseRef\" = \"U_SAPPaymentID\") " +
    " , \"U_InvoiceID\" = '{0}' " +
    " , \"U_refNum\" ='{2}' " +
    " , \"U_BatchResult\" = 'Success' " +
    " , \"U_BatchDate\"=CURRENT_DATE  " +
    " where \"U_DeliveryID\"='{1}' and \"U_command\"='cc:authonly-remain'", InvoiceID, DeliveryID, refNum);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void DuplicateTransLogForPartial(string amount, string deliveryid, string refnum)
    {
        int id = getNextTableID("@CCTRANS");
        string sql = string.Format("insert into \"@CCTRANS\" (\"DocEntry\", \"U_amount\", \"U_command\", \"U_OrderID\", \"U_customerID\", \"U_InvoiceID\",\"U_MethodID\", \"U_custNum\",\"U_CCAccountID\", \"U_recDate\") " +
    " select {3}, '{0}' , 'cc:authonly-remain' , \"U_OrderID\", \"U_customerID\", '',\"U_MethodID\", \"U_custNum\",\"U_CCAccountID\", CURRENT_DATE " +
    " from \"@CCTRANS\" where \"U_DeliveryID\" = '{1}' and \"U_command\" in ('cc:authonly', 'cc:authonly-remain') and \"U_refNum\" = '{2}'", amount, deliveryid, refnum, id);
        try
        {
            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }

    }
    public void CCTRANAdjustWithSQId(string refNum, string QuoteID, string CustomerID, string amount)
    {
        string sql = string.Format("update \"@CCTRANS\" set \"U_QuoteID\" = '{1}' where \"U_refNum\" = '{0}' and \"U_customerID\"='{2}' and \"U_amount\" = '{3}'", refNum, QuoteID, CustomerID, amount);
        try
        {
            execute(sql);
            //  refresh(db.CCTRANs);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void CCTRANASetOrderIDwithQuoteID()
    {
        string sql = string.Format("update \"@CCTRANS\" set \"U_OrderID\" = CAST(b.\"DocEntry\" AS VARCHAR) FROM \"@CCTRANS\" a, RDR1 b where a.\"U_QuoteID\" = CAST(b.\"BaseEntry\" AS VARCHAR) and b.\"BaseType\" = 23 and a.\"U_OrderID\" = ''");
        try
        {
            execute(sql);
            // refresh(db.CCTRANs);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public List<int> getSOInvoice(int soID)
    {
        List<int> list = new List<int>();
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string sql = string.Format("select distinct \"TrgetEntry\" from RDR1 where \"TargetType\"=13 and \"DocEntry\" ={0}", soID);
        try
        {

            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                list.Add((int)oRS.Fields.Item(0).Value);
                oRS.MoveNext();
            }
            sql = string.Format("select distinct b.\"TrgetEntry\" from RDR1 a, DLN1 b where a.\"TargetType\"=15 and a.\"TrgetEntry\" = b.\"DocEntry\" and b.\"TargetType\"=13 and a.\"DocEntry\" ={0}", soID);
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                list.Add((int)oRS.Fields.Item(0).Value);
                oRS.MoveNext();
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;

        }
        return list;
    }
    public int getUserID(string username)
    {
        int id = 0;
        string sql = string.Format("select USERID from OUSR where USER_CODE = '{0}'", username);
        try
        {

            return getSQLInt(sql);


        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return id;

    }
    public void updatePaymentTerm(string cardcode)
    {
        string sql = string.Format("Update OCTG set \"OpenRcpt\" = 'N' from OCRD a , OCTG b where a.\"GroupNum\" = b.\"GroupNum\" and a.\"CardCode\" = '{0}'", cardcode);
        try
        {

            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public string GetSalesOrderDocNum(string InvoiceNum)
    {
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string sql = string.Format("select a.\"BaseEntry\", a.\"BaseType\" from INV1 a, OINV b where a.\"DocEntry\" = b.\"DocEntry\" and b.\"DocNum\" = {0}", InvoiceNum);
        try
        {
            oRS.DoQuery(sql);
            if (oRS.EoF)
                return InvoiceNum;
            int BaseEntry = (int)oRS.Fields.Item(0).Value;
            int BaseType = (int)oRS.Fields.Item(1).Value;
            if (BaseType == 17)
            {
                sql = string.Format("select \"DocNum\" from ORDR where \"DocEntry\" = {0}", BaseEntry);
                oRS.DoQuery(sql);
                if (oRS.EoF)
                    return InvoiceNum;
                return ((int)oRS.Fields.Item(0).Value).ToString();
            }
            if (BaseType == 15)
            {
                sql = string.Format("select b.\"DocNum\" from DLN1 a, ORDR b where a.\"BaseEntry\" = b.\"DocEntry\" and a.\"DocEntry\" = {0}", BaseEntry);
                oRS.DoQuery(sql);
                if (oRS.EoF)
                    return InvoiceNum;
                return ((int)oRS.Fields.Item(0).Value).ToString();
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;

        }
        return InvoiceNum;
    }
    public void removePaymentMethods(string cardcode)
    {
        string sql = string.Format("delete from \"@CCCUST\" where \"U_CustomerID\" = '{0}' ", cardcode);
        try
        {

            execute(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }

    }
    public void updateTransInvoiceID(string invoiceID, string customerID)
    {
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string sql = string.Format("select \"DocEntry\" from RDR1 where \"TargetType\" = 13 and \"TrgetEntry\" = {0} ", invoiceID);
        try
        {
            oRS.DoQuery(sql);
            if (oRS.EoF)
            {
                sql = string.Format("select a.\"DocEntry\" from RDR1 a, DLN1 b where a.\"TargetType\" = 15 and a.\"TrgetEntry\" = b.\"DocEntry\" and b.\"TargetType\" = 13 and b.\"TrgetEntry\" = {0} ", invoiceID);
                oRS.DoQuery(sql);
                if (oRS.EoF)
                    return;
            }
            int OID = (int)oRS.Fields.Item(0).Value;
            sql = string.Format("UPDATE \"@CCTRANS\" set \"U_InvoiceID\" = '{1}', \"U_customerID\" = '{2}' where \"U_OrderID\" = '{0}' and \"U_InvoiceID\" in( '', null)", OID, invoiceID, customerID);
            execute(sql);
            sql = string.Format("UPDATE \"@CCTRANS\" set \"U_InvoiceID\" = '{0}', \"U_customerID\" = '{1}' where \"U_OrderID\" in (select CAST(\"BaseEntry\" AS NVARCHAR) from INV1 where \"DocEntry\"={0} and \"BaseType\" = 17) ", invoiceID, customerID);
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;

        }

    }
    public void updateTransCustomerID(string OID, string CustomerID)
    {
        string sql = string.Format("UPDATE \"@CCTRANS\" set \"U_customerID\" = '{1}' where \"U_OrderID\" = '{0}' ", OID, CustomerID);
        try
        {

            execute(sql);


        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }

    }
    public void updateTransCMCustomerID(string CRID)
    {

        string sql = string.Format("select \"DocEntry\" from INV1 where \"TargetType\" = 14 and \"TrgetEntry\" = {0} ", CRID);
        try
        {


            int InvID = getSQLInt(sql);
            sql = string.Format("UPDATE \"@CCTRANS\" set \"U_CreditMemoID\"= '{1}', \"U_customerID\" = '{2}' where \"U_InvoiceID\" = '{0}' ", InvID, CRID, getCustomerID());
            trace(sql);
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void BatchPreAuthUpdateTrans(int OID, string CustomerID, string refNum)
    {

        string sql = string.Format("Update \"@CCTRANS\" set \"U_OrderID\" = '{0}' where \"U_customerID\" = '{1}' and \"U_refNum\" = '{2}' and \"U_command\"='cc:authonly' ", OID, CustomerID, refNum);
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public string getDocNumAdded(string customerID, string form)
    {
        string tbname = "ORDR";
        if (form == FORMCREDITMEMO)
            tbname = "ORIN";
        if (form == FORMINVOICE)
            tbname = "OINV";
        if (form == FORMDOWNPAYMENT)
            tbname = "ODPI";
        if (form == FORMSALESQUOTE)
            tbname = "OQUT";
        if (form == FORMINCOMINGPAYMENT)
            tbname = "ORCT";

        string sql = string.Format("select MAX(\"DocNum\") from {1} where \"CardCode\" = '{0}' ", customerID, tbname);
        try
        {
            trace(sql);

            return getSQLInt(sql).ToString();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return "";
    }


    #region SAP Function 
    long v_RetVal;
    int v_ErrCode;
    string v_ErrMsg = "";
    string FileName, defFileName;
    public string B1ClientName = "";


    public string FindFile()
    {
        System.Threading.Thread ShowFolderBrowserThread = null;
        try
        {
            ShowFolderBrowserThread = new System.Threading.Thread(ShowFolderBrowser);
            if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
            {
                ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA);
                ShowFolderBrowserThread.Start();
            }
            else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
            {
                ShowFolderBrowserThread.Start();
                ShowFolderBrowserThread.Join();

            }
            Thread.Sleep(5000);
            while (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Running)
            {
                System.Windows.Forms.Application.DoEvents();
            }
            if (!string.IsNullOrEmpty(FileName))
            {
                return FileName;
            }
        }
        catch (Exception ex)
        {
            SBO_Application.MessageBox("FileFile" + ex.Message);
        }

        return "";

    }

    public string SaveFile(string defName)
    {

        defFileName = defName;
        System.Threading.Thread ShowFolderBrowserThread = null;
        try
        {
            ShowFolderBrowserThread = new System.Threading.Thread(SaveFileBrowser);

            if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
            {
                ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA);
                ShowFolderBrowserThread.Start();
            }
            else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
            {
                ShowFolderBrowserThread.Start();
                ShowFolderBrowserThread.Join();

            }
            Thread.Sleep(5000);

            while (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Running)
            {
                System.Windows.Forms.Application.DoEvents();
            }
            if (!string.IsNullOrEmpty(FileName))
            {
                return FileName;
            }
        }
        catch (Exception ex)
        {
            SBO_Application.MessageBox("FileFile" + ex.Message);
        }

        return "";

    }

    public void ShowFolderBrowser()
    {
        System.Diagnostics.Process[] MyProcs = null;
        dynamic UserName = Environment.UserName;
        int CallingWindowSAP = 0;
        FileName = "";
        OpenFileDialog OpenFile = new OpenFileDialog();
        string newTitle = SBO_Application.Desktop.Title;
        System.Diagnostics.Process oProcess = System.Diagnostics.Process.GetCurrentProcess();
        B1ClientName = newTitle + " " + oProcess.Id.ToString();
        try
        {
            OpenFile.Multiselect = false;
            OpenFile.Filter = "All files(*.)|*.*";
            int filterindex = 0;
            try
            {
                filterindex = 0;
            }
            catch (Exception ex)
            {
            }

            OpenFile.FilterIndex = filterindex;

            OpenFile.RestoreDirectory = true;
            MyProcs = System.Diagnostics.Process.GetProcessesByName("SAP Business One");
            for (int i = 0; i <= MyProcs.GetLength(0); i++)
            {
                if (MyProcs[i].MainWindowTitle == B1ClientName)
                {
                    CallingWindowSAP = i;
                    goto NEXT_STEP;
                }
            }
            SBO_Application.MessageBox("Unable to determine Running processes by UserName!");
            OpenFile.Dispose();
            GC.Collect();
            return;
            NEXT_STEP:
            if (MyProcs.Length == 1)
            {
                for (int i = 0; i <= MyProcs.Length - 1; i++)
                {
                    WindowWrapper MyWindow = new WindowWrapper(MyProcs[i].MainWindowHandle);
                    DialogResult ret = OpenFile.ShowDialog(MyWindow);

                    if (ret == DialogResult.OK)
                    {
                        FileName = OpenFile.FileName;
                        OpenFile.Dispose();
                    }
                    else
                    {
                        System.Windows.Forms.Application.ExitThread();
                    }
                }
            }
            else if (MyProcs.Length > 1)
            {
                for (int i = 0; i <= MyProcs.Length - 1; i++)
                {
                    if (CallingWindowSAP == i)
                    {
                        WindowWrapper MyWindow = new WindowWrapper(MyProcs[i].MainWindowHandle);
                        DialogResult ret = OpenFile.ShowDialog(MyWindow);

                        if (ret == DialogResult.OK)
                        {
                            FileName = OpenFile.FileName;
                            OpenFile.Dispose();
                        }
                        else
                        {
                            System.Windows.Forms.Application.ExitThread();
                        }
                    }
                }
            }
            else
            {
                SBO_Application.MessageBox("More than 1 SAP B1 is started!");
            }
        }
        catch (Exception ex)
        {
            SBO_Application.StatusBar.SetText(ex.Message);
            FileName = "";
        }
        finally
        {
            OpenFile.Dispose();
            GC.Collect();
        }

    }
    public void SaveFileBrowser()
    {
        System.Diagnostics.Process[] MyProcs = null;
        dynamic UserName = Environment.UserName;

        FileName = "";
        SaveFileDialog saveFile = new SaveFileDialog();
        saveFile.FileName = defFileName;
        try
        {
            MyProcs = System.Diagnostics.Process.GetProcessesByName("SAP Business One");

            for (int i = 0; i <= MyProcs.GetLength(1); i++)
            {
                if (GetProcessUserName(MyProcs[i]) == UserName)
                {
                    goto NEXT_STEP;
                }
            }
            SBO_Application.MessageBox("Unable to determine Running processes by UserName!");
            saveFile.Dispose();
            GC.Collect();
            return;
            NEXT_STEP:
            if (MyProcs.Length == 1)
            {

                for (int i = 0; i <= MyProcs.Length - 1; i++)
                {
                    WindowWrapper MyWindow = new WindowWrapper(MyProcs[i].MainWindowHandle);
                    DialogResult ret = saveFile.ShowDialog(MyWindow);

                    if (ret == DialogResult.OK)
                    {
                        FileName = saveFile.FileName;
                        saveFile.Dispose();
                    }
                    else
                    {
                        System.Windows.Forms.Application.ExitThread();
                    }
                }
            }
            else
            {
                SBO_Application.MessageBox("More than 1 SAP B1 is started!");
            }
        }
        catch (Exception ex)
        {
            SBO_Application.StatusBar.SetText(ex.Message);
            FileName = "";
        }
        finally
        {
            saveFile.Dispose();
            GC.Collect();
        }

    }
    private string GetProcessUserName(System.Diagnostics.Process Process)
    {
        string strResult = "";
        /*
        ObjectQuery sq = new ObjectQuery("Select * from Win32_Process Where ProcessID = '" + Process.Id + "'");
        ManagementObjectSearcher searcher = new ManagementObjectSearcher(sq);


        if (searcher.Get().Count == 0)
            return null;

        foreach (ManagementObject oReturn in searcher.Get())
        {
            string[] o = new string[2];

            //Invoke the method and populate the o var with the user name and domain                         
            oReturn.InvokeMethod("GetOwner", (object[])o);
            strResult = o[0];
            return o[0];
        }
        */
        return strResult;


    }

    public class WindowWrapper : System.Windows.Forms.IWin32Window
    {

        private IntPtr _hwnd;
        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }

        public System.IntPtr Handle
        {
            get { return _hwnd; }
        }

    }

    public void addFms(string frmId, string itmId, string colID, string query)
    {

        int queryId = 0;
        int fmsId = 0;

        SAPbobsCOM.Recordset oRecSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            oRecSet.DoQuery("Select \"QueryId\", \"IndexID\" From \"CSHS\" Where \"FormID\" = '" + frmId + "' And \"ItemID\" = '" + itmId + "' And \"ColID\" = '" + colID + "'");
            if (!oRecSet.EoF)
            {
                queryId = Convert.ToInt32(oRecSet.Fields.Item("QueryId").Value);
                fmsId = Convert.ToInt32(oRecSet.Fields.Item("IndexID").Value);
                var sqll = "Update \"OUQR\" Set \"QString\" = '" + query + "' Where \"IntrnalKey\" = '" + queryId.ToString() + "'";
                oRecSet.DoQuery(sqll);

            }
            else
            {

                oRecSet.DoQuery("Select IFNULL(Max(\"IndexID\"),0) + 1 As fmsId From \"CSHS\"");


                fmsId = Convert.ToInt32(oRecSet.Fields.Item("fmsId").Value);
                queryId = addQuery(query, "Fms_" + frmId + "_" + itmId + "_" + colID);

                string strS = "INSERT Into \"CSHS\" (\"FormID\" , \"ItemID\" ,\"ColID\" ,\"ActionT\" ,\"QueryId\" ,\"IndexID\" ,\"Refresh\"  ,\"FrceRfrsh\" ,\"ByField\") ";
                strS += " Values ('" + frmId + "','" + itmId + "','" + colID + "','2','" + queryId.ToString() + "','" + fmsId.ToString() + "','N','N','N')";
                oRecSet.DoQuery(strS);
            }

            oRecSet = null;
        }
        catch (Exception ex)
        {
            SBO_Application.SetStatusBarMessage("Error in creating formatted search" + "Fms_" + frmId + "_" + itmId + "_" + colID + ex.Message);
        }

    }
    public void MsgWarning(string pMessage)
    {
        try
        {
            SBO_Application.StatusBar.SetText(pMessage, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }

    public int addQuery(string strQuery, string queryName, String Category = "EbizCharge")
    {
        int queryId = 0;
        int catId = 0;
        if (queryName.Equals("Fms_EmailEsti_txtBpCode_-1") || queryName.Equals("Fms_SyncUp_edBP_-1"))
        {
            queryName = "Select Customer";
        }//edBPId
        if (queryName.Equals("Fms_SyncUp_edBPId_-1") || queryName.Equals("Fms_EmailPaySO_txtBpCode_-1"))
        {
            queryName = "Select Business Partner";
        }
        SAPbobsCOM.Recordset oRecSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        oRecSet.DoQuery(String.Format("Select \"CategoryId\" From \"OQCN\" Where \"CatName\" = '{0}'", Category));
        if (!oRecSet.EoF)
        {
            catId = (int)oRecSet.Fields.Item("CategoryId").Value;
        }
        else
        {
            addQryCategor(Category);
            oRecSet.DoQuery(String.Format("Select \"CategoryId\" From \"OQCN\" Where \"CatName\" = '{0}'", Category));
            if (!oRecSet.EoF)
            {
                catId = (int)oRecSet.Fields.Item("CategoryId").Value;
            }
        }

        oRecSet.DoQuery("Select \"IntrnalKey\" As qId From \"OUQR\" Where \"QName\" = '" + queryName + "'");
        if (!oRecSet.EoF)
        {
            queryId = Convert.ToInt32(oRecSet.Fields.Item("qId").Value);
        }
        else
        {

            {
                oRecSet.DoQuery("Select IFNULL(Max(\"IntrnalKey\"),0) + 1 As newId From \"OUQR\"");
            }

            queryId = Convert.ToInt32(oRecSet.Fields.Item("newId").Value);
            string sQuery = " Insert Into \"OUQR\" (\"IntrnalKey\" ,\"QCategory\" ,\"QName\" ,\"QString\" ,\"QType\" ) ";
            sQuery += " Values ('" + queryId.ToString() + "','" + catId.ToString() + "','" + queryName + "','" + strQuery + "','W')";
            oRecSet.DoQuery(sQuery);
        }
        oRecSet = null;

        return queryId;
    }
    public void addQryCategor(string catName)
    {
        try
        {
            SAPbobsCOM.QueryCategories qCat = default(SAPbobsCOM.QueryCategories);
            qCat = (SAPbobsCOM.QueryCategories)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories);
            qCat.Name = catName;
            qCat.Add();
        }
        catch { }
    }
    public bool AddColumns(string TableName, string Name, string Description, SAPbobsCOM.BoFieldTypes Type, int Size = 0, SAPbobsCOM.BoFldSubTypes SubType = SAPbobsCOM.BoFldSubTypes.st_None, string LinkedTable = "", string[,] LOV = null, string DefV = "")
    {
        bool outResult = false;
        try
        {
            SAPbobsCOM.UserFieldsMD v_UserField = default(SAPbobsCOM.UserFieldsMD);

            if (TableName.StartsWith("@") == true)
            {
                if (!ColumnExists(TableName, Name))
                {
                    v_UserField = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                    v_UserField.TableName = TableName;
                    v_UserField.Name = Name;
                    if (!string.IsNullOrEmpty(DefV))
                    {
                        v_UserField.DefaultValue = DefV;
                    }

                    if (LOV == null)
                    {
                    }
                    else
                    {
                        for (int k = 0; k <= LOV.Length - 1; k++)
                        {
                            v_UserField.ValidValues.Value = LOV[k, 0];
                            v_UserField.ValidValues.Value = LOV[k, 1];
                            v_UserField.ValidValues.Add();
                        }

                    }

                    v_UserField.Description = Description;
                    v_UserField.Type = Type;
                    if (Type != SAPbobsCOM.BoFieldTypes.db_Date)
                    {
                        if (Size != 0)
                        {
                            v_UserField.Size = Convert.ToInt16(Size);
                            v_UserField.EditSize = Convert.ToInt16(Size);
                        }
                    }
                    if (SubType != SAPbobsCOM.BoFldSubTypes.st_None)
                    {
                        v_UserField.SubType = SubType;
                    }
                    if (!string.IsNullOrEmpty(LinkedTable))
                        v_UserField.LinkedTable = LinkedTable;
                    v_RetVal = v_UserField.Add();
                    if (v_RetVal != 0)
                    {
                        oCompany.GetLastError(out v_ErrCode, out v_ErrMsg);
                        SBO_Application.StatusBar.SetText("Failed to add UserField " + Description + " - " + v_ErrCode + " " + v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                    else
                    {
                        //       SBO_Application.StatusBar.SetText("[@" + TableName + "] - " + Description + " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        outResult = true;
                        return true;
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField);
                    v_UserField = null;
                }
                else
                {
                    return false;
                }
            }


            if (TableName.StartsWith("@") == false)
            {
                if (!UDFExists(TableName, Name))
                {
                    v_UserField = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                    v_UserField.TableName = TableName;
                    v_UserField.Name = Name;
                    if (!string.IsNullOrEmpty(DefV))
                    {
                        v_UserField.DefaultValue = DefV;
                    }

                    if (LOV == null)
                    {
                    }
                    else
                    {
                        for (int k = 0; k <= LOV.Length / 2 - 1; k++)
                        {
                            v_UserField.ValidValues.Value = LOV[k, 0];
                            v_UserField.ValidValues.Description = LOV[k, 1];
                            v_UserField.ValidValues.Add();
                        }

                    }
                    v_UserField.Description = Description;
                    v_UserField.Type = Type;
                    if (Type != SAPbobsCOM.BoFieldTypes.db_Date)
                    {
                        if (Size != 0)
                        {
                            v_UserField.Size = Size;
                            v_UserField.EditSize = Size;
                        }
                    }
                    if (SubType != SAPbobsCOM.BoFldSubTypes.st_None)
                    {
                        v_UserField.SubType = SubType;
                    }
                    if (!string.IsNullOrEmpty(LinkedTable))
                        v_UserField.LinkedTable = LinkedTable;
                    v_RetVal = v_UserField.Add();
                    if (v_RetVal != 0)
                    {
                        oCompany.GetLastError(out v_ErrCode, out v_ErrMsg);
                        SBO_Application.StatusBar.SetText("Failed to add UserField " + Description + " - " + v_ErrCode + " " + v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                    else
                    {
                        //       SBO_Application.StatusBar.SetText("[@" + TableName + "] - " + Description + " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        outResult = true;
                        return true;
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField);
                    v_UserField = null;
                }
                else
                {
                    return false;
                }
            }
        }
        catch (Exception ex)
        {
            SBO_Application.StatusBar.SetText("Failed to Add Columns : " + ex.Message);
        }
        finally
        {
        }
        return outResult;
    }
    public bool ColumnExists(string TableName, string FieldID)
    {
        bool oFlag = true;
        try
        {
            SAPbobsCOM.Recordset rsetField = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string s = "Select 1 From \"CUFD\" Where \"TableID\" ='" + TableName.Trim() + "' And \"AliasID\" ='" + FieldID.Trim() + "'";
            rsetField.DoQuery("Select 1 From \"CUFD\" Where \"TableID\" ='" + TableName.Trim() + "' And \"AliasID\" ='" + FieldID.Trim() + "'");
            if (rsetField.EoF)
                oFlag = false;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rsetField);
            rsetField = null;
            GC.Collect();
            return oFlag;
        }
        catch (Exception ex)
        {
            SBO_Application.StatusBar.SetText("Failed to Column Exists : " + ex.Message);
        }
        finally
        {
        }
        return oFlag;
    }

    public bool UDFExists(string TableName, string FieldID)
    {
        bool outResult = false;
        try
        {
            SAPbobsCOM.Recordset rsetUDF = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            bool oFlag = true;

            rsetUDF.DoQuery("Select 1 From \"CUFD\" Where \"TableID\" ='" + TableName.Trim() + "' And \"AliasID\" = '" + FieldID.Trim() + "'");
            if (rsetUDF.EoF)
                oFlag = false;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rsetUDF);
            rsetUDF = null;
            outResult = oFlag;
            GC.Collect();
            return oFlag;
        }
        catch (Exception ex)
        {
            SBO_Application.StatusBar.SetText("Failed to UDF Exisits : " + ex.Message);
        }
        finally
        {
        }
        return outResult;
    }
    public void ShowMsgCustom(string strMsg, string msgTime = "S", string errType = "W")
    {
        SAPbouiCOM.BoMessageTime time = default(SAPbouiCOM.BoMessageTime);
        SAPbouiCOM.BoStatusBarMessageType msgType = default(SAPbouiCOM.BoStatusBarMessageType);
        switch (errType.ToUpper())
        {
            case "E":
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Error;
                break;
            case "W":
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Warning;
                break;
            case "N":
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_None;
                break;
            case "S":
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Success;
                break;
            default:
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Warning;
                break;
        }
        switch (msgTime.ToUpper())
        {
            case "M":
                time = SAPbouiCOM.BoMessageTime.bmt_Medium;
                break;
            case "S":
                time = SAPbouiCOM.BoMessageTime.bmt_Short;
                break;
            case "L":
                time = SAPbouiCOM.BoMessageTime.bmt_Long;
                break;
            default:
                time = SAPbouiCOM.BoMessageTime.bmt_Medium;
                break;
        }
        SBO_Application.StatusBar.SetText(strMsg, time, msgType);
    }
    public void DoOpenLinkedObjectForm(string FormUniqueID, string ActivateMenuItem, string FindItemUID, string FindItemUIDValue)
    {
        try
        {
            SAPbouiCOM.Form oForm = null;
            bool Bool = false;

            for (int frm = 0; frm <= SBO_Application.Forms.Count - 1; frm++)
            {
                if (SBO_Application.Forms.Item(frm).UniqueID == FormUniqueID)
                {
                    oForm = SBO_Application.Forms.Item(FormUniqueID);
                    oForm.Close();
                    break; // TODO: might not be correct. Was : Exit For
                }
            }
            if (Bool == false)
            {
                SBO_Application.ActivateMenuItem(ActivateMenuItem);
                oForm = SBO_Application.Forms.ActiveForm;
                if (oForm.TypeEx != FormUniqueID)
                    oForm = SBO_Application.Forms.GetForm(FormUniqueID, 1);
                string oformtype = oForm.Type.ToString();
                oForm.Select();
                oForm.Freeze(true);
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                SAPbouiCOM.EditText I = (SAPbouiCOM.EditText)oForm.Items.Item(FindItemUID).Specific;
                I.Item.Enabled = true;
                I.Value = FindItemUIDValue.Trim();
                oForm.Items.Item("1").Click();
                oForm.Freeze(false);

            }

        }
        catch (Exception ex)
        {
            //GolabalVariables.oGFun.StatusBarErrorMsg("" + ex.Message);
        }
        finally
        {
        }
    }

    public bool AddTable(string TableName, string TableDescription, SAPbobsCOM.BoUTBTableType TableType)
    {
        bool outResult = false;
        try
        {

            SAPbobsCOM.UserTablesMD v_UserTableMD = default(SAPbobsCOM.UserTablesMD);

            GC.Collect();
            if (!TableExists(TableName))
            {
                SBO_Application.StatusBar.SetText("Creating Table " + TableName + " ...................", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                v_UserTableMD = (SAPbobsCOM.UserTablesMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                v_UserTableMD.TableName = TableName;
                v_UserTableMD.TableDescription = TableDescription;
                v_UserTableMD.TableType = TableType;
                v_RetVal = v_UserTableMD.Add();
                if (v_RetVal != 0)
                {
                    oCompany.GetLastError(out v_ErrCode, out v_ErrMsg);
                    SBO_Application.StatusBar.SetText("Failed to Create Table " + TableName + v_ErrCode + " " + v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD);
                    v_UserTableMD = null;
                    GC.Collect();
                    return false;
                }
                else
                {
                    SBO_Application.StatusBar.SetText("[@" + TableName + "] - " + TableDescription + " created successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD);
                    v_UserTableMD = null;
                    outResult = true;
                    GC.Collect();
                    return true;
                }
            }
            else
            {
                GC.Collect();
                return false;
            }
        }
        catch (Exception ex)
        {
            SBO_Application.StatusBar.SetText("Failed to Add Table : " + ex.Message);
        }
        finally
        {
        }
        return outResult;
    }
    public bool TableExists(string TableName)
    {
        bool outResult = false;
        try
        {
            SAPbobsCOM.UserTablesMD oTables = default(SAPbobsCOM.UserTablesMD);
            bool oFlag = false;

            oTables = (SAPbobsCOM.UserTablesMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            oFlag = oTables.GetByKey(TableName);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oTables);
            outResult = oFlag;
            return oFlag;
        }
        catch (Exception ex)
        {
            SBO_Application.StatusBar.SetText("Failed to Table Exists : " + ex.Message);
        }
        finally
        {
        }
        return outResult;
    }

    #endregion
    public void insert(ref CCCUST cccust)
    {
        string sql = string.Format("Insert Into CCCUST (active, [CustNum], [MethodID],[methodDescription]," +
    "[cardNum],	[expDate],	[routingNumber] ,[checkingAccount], " +
    "[firstName],[lastName],[email],[street],[city],[state]," +
    "[zip],[cardType],[GroupName],[CardName],[CCAccountID],Declined,CustomerID,methodName,[default], CardCode) values('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}', '{17}','{18}','{19}','{20}','{21}','{22}', '{23}')",
                "Y",
                cccust.CustNum,
                cccust.MethodID,
                Truncate(cccust.methodDescription == null ? "" : cccust.methodDescription.Replace("'", "''"), 250),
                cccust.cardNum,
                cccust.expDate,
                cccust.routingNumber,
                cccust.checkingAccount,
                Truncate(cccust.firstName == null ? "" : cccust.firstName.Replace("'", "''"), 50),
                Truncate(cccust.lastName == null ? "" : cccust.lastName.Replace("'", "''"), 255),
                Truncate(cccust.email == null ? "" : cccust.email.Replace("'", "''"), 80),
                Truncate(cccust.street == null ? "" : cccust.street.Replace("'", "''"), 80),
                Truncate(cccust.city == null ? "" : cccust.city.Replace("'", "''"), 50),
                cccust.state,
                cccust.zip,
                cccust.cardType,
                Truncate(cccust.GroupName == null ? "" : cccust.GroupName.Replace("'", "''"), 80),
                cccust.CardName == null ? "" : cccust.CardName.Replace("'", "''"),
                cccust.CCAccountID,
                cccust.Declined,
                cccust.CustomerID,
                Truncate(cccust.methodName == null ? "" : cccust.methodName.Replace("'", "''"), 255),
                cccust.@default,
                cccust.cardCode  // used as currency code
    );

        try
        {
            execute(sql);
            cccust.recID = getCCCustRecID(cccust.CustomerID);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }

    }
    public int getCCCustRecID(string id)
    {
        int nRet = 0;
        string sql = string.Format("select MAX(recID) from CCCUST (NOLOCK) where CustomerID = '{0}'", id);

        try
        {
            return getSQLInt(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return nRet;
    }
    public int getNextCCCUSTID()
    {
        int rtn = 0;
        string sql = string.Format("select MAX(RecID) from CCCUST");
        try
        {

            return getSQLInt(sql) + 1;
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return rtn;
    }
    public void updateCCConnectInvoice(string id, string paidamt, string bal)
    {
        string status = CONNECT_STATUS_PAID;
        string pstr = getMoneyString(paidamt);
        bal = getMoneyString(bal);
        if (getDoubleValue(bal) != 0)
        {
            if (getDoubleValue(pstr) != 0)
                status = CONNECT_STATUS_PARTIAL;
            else
                status = CONNECT_STATUS_UPLOADED;
        }
        string sql = string.Format("update [@CCINVLOG[ set [U_PaidAmt[ = '{0}', [U_Balance[ = '{1}', [U_UpdateDt[ = current_timestamp, [U_PaymentStatus[='{3}' where [U_DocNum[ = '{2}'", paidamt, bal, id, status).Replace("[", "\"");
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public void updateCCConnectSalesOrder(string id, string paidamt, string bal, String Application = "")
    {
        string status = CONNECT_STATUS_PAID;
        string pstr = getMoneyString(paidamt);
        bal = getMoneyString(bal);
        if (getDoubleValue(bal) != 0)
        {
            if (getDoubleValue(pstr) != 0)
                status = CONNECT_STATUS_PARTIAL;
            else
                status = CONNECT_STATUS_UPLOADED;
        }
        string sql = string.Format("update [@CCSOLOG[ set [U_PaidAmt[ = '{0}', [U_Balance[ = '{1}', [U_UpdateDt[ = current_timestamp, [U_PaymentStatus[='{3}' where [U_DocNum[ = '{2}'", paidamt, bal, id, status).Replace("[", "\"");
        try
        {
            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public string GetCardNumber(string CustomerID)
    {
        string sRtn = "XXXX";
        string sql = string.Format("select top 1 \"U_cardNum\" from \"@CCCUST\"  where \"U_CustomerID\" = '{0}'", CustomerID);
        try
        {
            return getSQLString(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return sRtn;
    }
    public int getBranchIdFromODPI(string docentry)
    {

        string sql = "";
        try
        {
            sql = string.Format("select BPLId from ODPI (NOLOCK) where DocEntry  = '{0}'", docentry);
            return getSQLInt(sql);
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return 0;
    }
}

