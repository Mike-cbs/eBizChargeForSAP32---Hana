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
using eBizChargeForSAP.ServiceReference1;

partial class SAP  {

     string INVDocEntry = "";
        string DPDocEntry = "";
        string SODocEntry = "";
        string SQDocEntry = "";
        string CMDocEntry = "";   
    
    private  TransactionRequestObject createRequest(SAPbouiCOM.Form form, SAPCust cust)
    {
        form = theActiveForm;
        TransactionRequestObject req = new TransactionRequestObject();
        try
        {
            req.RecurringBilling = new GwRecurringBilling();
            req.BillingAddress = new GwAddress();
            req.AccountHolder = "";
            req.AuthCode = "";
            req.ClientIP = getIP();
            req.CustomerID = getCustomerID();
            req.CheckData = new CheckData();
            req.CheckData.Account = "";
            req.CheckData.AccountType = "";
            req.Details = new TransactionDetail();
            LineItem[] li = null;
            setLineItem(form, ref li);
            req.LineItems = li;
            if (theActiveForm.TypeEx == FORMINCOMINGPAYMENT)
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)theActiveForm.Items.Item("20").Specific;
                string s = "";
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    if (getMatrixSelect(oMatrix, "10000127", i))
                    {
                        req.Details.Invoice = getMatrixItem(oMatrix, "1", i);
                        s = s + req.Details.Invoice + ", ";
                    }
                }
                if (s != "")
                    s = (s + "\r\n").Replace(", \r\n", "\r\n");
                if (req.Details.Invoice == "" || req.Details.Invoice == null)
                {
                    req.Details.Invoice = getCustomerID();
                    req.Details.OrderID = req.Details.Invoice;
                }
                else
                {
                    req.Details.OrderID = GetSalesOrderDocNum(req.Details.Invoice);
                }
                req.Details.PONum = req.Details.PONum;
                req.Details.Subtotal = 0; // req.Details.Amount;
                req.Details.Shipping = 0;
                req.Details.Discount = 0;
                req.Details.Tax = 0;
                req.LineItems[0].UnitPrice = req.Details.Amount.ToString();
                if (s != "")
                    req.LineItems[0].ProductName += ". Invoice(s): " + s;
                return req;
            }
            req.Details.Amount = getItemDoubleValue(form,fidBalance);
            req.Details.Discount = getItemDoubleValue(form, fidDiscount);
            req.Details.Shipping = getItemDoubleValue(form, fidFreight);
            req.Details.Tax = getItemDoubleValue(form, fidTax);
            req.Details.Invoice = getFormItemVal(form, fidRecordID);
            req.Details.OrderID = getFormItemVal(form, fidRecordID);
            req.Details.PONum = getFormItemVal(form, "14");
            if (req.Details.PONum=="")
                req.Details.PONum = getFormItemVal(form, fidRecordID);
            if (form.TypeEx == FORMINVOICE)
            {
                if(req.Details.Invoice != "" && req.Details.Invoice != null)
                    req.Details.OrderID = GetSalesOrderDocNum(req.Details.Invoice);
            }
            req.Details.Description = getFormItemVal(form,fidDesc);
            req.Details.Subtotal = 0; // getItemDoubleValue(form, fidSubTotal);
            req.BillingAddress = new GwAddress();
            req.BillingAddress.Company = "";
            try
            {
                req.BillingAddress.Email = cust.cccust.email;
            }
            catch (Exception) { }
            req.BillingAddress.FirstName = cust.custObj.BillingAddress.FirstName;
            req.BillingAddress.LastName = cust.custObj.BillingAddress.LastName;
            req.BillingAddress.Street = cust.custObj.BillingAddress.Address1;
            req.BillingAddress.City = cust.custObj.BillingAddress.City;
            req.BillingAddress.State = cust.custObj.BillingAddress.State;
            req.BillingAddress.Zip = cust.custObj.BillingAddress.ZipCode;
            req.CreditCardData = new CreditCardData();
            req.ShippingAddress = new GwAddress();
            req.ShippingAddress.Company = "";
            req.ShippingAddress.Email = cust.cccust.email;
            req.ShippingAddress.FirstName = cust.custObj.BillingAddress.FirstName;
            req.ShippingAddress.LastName = cust.custObj.BillingAddress.LastName;
            req.ShippingAddress.Street = cust.custObj.BillingAddress.Address1;
            req.ShippingAddress.City = cust.custObj.BillingAddress.City;
            req.ShippingAddress.State = cust.custObj.BillingAddress.State;
            req.ShippingAddress.Zip = cust.custObj.BillingAddress.ZipCode;
          
            req.CustReceipt = cfgsendCustomerEmail == "Y" ? true : false;
            
        }
        catch(Exception ex)
        {
            errorLog(ex);
        }
        return req;
    }
    private CustomerTransactionRequest createCustomRequest(SAPbouiCOM.Form form)
    {

        form = theActiveForm;
        CustomerTransactionRequest req = new CustomerTransactionRequest();
        try
        {

            req.CustReceipt = cfgsendCustomerEmail == "Y" ? true : false;
            // req.CustReceiptSpecified = false;  //true;
            req.MerchReceipt = true;
            //  req.MerchReceiptSpecified = false; //true;
            req.MerchReceiptEmail = cfgmerchantEmail;
            req.CustReceiptName = "vterm_customer";
            req.MerchReceiptName = cfgMerchantReceipt;
            req.ClientIP = getIP();
            req.Details = new TransactionDetail();
            req.Details.Clerk = oCompany.UserName;
            req.Details.AllowPartialAuth = true;
            LineItem[] li = null;
            setLineItem(form, ref li);
            req.LineItems = li;
            if (theActiveForm.TypeEx == FORMINCOMINGPAYMENT)
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)theActiveForm.Items.Item("20").Specific;
                string s = "";
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    if (getMatrixSelect(oMatrix, "10000127", i))
                    {
                        req.Details.Invoice = getMatrixItem(oMatrix, "1", i);
                        s = s + req.Details.Invoice + ", ";
                    }
                }
                if (s != "")
                    s = (s + "\r\n").Replace(", \r\n", "\r\n");
                if (req.Details.Invoice == "" || req.Details.Invoice== null)
                {
                    req.Details.Invoice = getCustomerID();
                    req.Details.OrderID = req.Details.Invoice;
                }
                else
                {
                    req.Details.OrderID = GetSalesOrderDocNum(req.Details.Invoice);
                }
                req.Details.PONum = req.Details.PONum;
                req.Details.Subtotal = 0; // req.Details.Amount;
                req.Details.Shipping = 0;
                req.Details.Discount = 0;
                req.Details.Tax = 0;
                req.LineItems[0].UnitPrice = req.Details.Amount.ToString();
                if (s != "")
                    req.LineItems[0].ProductName += ". Invoice(s): " + s;
                return req;
            }

            req.Details.Amount = getItemDoubleValue(form, fidBalance);

            req.Details.Discount = getItemDoubleValue(form, fidDiscount);

            req.Details.Shipping = getItemDoubleValue(form, fidFreight);

            req.Details.Tax = getItemDoubleValue(form, fidTax);

            req.Details.Invoice = getFormItemVal(form, fidRecordID);
            req.Details.OrderID = getFormItemVal(form, fidRecordID);
            req.Details.PONum = getFormItemVal(form, "14");
            if (req.Details.PONum == "")
                req.Details.PONum = getFormItemVal(form, fidRecordID);
            if (form.TypeEx == FORMINVOICE)
            {
                if(req.Details.Invoice != "" && req.Details.Invoice != null)
                   req.Details.OrderID = GetSalesOrderDocNum(req.Details.Invoice);
            }
            req.Details.ShipFromZip = "00000";
            req.Details.Description = getFormItemVal(form, fidDesc);
            req.Details.Subtotal = 0; // getItemDoubleValue(form, fidSubTotal);

            req.isRecurring = false;


            req.CustReceipt = cfgsendCustomerEmail == "Y" ? true : false;

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return req;
    }
    private TransactionRequestObject createRequest()
    {
        TransactionRequestObject req = new TransactionRequestObject();
        try
        {
            req.RecurringBilling = new GwRecurringBilling();

            req.AccountHolder = "";
            req.AuthCode = "";
            req.ClientIP = getIP();
            req.CustomerID = "";
            req.CheckData = new CheckData();
            req.CheckData.Account = "";
            req.CheckData.AccountType = "";
            req.Details = new TransactionDetail();
            req.Details.Amount = 0;
            
            req.Details.Discount = 0;
           
            req.Details.Shipping = 0;
            
            req.Details.Tax = 0;
          
            req.Details.Invoice = "";
            req.Details.OrderID = "";
            req.Details.PONum = "";
            req.Details.Description = "";
            req.Details.Subtotal = 0;
           
            req.BillingAddress = new GwAddress();
            req.BillingAddress.Company = "";
            req.BillingAddress.Email = "";
            req.BillingAddress.FirstName = "";
            req.BillingAddress.LastName = "";
            req.BillingAddress.Street = "";
            req.BillingAddress.City = "";
            req.BillingAddress.State = "";
            req.BillingAddress.Zip = "";
            req.CreditCardData = new CreditCardData();
            req.ShippingAddress = new GwAddress();
            req.ShippingAddress.Company = "";
            req.ShippingAddress.Email = "";
            req.ShippingAddress.FirstName = "";
            req.ShippingAddress.LastName = "";
            req.ShippingAddress.Street = "";
            req.ShippingAddress.City = "";
            req.ShippingAddress.State = "";
            req.ShippingAddress.Zip = "";
            LineItem[] li = null;
            setLineItem(null, ref li);
            req.LineItems = li;
            req.CustReceipt = cfgsendCustomerEmail == "Y" ? true : false;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return req;
    }
    private string getFirstName(string name)
    {
        string ret = name;
        try
        {
            string[] names = name.Split(' ');
            return names[0];
        }catch(Exception)
        {

        }
        return ret;
    }
    private string getLastName(string name)
    {
        string ret = name;
        try
        {
            string[] names = name.Split(' ');
            if (names.Length == 1)
                return names[0];
            else
                return names[1];
        }
        catch (Exception)
        {

        }
        return ret;
    }
    
    CreditCardData swipedCard = null;
    public void handleSwipe(SAPbouiCOM.Form form)
    {
        try
        {
            string s = getFormItemVal(form, editCardNum);
            if (s == "")
                return;
            if (s[0] != '%')
            {
                if (swipedCard != null)
                    if (s != swipedCard.CardNumber)
                        swipedCard = null;
                return;
            }
            swipedCard = new CreditCardData();
            swipedCard.CardPresent = true;
            
            swipedCard.MagSupport = "Yes";
            swipedCard.MagStripe = s;  // "enc://" + Base64Encode(s);
            char token = '&';
            for(int i = 16 ; i < 20; i++)
            {
                if(!Char.IsLetterOrDigit(s[i]) && s[i] != ' ')
                {
                    token = s[i];
                }
            }
            string[] sTrack = s.Split(token);
            if (sTrack.Length > 1)
            {
                string CardNumber = sTrack[0].Replace("%B", "");
                trace("swiped card token: " + token);
                //get name
                string name = sTrack[1];
                for (int i = 0; i < name.Length; i++)
                {
                    if (!Char.IsLetterOrDigit(name[i]) && name[i] != ' ')
                    {
                        token = name[i];
                    }
                }
                string fname = "";
                string lname = "";
                string[] n = name.Split(token);
                fname = n[1].Trim();
                lname = n[0].Trim();
                string expd = sTrack[2]; 
                string CardExpiration = expd.Substring(2, 2) + expd.Substring(0, 2);
                swipedCard.CardNumber = CardNumber;
                swipedCard.CardExpiration = CardExpiration;
                setFormEditVal(form, editCardNum, CardNumber);
                setFormEditVal(form, editCardExpDate, CardExpiration);
                if (fname != "" && lname != "")
                {
                    setFormEditVal(form, editHolderName, fname + " " + lname);
                }
            }
        }
        catch (Exception ex)
        {
            swipedCard = null;
            errorLog(ex);
        }
    }
    public bool AddPayment(double total, SAPCust cust, string invID = "", string discountPercent = "")
    {
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
        SAPbobsCOM.Payments oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);

        try
        {
            if (invID == "")
                invID = INVDocEntry;
            if (invID == "")
            {
                showMessage("DocEntry for Invoice not found");
                return false;
            }
            trace(string.Format("Add payment to Invoice ID: {0}, amount: {1}, discount={2}%", invID, total,discountPercent));
              if (!oDoc.GetByKey(int.Parse(invID)))
            {
                showMessage(oCompany.GetLastErrorDescription());
                return false;
            }
            bool isFC = false;
            double dtotal = oDoc.DocTotal;
            double dtotalFC = oDoc.DocTotalFc;
            if (dtotalFC != 0)
                isFC = true;
              oPmt.DocCurrency = oDoc.DocCurrency;
            oPmt.TaxDate = DateTime.Now;
            int bplid = getBranchIdFromINV(invID); 
            if (bplid!= 0)
                oPmt.BPLID = bplid;
            
            string sql = "select \"InstlmntID\", \"InsTotal\", \"InsTotalFC\", \"Paid\", \"PaidFC\" from INV6 where \"Status\" = 'O' and \"DocEntry\" = " + invID;
            oRS.DoQuery(sql);
            double remain = total;
            while (!oRS.EoF)
            {
                int instID = (int)oRS.Fields.Item(0).Value;
                double InsTotal = (double)oRS.Fields.Item(1).Value;
                double InsTotalFC = (double)oRS.Fields.Item(2).Value;
                double Paid = (double)oRS.Fields.Item(3).Value;
                double PaidFC = (double)oRS.Fields.Item(4).Value;
                double bal = InsTotal - Paid;
                
                if (isFC)
                    bal = InsTotalFC - PaidFC;
                trace(string.Format("AddPayment InstID {0},{1},{2},{3},{4}", instID, InsTotal, InsTotalFC, Paid, PaidFC));
                if(bal == 0)
                {
                   if (SBO_Application.MessageBox("Invoice is already paid.  Do you want to void this transaction?", 1, "Yes", "No") == 1)
                   {
                        voidCCTRANS(confirmNum);
                   }
                    return true;

                }
                if (remain >= bal)
                {
                    oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice;
                    if (isFC)
                        oPmt.Invoices.AppliedFC = bal;
                    else
                        oPmt.Invoices.SumApplied = bal;
                    trace("add payment loop remain >= bal;  " + remain + "," + bal);

                    oPmt.Invoices.InstallmentId = (int)oRS.Fields.Item(0).Value;
                    oPmt.Invoices.DocEntry = int.Parse(invID);
                    if (discountPercent != "")
                        oPmt.Invoices.DiscountPercent = getDoubleValue(discountPercent);
                    oPmt.Invoices.Add();
                }
                else if (remain > 0)
                {
                    oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice;
                    if (isFC)
                        oPmt.Invoices.AppliedFC = remain;
                    else
                        oPmt.Invoices.SumApplied = remain;
                    trace("add payment loop remain > 0;  remain=" + remain + ", bal=" + bal);
                    oPmt.Invoices.InstallmentId = (int)oRS.Fields.Item(0).Value;
                    oPmt.Invoices.DocEntry = int.Parse(invID);
                    if (discountPercent != "")
                        oPmt.Invoices.DiscountPercent = getDoubleValue(discountPercent);
                    oPmt.Invoices.Add();
                }
                remain = Math.Round(remain - bal, 2);
                oRS.MoveNext();
            }
            oPmt.CardCode = oDoc.CardCode;
            if (cust.paymentDate == null)
                oPmt.DocDate = DateTime.Now;
            else
                oPmt.DocDate = cust.paymentDate;

            //oPmt.Remarks = "eBizCharge " + confirmNum;
            //oPmt.TransferDate = DateTime.Now;
            oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
            AddCreditCardPayment(oPmt, cust, total, true);
            int r = oPmt.Add();
            if (r != 0)
            {
                errorLog("add payment error: " + oCompany.GetLastErrorDescription());
                
                showMessage(oCompany.GetLastErrorDescription());
                voidCustomer(confirmNum);
                return false;
            }
            else
            {
                CCTRANUpdateIncomingPaymentID(confirmNum);
                showStatus("eBizCharge: incoming payment to invoice: " + INVDocEntry + " added.  Amount: " + total, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            return false;
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oPmt);
            oRS = null;
            oDoc = null;
            oPmt = null;
        }
        return true;
    }
    public bool AddPaymentOnAccount(double amount, SAPCust cust, string account)
    {
        SAPbobsCOM.Payments oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);

        try
        {
           
            account = account.Substring(0, account.IndexOf("-"));
              try
            {
                string bplid = getFormItemVal(theActiveForm, "1320002037");
                if (bplid != "")
                {
                    trace("AddPaymentOnAccount Branch ID: " + bplid);
                    oPmt.BPLID = int.Parse(bplid);
                }
            }
            catch (Exception) { }
            oPmt.CardCode = cust.cccust.CustomerID;
            if (cust.paymentDate == null)
                oPmt.DocDate = DateTime.Now;
            else
                oPmt.DocDate = cust.paymentDate;
            oPmt.DocRate = 0;
            oPmt.DocTypte = 0;
            oPmt.Remarks = "eBizCharge " + confirmNum;
            oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
            oPmt.JournalRemarks = "Incoming Payment - " + cust.cccust.CustomerID;
            AddCreditCardPayment(oPmt, cust, amount, true);
            int r = oPmt.Add();
            if (r != 0)
            {
                errorLog("add payment on account failed." + getErrorString());
                showMessage(getErrorString());
                voidCustomer(confirmNum);
                return false;
            }
            else
                showStatus("eBizCharge: incoming payment on account: " + account + " added.  Amount: " + amount, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
           
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oPmt);
         
            oPmt = null;
        }
        return true;
    }
   
    public bool AddDownPayment(double total, int id, SAPCust cust)
    {
        SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments);
        SAPbobsCOM.Payments oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);

        try
        {
             if (!oDoc.GetByKey(id))
            {
                showMessage(getErrorString());
                return false;
            }
             oPmt.DocCurrency = oDoc.DocCurrency;
            int bplid = getBranchIdFromOPDI(id.ToString());  
            if (bplid != 0)
                oPmt.BPLID = bplid;
            oPmt.Invoices.DocEntry = id;
            oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_DownPayment;
        
            oPmt.CardCode = oDoc.CardCode;
            oPmt.DocDate = DateTime.Now;
            oPmt.Remarks = "eBizCharge Down Payment " + confirmNum;
           
            oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
            if (oDoc.DocTotalFc != 0)
                oPmt.Invoices.AppliedFC = total;
            else
                oPmt.Invoices.SumApplied = total;
            AddCreditCardPayment(oPmt, cust, total, true);
            int r = oPmt.Add();
            if (r != 0)
            {

                errorLog("Add downpayment failed. " + getErrorString());
                showMessage(getErrorString());
                voidCustomer(confirmNum);
                return false;
            }
            CCTRANUpdateIncomingPaymentID(confirmNum);
            showStatus("eBizCharge: incoming payment to invoice: " + id + " added.  Amount: " + oDoc.DocTotal, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
           
            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oPmt);
         
            oDoc = null;
            oPmt = null;
        }
        return false;
    }
    public bool AddDownPayment(string id, SAPCust cust)
    {
        try
        {
            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments);
            int key = int.Parse(id);
            if (!oDoc.GetByKey(key))
            {
                errorLog("get down payment failed. key=" + id + ". " + getErrorString());
                showMessage(getErrorString());
                return false;
            }
            SAPbobsCOM.Payments oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            int bplid = getBranchIdFromOPDI(id);
            if (bplid != 0)
                oPmt.BPLID = bplid;
            oPmt.Invoices.DocEntry = key;
            oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_DownPayment;
            oPmt.CardCode = oDoc.CardCode;
            oPmt.DocDate = DateTime.Now;
            oPmt.Remarks = "eBizCharge Down Payment " + confirmNum;
            //oPmt.CashSum = oDoc.DocTotal;
            oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
         //   string acct = GetConfig("SAPAccountForIncomingPayment");
           // if (acct != "")
                //oPmt.CashAccount = acct;
            AddCreditCardPayment(oPmt, cust, oDoc.DocTotal, true);
            int r = oPmt.Add();
            if (r != 0)
            {

                errorLog("Add downpayment failed." + getErrorString());
                showMessage(getErrorString());
                voidCustomer(confirmNum);
                return false;
            }
            CCTRANUpdateIncomingPaymentID(confirmNum);
            showStatus("eBizCharge: incoming payment to down payment invoice: " + id + " added.  Amount: " + oDoc.DocTotal, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    } 
    public bool AddPayment(string id, SAPCust cust)
    {
        try
        {
            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            int key = int.Parse(id);
            if (!oDoc.GetByKey(key))
            {
                showMessage(getErrorString());
                return false;
            }
            SAPbobsCOM.Payments oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            int bplid = getBranchIdFromINV(id); 
            if (bplid != 0)
                oPmt.BPLID = bplid;
            

            oPmt.Invoices.DocEntry = key;
            oPmt.CardCode = oDoc.CardCode;
            if (cust.paymentDate == null)
                oPmt.DocDate = DateTime.Now;
            else
                oPmt.DocDate = cust.paymentDate;
            oPmt.Remarks = "eBizCharge " + confirmNum;
            //oPmt.CashSum = oDoc.DocTotal;
            oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
            //string acct = GetConfig("SAPAccountForIncomingPayment");
            //if (acct != "")
                //oPmt.CashAccount = acct;
            AddCreditCardPayment(oPmt, cust, oDoc.DocTotal - getInvoicePaidToDate(oDoc.DocEntry), true);
            int r = oPmt.Add();
            if (r != 0)
            {
                errorLog("Add payment failed. " + getErrorString());
                showMessage(getErrorString());
                voidCustomer(confirmNum);
                return false;
            }

            showStatus("eBizCharge: incoming payment to invoice: " + id + " added.  Amount: " + oDoc.DocTotal, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    public int GetCreditCardCodeFromTrans()
    {
        int code = -1;
        string s = "";
        try
        {
            

            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sql = string.Format("select U_CCAccountID from \"@CCTRANS\" where U_refNum = '{0}'", confirmNum);
            oRS.DoQuery(sql);
            int count = oRS.Fields.Count;
            if (count > 0)
            {
                int n = (int)oRS.Fields.Item(0).Value;
                return n;
            }
            else
                errorLog(sql + " not found");

        }
        catch (Exception ex)
        {
            errorLog(s);
            errorLog(ex);
        }
        return code;
    }
  
   
    public DateTime getExpDate(SAPCust cust)
    {
        return getExpDate(cust.custObj);
    }
    public DateTime getExpDate( Customer custObj)
    {
        DateTime dt = DateTime.Today.AddYears(2);
        try
        {
            // "2017-07";
            string d = custObj.PaymentMethodProfiles[0].CardExpiration;
            string s = string.Format("{0}/01/{1}", d.Substring(5, 2), d.Substring(0, 4));
            dt = DateTime.Parse(s);
        }
        catch (Exception ex)
        {
            errorLog(custObj.PaymentMethodProfiles[0].CardExpiration);
            errorLog(ex);
            dt = DateTime.Today.AddYears(2);
        }
        return dt;
    }
    public DateTime getExpDate(string d)
    {
        DateTime dt = DateTime.Today.AddYears(2);
        try
        {
            
            string s = string.Format("{0}/01/{1}", d.Substring(5, 2), d.Substring(0, 4));
            dt = DateTime.Parse(s);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            dt = DateTime.Today.AddYears(2);
        }
        return dt;
    }
     public void AddCreditCardPayment(SAPbobsCOM.Payments vPay, SAPCust cust, double total, bool bIncoming)
    {
        try
        {
            if (cust.custObj.PaymentMethodProfiles[0].MethodType != "check")
                AddCreditCardPayment(vPay, cust.cccust.CustomerID, cust.custObj.PaymentMethodProfiles[0].CardNumber, cust.cccust.cardType, total, cust.cccust.CCAccountID, bIncoming);
            else
                AddCreditCardPayment(vPay, cust.cccust.CustomerID, cust.custObj.PaymentMethodProfiles[0].Account, "eCheck", total, cust.cccust.CCAccountID, bIncoming);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
     public void AddCreditCardPayment(SAPbobsCOM.Payments vPay, string customerID, string cardNumber, string type, double total, string CCAccountID, bool bIncoming)
     {
          try
         {
             vPay.CreditCards.AdditionalPaymentSum = 0;
             vPay.CreditCards.CardValidUntil = DateTime.Today.AddYears(2);
             int n = 0;
             try
             {
                 int.TryParse(CCAccountID, out n);
                 string cardcode = getCardAcctCode(n);
                 if (cardcode != CCAccountID)
                 {
                    string tbname = "OINV";
                    if (vPay.Invoices.InvoiceType == SAPbobsCOM.BoRcptInvTypes.it_DownPayment)
                        tbname = "ODPI";
                     string currency = getCurrency(tbname, vPay.Invoices.DocEntry.ToString());
                     string group = getGroupName(customerID);
                     if (type == "" || type == null)
                         type = "eCheck";
                     string acctid = "";
                    
                     string cardname = getCardName(group, type, currency, ref acctid, vPay.BPLID.ToString());
                     if (acctid != "")
                         n = int.Parse(acctid);
                     else
                         n = int.Parse(cardcode);
                 }
             }
             catch (Exception) { }
             if (confirmNum == null || confirmNum == "")
                 confirmNum = "No Trans";
             vPay.CreditCards.CreditCard = n;
             vPay.CreditCards.CreditCardNumber = cardNumber;
             //vPay.CreditCards.CreditCur = "EUR"
             //vPay.CreditCards.CreditRate = 0
             vPay.CreditCards.CreditSum = total;
             //vPay.CreditCards.CreditType = 1;
             vPay.CreditCards.FirstPaymentDue = DateTime.Now;
             vPay.CreditCards.FirstPaymentSum = total;
             vPay.CreditCards.NumOfCreditPayments = 1;
             vPay.CreditCards.NumOfPayments = 1;
             vPay.CreditCards.VoucherNum = confirmNum;
             vPay.CreditCards.ConfirmationNum = confirmNum;
             vPay.CreditCards.OwnerIdNum = customerID;
             // vPay.CreditCards.OwnerPhone = "383838888"
             vPay.CreditCards.PaymentMethodCode = getPaymentTypeCode(n);
             vPay.CreditCards.Add();
             confirmNum = "";
            trace(string.Format("Add credit card payment of: {0} on card: {1}, confirmNum: {2}, Customer: {3}", total, cardNumber, confirmNum, customerID));

        }
        catch (Exception ex)
         {
              errorLog(ex);
         }
     }
     public bool isDPPaid(string id)
     {
        SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments);

        try
        {
              if (!oDoc.GetByKey(int.Parse(id)))
                 return false;
             if (oDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Close)
                 return true;

         }
         catch (Exception ex)
         {
             errorLog(ex);
        }finally
        {
             System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
             oDoc = null;
            
        }
     
        return false;
     }
    public bool createInvoice(SAPbouiCOM.Form form, SAPCust sapcust)
    {
        SAPbobsCOM.Documents oInvDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
        try
        {
            INVDocEntry = form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0);
            if (INVDocEntry != "")
                return true;

              int nSOID = 0;
            createDocLine(form, ref oInvDoc, ref nSOID);
            SAPbobsCOM.DownPaymentsToDraw dp2d = oInvDoc.DownPaymentsToDraw;
            List<string> list = getDownpaymentToDraw(nSOID.ToString());
            int i = 0;
            foreach (string dp in list)
            {
                if (isDPPaid(dp))
                {
                    dp2d.Add();
                    dp2d.SetCurrentLine(i);
                    dp2d.DocEntry = int.Parse(dp);
                    i++;
                }
            }
            int r = oInvDoc.Add();
            if (r != 0)
            {
                errorLog(getErrorString() + ", form: " + form.TypeEx);
                bAuto = false;
                SBO_Application.MessageBox(getErrorString());
               // voidCustomer(confirmNum);
                return false;
            }
            INVDocEntry = (getNextTableID("OINV") - 1).ToString();
            if(theActiveForm.TypeEx=="60090")  // Invoice + Payment
            {
                SAPbobsCOM.Recordset oCustomerRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oCustomerRS.DoQuery("UPDATE OINV SET IsICT = 'Y' WHERE DocEntry=" + INVDocEntry);
            }
            if (theActiveForm.TypeEx == "60091")  // Reserve Invoice
            {
                SAPbobsCOM.Recordset oCustomerRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oCustomerRS.DoQuery("UPDATE OINV SET UpdInvnt = 'C' WHERE DocEntry=" + INVDocEntry);
            }
            showStatus("Invoice added ,inv id=" + INVDocEntry + ", form: " + form.TypeEx, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
            return false;
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvDoc);
            oInvDoc = null;

        }

    }
   
    public bool  createSO(SAPbouiCOM.Form form, SAPCust sapcust, ref SAPbobsCOM.Documents oDoc)
    {
        try
        {
            SODocEntry = form.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);
            if (SODocEntry != "")
            {
                oDoc.GetByKey(int.Parse(SODocEntry));
                return true;
            }
            int nSOID = 0;
            createDocLine(form, ref oDoc, ref nSOID);
           
            int r = oDoc.Add();
            if (r != 0)
            {
                bAuto = false;
                errorLog("Create SO failed. " + getErrorString());
                SBO_Application.MessageBox(getErrorString());
                voidCustomer(confirmNum);
                return false;
            }
           
            SODocEntry = (getNextTableID("ORDR") - 1).ToString();
            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    public bool createDocLine(SAPbouiCOM.Form form, ref SAPbobsCOM.Documents oDoc, ref int SOID )
    {
        try
        {
            string disc = getFormItemVal(form, "24");
            if (disc != "")
            {
                double d = 0;
                double.TryParse(disc, out d);
                oDoc.DiscountPercent = d;

            }
            double rounding = getItemDoubleValue(form, "103");
            if (rounding > 0)
            {
                oDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES;
                oDoc.RoundingDiffAmount = rounding;
            }
            string s = getFormItemVal(form, "8");
            int sNum = int.Parse(s);
            string manual = getFormItemVal(form, "88");
            if (manual.Trim() == "-1")
                oDoc.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES;
            oDoc.DocNum = sNum;
            s = getCustomerID();
            oDoc.CardCode = s;
            s = getFormItemVal(form, "54");
            oDoc.CardName = s;
            oDoc.DocDueDate = DateTime.Today.AddDays(5);
            oDoc.DocDate = DateTime.Today;
            try
            {
                s = getFormItemVal(form, "20");
                oDoc.SalesPersonCode = int.Parse(s);
            }
            catch (Exception)
            { }
            try
            {
                s = getFormItemVal(form, "89");
                if (s != "")
                {
                    oDoc.Expenses.ExpenseCode = getExpenseCode();
                    oDoc.Expenses.TaxCode = getEmptTaxCode();
                    oDoc.Expenses.LineTotal = getItemDoubleValue(form, "89");
                    oDoc.Expenses.Add();
                }
            }
            catch (Exception)
            { }
           
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item(fidMatrix).Specific;
            if (oMatrix.RowCount >= 2)
            {
                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                for (int i = 1; i < oMatrix.RowCount; i++)
                {
                    s = getMatrixItem(oMatrix, "43", i);
                    if (s != "-1" && s != "")
                    {
                        oDoc.Lines.BaseType = int.Parse(s);

                        s = getMatrixItem(oMatrix, "45", i);
                        oDoc.Lines.BaseEntry = int.Parse(s);
                        oDoc.Lines.BaseLine = i - 1;
                        if (oDoc.Lines.BaseType == 17)
                            SOID = oDoc.Lines.BaseEntry;
                        else if (oDoc.Lines.BaseType == 15)
                            SOID = getSOFromDelivery(oDoc.Lines.BaseEntry.ToString());

                    }
                    s = getMatrixItem(oMatrix, "1", i);
                    if (s != "")
                    {
                       
                        s = getMatrixItem(oMatrix, "23", i);

                        s = getMatrixItem(oMatrix, "1", i);
                        oDoc.Lines.ItemCode = s;

                        s = getMatrixItem(oMatrix, "3", i);
                        oDoc.Lines.ItemDescription = s;

                        s = getMatrixItem(oMatrix, "6", i);
                        oDoc.Lines.SerialNum = s;



                        s = getMatrixItem(oMatrix, "14", i);
                        oDoc.Lines.UnitPrice = getDoubleValue(getMoneyString(s));

                        s = getMatrixItem(oMatrix, "160", i);
                       // if (s == "")
                            //s = defaultTaxCode;  // sapcust.custObj.BillingAddress.State;
                        oDoc.Lines.TaxCode = s;

                        s = getMatrixItem(oMatrix, "21", i);
                        oDoc.Lines.LineTotal = getDoubleValue(getMoneyString(s));

                        //     s = getMatrixItem(oMatrix, "11", i);
                        //int q = int.Parse(s);
                        
                        int q = (int)(oDoc.Lines.LineTotal / oDoc.Lines.UnitPrice);
                        oDoc.Lines.Quantity = q;
                        //string str = string.Format("q = {0}, total={1}, price={2}", q, oDoc.Lines.LineTotal, oDoc.Lines.UnitPrice);
                       // errorLog(str);

                        s = getMatrixItem(oMatrix, "15", i);
                        oDoc.Lines.DiscountPercent = getDoubleValue(s);
                        oDoc.Lines.Add();
                    }else
                    {
                        s = getMatrixItem(oMatrix, "258", i);
                        oDoc.SpecialLines.LineType = SAPbobsCOM.BoDocSpecialLineType.dslt_Text;
                        oDoc.SpecialLines.LineText = s;
                        oDoc.SpecialLines.AfterLineNumber = i - 2;
                        oDoc.SpecialLines.Add();
                    }
                }

            }
            oMatrix = (SAPbouiCOM.Matrix)form.Items.Item("39").Specific;
            if (oMatrix.RowCount >= 2)
            {
                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
                for (int i = 1; i < oMatrix.RowCount; i++)
                {
                    s = getMatrixItem(oMatrix, "23", i);
                    if (s != "-1" && s != "")
                    {
                        oDoc.Lines.BaseType = int.Parse(s);
                        s = getMatrixItem(oMatrix, "25", i);
                        oDoc.Lines.BaseEntry = int.Parse(s);
                        oDoc.Lines.BaseLine = i - 1;
                        if (oDoc.Lines.BaseType == 17)
                            SOID = oDoc.Lines.BaseEntry;
                        else if (oDoc.Lines.BaseType == 15)
                            SOID = getSOFromDelivery(oDoc.Lines.BaseEntry.ToString());
                    }
                    s = getMatrixItem(oMatrix, "257", i);
                    if (s != "T")
                    {
                        s = getMatrixItem(oMatrix, "1", i);
                        oDoc.Lines.ItemDescription = s;

                        s = getMatrixItem(oMatrix, "94", i);
                        oDoc.Lines.AccountCode = getAcctCode(s);


                        s = getMatrixItem(oMatrix, "95", i);
                        oDoc.Lines.TaxCode = s;

                        s = getMatrixItem(oMatrix, "12", i);
                        oDoc.Lines.LineTotal = getDoubleValue(getMoneyString(s));

                        oDoc.Lines.Add();
                    }
                    else
                    {
                        s = getMatrixItem(oMatrix, "258", i);
                        oDoc.SpecialLines.LineType = SAPbobsCOM.BoDocSpecialLineType.dslt_Text;
                        oDoc.SpecialLines.LineText = s;
                        oDoc.SpecialLines.AfterLineNumber = i - 2;
                        oDoc.SpecialLines.Add();
                    }
                }

            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    public bool createDownpaymentInvoice(SAPbouiCOM.Form form, SAPCust sapcust)
    {
        SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments);

        try
        {
            DPDocEntry = form.DataSources.DBDataSources.Item("ODPI").GetValue("DocEntry", 0);
            if (DPDocEntry != "")
                return true;
               oDoc.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptInvoice;
            
            try
            {
                oDoc.DownPaymentPercentage = getDoubleValue(getFormItemVal(form, "205"));
            }catch(Exception)
            {
                oDoc.DownPaymentPercentage = 100;
            }

            int nSOID = 0;
            createDocLine(form, ref oDoc, ref nSOID);
            
            int r = oDoc.Add();
            if (r != 0)
            {
                bAuto = false;
                errorLog("Create Down Payment Invoice. " + getErrorString());
                SBO_Application.MessageBox(getErrorString());
                voidCustomer(confirmNum);
                return false;
            }
            DPDocEntry = (getNextTableID("ODPI") - 1).ToString();
            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
            oDoc = null;
        }
        return false;
    }
    public bool createDownpaymentInvoice(SAPbouiCOM.Form form, SAPCust sapcust, double amount, SAPbobsCOM.Documents oSODoc)
    {
        try
        {
            if (form.TypeEx == FORMDOWNPAYMENT)
            {
                DPDocEntry = form.DataSources.DBDataSources.Item("ODPI").GetValue("DocEntry", 0);
                if (DPDocEntry != "")
                    return true;
            }
            SAPbobsCOM.Documents oDPM = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments);
            int bplid = getBranchIdFromORDR(oSODoc.DocEntry.ToString()); 
            if (bplid != 0)
                oDPM.BPL_IDAssignedToInvoice = bplid;
            int docNum = 0;
            if (form.TypeEx == FORMSALESORDER)
                docNum = getNextTableNum("ODPI");
            else
                docNum = int.Parse(getFormItemVal(form, "8"));
            oDPM.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptInvoice;
            oDPM.DocCurrency = oSODoc.DocCurrency;
            if (docNum == 0)
                docNum = 1;
            oDPM.DocNum = docNum;
            string s = oSODoc.CardCode;
            int soID = oSODoc.DocEntry;
            if (soID == 0)
                soID = getDocEntry("17",oSODoc.DocNum.ToString(), s);
            oDPM.CardCode = s;
            s = oSODoc.CardName;
            oDPM.CardName = s;
            oDPM.Comments = string.Format("Based On Sales Orders {0}.", oSODoc.DocNum);
            oDPM.DocDueDate = oSODoc.DocDueDate;
            oDPM.DocDate = DateTime.Today;
            oDPM.DocType = oSODoc.DocType;
            oDPM.DeferredTax = oSODoc.DeferredTax;
            oDPM.DiscountPercent = oSODoc.DiscountPercent;
            oDPM.DocCurrency = oSODoc.DocCurrency;
            oDPM.SalesPersonCode = oSODoc.SalesPersonCode;
            SAPbobsCOM.Document_Lines olines = oSODoc.Lines;
            double total = 0;
            bool bFC = false;
            for (int i = 0; i < olines.Count; i++)
            {
                olines.SetCurrentLine(i);
                oDPM.Lines.BaseEntry = soID;
                oDPM.Lines.BaseLine = olines.LineNum;
                oDPM.Lines.BaseType = 17;
                if(oSODoc.DocType == SAPbobsCOM.BoDocumentTypes.dDocument_Items)
                {
                    oDPM.Lines.ItemCode = olines.ItemCode;
                    oDPM.Lines.SerialNum = olines.SerialNum;
                    oDPM.Lines.Quantity = olines.Quantity;
                    oDPM.Lines.Price = olines.Price;
                    oDPM.Lines.DiscountPercent = olines.DiscountPercent;
                
                }
                else
                {
                    oDPM.Lines.AccountCode = olines.AccountCode;
                
                }
                oDPM.Lines.ItemDescription = olines.ItemDescription;
                oDPM.Lines.TaxCode = olines.TaxCode;
                oDPM.Lines.LineTotal = olines.LineTotal;
                oDPM.Lines.TaxTotal = olines.TaxTotal;
                oDPM.Lines.TaxPercentagePerRow = olines.TaxPercentagePerRow;
                oDPM.Lines.Currency = olines.Currency;
                oDPM.Lines.RowTotalFC = olines.RowTotalFC;
                if (oDPM.Lines.RowTotalFC != 0)
                {
                    bFC = true;
                    total += olines.RowTotalFC + olines.AppliedTaxFC;
                }
                else
                {

                    total += olines.LineTotal + olines.AppliedTax;
                }

                oDPM.Lines.Add();
               
            }
            double percent = 0;
            if (!bFC)
            {
            
                oDPM.DownPaymentAmount = amount;
            }
            else
            {
                oDPM.DownPaymentAmountFC = amount;
               

            }
            percent = (amount / total) * 100;
            oDPM.DownPaymentStatus = SAPbobsCOM.BoSoStatus.so_Closed;



            oDPM.DownPaymentPercentage = percent;
            int r = oDPM.Add();
            if (r != 0)
            {
                bAuto = false;
                errorLog("Create DPI failed. " + getErrorString() + "\r\nsoID=" + soID.ToString());
                SBO_Application.MessageBox(getErrorString());
                voidCustomer(confirmNum);
                return false;
            }

            DPDocEntry = (getNextTableID("ODPI") - 1).ToString();
        
                      
            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    public bool createCreditMemo(SAPbouiCOM.Form form, SAPCust sapcust)
    {
        try
        {
            CMDocEntry = form.DataSources.DBDataSources.Item("ORIN").GetValue("DocEntry", 0);
            if (CMDocEntry != "")
                return true;
            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
            int nSOID = 0;
            createDocLine(form, ref oDoc, ref nSOID);
            int r = oDoc.Add();
            if (r != 0)
            {
                bAuto = false;
                errorLog("Failed to credit Credit memo. " + getErrorString());
                SBO_Application.MessageBox(getErrorString());
                voidCustomer(confirmNum);
                return false;
            }
            CMDocEntry = (getNextTableID("ORIN") - 1).ToString();
            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    public double getCreditMemoTotal(int id)
    {
        double dRet = 0;
        try
        {
            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
           
            if (!oDoc.GetByKey(id))
                return dRet;

            return oDoc.DocTotal;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return dRet;
    }
    public bool AddOutgoingPayment(SAPbouiCOM.Form form, double total, SAPCust cust)
    {
        try
        {
            if(CMDocEntry == "")
            {
                SBO_Application.MessageBox("DocEntry for Credit Memo not found.");
                return false;
            }
            SAPbobsCOM.Documents oDoc = null;
            oDoc =(SAPbobsCOM.Documents) oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
            int sNum = int.Parse(CMDocEntry);
            if (!oDoc.GetByKey(sNum))
            {
                bAuto = false;
                SBO_Application.MessageBox(getErrorString());
                return false;
            }
            SAPbobsCOM.Payments oPmt;
            if (cfgNegativeIncomingPayment == "Y")
            {
                oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                total = -total;
            }
            else
            {
                oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
   
            }
            int bplid = getBranchIdFromCM(CMDocEntry);
            if (bplid != 0)
                oPmt.BPLID = bplid;
            oPmt.DocCurrency = oDoc.DocCurrency;
            oPmt.Invoices.DocEntry = sNum;
            oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote;
            if (oDoc.DocTotalFc > 0)
                oPmt.Invoices.AppliedFC = total;
            else
                oPmt.Invoices.SumApplied = total;
            //oPmt.CashAccount = GetConfig("SAPAccountForOutgoingPayment");
            oPmt.CardCode = oDoc.CardCode;
            //oPmt.Reference2 = confirmNum;
            oPmt.DocDate = DateTime.Now;
            oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
            //oPmt.CashSum = total;
            oPmt.Remarks = "eBizCharge " + confirmNum;
            AddCreditCardPayment(oPmt, cust, total, false);
            int r = oPmt.Add();
            if (r != 0)
            {
                errorLog("Failed to add out going payment. " + getErrorString());
                bAuto = false;
                SBO_Application.MessageBox(getErrorString());
                voidCustomer(confirmNum);
                return false;
            }
            if (cfgNegativeIncomingPayment == "Y")
            {
                CCTRANUpdateIncomingPaymentID(confirmNum);
            }
            else
            {
                CCTRANUpdateOutgoingPaymentID(confirmNum);

            }
           


            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    public bool AddOutgoingPayment(int docID, string customerID, string cardNum, string type, double total, string CCAccountID)
    {
        try
        {
            SAPbobsCOM.Documents oDoc = null;
            oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
          
            if (!oDoc.GetByKey(docID))
            {
                bAuto = false;
                SBO_Application.MessageBox(getErrorString());
                return false;
            }
            SAPbobsCOM.Payments oPmt;
            if (cfgNegativeIncomingPayment == "Y")
            {
                oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                total = -total;
            }
            else
            {
                oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                
            }
            int bplid = getBranchIdFromCM(docID.ToString());
            if (bplid != 0)
                oPmt.BPLID = bplid;
            oPmt.DocCurrency = oDoc.DocCurrency;
            oPmt.Invoices.DocEntry = docID;
            oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote;
            if(oDoc.DocTotalFc > 0)
                oPmt.Invoices.AppliedFC = Math.Round(total, 2); 
            else
                oPmt.Invoices.SumApplied = Math.Round(total, 2);
            //oPmt.CashAccount = GetConfig("SAPAccountForOutgoingPayment");
            oPmt.CardCode = oDoc.CardCode;
            //oPmt.Reference2 = confirmNum;
            oPmt.DocDate = DateTime.Now;
            oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
            oPmt.Remarks = "eBizCharge " + confirmNum;
            AddCreditCardPayment(oPmt, customerID, cardNum, type, total, CCAccountID, false);
            int r = oPmt.Add();
            if (r != 0)
            {
                errorLog("Failed to create outgoing payment. " + getErrorString());
                bAuto = false;
                SBO_Application.MessageBox(getErrorString());
                voidCustomer(confirmNum);
                return false;
            }
            if (cfgNegativeIncomingPayment == "Y")
            {
                CCTRANUpdateIncomingPaymentID(confirmNum);
            }
            else
            {
                CCTRANUpdateOutgoingPaymentID(confirmNum);

            }
           
    
            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    public bool AddOutgoingPayment(string id, SAPCust cust)
    {
        try
        {
            SAPbobsCOM.Documents oDoc = null;
            oDoc =( SAPbobsCOM.Documents) oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
            int sNum = int.Parse(id);
            if (!oDoc.GetByKey(sNum))
            {
                bAuto = false;
                SBO_Application.MessageBox(getErrorString());
                return false;
            }
            double total = oDoc.DocTotal;
            SAPbobsCOM.Payments oPmt;
            if (cfgNegativeIncomingPayment == "Y")
            {
                oPmt =(SAPbobsCOM.Payments) oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                total = -total;
            }else
            {
                oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);

            }
            int bplid = getBranchIdFromCM(id); 
            if (bplid != 0)
                oPmt.BPLID = bplid;
            oPmt.DocCurrency = oDoc.DocCurrency;
            oPmt.Invoices.DocEntry = sNum;
            oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote;
            if (oDoc.DocTotalFc > 0)
                oPmt.Invoices.AppliedFC = total;
            else
                oPmt.Invoices.SumApplied = total;
            //oPmt.CashAccount = GetConfig("SAPAccountForOutgoingPayment");
            oPmt.CardCode = oDoc.CardCode;
           // errorLog(confirmNum);
           // oPmt.Reference2 = confirmNum;
            oPmt.DocDate = DateTime.Now;
            oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
           // oPmt.CashSum = oDoc.DocTotal;
            oPmt.Remarks = "eBizCharge " + confirmNum;
            AddCreditCardPayment(oPmt, cust, total, true);
            int r = oPmt.Add();
            if (r != 0)
            {
                errorLog("Add outgoing error. " + getErrorString());
                bAuto = false;
                SBO_Application.MessageBox(getErrorString());
                return false;
            }
            if (cfgNegativeIncomingPayment == "Y")
            {
                CCTRANUpdateIncomingPaymentID(confirmNum);
            }
            else
            {
                CCTRANUpdateOutgoingPaymentID(confirmNum);

            }
           
            showStatus("eBizCharge: Outcoming payment to credit memo: " + id + " added.  Amount: " + oDoc.DocTotal,SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    public void SetInvoiceTranLog()
    {
        try
        {
            int invID = getNextTableID("OINV") - 1;
      
            SAPbobsCOM.Documents oInvDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            if (!oInvDoc.GetByKey(invID))
                return;
            SAPbobsCOM.DownPaymentsToDraw dp = oInvDoc.DownPaymentsToDraw;
            for(int i=0; i < dp.Count; i++)
            {
                dp.SetCurrentLine(i);
                updateDownPaymentTranLogInvId(invID, dp.DocEntry);
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
       
    }
    public int getSOFromDelivery(string id)
    {
        try
        {
            SAPbobsCOM.Documents oDelivery = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
            oDelivery.GetByKey(int.Parse(id));
            SAPbobsCOM.Document_Lines oDlines = oDelivery.Lines;
            oDlines.SetCurrentLine(0);
            if (oDlines.BaseType == 17)
            {
                return oDlines.BaseEntry;
            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return 0;
    }
   public void updateCCTRAN(SAPbobsCOM.BoObjectTypes type, string CCAccountID)
   {
      
       try
       {
           updateCCTRANSWithPaymentId(confirmNum, CCAccountID);
           if (type == SAPbobsCOM.BoObjectTypes.oChartOfAccounts)
               return;
           SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(type);
           switch(type)
           {
               case SAPbobsCOM.BoObjectTypes.oOrders:
                   if (!oDoc.GetByKey(int.Parse(SODocEntry)))
                       return;
                   CCTRANAdjustWithSOId(confirmNum, SODocEntry);
                   break;
               case SAPbobsCOM.BoObjectTypes.oInvoices:
                   if (!oDoc.GetByKey(int.Parse(INVDocEntry)))
                       return;
                   CCTRANAdjustWithINVId(confirmNum, INVDocEntry);
                   break;
               case SAPbobsCOM.BoObjectTypes.oDownPayments:
                   if (!oDoc.GetByKey(int.Parse(DPDocEntry)))
                       return;
                   CCTRANAdjustWithDPId(confirmNum, DPDocEntry);
                   break;
               case SAPbobsCOM.BoObjectTypes.oCreditNotes:
                   if (!oDoc.GetByKey(int.Parse(CMDocEntry)))
                       return;
                   CCTRANAdjustWithCMId(confirmNum, CMDocEntry);
                   break;
           }
           SAPbobsCOM.Document_Lines olines = oDoc.Lines;
           olines.SetCurrentLine(0);
    
           if(olines.BaseType == 17)  //base on so
           {
               CCTRANAdjustWithSOId(confirmNum, olines.BaseEntry.ToString());
               if(type == SAPbobsCOM.BoObjectTypes.oInvoices)
               {
                   CCTRANAdjustSOWithInvId(oDoc.DocEntry, olines.BaseEntry);
               }else if(type == SAPbobsCOM.BoObjectTypes.oDownPayments)
               {
                   CCTRANAdjustSOWithDPId(oDoc.DocEntry, olines.BaseEntry);
               }

           }else if(olines.BaseType == 15) // base on Delivery
           {
               SAPbobsCOM.Documents oDelivery = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
               oDelivery.GetByKey(olines.BaseEntry);
               SAPbobsCOM.Document_Lines oDlines = oDelivery.Lines;
               oDlines.SetCurrentLine(0);
               if(oDlines.BaseType == 17)
               {
                   CCTRANAdjustSOWithInvId(oDoc.DocEntry, oDlines.BaseEntry);
               }
           }else if(olines.BaseType == 13)  //From Invoice this only possible if oDoc is a Credit memo
           {
               if (type == SAPbobsCOM.BoObjectTypes.oCreditNotes)
                   CCTRANAdjustInvWithCMId(oDoc.DocEntry, olines.BaseEntry);
           }
                 
               

              
           
       }catch(Exception ex)
       {
           errorLog(ex);
       }
   }
    public string getErrorString()
   {
       string str = "";
        try
        {
            str = "B1 Error: " + oCompany.GetLastErrorDescription();
        }
        catch(Exception){}
        return str;
   }
    public double getShipping(SAPbobsCOM.Documents oDoc)
    {
        double shipping = 0;
        try
        {
            SAPbobsCOM.DocumentsAdditionalExpenses expLines = oDoc.Expenses;

            for (int n = 0; n < expLines.Count; n++)
            {
                expLines.SetCurrentLine(n);
                if (expLines.ExpenseCode != 0)
                {

                    shipping += oDoc.Expenses.LineTotal;
                }
            }
        }
        catch (Exception)
        { }
        return Math.Round(shipping, 2);
    }
    public void setApprovedUDF(bool approved)
    {
        try
        {
            if (cfgApprUDF == "" || cfgApprUDF == null || theActiveForm.TypeEx != FORMSALESORDER)
                return;

            string soid = theActiveForm.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);
            trace(string.Format("Approval UDF: {0}; Approved={1}; SOID={2}", cfgApprUDF, approved, soid));
            SAPbobsCOM.Documents oSO = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
            if (!oSO.GetByKey(int.Parse(soid)))
                return;
            if (approved)
                oSO.UserFields.Fields.Item("U_" + cfgApprUDF).Value = "Approved";
            else
                oSO.UserFields.Fields.Item("U_" + cfgApprUDF).Value = "";
            oSO.Update();
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    public void setCreditCardUDF(string cardnum)
    {
        SAPbobsCOM.Documents oDoc = null;
        string id;
        try
        {

            if (cfgCardNumUDF == "" || cfgCardNumUDF == null)
                return;
            if (theActiveForm.TypeEx == FORMSALESORDER)
            {
                id = theActiveForm.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);
                trace(string.Format("Credit card UDF in SO: {0}; Card Num={1}; ID={2}", cfgCardNumUDF, cardnum, id));

                SAPbobsCOM.Documents oSO = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                if (!oSO.GetByKey(int.Parse(id)))
                    return;
                oSO.UserFields.Fields.Item("U_" + cfgCardNumUDF).Value = cardnum;
                oSO.Update();
            }
            if (theActiveForm.TypeEx == FORMSALESORDER)
            {
                id = theActiveForm.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0);
                trace(string.Format("Credit card UDF in INV: {0}; Card Num={1}; ID={2}", cfgCardNumUDF, cardnum, id));

                oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                if (!oDoc.GetByKey(int.Parse(id)))
                    return;

            }
            if (theActiveForm.TypeEx == FORMINCOMINGPAYMENT)
            {
                id = (getNextTableID("ORCT") - 1).ToString();
                trace(string.Format("Credit card UDF in ORCT: {0}; Card Num={1}; ID={2}", cfgCardNumUDF, cardnum, id));

                oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                if (!oDoc.GetByKey(int.Parse(id)))
                    return;

            }
            oDoc.UserFields.Fields.Item("U_" + cfgCardNumUDF).Value = cardnum;
            oDoc.Update();
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        finally
        {
            if (oDoc != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
                oDoc = null;
            }
        }
    }
    public void setCreditCardUDF(string cardnum, int invID)
    {
        try
        {
            if (cfgCardNumUDF == "" || cfgCardNumUDF == null)
                return;
            trace(string.Format("Set Credit card UDF: {0}; Card Num={1}; InvoiceID={2}", cfgCardNumUDF, cardnum, invID));

            SAPbobsCOM.Documents oSO = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            if (!oSO.GetByKey(invID))
                return;
            oSO.UserFields.Fields.Item("U_" + cfgCardNumUDF).Value = cardnum;
            oSO.Update();
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private SecurityToken getTokenEx(string custid, string group, string currency, string branchid)
    {
        string cardcode = "";
        try
        {
            if (currency == "")
                currency = getSQLString(string.Format("select  \"Currency\" from OCRD where \"CardCode\" = '{0}'", custid));
            if (group == "")
                group = getGroupName(custid);
            if (branchid == "")
                branchid = getBranchIdByCustomer(custid).ToString();
            string acctname = getAccountName(custid);
            string cardName = getCardName(group, "", currency, ref cardcode, acctname, branchid);

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return getToken(cardcode);
    }
    public void AddCreditCardPayment(SAPbobsCOM.Payments vPay, SAPCust cust, double total, bool bIncoming, string refNum)
    {
        try
        {
            if (cust.custObj.PaymentMethodProfiles[0].MethodType != "check")
                AddCreditCardPayment(vPay, cust.cccust.CustomerID, cust.custObj.PaymentMethodProfiles[0].CardNumber, cust.cccust.cardType, total, cust.cccust.CCAccountID, bIncoming, refNum);
            else
                AddCreditCardPayment(vPay, cust.cccust.CustomerID, cust.custObj.PaymentMethodProfiles[0].Account, "eCheck", total, cust.cccust.CCAccountID, bIncoming, refNum);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    public void AddCreditCardPayment(SAPbobsCOM.Payments vPay, string customerID, string cardNumber, string type, double total, string CCAccountID, bool bIncoming, string refNum)
    {
        try
        {
            if (refNum == null || refNum == "")
                refNum = "No Trans";
            vPay.CreditCards.AdditionalPaymentSum = 0;
            vPay.CreditCards.CardValidUntil = DateTime.Today.AddYears(2);
            int n = 0;
            try
            {
                int.TryParse(CCAccountID, out n);
                string cardcode = getCardAcctCode(n);
                if (cardcode != CCAccountID)
                {
                    string tbname = "OINV";
                    if (vPay.Invoices.InvoiceType == SAPbobsCOM.BoRcptInvTypes.it_DownPayment)
                        tbname = "ODPI";
                    string currency = getCurrency(tbname, vPay.Invoices.DocEntry.ToString());
                    string group = getGroupName(customerID);
                    if (type == "" || type == null)
                        type = "eCheck";
                    string acctid = "";
                    string bid = getBranchID();
                    string acctname = getAccountName(customerID);
                    string cardname = getCardName(group, type, currency, ref acctid, acctname, vPay.BPLID.ToString());
                    if (acctid != "")
                        n = int.Parse(acctid);
                    else
                        n = int.Parse(cardcode);
                }
            }
            catch (Exception) { }
            if (n == 0)
            {

                n = getSQLInt("select \"CreditCard\" from  OCRC");
                trace("Cannot resolve card code.  Try the first card code and get. n=" + n);
            }
            vPay.CreditCards.CreditCard = n;
            if (cardNumber == "" || cardNumber == null)
                cardNumber = GetCardNumber(customerID);
            if (cardNumber.Length > 4)
                cardNumber = cardNumber.Substring(cardNumber.Length - 4);
            vPay.CreditCards.CreditCardNumber = cardNumber;
            //vPay.CreditCards.CreditCur = "EUR"
            //vPay.CreditCards.CreditRate = 0
            vPay.CreditCards.CreditSum = total;
            //vPay.CreditCards.CreditType = 1;
            vPay.CreditCards.FirstPaymentDue = DateTime.Now;
            vPay.CreditCards.FirstPaymentSum = total;
            vPay.CreditCards.NumOfCreditPayments = 1;
            vPay.CreditCards.NumOfPayments = 1;
            vPay.CreditCards.VoucherNum = refNum;
            vPay.CreditCards.ConfirmationNum = refNum;
            vPay.CreditCards.OwnerIdNum = customerID;
            if (cfgNegativeIncomingPayment == "N" && bIncoming == false && cfgOtherOutgoingAccount != "")
            {
                vPay.CreditCards.CreditAcct = cfgOtherOutgoingAccount;
                trace("CreditCards.CreditAcct=" + cfgOtherOutgoingAccount);
            }
            // vPay.CreditCards.OwnerPhone = "383838888"
            int pmCode = getPaymentTypeCode(n);
            vPay.CreditCards.PaymentMethodCode = pmCode;
            vPay.CreditCards.Add();
            trace(string.Format("Add credit card payment of: {0} on card: {1}, refNum: {2}, Customer: {3}, Card code: {4}, Payment Method Code:{5}", total, cardNumber, refNum, customerID, n, pmCode));
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    public bool AddPaymentOnAccount(string refNum, double amount, string custid)
    {
        try
        {

            SAPbobsCOM.Payments oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            try
            {

            }
            catch (Exception) { }
            oPmt.CardCode = custid;
            oPmt.DocDate = DateTime.Now;

            oPmt.DocRate = 0;
            oPmt.DocTypte = 0;
            oPmt.Remarks = "eBizCharge " + refNum;
            oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;

            // oPmt.JournalRemarks = "Incoming Payment - " + cust.cccust.CustomerID;
            SAPCust cust = getCustomerByID(custid, "");
            AddCreditCardPayment(oPmt, cust, amount, true, refNum);
            int r = oPmt.Add();
            if (r != 0)
            {
                errorLog(oCompany.GetLastErrorDescription());
                showMessage(oCompany.GetLastErrorDescription());
                voidCustomer(refNum);
                return false;
            }
            else
                showStatus("eBizCharge: incoming payment on account: " + custid + " added.  Amount: " + amount, SAPbouiCOM.BoMessageTime.bmt_Medium, false);

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
} 

