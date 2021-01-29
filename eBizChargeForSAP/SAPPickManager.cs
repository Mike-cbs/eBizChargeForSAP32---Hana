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
using System.Linq;
using eBizChargeForSAP.ServiceReference1;

partial class SAP
{
        private void PickManagerFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                if (pVal.BeforeAction)
                {

                }
                else
                {
                    switch (pVal.EventType)
                    {
                       
                        case SAPbouiCOM.BoEventTypes.et_CLICK:
                           
                            break;
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            {
                                switch (pVal.ItemUID)
                                {
                                    case "11":  //btn Release To pick
                                        BatchPreauth(form);
                                        break;
                                }
                            }
                            break;
                    }

                }
            }
            catch (Exception ex)
            {
                errorLog(ex);
            }

        }
    public void BatchPreauth(SAPbouiCOM.Form form)
        {
            try {
                if (cfgBatchPreauth != "Y")
                    return;
                if (SBO_Application.MessageBox("Run pre-auth on selected orders?", 1, "Yes", "No") != 1)
                    return;
                                
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item("10").Specific;
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    SAPbouiCOM.CheckBox check = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("1").Cells.Item(i).Specific;
                    if(check.Checked)
                    {
                        string docnum = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(i).Specific).Value;
                        string customerid = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("10").Cells.Item(i).Specific).Value;
                        int docentry = getDocEntry("17", docnum, customerid);
                        SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                        if(oDoc.GetByKey(docentry))
                        {
                            if(oDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Close)
                                SBO_Application.SetStatusBarMessage("eBizCharge Batch pre-auth sales order : " + docnum + " is closed.", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                            else
                            {
                                SAPCust sapcust = getCustomerByID(customerid, "");
                                if (sapcust == null)
                                    SBO_Application.SetStatusBarMessage("eBizCharge Batch pre-auth sales order : " + docnum + " error customer: " + customerid + " has no payment method.", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                                else
                                {
                                    if (isOrderPreauthed(docentry.ToString()))
                                        SBO_Application.SetStatusBarMessage("eBizCharge Batch pre-auth sales order : " + docnum + " order already pre-authed", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                                    else
                                    {
                                        BatchPreauth_PreAuthCustomer(oDoc, sapcust);
                                    }
                                }
                            }

                        }else
                        {
                            SBO_Application.SetStatusBarMessage("eBizCharge Batch pre-auth can not retrieve sales order : " + docnum + ". " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                        }

                    }
                }
            
            }
            catch (Exception ex)
            {

            }
        }
    public bool BatchPreauth_PreAuthCustomer(SAPbobsCOM.Documents doc, SAPCust sapcust)
    {
        try
        {

            CustomerTransactionRequest req = Batch_createCustomRequest(doc);
            req.CustReceipt = cfgsendCustomerEmail == "Y" ? true : false;
            req.CustReceiptName = "vterm_customer";
            req.CustReceiptEmail = sapcust.custObj.Email;
            
            req.Command = "cc:authonly";
            if (req.Command == "cc:authonly" && cfgPreAuthEmail != "Y")
                req.CustReceipt = false;
            if (req.CustReceiptEmail == "")
                req.CustReceipt = false;
            string refNum = "";
            if (Batch_runCustomRequest(req, sapcust, doc.DocEntry.ToString(), "" , ref refNum))
            {
                confirmNum = refNum;
                BatchPreAuthUpdateTrans(doc.DocEntry, doc.CardCode, refNum);
                SBO_Application.SetStatusBarMessage("eBizCharge Batch pre-auth sales order : " + doc.DocNum + " approved. Approval Code:" + refNum + ", Amount:" + doc.DocTotal.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
            else
            {
                return false;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    private CustomerTransactionRequest Batch_createCustomRequest(SAPbobsCOM.Documents doc)
    {


        CustomerTransactionRequest req = new CustomerTransactionRequest();
        try
        {

            req.ClientIP = getIP();
            req.CustReceipt = cfgsendCustomerEmail == "Y" ? true : false;
             //true;
            req.MerchReceipt = true;
            //true;
            req.MerchReceiptEmail = cfgmerchantEmail;
            req.Details = new TransactionDetail();
            req.Details.Clerk = oCompany.UserName;
            req.Details.AllowPartialAuth = true;
            
            req.Details.Amount = Math.Round(doc.DocTotal, 2);
            
            req.Details.Discount = Math.Round(doc.TotalDiscount,2);
           
            req.Details.Shipping = getShipping(doc);
            
            req.Details.Tax = doc.VatSum;
          
            req.Details.Invoice = doc.DocNum.ToString();
            req.Details.OrderID = GetSalesOrderDocNum(req.Details.Invoice);
            req.Details.PONum = doc.NumAtCard;
            if (req.Details.PONum == "" || req.Details.PONum == null)
                req.Details.PONum = doc.DocNum.ToString();
            req.Details.ShipFromZip = getZipFromAddr(doc.Address2);
            req.Details.Description = doc.Comments;
            req.Details.Subtotal = 0; // req.Details.Subtotal = Math.Round(req.Details.Amount - (req.Details.Shipping + req.Details.Tax - req.Details.Discount), 2);
           
            SAPbobsCOM.Document_Lines olines = doc.Lines;
            req.LineItems = new LineItem[olines.Count];
       
            for (int i = 0; i < olines.Count; i++)
            {
                olines.SetCurrentLine(i);
                req.LineItems[i] = new LineItem();
                req.LineItems[i].CommodityCode = merchantData.ItemCommodityCode == null ? olines.ItemCode : merchantData.ItemCommodityCode;
                req.LineItems[i].Qty =  olines.Quantity.ToString();
                req.LineItems[i].UnitPrice = olines.Price.ToString();
                req.LineItems[i].ProductName = olines.ItemDescription;
                req.LineItems[i].UnitOfMeasure = "EACH";
                req.LineItems[i].SKU = olines.ItemCode;
                req.LineItems[i].CommodityCode = olines.ItemCode;
                req.LineItems[i].Taxable = true;
                //req.LineItems[i].TaxableSpecified = true;
                req.LineItems[i].TaxRate = olines.TaxPercentagePerRow.ToString();
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return req;
    }
    private TransactionRequestObject Batch_createRequest(SAPbobsCOM.Documents doc)
    {


        TransactionRequestObject req = new TransactionRequestObject();
        try
        {
            req.BillingAddress = new GwAddress();
            req.ClientIP = getIP();
            req.CustReceipt = false;
             //true;
            req.Details = new TransactionDetail();
            req.Details.Clerk = oCompany.UserName;
            req.Details.AllowPartialAuth = true;
            
            req.Details.Amount = Math.Round(doc.DocTotal, 2);
            
            req.Details.Discount = Math.Round(doc.TotalDiscount, 2);
           
            req.Details.Shipping = getShipping(doc);
            
            req.Details.Tax = doc.VatSum;
          
            req.Details.Invoice = doc.DocNum.ToString();
            req.Details.OrderID = GetSalesOrderDocNum(req.Details.Invoice);
            req.Details.PONum = doc.NumAtCard;
            if (req.Details.PONum == "" || req.Details.PONum == null)
                req.Details.PONum = doc.DocNum.ToString();
            req.Details.ShipFromZip = getZipFromAddr(doc.Address2);
            req.Details.Description = doc.Comments;
            req.Details.Subtotal = 0; // req.Details.Subtotal = Math.Round(req.Details.Amount - (req.Details.Shipping + req.Details.Tax - req.Details.Discount), 2);
           
            SAPbobsCOM.Document_Lines olines = doc.Lines;
            req.LineItems = new LineItem[olines.Count];

            for (int i = 0; i < olines.Count; i++)
            {
                olines.SetCurrentLine(i);
                req.LineItems[i] = new LineItem();
                req.LineItems[i].CommodityCode = merchantData.ItemCommodityCode == null ? olines.ItemCode : merchantData.ItemCommodityCode;
                req.LineItems[i].Qty = olines.Quantity.ToString();
                req.LineItems[i].UnitPrice = olines.Price.ToString();
                req.LineItems[i].ProductName = olines.ItemDescription;
                req.LineItems[i].UnitOfMeasure = "EACH";
                req.LineItems[i].SKU = olines.ItemCode;
                req.LineItems[i].CommodityCode = olines.ItemCode;
                req.LineItems[i].Taxable = true;
               // req.LineItems[i].TaxableSpecified = true;
                req.LineItems[i].TaxRate = olines.TaxPercentagePerRow.ToString();
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return req;
    }
    private bool Batch_runCustomRequest(CustomerTransactionRequest req, SAPCust sapcust, string orderid, string invoiceid , ref string refNum)
    {

        try
        {
            req.CustReceipt = cfgsendCustomerEmail == "Y" ? true : false;
            req.CustReceiptName = "vterm_customer";
            req.CustReceiptEmail = sapcust.custObj.Email;
            //form.Freeze(true);
            req.Details.Subtotal = 0;
            trace("Batch_runCustomRequest: command=" + req.Command + ", amount = " + req.Details.Amount + ", tax=" + req.Details.Tax + ",Subtotal=" + req.Details.Subtotal + ",shipping=" + req.Details.Shipping + ",Discount=" + req.Details.Discount);

            SecurityToken token = getToken(sapcust.cccust.CCAccountID);
            TransactionResponse resp = new TransactionResponse();
         
            resp = ebiz.runCustomerTransaction(token, sapcust.custObj.CustomerToken, sapcust.custObj.PaymentMethodProfiles[0].MethodID, req);
            trace("Batch_runCustomRequest: refNum = " + resp.RefNum + ", error=" + resp.Error);


            if (sapcust != null)
            {
                refNum = resp.RefNum;
                try
                {
                    if (resp.ResultCode == "D")
                    {

                        updateCustDeclined(sapcust.cccust.recID, "Y");

                    }
                    if (resp.ResultCode == "A")
                    {
                        updateCustDeclined(sapcust.cccust.recID, "N");

                    }
                }
                catch (Exception ex)
                {
                    errorLog(ex);
                }
            }
       
            logCustomTransaction(req, resp, sapcust, orderid, invoiceid);
            return Batch_HandleResponse(resp, sapcust, ref refNum);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            SBO_Application.MessageBox(ex.Message);
        }
        return false;
    }
    private bool Batch_HandleResponse(TransactionResponse resp, SAPCust sapcust, ref string refNum)
    {
        try
        {
            if (resp.ResultCode == "A")
            {
                return true;
            }
            else if (resp.Error.ToLower() == "approved")
            {
                return true;
            }
            else
            {
                string msg = "";
                if (sapcust != null)
                {
                    if (sapcust.custObj.PaymentMethodProfiles[0].MethodType == "check")
                        msg = "Can not process your transaction for checking account: " + sapcust.custObj.PaymentMethodProfiles[0].Account + ".\r\n" + resp.Error;
                    else
                        msg = "Can not process your transaction for card: " + sapcust.custObj.PaymentMethodProfiles[0].CardNumber + ".\r\n" + resp.Error;
                }
                else
                {
                    msg = "Can not process your transaction for card.\r\n" + resp.Error;
                }

                SBO_Application.SetStatusBarMessage("eBizCharge Batch processing failed. " + msg, SAPbouiCOM.BoMessageTime.bmt_Medium, true);

              
                return false;
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    string getZipFromAddr(string addr)
    {
        string zip = "";
        try
        {
            addr = addr.Replace("-", " ");
            string[] l = addr.Split(' ');
            foreach(string s in l)
            {
                if(s.Length == 5)
                {
                    if (s.All(char.IsDigit))
                        zip = s;
                }
            }


        }catch(Exception)
        { }
        return zip;
    }
}
