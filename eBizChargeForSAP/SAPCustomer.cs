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
using System.Diagnostics;
using eBizChargeForSAP.ServiceReference1;

partial class SAP  {

    public bool AddCustomer( ref SAPCust sapcust)
    {
        try
        {
           // refresh(db.CCCUSTs);
            if (sapcust == null)
            {
                showStatus("Can not add credit card. Invalid eBizCharge Customer object");
                return false;
            }
           
            string custID = sapcust.cccust.CustomerID;
            string CCAccountID = sapcust.cccust.CCAccountID;
            SecurityToken token = getToken(sapcust.cccust.CCAccountID);
            if (sapcust.cccust == null)
                sapcust.cccust = new CCCUST();
            if (IsWalkIn(custID))
                sapcust.custObj.CustomerId = sapcust.custObj.CustomerId + "wi" + DateTime.Now.ToString("yyMMddhhmmss");
            else
                sapcust.custObj.CustomerId = sapcust.cccust.CustomerID;
            trace("addCustomer securityKey:" + token.SecurityId + ", CCAccountID = " + CCAccountID + ", CustomerID=" + sapcust.custObj.CustomerId);

            sapcust.cccust.CCAccountID = CCAccountID;
            sapcust.cccust.CustomerID = custID;
            sapcust.cccust.active = 'Y';
            sapcust.cccust.email = sapcust.custObj.Email;
            sapcust.cccust.firstName = sapcust.custObj.BillingAddress.FirstName;
            sapcust.cccust.lastName = sapcust.custObj.BillingAddress.LastName;
            sapcust.cccust.street = sapcust.custObj.BillingAddress.Address1;
            sapcust.cccust.city = sapcust.custObj.BillingAddress.City;
            sapcust.cccust.state = sapcust.custObj.BillingAddress.State;
            sapcust.cccust.zip = sapcust.custObj.BillingAddress.ZipCode;
            sapcust.cccust.expDate = sapcust.custObj.PaymentMethodProfiles[0].CardExpiration;
            sapcust.cccust.cardCode = sapcust.custObj.PaymentMethodProfiles[0].CardCode;
            sapcust.cccust.routingNumber = sapcust.custObj.PaymentMethodProfiles[0].Routing;
            sapcust.cccust.checkingAccount = sapcust.custObj.PaymentMethodProfiles[0].Account;
            sapcust.custObj.BillingAddress.CompanyName= getGroupName(custID);

            sapcust.custObj.PaymentMethodProfiles[0].MethodID = "";
            sapcust.custObj.PaymentMethodProfiles[0].CardNumber = sapcust.custObj.PaymentMethodProfiles[0].CardNumber.Replace("X", ""); 
             Customer c = null;
            try
            {
                c = ebiz.GetCustomer(token, sapcust.custObj.CustomerId, "");
            
            }catch(Exception ex)
            {
                if (ex.Message.ToLower() != "not found")
                {
                    trace("Add customer on ebiz.GetCustomer exception. " + ex.Message);
                    showMessage("add Customer error: " + ex.Message);
                    return false;
                }
            }
            if (c == null)
            {
                CustomerResponse ret = ebiz.AddCustomer(token, sapcust.custObj);
                if (ret.ErrorCode != 0)
                {
                      trace("Add Customer error. " + ret.Error);
                    showMessage(ret.Error);
                    return false;
                }else
                    c = ebiz.GetCustomer(token, ret.CustomerId, ret.CustomerInternalId);
            }
            sapcust.custObj.PaymentMethodProfiles[0].AvsStreet = sapcust.custObj.BillingAddress.Address1;
            sapcust.custObj.PaymentMethodProfiles[0].AvsZip = sapcust.custObj.BillingAddress.ZipCode;

            string mid = ebiz.AddCustomerPaymentMethodProfile(token, c.CustomerInternalId, sapcust.custObj.PaymentMethodProfiles[0]);
            trace("add payment method returns: " + mid);
            c = ebiz.GetCustomer(token, sapcust.custObj.CustomerId, "");
            sapcust.custObj.CustomerToken = c.CustomerToken;
            foreach (PaymentMethodProfile p in c.PaymentMethodProfiles)
            {
                if (p.MethodID == mid)
                {
                    sapcust.custObj.PaymentMethodProfiles[0] = p;
                    trace("Customer Payment profile: " + sapcust.custObj.PaymentMethodProfiles[0].CardNumber);
                }
            }

            sapcust.cccust.CustNum = sapcust.custObj.CustomerToken;
            if (sapcust.custObj.PaymentMethodProfiles.Count() > 0)
                sapcust.cccust.MethodID = sapcust.custObj.PaymentMethodProfiles[0].MethodID;
            int id = getNextTableID("@CCCUST");
            string desc = sapcust.custObj.BillingAddress.FirstName + " " + sapcust.custObj.BillingAddress.LastName;
            if (sapcust.custObj.PaymentMethodProfiles[0].MethodName != null && sapcust.custObj.PaymentMethodProfiles[0].MethodName != "")
                desc = sapcust.custObj.PaymentMethodProfiles[0].MethodName;
            sapcust.cccust.methodName = sapcust.custObj.PaymentMethodProfiles[0].MethodName;
            if (sapcust.custObj.PaymentMethodProfiles[0].MethodType == "check")
            {
                sapcust.cccust.methodDescription = id.ToString() + "_" + sapcust.custObj.PaymentMethodProfiles[0].Routing + " " + sapcust.custObj.PaymentMethodProfiles[0].Account + "(" + desc + ")";
                sapcust.cccust.checkingAccount = sapcust.custObj.PaymentMethodProfiles[0].Account;
                sapcust.cccust.routingNumber = sapcust.custObj.PaymentMethodProfiles[0].Routing;
            }
            else
            {
                sapcust.cccust.cardNum = sapcust.custObj.PaymentMethodProfiles[0].CardNumber;
                sapcust.cccust.methodDescription = id.ToString() + "_" + sapcust.custObj.PaymentMethodProfiles[0].CardNumber + " " + sapcust.custObj.PaymentMethodProfiles[0].CardExpiration + "(" + desc + ")";
                sapcust.cccust.cardType = sapcust.custObj.PaymentMethodProfiles[0].CardType;
            }
            if (!IsWalkIn(custID))
            {
                insert(sapcust.cccust);
                showStatus("eBizCharge credit card " + sapcust.custObj.PaymentMethodProfiles[0].CardNumber + " added.  Customer ID:" + sapcust.cccust.CustomerID, SAPbouiCOM.BoMessageTime.bmt_Medium, false);

            }
            else
                showStatus("Walk in customer.  card no. " + sapcust.custObj.PaymentMethodProfiles[0].CardNumber + ".  Customer ID:" + sapcust.cccust.CustomerID, SAPbouiCOM.BoMessageTime.bmt_Medium, false);


            return true ;
        }
        catch (Exception ex)
        {
            trace("Add customer exception. " + ex.Message);

            if (sapcust.custObj == null)
            {
                showMessage("Add customer failed.\r\n" + ex.Message);
            }
            else
            {
                if (sapcust.custObj.PaymentMethodProfiles.Count() > 0)
                {
                    if(sapcust.custObj.PaymentMethodProfiles[0].MethodType == "check")
                        showMessage("Add customer for account: " + sapcust.custObj.PaymentMethodProfiles[0].Account + " failed.\r\n" + ex.Message);
                    else
                        showMessage("Add customer for card: " + sapcust.custObj.PaymentMethodProfiles[0].CardNumber + " failed.\r\n" + ex.Message);
                }
                else
                    showMessage("Add customer failed.\r\n" + ex.Message);
            }
        }
        return false;
    }
   
    public  Customer getCustomerByCustNum(string CustNum)
    {
        try
        {
            string id = "";
            string cardcode = getCardAcctByCustNum(CustNum, ref id);
            SecurityToken token = getToken(cardcode);

            return ebiz.GetCustomer(token,id, "");
  
        }
        catch (Exception ex)
        {
            showMessage(ex.Message);
        }
        return null;
    }
    public  Customer AddCustomer(string id)
    {
        try
        {
            SAPbobsCOM.Recordset oCustomerRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oCustomerRS.DoQuery(string.Format("Select \"CardName\",\"Address\",\"City\",\"State1\",\"ZipCode\", \"MailAddres\",\"MailCity\",\"State2\",\"MailZipCod\" from OCRD where \"CardCode\" = '{0}'", id));
            string cardcode = getCardAcctByCustomerID(id);
            SecurityToken token = getToken(cardcode);
             Customer CustomerData = new  Customer();
            CustomerData.BillingAddress = new Address();
            CustomerData.CustomerId = id;
            CustomerData.BillingAddress.FirstName =getFirstName( GetFieldVal(oCustomerRS, 0));
            CustomerData.BillingAddress.LastName = getLastName(GetFieldVal(oCustomerRS, 0));
            CustomerData.BillingAddress.Address1 = GetFieldVal(oCustomerRS, 1);
            CustomerData.BillingAddress.City = GetFieldVal(oCustomerRS, 2);
            CustomerData.BillingAddress.State = GetFieldVal(oCustomerRS, 3);
            CustomerData.BillingAddress.ZipCode = GetFieldVal(oCustomerRS, 4);
            CustomerResponse ret = ebiz.AddCustomer(token, CustomerData);
            theSAPCustomer = getCustomer(id);
           
        }
        catch (Exception ex)
        {
            showMessage(ex.Message);
            errorLog(ex);
        }
        return theSAPCustomer;
    }

    public bool updateCustomer(ref SAPCust sapcust)
    {
        bool bResult = true;
        try
        {
            SecurityToken token = getToken(sapcust.cccust.CCAccountID);
            try
            {
                sapcust.custObj.PaymentMethodProfiles[0].AvsZip = sapcust.custObj.BillingAddress.ZipCode;
                sapcust.custObj.PaymentMethodProfiles[0].AvsStreet = sapcust.custObj.BillingAddress.Address1;
                if (!ebiz.UpdateCustomerPaymentMethodProfile(token, sapcust.custObj.CustomerToken, sapcust.custObj.PaymentMethodProfiles[0]))
                    showMessage("Failed to update payment method");
            }
            catch (Exception)
            {

            }
            Customer custObj = ebiz.GetCustomer(token, sapcust.cccust.CustomerID, "");
            custObj.BillingAddress = sapcust.custObj.BillingAddress;
            custObj.Email = sapcust.custObj.Email;
            custObj.LastName = sapcust.custObj.BillingAddress.LastName;
            custObj.FirstName = sapcust.custObj.BillingAddress.FirstName;
            CustomerResponse ret = ebiz.UpdateCustomer(token, custObj, custObj.CustomerId, custObj.CustomerInternalId);
            if (ret.StatusCode != 1)
                bResult = false;
            if (bResult)
            {
                custObj = ebiz.GetCustomer(token, sapcust.cccust.CustomerID, "");
                foreach (PaymentMethodProfile pm in custObj.PaymentMethodProfiles)
                {
                    if (pm.MethodID == sapcust.custObj.PaymentMethodProfiles[0].MethodID)
                    {
                        sapcust.custObj.PaymentMethodProfiles[0] = pm;
                    }
                }
                sapcust.cccust.email = custObj.Email;
                sapcust.cccust.firstName = custObj.BillingAddress.FirstName;
                sapcust.cccust.lastName = custObj.BillingAddress.LastName;
                sapcust.cccust.street = custObj.BillingAddress.Address1;
                sapcust.cccust.city = custObj.BillingAddress.City;
                sapcust.cccust.state = custObj.BillingAddress.State;
                sapcust.cccust.zip = custObj.BillingAddress.ZipCode;

                sapcust.cccust.cardType = sapcust.custObj.PaymentMethodProfiles[0].CardType;
                sapcust.cccust.cardNum = sapcust.custObj.PaymentMethodProfiles[0].CardNumber;
                sapcust.cccust.MethodID = sapcust.custObj.PaymentMethodProfiles[0].MethodID;
                sapcust.cccust.cardCode = "***";
                sapcust.cccust.routingNumber = sapcust.custObj.PaymentMethodProfiles[0].Routing;
                sapcust.cccust.checkingAccount = sapcust.custObj.PaymentMethodProfiles[0].Account;
                string acctID = "";
                string desc = sapcust.custObj.BillingAddress.FirstName + " " + sapcust.custObj.BillingAddress.LastName;
                if (sapcust.custObj.PaymentMethodProfiles[0].MethodName != null && sapcust.custObj.PaymentMethodProfiles[0].MethodName != "")
                    desc = sapcust.custObj.PaymentMethodProfiles[0].MethodName;
                if (sapcust.custObj.PaymentMethodProfiles[0].MethodType == "check")
                {
                    sapcust.cccust.methodDescription = sapcust.cccust.recID.ToString() + "_" + sapcust.custObj.PaymentMethodProfiles[0].Routing + " " + sapcust.custObj.PaymentMethodProfiles[0].Account + "(" + desc + ")";
                }
                else
                {
                    sapcust.cccust.methodDescription = sapcust.cccust.recID.ToString() + "_" + sapcust.custObj.PaymentMethodProfiles[0].CardNumber + " " + sapcust.custObj.PaymentMethodProfiles[0].CardExpiration + "(" + desc + ")";
                }
                update(sapcust.cccust);
            }
        }
        catch (Exception ex)
        {
            trace("Payment method update failed.\r\n" + ex.Message);
            bResult = false;
        }
        return bResult;
    }
    public  Customer getCustomer(string CustomerID)
    {
         Customer custObj = new  Customer();
        try
        {
            SecurityToken token = getToken();
    
            custObj = ebiz.GetCustomer(token, CustomerID,"");
        }catch(Exception ex)
        {
           errorLog(ex);            
        }
        return custObj;
    }

  
     
    public string getSAPCustomerID(SAPbouiCOM.Form form)
    {
        string id = "";
        try
        {
            SAPbouiCOM.EditText customer = (SAPbouiCOM.EditText)form.Items.Item(fidCustID).Specific;
            return customer.Value;
        }catch(Exception)
        {

        }
        return id;
    }
   
    public void deleteCustomer(string custNum)
    {
        try
        {
            /*
             SecurityToken token = getToken();

            ebiz.deleteCustomer(token, custNum);
            */
        }catch(Exception ex)
        {
            errorLog(ex);
        }

    }
  
    private bool runCustomRequest(SAPbouiCOM.Form form, CustomerTransactionRequest req, SAPCust sapcust, ref string refNum)
    {

        try
        {
            
            SecurityToken token = getToken(sapcust.cccust.CCAccountID);
            TransactionResponse resp = new TransactionResponse();
            if (swipedCard != null)
            {
                swipedCard.CardCode = req.CardCode;
                swipedCard.AvsStreet = sapcust.custObj.BillingAddress.Address1;
                swipedCard.AvsZip = sapcust.custObj.BillingAddress.ZipCode;
                TransactionRequestObject reqT = new TransactionRequestObject();
                reqT.RecurringBilling = new GwRecurringBilling();
                reqT.AccountHolder = "";
                reqT.AuthCode = "";
                reqT.Command = req.Command;
                reqT.ClientIP = req.ClientIP;
                reqT.CustomerID = sapcust.cccust.CustomerID;
                reqT.CheckData = new CheckData();
                reqT.CheckData.Account = "";
                reqT.CheckData.AccountType = "";
                reqT.Details = req.Details;
                GwAddress a = new GwAddress();
                copyAddress(ref a, sapcust.custObj.BillingAddress);
                reqT.BillingAddress = a;
                reqT.CreditCardData = swipedCard;
                reqT.ShippingAddress = a;
                reqT.LineItems = req.LineItems;
                trace("runCustomRequest swipe card : command=" + req.Command + ", amount = " + req.Details.Amount + ", tax=" + req.Details.Tax + ",Subtotal=" + req.Details.Subtotal + ",shipping=" + req.Details.Shipping + ",Discount=" + req.Details.Discount);
                reqT.CustReceipt = cfgsendCustomerEmail == "Y" ? true : false;
                reqT.CustReceiptName = "vterm_customer";
                try
                {
                    reqT.BillingAddress.Email = sapcust.custObj.Email;
                }
                catch (Exception) { }
                if (reqT.Command == "cc:authonly" && cfgPreAuthEmail != "Y")
                    reqT.CustReceipt = false;
                if (a.Email == "" || a.Email == null)
                    reqT.CustReceipt = false;
                resp = ebiz.runTransaction(token, reqT);
                
                trace("runCustomRequest swipe card result=" + resp.Result + ",  refNum=" + resp.RefNum + ", Error = " + resp.Error);
            }
            else
            {
                req.Details.Subtotal = 0;
                  req.CustReceipt = cfgsendCustomerEmail == "Y" ? true : false;
                req.CustReceiptName = "vterm_customer";
                req.CustReceiptEmail = sapcust.custObj.Email;
                if (req.Command == "cc:authonly" && cfgPreAuthEmail != "Y")
                    req.CustReceipt = false;
                if (req.CustReceiptEmail == "")
                    req.CustReceipt = false;
                trace("runCustomRequest: command=" + req.Command + ", amount = " + req.Details.Amount + ", tax=" + req.Details.Tax + ",Subtotal=" + req.Details.Subtotal + ",shipping=" + req.Details.Shipping 
                    + ",Discount=" + req.Details.Discount + ",email = " + req.CustReceiptEmail + ", cfgPreAuthEmail = " + cfgPreAuthEmail);

                resp = ebiz.runCustomerTransaction(token, sapcust.custObj.CustomerToken, sapcust.custObj.PaymentMethodProfiles[0].MethodID, req);
                trace("runCustomRequest: result=" + resp.Result + ",  refNum=" + resp.RefNum + ", Error = " + resp.Error);

            }
            if (sapcust != null)
            {
                refNum = resp.RefNum;
                try
                {
                    if (resp.ResultCode == "D")
                    {
                        bDeclined = true;
                        updateCustDeclined(sapcust.cccust.recID, "Y");
                        setApprovedUDF(false);
                        if (req.CustReceiptEmail != "" && req.CustReceiptEmail != null && cfgCustomerReceipt != "")
                        {

                            try
                            {
                                EmailReceiptResponse emresp = ebiz.EmailReceipt(token, refNum, refNum, cfgCustomerReceipt, req.CustReceiptEmail);

                                if (emresp.ErrorCode == 0)
                                {
                                    showStatus("Declined receipt sent.", SAPbouiCOM.BoMessageTime.bmt_Medium,false);
                                }
                                else
                                    showStatus("Failed to send declined receipt.\r\n" + resp.Error, SAPbouiCOM.BoMessageTime.bmt_Medium, true);

                            }
                            catch (Exception ex2)
                            {
                                showStatus("Failed to send receipt.\r\n" + ex2.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                            }

                        }
                    }
                    if (resp.ResultCode == "A")
                    {
                        bDeclined = false;
                        updateCustDeclined(sapcust.cccust.recID, "N");
                        setApprovedUDF(true);
                    }
                }catch(Exception ex)
                {
                    errorLog(ex);
                }
            }
            if (resp.ResultCode == "A")
            {
                //sendReceipt(token, sapcust, resp);
                printReceipt(token, sapcust, resp);
               
                logCustomTransaction(req, resp, sapcust);
            }
            return HandleResponse(resp, sapcust, ref refNum);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            showMessage(ex.Message);
        }
        return false;
    }
    public void copyAddress(ref GwAddress a1, Address a2)
    {
        try
        {
            a1.Street = a2.Address1;
            a1.Street2 = a2.Address1;
            a1.City = a2.City;
            a1.State = a2.State;
            a1.Zip = a2.ZipCode;
            a1.LastName = a2.LastName;
            a1.FirstName = a2.FirstName;
            a1.Company = a2.CompanyName;
            a1.Country = a2.Country;
           
        }catch(Exception )
        {

        }
    }
    public void copyAddress(ref Address a1, GwAddress a2)
    {
        try
        {
            a1.Address1 = a2.Street;
            a1.Address2 = a2.Street2;
            a1.City = a2.City;
            a1.State = a2.State;
            a1.ZipCode = a2.Zip;
            a1.LastName = a2.LastName;
            a1.FirstName = a2.FirstName;
            a1.CompanyName = a2.Company;
            a1.Country = a2.Country;

        }
        catch (Exception)
        {

        }
    }
    private void sendReceipt(SecurityToken token, SAPCust sapcust, TransactionResponse resp)
    {
        try
        {
           // string target = "email";
            if (resp.ResultCode == "A")
            {
                /*
                if (cfgmerchantEmail != "" && cfgMerchantReceipt != "")
                {
                    try
                    {

                        //Receipt rcp = ebiz.getReceiptByName(token, cfgMerchantReceipt);
                        EmailReceiptResponse eresp = ebiz.EmailReceipt(token, resp.RefNum, resp.RefNum, cfgMerchantReceipt, cfgmerchantEmail);
                       
                    }
                    catch (Exception ex)
                    {
                        string err = cfgMerchantReceipt + " Not found.\r\n" + ex.Message + "\r\n";
                        try
                        {
                            Receipt[] receipts = ebiz.GetReceiptsList(token, target);
                            foreach (Receipt r in receipts)
                            {
                                err = err + r.ReceiptRefNum + "_" + r.Name + ";\r\n";
                            }
                        }
                        catch (Exception)
                        {
                        }
                        errorLog(err);
                    }
                }
                */
                if (cfgsendCustomerEmail == "Y" && sapcust.custObj.Email != null && sapcust.custObj.Email != "")
                {
                    try
                    {
                        //ebiz.emailTransactionReceiptByName(token, resp.RefNum, cfgCustomerReceipt, sapcust.custObj.Email);
                        EmailReceiptResponse eresp = ebiz.EmailReceipt(token, resp.RefNum, resp.RefNum, cfgCustomerReceipt, sapcust.custObj.Email);
                        trace("send receipt to: " + sapcust.custObj.Email);
                    }
                    catch (Exception ex)
                    {
                        string err = cfgCustomerReceipt + " Not found.\r\n" + ex.Message + "\r\n";
                       
                        errorLog(err);
                    }
                }
            }
            /*
            if (resp.ResultCode == "D")
            {
                if (cfgmerchantEmail != "" && cfgmerchantEmail != null && cfgDeclineMerchantReceipt != "" && cfgDeclineMerchantReceipt != null)
                {
                    try
                    {

                        //Receipt rcp = ebiz.getReceiptByName(token, cfgMerchantReceipt);
                        //ebiz.emailTransactionReceiptByName(token, resp.RefNum, cfgDeclineMerchantReceipt, cfgmerchantEmail);
                        //EmailReceiptResponse eresp = ebiz.EmailReceipt(token, resp.RefNum, resp.RefNum, cfgMerchantReceipt, cfgmerchantEmail);
                  
                    }
                    catch (Exception ex)
                    {
                        string err = cfgDeclineReceipt + " Not found.\r\n" + ex.Message + "\r\n";
                       
                        errorLog(err);
                    }
                }
                if (cfgsendCustomerEmail == "Y" && cfgDeclineReceipt != null && cfgDeclineReceipt != "" 
                    && sapcust.custObj.Email != null && sapcust.custObj.Email != "")
                {
                    try
                    {

                       // EmailReceiptResponse eresp = ebiz.EmailReceipt(token, resp.RefNum, resp.RefNum, cfgMerchantReceipt, cfgmerchantEmail);
                   
                    }
                    catch (Exception ex)
                    {
                        string err = cfgDeclineReceipt + " Not found.\r\n" + ex.Message + "\r\n";
                       
                        errorLog(err);
                    }
                }
            }*/

        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
    private void printReceipt(SecurityToken token, SAPCust sapcust, TransactionResponse resp)
    {
        try
        {
            if (cfgSaveReceipt != "Y")
                return;
            string folder = @"C:\CBS\";
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            folder = @"C:\CBS\Receipts\";
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            string filename = folder + resp.RefNum + ".html";

            string s = ebiz.RenderReceipt(token, resp.RefNum,null, "vterm", "HTML");
       
            if (s != "")
            {
                byte[] todecode = Convert.FromBase64String(s);
                s = string.Concat(System.Text.Encoding.UTF8.GetString(todecode));
                s = s.Replace("<tr><td align=right><b>AMOUNT:</b></td>", "<tr style='display:none'><td align=right><b>AMOUNT:</b></td>");
                s = s.Replace("<tr><td align=right><b>TAX:</b></td>", "<tr style='display:none'><td align=right><b>TAX:</b></td>");
 
                FileStream fs = new FileStream(filename, FileMode.Create, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs);
                sw.Write(s);
                sw.Close();
                System.Diagnostics.Process.Start(filename);
            }
            try
            {
                if(userEmail != "")
                    ebiz.EmailReceipt(token, resp.RefNum, null, "vterm", userEmail);

            }
            catch (Exception ex)
            {
                errorLog(ex);
            }

        }
        catch (Exception ex2)
        {
            errorLog(ex2);
        }
    }
    private bool runCustomRequest(TransactionRequestObject req, SAPCust sapcust,  ref string refNum)
    {

        try
        {

            //form.Freeze(true);
            string ccacountid = "";
            if (sapcust != null)
                ccacountid = sapcust.cccust.CCAccountID;
            if (refNum != "")
                ccacountid = getCCAcountIDByRefNum(refNum);
            req.Details.Subtotal = 0;
            SecurityToken token = getToken(ccacountid);
            trace("accid:" + ccacountid + ",securityID=" + token.SecurityId);
            
            try
            {
                if (req.BillingAddress == null)
                    req.BillingAddress = new GwAddress();
                req.BillingAddress.Email = sapcust.custObj.Email;
                req.CustReceipt = cfgsendCustomerEmail == "Y" ? true : false;
                req.CustReceiptName = "vterm_customer";
           
                if (req.Command == "cc:authonly" && cfgPreAuthEmail != "Y")
                    req.CustReceipt = false;
                if (sapcust.custObj.Email == "" || sapcust.custObj.Email == null)
                    req.CustReceipt = false;
            }
            catch (Exception) { }

            TransactionResponse resp = new TransactionResponse();
            trace("runTrans : command=" + req.Command + ", amount = " + req.Details.Amount + ", tax=" + req.Details.Tax + ",Subtotal=" + req.Details.Subtotal + ",shipping=" + req.Details.Shipping + ",Discount=" + req.Details.Discount);
            resp = ebiz.runTransaction(token, req);
            trace("runTrans result=" + resp.Result + ",  refNum=" + resp.RefNum + ", Error = " + resp.Error + ", ResultCode=" + resp.ResultCode + ", ErrorCode=" + resp.ErrorCode);
            if (sapcust != null)
            {
                if (resp.ResultCode == "D")
                {
                   updateCustDeclined(sapcust.cccust.recID, "Y");
                  
                }
                if (resp.ResultCode == "A")
                {
                  updateCustDeclined(sapcust.cccust.recID, "N");
                    printReceipt(token, sapcust, resp);
                }
            }
            logCustomTransaction(req, resp, sapcust);
            return HandleResponse(resp, sapcust,  ref refNum);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    private bool HandleResponse(TransactionResponse resp, SAPCust sapcust,  ref string refNum)
    {
        try
        {
            if (resp.ResultCode == "A")
            {
                bool berror = false;
                string msg = "Warning!\r\n\r\n";
                if (resp.AvsResultCode == "N" || resp.AvsResultCode == "NN" || resp.AvsResultCode == "NNN")
                {
                    berror = true;
                    msg = msg + "AVS verification failed.  " + resp.AvsResult + "\r\n";
                }
                if (resp.CardCodeResultCode == "N")
                {
                    berror = true;
                    msg = msg + "CVV2 verification failed.  " + resp.CardCodeResult + "\r\n";
                }
                if (berror)
                {
                    msg = msg + "\r\nDo you want to continue?";
                    if (SBO_Application.MessageBox(msg, 1, "Yes", "No") == 2)
                    {
                        voidCustomer(resp.RefNum, true);
                        return false;
                    }
                }
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

                if (msg.IndexOf("already been settled") >= 0)
                    return true;
                showMessage(msg);
                return false;
            }

        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    private void logCustomTransaction(CustomerTransactionRequest req, TransactionResponse resp, SAPCust sapcust, string orderid = "", string invoiceid = "")
    {
   
        try
        {
           
            CCTRAN cctran = new CCTRAN();
            cctran.customerID = sapcust.cccust.CustomerID;
            cctran.CCAccountID = 0;
            cctran.OrderID = orderid;
            cctran.InvoiceID = invoiceid;
            cctran.CreditMemoID = "";
            cctran.DownPaymentInvoiceID="";
            try
            {
                switch(theActiveForm.TypeEx)
                {
                    case FORMSALESORDER:
                        cctran.OrderID = theActiveForm.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);
                        trace("logCustomTransaction theActiveForm is sales order with id=" + cctran.OrderID);
                        break;
                    case FORMINVOICE:
                        cctran.InvoiceID = theActiveForm.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0);
                        trace("logCustomTransaction theActiveForm is Invoice with id=" + cctran.InvoiceID);
                        break;
                }

            }catch(Exception) { }


            if (sapcust != null)
            {
                int n = 0;
                int.TryParse(sapcust.cccust.CCAccountID, out n);
                cctran.CCAccountID = n;
                cctran.customerID = sapcust.cccust.CustomerID;
                cctran.MethodID = sapcust.custObj.PaymentMethodProfiles[0].MethodID;
                cctran.custNum = sapcust.cccust.CustNum;
                cctran.crCardNum = sapcust.custObj.PaymentMethodProfiles[0].CardNumber;
                cctran.CardHolder = sapcust.custObj.BillingAddress.FirstName + " " + sapcust.custObj.BillingAddress.LastName;
                if (sapcust.SaveCard == "N")
                {
                    updateCustState(sapcust.cccust.recID, "N");
                   // refresh(db.CCCUSTs);
                }
            }
            cctran.Description = req.Details.Description;
            cctran.recID = req.Details.Invoice;
            cctran.acsUrl = resp.AcsUrl;
            cctran.authAmount = resp.AuthAmount.ToString();
            cctran.authCode = resp.AuthCode;
            cctran.avsResult = resp.AvsResult;
            cctran.avsResultCode = resp.AvsResultCode;
            cctran.batchNum = resp.BatchNum;
            cctran.batchRef = resp.BatchRefNum;
            cctran.cardCodeResult = resp.CardCodeResult;
            cctran.cardCodeResultCode = resp.CardCodeResultCode;
            cctran.cardLevelResult = resp.CardLevelResult;
            cctran.cardLevelResultCode = resp.CardLevelResultCode;
            cctran.conversionRate = resp.ConversionRate.ToString();
            cctran.convertedAmount = resp.ConvertedAmount.ToString();
            cctran.convertedAmountCurrency = resp.ConvertedAmountCurrency.ToString();
            cctran.custNum = resp.CustNum;
            cctran.error = resp.Error;
            cctran.errorCode = resp.ErrorCode;
            cctran.isDuplicate = resp.isDuplicate.ToString();
            cctran.payload = resp.Payload;
            cctran.profilerScore = resp.ProfilerScore;
            cctran.profilerResponse = resp.ProfilerResponse;
            cctran.profilerReason = resp.ProfilerReason;
            cctran.refNum = resp.RefNum;
            cctran.remainingBalance = resp.RemainingBalance.ToString();
            cctran.result = resp.Result;
            cctran.resultCode = resp.ResultCode;
            cctran.status = resp.Status;
            cctran.statusCode = resp.StatusCode;
            cctran.vpasResultCode = resp.VpasResultCode;
            cctran.recDate = DateTime.Now;//Use local time not server time
            cctran.command = req.Command;
            cctran.amount = req.Details.Amount.ToString();
            insert(cctran);

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
   
    //for capture
    private void logCustomTransaction(TransactionRequestObject req, TransactionResponse resp, SAPCust sapcust)
    {
       
        try
        {

             CCTRAN cctran = new CCTRAN();
            
            cctran.OrderID = "";
            cctran.InvoiceID = "";
            cctran.CreditMemoID = "";
            if (sapcust != null)
            {
                int n = 0;
                int.TryParse(sapcust.cccust.CCAccountID, out n);
                cctran.CCAccountID = n;
                cctran.customerID = sapcust.cccust.CustomerID;
                cctran.MethodID = sapcust.custObj.PaymentMethodProfiles[0].MethodID;
                cctran.custNum = sapcust.cccust.CustNum;
                cctran.crCardNum = sapcust.custObj.PaymentMethodProfiles[0].CardNumber;
                if (sapcust.SaveCard == "N")
                {
                    updateCustState(sapcust.cccust.recID, "N");
                   
                }
            }
            cctran.Description = req.Details.Description;
            cctran.recID = req.Details.Invoice;
            cctran.acsUrl = resp.AcsUrl;
            cctran.authAmount = resp.AuthAmount.ToString();
            cctran.authCode = resp.AuthCode;
            cctran.avsResult = resp.AvsResult;
            cctran.avsResultCode = resp.AvsResultCode;
            cctran.batchNum = resp.BatchNum;
            cctran.batchRef = resp.BatchRefNum;
            cctran.cardCodeResult = resp.CardCodeResult;
            cctran.cardCodeResultCode = resp.CardCodeResultCode;
            cctran.cardLevelResult = resp.CardLevelResult;
            cctran.cardLevelResultCode = resp.CardLevelResultCode;
            cctran.conversionRate = resp.ConversionRate.ToString();
            cctran.convertedAmount = resp.ConvertedAmount.ToString();
            cctran.convertedAmountCurrency = resp.ConvertedAmountCurrency.ToString();
            cctran.custNum = resp.CustNum;
            cctran.error = resp.Error;
            cctran.errorCode = resp.ErrorCode;
            cctran.isDuplicate = resp.isDuplicate.ToString();
            cctran.payload = resp.Payload;
            cctran.profilerScore = resp.ProfilerScore;
            cctran.profilerResponse = resp.ProfilerResponse;
            cctran.profilerReason = resp.ProfilerReason;
            cctran.refNum = resp.RefNum;
            cctran.remainingBalance = resp.RemainingBalance.ToString();
            cctran.result = resp.Result;
            cctran.resultCode = resp.ResultCode;
            cctran.status = resp.Status;
            cctran.statusCode = resp.StatusCode;
            cctran.vpasResultCode = resp.VpasResultCode;
            cctran.recDate = DateTime.Now;//Use local time not server time
            cctran.command = req.Command;
            cctran.amount = req.Details.Amount.ToString();
            insert(cctran);
          

        }
        catch (Exception ex)
        {
            errorLog(ex);
    
        }

    }
    /*
    public int checkAccountID(string AccID, string cardNum, string type, bool incoming)
    {
        int n = 0;
        try
        {
            if(!int.TryParse(AccID, out n))
            {
                return GetCreditCardCode(cardNum,type, incoming);
            }
            else
            {
                if (isCardCodeExists(AccID))
                    return n;
                else
                    return GetCreditCardCode(cardNum, type, incoming);
            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return n;
    }
     * */
    private void logCustomTransactionCreditMemo(CustomerTransactionRequest req, TransactionResponse resp, string customerID, string methodID, string cardNum, string custNum, string creditmemoid, string CCAccountID)
    {

        try
        {

            CCTRAN cctran = new CCTRAN();
            cctran.OrderID = "";
            cctran.InvoiceID = "";
            cctran.CreditMemoID = creditmemoid;
            cctran.customerID = customerID;
            cctran.MethodID = methodID;
            cctran.custNum = custNum;
            
                int n = 0;
                int.TryParse(CCAccountID, out n);
                cctran.CCAccountID = n;
            
            cctran.crCardNum = cardNum;
            cctran.Description = req.Details.Description;
            cctran.recID = req.Details.Invoice;
            cctran.acsUrl = resp.AcsUrl;
            cctran.authAmount = resp.AuthAmount.ToString();
            cctran.authCode = resp.AuthCode;
            cctran.avsResult = resp.AvsResult;
            cctran.avsResultCode = resp.AvsResultCode;
            cctran.batchNum = resp.BatchNum;
            cctran.batchRef = resp.BatchRefNum;
            cctran.cardCodeResult = resp.CardCodeResult;
            cctran.cardCodeResultCode = resp.CardCodeResultCode;
            cctran.cardLevelResult = resp.CardLevelResult;
            cctran.cardLevelResultCode = resp.CardLevelResultCode;
            cctran.conversionRate = resp.ConversionRate.ToString();
            cctran.convertedAmount = resp.ConvertedAmount.ToString();
            cctran.convertedAmountCurrency = resp.ConvertedAmountCurrency.ToString();
            cctran.custNum = resp.CustNum;
            cctran.error = resp.Error;
            cctran.errorCode = resp.ErrorCode;
            cctran.isDuplicate = resp.isDuplicate.ToString();
            cctran.payload = resp.Payload;
            cctran.profilerScore = resp.ProfilerScore;
            cctran.profilerResponse = resp.ProfilerResponse;
            cctran.profilerReason = resp.ProfilerReason;
            cctran.refNum = resp.RefNum;
            cctran.remainingBalance = resp.RemainingBalance.ToString();
            cctran.result = resp.Result;
            cctran.resultCode = resp.ResultCode;
            cctran.status = resp.Status;
            cctran.statusCode = resp.StatusCode;
            cctran.vpasResultCode = resp.VpasResultCode;
            cctran.recDate = DateTime.Now;//Use local time not server time
            cctran.command = req.Command;
            cctran.amount = req.Details.Amount.ToString();
            insert(cctran);
            

        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }
   
    
    public bool ChargeCustomer(SAPbouiCOM.Form form, SAPCust sapcust)
    {
        try
        {
            confirmNum = "";
            string strAmt = "";
            if (!GetTotalAmount(form, ref strAmt))
                return false;

            if(!handleCCCust(ref sapcust))
                return false;
            CustomerTransactionRequest req = createCustomRequest(theActiveForm);
            req.CardCode = getFormItemVal(form, editCardCode);
            req.Details.Comments = sapcust.cccust.GroupName;
            // req.CustomFields[0].Field = 
            // req.CustomFields[0].Value = sapcust.cccust.GroupName;
            // req.CustomFields[1].Value = sapcust.cccust.CCAccountID;
           // req.CustReceiptName = sapcust.custObj.PaymentMethodProfiles[0].MethodName;
            //req.CustReceiptEmail = sapcust.custObj.Email;
            if (sapcust.cccust.zip != "" && sapcust.cccust.zip != null)
                req.Details.ShipFromZip = sapcust.cccust.zip;
            req.Details.Amount = getItemDoubleValue(form, editCAmount);
          
            if (req.Details.Amount < 0)
            {

                req.Command = "cc:credit";
                
            }
            else
                req.Command = "cc:sale";
            if (sapcust != null)
            {
                if (sapcust.custObj != null)
                {
                    if (sapcust.custObj.PaymentMethodProfiles.Count() > 0)
                    {
                        if (sapcust.custObj.PaymentMethodProfiles[0].MethodType == "check")
                        {
                            if (req.Details.Amount < 0)
                            {

                                req.Command = "checkcredit";
                             
                            }
                            else
                                req.Command = "check";
                        }
                    }
                }
            }
            if (req.Details.Amount < 0)
            { 
                req.Details.Amount = -req.Details.Amount;
            }
            if (runCustomRequest(form, req, sapcust,  ref confirmNum))
            {

                string msg = "Charge Completed. Approval Code:" + confirmNum + ", Amount:" + strAmt;
                double amt = getItemDoubleValue(form, editCAmount);
                if (amt < 0)
                    msg = msg.Replace("Charge ", "Refund ");

                showMessage(msg);
                return true;
            }
            else
                return false;

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
  
    public bool runCustomerTrans(string cmd, double dAmt, string custNum, string cardNum, string customerID, string creditmemoID, string CCAccountID)
    {
        try
        {
            confirmNum = "";
            errorLog(string.Format("{0}, {1}, {2}, {3}", cmd, CCAccountID, creditmemoID, customerID));
             Customer custObj = getCustomerByCustNum(custNum);
            if (custObj == null)
                return false;
            CustomerTransactionRequest req = createCustomRequest(theActiveForm);
           // req.CustReceiptName = custObj.PaymentMethodProfiles[0].MethodName;
            req.CustReceiptEmail = custObj.Email;
           
            req.Details.Amount = dAmt;
            req.Details.Subtotal =0;
            req.Details.Shipping = 0;
            req.Details.Tax = 0;
            req.Details.Discount = 0;
            req.Command = cmd;
            SecurityToken token = getToken(CCAccountID);
            TransactionResponse resp = new TransactionResponse();
            resp = ebiz.runCustomerTransaction(token, custObj.CustomerToken, custObj.PaymentMethodProfiles[0].MethodID, req);
            confirmNum = resp.RefNum;
            logCustomTransactionCreditMemo(req, resp, customerID, custObj.PaymentMethodProfiles[0].MethodID, cardNum, custNum, creditmemoID, CCAccountID);
            if (resp.ErrorCode == "0")
            {
                return true;
            }
            else
            {
                errorLog(resp.Error);
                showMessage("Can not process your transaction.\r\n" + resp.Error);
                return false;
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            showMessage("Can not process your transaction.\r\n" + ex.Message);
            return false;
        }

    }
    public bool voidCustomer(string confNum, bool bError = true)
    {
        try
        {
            if (confNum == "")
                return false;
            if (bError && !bRunJob)
            {
                if (SBO_Application.MessageBox("Error occurred.  Do you want to void this transaction?", 1, "Yes", "No") != 1)
                    return false;
            }
            TransactionRequestObject req = createRequest();
            req.RefNum = confNum;
            req.Command = "cc:void";
            string refNum = confNum;
            if (runCustomRequest(req, null,  ref refNum))
            {
                if (bError)
                    voidErrorCCTRANS(refNum);
                else
                    voidCCTRANS(refNum);
                showMessage("Transaction voided. Reference Number:" + refNum);

            }
            else
                return false;

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    public bool voidTransaction(string confNum)
    {
        try
        {
            if (confNum == "")
                return false;
            TransactionRequestObject req = createRequest();
            req.RefNum = confNum;
            req.Command = "cc:void";
            string refNum = confNum;
            if (runCustomRequest(req, null, ref refNum))
            {

                showMessage("Transaction voided. Reference Number:" + refNum);
            }
            else
                return false;

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    public bool CaptureCustomer(SAPbouiCOM.Form form, SAPCust sapcust, CCTRAN ccTran, double amount)
    {
        try
        {
          //  if (!handleCCCust(ref sapcust))
                //return false;
            TransactionRequestObject req = createRequest(theActiveForm, sapcust);
            req.RefNum = ccTran.refNum;
            req.Command = "cc:capture";
            req.Details.Amount = amount;  // getDoubleValue(ccTran.amount);
            string refNum = ccTran.refNum;
            if (runCustomRequest(req, sapcust, ref refNum))
            {

                confirmNum = refNum;
                CaptureUpdateCCTrans(ccTran, refNum);
              
               // string remark = GetRemark(theActiveForm);
               // remark += string.Format(strHeaderCaptured + "ref. No.:{0} **\r\n", refNum, DateTime.Now.ToString("MM-dd-yyy hh:mm"));
               // SetRemark(theActiveForm, remark);
                showMessage("Transaction Captured. Reference Number:" + refNum);



            }
            else
                return false;
             
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    public bool PreAuthCustomer(SAPbouiCOM.Form form, SAPCust sapcust)
    {
        try
        {
            string strAmt = "";
            if (!GetTotalAmount(form, ref strAmt))
                return false;
            if (!handleCCCust(ref sapcust))
                return false;
            CustomerTransactionRequest req = createCustomRequest(theActiveForm);
            
            req.CardCode = getFormItemVal(form, editCardCode);
          //  req.CustReceiptName = sapcust.custObj.PaymentMethodProfiles[0].MethodName;
            req.CustReceiptEmail = sapcust.custObj.Email;
            req.Details.Amount = getItemDoubleValue(form, editCAmount);
            if (sapcust.cccust.zip != "" && sapcust.cccust.zip != null)
                req.Details.ShipFromZip = sapcust.cccust.zip;
            if (req.Details.Discount < 0)
            {
                req.Details.Discount = 0;
            }
            if (req.Details.Amount < (req.Details.Shipping + req.Details.Tax - req.Details.Discount))
            {
                req.Details.Shipping = 0;
                req.Details.Tax = 0;
                req.Details.Discount = 0;
            }


            req.Command = "cc:authonly";
            string refNum = "";
            if (runCustomRequest(form, req, sapcust, ref refNum))
            {
                confirmNum = refNum;
                //string remark = GetRemark(theActiveForm);
                // remark += string.Format(strHeaderApproved + "ref. No.:{0} Amount:{1} {2}**\r\n", refNum, strAmt, DateTime.Now.ToString("MM-dd-yyy hh:mm"));
                //SetRemark(theActiveForm, remark);
                showMessage("Approved. Approval Code:" + refNum + ", Amount:" + strAmt);

            }
            else
                return false;

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    public bool handleCCCust(ref SAPCust sapcust)
    {
        try
        {
        
            if (sapcust == null)
                return false;
            if (sapcust.cccust == null)
                return false;
            if (sapcust.custObj == null)
                return false;

            if (sapcust.custObj.CustomerToken == "" || sapcust.custObj.CustomerToken == null)
                return AddCustomer(ref sapcust);
            if (sapcust.Changed())
            {
                if (sapcust.cccust == null)
                    return false;
                if (sapcust.custObj == null)
                    return false;
                sapcust.custObj.CustomerId = sapcust.cccust.CustomerID + "_" + sapcust.cccust.recID;
                sapcust.custObj.BillingAddress.CompanyName = sapcust.cccust.GroupName;
                 Customer cobj = sapcust.custObj;
                if (cobj == null)
                    return false;
                sapcust.cccust.email = cobj.Email;
                sapcust.cccust.firstName = cobj.BillingAddress.FirstName;
                sapcust.cccust.lastName = cobj.BillingAddress.LastName;
                sapcust.cccust.street = cobj.BillingAddress.Address1;
                sapcust.cccust.city = cobj.BillingAddress.City;
                sapcust.cccust.state = cobj.BillingAddress.State;
                sapcust.cccust.zip = cobj.BillingAddress.ZipCode;
                sapcust.cccust.expDate = cobj.PaymentMethodProfiles[0].CardExpiration;
                sapcust.cccust.cardCode = cobj.PaymentMethodProfiles[0].CardCode;
                sapcust.cccust.routingNumber = cobj.PaymentMethodProfiles[0].Routing;
                sapcust.cccust.checkingAccount = cobj.PaymentMethodProfiles[0].Account;
                if (cobj.PaymentMethodProfiles[0].MethodType == "check")
                    sapcust.cccust.methodDescription = sapcust.cccust.recID.ToString() + "_" + cobj.PaymentMethodProfiles[0].Routing + " " + cobj.PaymentMethodProfiles[0].Account + "(" + cobj.BillingAddress.FirstName + " " + cobj.BillingAddress.LastName + ")";
                else
                    sapcust.cccust.methodDescription = sapcust.cccust.recID.ToString() + "_" + cobj.PaymentMethodProfiles[0].CardNumber + " " + cobj.PaymentMethodProfiles[0].CardExpiration + "(" + cobj.BillingAddress.FirstName + " " + cobj.BillingAddress.LastName + ")";
                sapcust.cccust.cardType = cobj.PaymentMethodProfiles[0].CardType;
                sapcust.cccust.Declined = 'N';
                sapcust.cccust.cardCode = "***";
                sapcust.cccust.cardNum = cobj.PaymentMethodProfiles[0].CardNumber;
                sapcust.cccust.routingNumber = cobj.PaymentMethodProfiles[0].Routing;
                sapcust.cccust.checkingAccount = cobj.PaymentMethodProfiles[0].Account;
                update(sapcust.cccust);
            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    public bool CreditCustomer(SAPbouiCOM.Form form, SAPCust sapcust)
    {
        try
        {

            string strAmt = "";
            if (!GetTotalAmount(form, ref strAmt))
                return false;
            if(!handleCCCust(ref sapcust))
                return false;
            CustomerTransactionRequest req = createCustomRequest(theActiveForm);
            req.CardCode = getFormItemVal(form, editCardCode);
            //req.CustReceiptName = sapcust.custObj.PaymentMethodProfiles[0].MethodName;
           // req.CustReceiptEmail = sapcust.custObj.Email;
            double a = req.Details.Amount;
            req.Details.Amount = getItemDoubleValue(form, editCAmount);
           // req.Details.Subtotal = req.Details.Amount;
          
            if (req.Details.Tax == 0)
                req.Details.Tax = Math.Round(req.Details.Amount * getDoubleValue(cfgDefaultTaxRate), 2);
          
            if (sapcust.cccust.zip != "" && sapcust.cccust.zip != null)
                req.Details.ShipFromZip = sapcust.cccust.zip;
            req.Command = "cc:credit";
            if(sapcust != null)
            {
                if(sapcust.custObj != null)
                {
                    if(sapcust.custObj.PaymentMethodProfiles.Count() > 0)
                    {
                        if(sapcust.custObj.PaymentMethodProfiles[0].MethodType == "check")
                            req.Command = "checkcredit";
                    }
                }
            }
            string refNum = "";
            if (runCustomRequest(form, req, sapcust, ref refNum))
            {
                confirmNum = refNum;
               // string remark = GetRemark(form);
               // remark += string.Format(strHeaderCredited + "ref. No.:{0} Amount:{1} {2}**\r\n", refNum, strAmt, DateTime.Now.ToString("MM-dd-yyy hh:mm"));
                //SetRemark(form, remark);
                showMessage("Transaction Completed. Approval Code:" + refNum + ", Amount:" + strAmt);
                
            }
            else
                return false;

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
  
  
    private string getBranchID()
    {
        try
        {
            if(bRunJob)
                return "0";
            if (!IsBranchEnabled())
                return "0";
            string Custid = getCustomerID();
            string DocNum = getDocNum();
            int docEntry = 0;
            if (theActiveForm == null)
            {
                return getBranchIdByCustomer(Custid).ToString();
            }
            switch (theActiveForm.TypeEx)
            {
                case FORMSALESQUOTE:
                    docEntry = getDocEntry("23", DocNum, Custid);
                    if(docEntry==0)
                        return getBranchIdByCustomer(Custid).ToString();
                    return getBranchIdFromQUT(docEntry.ToString()).ToString();
                case FORMCREDITMEMO:
                    docEntry = getDocEntry("14", DocNum, Custid);
                    if (docEntry == 0)
                        return getBranchIdByCustomer(Custid).ToString();
                    return getBranchIdFromCM(docEntry.ToString()).ToString();
                case FORMINVOICE:
                    docEntry = getDocEntry("13", DocNum, Custid);
                    if (docEntry == 0)
                        return getBranchIdByCustomer(Custid).ToString();

                    return getBranchIdFromINV(docEntry.ToString()).ToString();
                case FORMSALESORDER:
                    docEntry = getDocEntry("17", DocNum, Custid);
                    if (docEntry == 0)
                        return getBranchIdByCustomer(Custid).ToString();

                   return getBranchIdFromORDR(docEntry.ToString()).ToString();
                case FORMINCOMINGPAYMENT:
                    string bplid = getFormItemVal(theActiveForm, "1320002037").Trim();
                    if (bplid != "")
                    {
                        trace("Create Payment from incomming payment. Branch ID: " + bplid);
                        return bplid;
                    }
                    return "";
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return "";
    }
    private string getDocNum(SAPbouiCOM.Form form = null)
    {
        if (form == null)
            form = theActiveForm;
        if (form == null)
            return "";
        try
        {
            switch (form.TypeEx)
            {
                case FORMINCOMINGPAYMENT:
                    return "";
                default:
                    return getFormItemVal(form, fidRecordID);

            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return "";
    }
    private string getCustomerID(SAPbouiCOM.Form form = null)
    {
        if (form == null)
            form = theActiveForm;
        if (form == null)
            return "";
        if (form.TypeEx == FORMPAYMENTMEANS || form.TypeEx == formCredit)
            form = theActiveForm;
        try
        {
            switch (form.TypeEx)
            {
                case FORMBPARTNER:
                case FORMINCOMINGPAYMENT:
                    return getFormItemVal(form, "5");
                default:
                    return getFormItemVal(form, "4");

            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return "";
    }
    private void setLineItem(SAPbouiCOM.Form form, ref LineItem[] li)
    {
        try
        {
            if (theActiveForm.TypeEx == FORMINCOMINGPAYMENT || form == null)
            {
                li = new LineItem[1];
                li[0] = new LineItem();

                try
                {
                    li[0].CommodityCode = merchantData.ItemCommodityCode == null ? "12345678" : merchantData.ItemCommodityCode;
                    li[0].Qty = (merchantData.ItemQty == null ? 1 : merchantData.ItemQty).ToString();
                    li[0].UnitPrice = (merchantData.ItemUnitPrice == null ? 1 : merchantData.ItemUnitPrice).ToString();
                    //li[0].TaxableSpecified = true;
                    li[0].Taxable = (bool)(merchantData.ItemTaxable == null ? true : merchantData.ItemTaxable);
                    li[0].TaxRate = (merchantData.ItemTaxRate == null ? decimal.Parse(cfgDefaultTaxRate) : merchantData.ItemTaxRate).ToString();
                    li[0].ProductName = merchantData.ItemName == null ? "SAP B1 Hana Incoming Payment Batch Charge" : merchantData.ItemName;
                    li[0].DiscountRate = (merchantData.ItemDiscountRate == null ? 0 : merchantData.ItemDiscountRate).ToString();
                    li[0].UnitOfMeasure = merchantData.ItemUnitOfMeasure == null ? "EA" : merchantData.ItemUnitOfMeasure;
                    li[0].SKU = merchantData.ItemSKU == null ? "SKU1" : merchantData.ItemSKU;
                }
                catch (Exception)
                {
                    li[0].CommodityCode = "12345678";
                    li[0].Qty = "1";
                    li[0].UnitPrice = "1";
                    //li[0].TaxableSpecified = true;
                    li[0].Taxable = true;
                    li[0].TaxRate = cfgDefaultTaxRate;
                    li[0].ProductName = "SAP B1 Hana Incoming Payment Batch Charge";
                    li[0].DiscountRate = "0";
                    li[0].UnitOfMeasure = "EA";
                    li[0].SKU = "SKU1";
                }
                li[0].Description = li[0].ProductName;
            }
            else
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)theActiveForm.Items.Item(fidMatrix).Specific;
                if (theActiveForm.PaneLevel == 2)
                    oMatrix = (SAPbouiCOM.Matrix)theActiveForm.Items.Item("39").Specific;
                // errorLog(oMatrix.RowCount.ToString());
                if (oMatrix.RowCount >= 1)
                {
                    li = new LineItem[oMatrix.RowCount];
                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {

                        //string msg = "row=" + i.ToString() + "\r\n";
                        if (getMatrixItem(oMatrix, "1", i) != "")
                        {
                            li[i - 1] = new LineItem();
                           // li[i - 1].TaxableSpecified = true;
                            li[i - 1].Taxable = true;
                            li[i - 1].CommodityCode = merchantData.ItemCommodityCode == null ? getMatrixItem(oMatrix, "1", i) : merchantData.ItemCommodityCode;
                            li[i - 1].Qty = getMatrixItem(oMatrix, "11", i);
                            li[i - 1].UnitPrice = getMoneyString(getMatrixItem(oMatrix, "14", i));
                            string taxrate = getMatrixItem(oMatrix, "19", i);
                            if (taxrate == "")
                                taxrate = cfgDefaultTaxRate;
                            li[i - 1].TaxRate = taxrate; //;
                            li[i - 1].ProductName = getMatrixItem(oMatrix, "3", i);
                            li[i - 1].DiscountRate = getMatrixItem(oMatrix, "15", i);
                            li[i - 1].UnitOfMeasure = "EACH";
                            li[i - 1].SKU = getMatrixItem(oMatrix, "1", i);

                            li[i - 1].ProductRefNum = getMatrixItem(oMatrix, "1", i);
                            if (theActiveForm.PaneLevel == 2)
                            {
                                li[i - 1].CommodityCode = merchantData.ItemCommodityCode == null ? getMatrixItem(oMatrix, "94", i) : merchantData.ItemCommodityCode;
                                li[i - 1].Qty = "1";
                                li[i - 1].UnitPrice = getMoneyString(getMatrixItem(oMatrix, "12", i));
                                li[i - 1].TaxRate = "";
                                li[i - 1].ProductName = getMatrixItem(oMatrix, "1", i);
                                li[i - 1].DiscountRate = "";
                                li[i - 1].UnitOfMeasure = "EACH";
                                li[i - 1].SKU = "Service";
                            }
                            li[i - 1].Description = li[i - 1].ProductName;
                        }
                        //         errorLog("req.LineItems[i - 1].SKU = " + req.LineItems[i - 1].SKU);
                    }

                }
            }
        }
        catch (Exception)
        {

        }
    }
    
}

