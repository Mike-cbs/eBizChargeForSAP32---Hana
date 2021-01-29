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



partial class SAP  { 
    
   public bool PreAuth(SAPbouiCOM.Form form)
    {
        try
        {
            string strAmt = ""; 
            string cardNo = "";
            string exp = "";
            string code = "";
            if (!VerifyInput(form, ref strAmt, ref cardNo, ref exp, ref code))
                return false;
            ueSecurityToken token = getToken();
            TransactionRequestObject req = createRequest(form);
            req.Details.Amount = double.Parse(strAmt);
            req.CreditCardData.CardCode = code;
            req.CreditCardData.CardNumber = cardNo;
            req.CreditCardData.CardExpiration = exp;
            req.Command = "cc:authonly";
            string refNum = "";
            if (runRequest(form, req, ref refNum))
            {
               
                string remark = GetRemark(form);
                remark += string.Format(strHeaderApproved + "ref. No.:{0} Amount:{1} {2}**\r\n", refNum, strAmt, DateTime.Now.ToString("MM-dd-yyy hh:mm"));
                SetRemark(form, remark);
                SBO_Application.MessageBox("Approved. Reference Number:" + refNum + ", Amount:" + strAmt, 1, "Ok", "", "");

                ClearInput(form);
              
            }
          
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    public bool Capture(SAPbouiCOM.Form form)
    {
        try
        {
            string strRef = "";
            string amount = "";
            if (!GetRefNoAndAmountFromRemark(form, ref strRef, ref amount))
            {
                SBO_Application.MessageBox("Refrence number not found.", 1, "Ok", "", "");
                return false;
            }
            ueSecurityToken token = getToken();
            TransactionRequestObject req = createRequest(form);
            req.RefNum = strRef;
            req.Command = "cc:capture";
            string refNum = "";
            if (runRequest(form, req, ref refNum))
            {
                string remark = GetRemark(form);
                remark += string.Format(strHeaderCaptured + "ref. No.:{0} **\r\n",refNum,  DateTime.Now.ToString("MM-dd-yyy hh:mm"));
                SetRemark(form, remark);
                SBO_Application.MessageBox("Transaction Captured. Reference Number:" + refNum , 1, "Ok", "", "");
                //AddPayment(form, double.Parse(amount));

            }
            
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    public bool Charge(SAPbouiCOM.Form form)
    {
        try
        {
            string strAmt = ""; 
            string cardNo = "";
            string exp = "";
            string code = "";
            if (!VerifyInput(form, ref strAmt, ref cardNo, ref exp, ref code))
                return false;
            ueSecurityToken token = getToken();
            TransactionRequestObject req = createRequest(form);
            req.Details.Amount = double.Parse(strAmt);
            req.CreditCardData.CardCode = code;
            req.CreditCardData.CardNumber = cardNo;
            req.CreditCardData.CardExpiration = exp;
            req.Command = "cc:sale";
            string refNum = "";
            if (runRequest(form, req,ref refNum))
            {
                string remark = GetRemark(form);
                remark += string.Format(strHeaderCharged + "ref. No.:{0} Amount:{1} {2}**\r\n", refNum, strAmt, DateTime.Now.ToString("MM-dd-yyy hh:mm"));
                SetRemark(form, remark);
                SBO_Application.MessageBox("Charge Completed. Reference Number:" + refNum + ", Amount:" + strAmt, 1, "Ok", "", "");
                ClearInput(form);
                AddPayment(form, req.Details.Amount);

            }
            
          
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    public bool Credit(SAPbouiCOM.Form form)
    {
        try
        {
           
            string strAmt = "";
            string cardNo = "";
            string exp = "";
            string code = "";
            if (!VerifyInput(form, ref strAmt, ref cardNo, ref exp, ref code))
                return false;
            
            TransactionRequestObject req = createRequest(form);
            req.Details.Amount = double.Parse(strAmt);
            req.CreditCardData.CardCode = code;
            req.CreditCardData.CardNumber = cardNo;
            req.CreditCardData.CardExpiration = exp;
            req.Command = "cc:credit";
            string refNum = "";
            if (runRequest(form, req,ref refNum))
            {
                string remark = GetRemark(form);
                remark += string.Format(strHeaderCredited + "ref. No.:{0} Amount:{1} {2}**\r\n", refNum,strAmt, DateTime.Now.ToString("MM-dd-yyy hh:mm"));
                SetRemark(form, remark);
                SBO_Application.MessageBox("Transaction Completed. Reference Number:" + refNum + ", Amount:" + strAmt, 1, "Ok", "", "");
                ClearInput(form);
                AddOutgoingPayment(form, req.Details.Amount);
                
            }
          
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    private bool runRequest(SAPbouiCOM.Form form, TransactionRequestObject req, ref string refNum)
    {
        
        try
        {
           // form.Freeze(true);
            ueSecurityToken token = getToken();
            TransactionResponse resp = new TransactionResponse();
            resp = ebiz.runTransaction(token, req);
            //form.Freeze(false);
            if (resp.ErrorCode == "0")
            {
                AddCustomer(req);
                refNum = resp.RefNum;
                logTransaction(req, resp);
                
                return true;
            }
            else
            {
                errorLog(resp.Error);
                SBO_Application.MessageBox("Can not process your transaction.\r\n" + resp.Error, 1, "Ok", "", "");
                return false;
            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    private void logTransaction(TransactionRequestObject req, TransactionResponse resp)
    {
        string sql = "";
        try
        {
            string cardLastFour = "";
            try
            {
                if (req.CreditCardData.CardNumber == null)
                    cardLastFour = "";
                else
                    cardLastFour = req.CreditCardData.CardNumber.Substring(req.CreditCardData.CardNumber.Length - 4);
            }catch(Exception)
            {}
            
            db.CommandTimeout = 30000;
            CCTRAN cctran = new CCTRAN();
            cctran.customerID= req.CustomerID;
            cctran.customerName=req.BillingAddress.Company;
            cctran.crCardNum=cardLastFour;
            cctran.Description=req.Details.Description;
            cctran.recID=req.Details.Invoice;
            cctran.acsUrl=resp.AcsUrl;
            cctran.authAmount=resp.AuthAmount.ToString();
            cctran.authCode=resp.AuthCode;
            cctran.avsResult=resp.AvsResult;
            cctran.avsResultCode=resp.AvsResultCode;
            cctran.batchNum=resp.BatchNum;
            cctran.batchRef=resp.BatchRefNum;
            cctran.cardCodeResult=resp.CardCodeResult;
            cctran.cardCodeResultCode=resp.CardCodeResultCode;
            cctran.cardLevelResult=resp.CardLevelResult;
            cctran.cardLevelResultCode=resp.CardLevelResultCode;
            cctran.conversionRate=resp.ConversionRate.ToString();
            cctran.convertedAmount=resp.ConvertedAmount.ToString();
            cctran.convertedAmountCurrency=resp.ConvertedAmountCurrency.ToString();
            cctran.custNum=resp.CustNum;
            cctran.error=resp.Error;
            cctran.errorCode=resp.ErrorCode;
            cctran.isDuplicate=resp.isDuplicate.ToString();
            cctran.payload=resp.Payload;
            cctran.profilerScore=resp.ProfilerScore;
            cctran.profilerResponse=resp.ProfilerResponse;
            cctran.profilerReason= resp.ProfilerReason;
            cctran.refNum=resp.RefNum;
            cctran.remainingBalance=resp.RemainingBalance.ToString();
            cctran.result=resp.Result;
            cctran.resultCode=resp.ResultCode;
            cctran.status=resp.Status;
            cctran.statusCode=resp.StatusCode;
            cctran.vpasResultCode=resp.VpasResultCode;
            cctran.recDate=DateTime.Now;//Use local time not server time
            cctran.command=req.Command;
            cctran.amount = req.Details.Amount.ToString();
            db.CCTRANs.InsertOnSubmit(cctran);
            db.SubmitChanges();
           
        }catch(Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }

    }
    private  TransactionRequestObject createRequest(SAPbouiCOM.Form form)
    {
        SAPbouiCOM.EditText customer = form.Items.Item(fidCustID).Specific;
        TransactionRequestObject req = new TransactionRequestObject();
        try
        {
            req.RecurringBilling = new RecurringBilling();
           
            req.AccountHolder = "";
            req.AuthCode = "";
            req.ClientIP = getIP();
            req.CustomerID = customer.Value;
            req.CheckData = new CheckData();
            req.CheckData.Account = "";
            req.CheckData.AccountType = "";
            req.Details = new TransactionDetail();
            req.Details.Amount = getItemDoubleValue(form,fidBalance);
            req.Details.AmountSpecified = true;
            req.Details.Discount = getItemDoubleValue(form, fidDiscount);
            req.Details.DiscountSpecified = true;
            req.Details.Shipping = getItemDoubleValue(form, fidFreight);
            req.Details.ShippingSpecified = true;
            req.Details.Tax = getItemDoubleValue(form, fidTax);
            req.Details.TaxSpecified = true;
            req.Details.Invoice = getItemValue(form, fidRecordID);
            req.Details.OrderID = getItemValue(form, fidRecordID);
            req.Details.PONum = getItemValue(form, fidRecordID);
            req.Details.Description = getItemValue(form,fidDesc);
            req.Details.Subtotal = getItemDoubleValue(form, fidSubTotal);
            req.Details.SubtotalSpecified = true;
            req.BillingAddress = new Address();
            req.BillingAddress.Company = "";
            req.BillingAddress.FirstName = theSAPCustomer.BillingAddress.FirstName;
            req.BillingAddress.LastName =theSAPCustomer.BillingAddress.LastName;
            req.BillingAddress.Street = theSAPCustomer.BillingAddress.Street;
            req.BillingAddress.City = theSAPCustomer.BillingAddress.City;
            req.BillingAddress.State = theSAPCustomer.BillingAddress.State;
            req.BillingAddress.Zip =  theSAPCustomer.BillingAddress.Zip;
            req.CreditCardData = new CreditCardData();
            req.ShippingAddress = new Address();
            req.ShippingAddress.Company = "";
            req.ShippingAddress.FirstName = theSAPCustomer.BillingAddress.FirstName;
            req.ShippingAddress.LastName = theSAPCustomer.BillingAddress.LastName;
            req.ShippingAddress.Street = theSAPCustomer.BillingAddress.Street;
            req.ShippingAddress.City = theSAPCustomer.BillingAddress.City;
            req.ShippingAddress.State = theSAPCustomer.BillingAddress.State;
            req.ShippingAddress.Zip = theSAPCustomer.BillingAddress.Zip;
            SAPbouiCOM.Matrix oMatrix = form.Items.Item(fidMatrix).Specific;
            if (oMatrix.RowCount >= 2)
            {
                req.LineItems = new LineItem[oMatrix.RowCount];
                for (int i = 1; i < oMatrix.RowCount; i++)
                {
                    //string msg = "row=" + i.ToString() + "\r\n";
                    req.LineItems[i - 1] = new LineItem();
                    req.LineItems[i - 1].CommodityCode = getMatrixItem(oMatrix, GetConfig("SAPFIDColDiscountRate"), i);
                    req.LineItems[i - 1].Qty = getMatrixItem(oMatrix, GetConfig("SAPFIDColQty"), i);
                    req.LineItems[i - 1].UnitPrice = getMatrixItem(oMatrix, GetConfig("SAPFIDColUnitPrice"), i);
                    req.LineItems[i - 1].TaxRate = getMatrixItem(oMatrix, GetConfig("SAPFIDColTaxRate"), i);
                    req.LineItems[i - 1].ProductName = getMatrixItem(oMatrix, GetConfig("SAPFIDColProductName"), i);
                    req.LineItems[i - 1].DiscountRate = getMatrixItem(oMatrix, GetConfig("SAPFIDColDiscountRate"), i);
                    req.LineItems[i - 1].UnitOfMeasure = "EACH";
                    req.LineItems[i - 1].ProductRefNum = getMatrixItem(oMatrix, GetConfig("SAPFIDColProductRefNum"), i);
                    req.LineItems[i - 1].SKU = getMatrixItem(oMatrix, GetConfig("SAPFIDColProductRefNum"), i);
                    /*
                    for(int n=0; n < oMatrix.Columns.Count; n++)
                    {
                        string v = getMatrixItem(oMatrix, n.ToString(), i);
                        msg = msg + n.ToString() + ":" + v + "\r\n";
                    }
                    errorLog(msg);
                    */ 
                }
                
            }
           


        }catch(Exception ex)
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
    private ueSecurityToken getToken()
    {
        ueSecurityToken token = new ueSecurityToken();
        try
        {
            token.ClientIP = getIP();
            token.SourceKey = sourceKey;
            token.PinHash = new ueHash();
            token.PinHash.Seed = DateTime.Now.ToString("yyyyMMddhhmmssssss");
            string prehashvalue = string.Concat(token.SourceKey, token.PinHash.Seed, pin);  // combine data into single string

            token.PinHash.Type = "md5";
            token.PinHash.HashValue = GenerateHash(prehashvalue);
        }catch(Exception ex)
        { errorLog(ex); }
        return token;    
    }
    private static string GenerateHash(string input)
    {
        // Create a new instance of the MD5CryptoServiceProvider object.
        MD5 md5Hasher = MD5.Create();

        // Convert the input string to a byte array and compute the hash.
        byte[] data = md5Hasher.ComputeHash(Encoding.Default.GetBytes(input));

        // Create a new Stringbuilder to collect the bytes
        // and create a string.
        StringBuilder sBuilder = new StringBuilder();

        // Loop through each byte of the hashed data 
        // and format each one as a hexadecimal string.
        for (int i = 0; i < data.Length; i++)
        {
            sBuilder.Append(data[i].ToString("x2"));
        }

        // Return the hexadecimal string.
        return sBuilder.ToString();
    }
   
    private string getIP()
    {
        try
        {
            IPAddress[] ips;
            string hostname = Dns.GetHostName();
            ips = Dns.GetHostAddresses(hostname);
            foreach (IPAddress ip in ips)
            {
                if(ip.AddressFamily== System.Net.Sockets.AddressFamily.InterNetwork)
                    return ip.ToString();
            }
        }catch(Exception ex)
        {
            errorLog(ex);

        }
        return "";
    }
    public void handleSwipe(SAPbouiCOM.Form form)
    {
        try
        {
            string s = getFormEditVal(form, editCardNum);
            string track = s.Replace(";", "");
            string[] sTrack = track.Split('?');
            if (sTrack.Length > 1)
            {
                //get name
                string fname = "";
                string lname = "";
                if (sTrack[0].IndexOf('^') > 0)
                {
                    string[] n = sTrack[0].Split('^');
                    if (n.Length > 2)
                    {
                        string name = n[1];
                        string[] l = name.Split('/');
                        if (l.Length > 1)
                        {
                            fname = l[1].Trim();
                            lname = l[0].Trim();
                        }
                    }
                }
                if (fname != "" && lname != "")
                {
                    if (fname != theSAPCustomer.BillingAddress.FirstName || lname != theSAPCustomer.BillingAddress.LastName)
                    {
                        theSAPCustomer.BillingAddress.FirstName = fname;
                        theSAPCustomer.BillingAddress.LastName = lname;
                    }
                }
                //get car num
                if (sTrack[1].IndexOf("=") > 0)
                {
                    string[] f = sTrack[1].Split('=');
                    string CardNumber = f[0].Trim();
                    string CardExpiration = f[1].Substring(2, 2) + f[1].Substring(0, 2);
                    setFormEditVal(form, editCardNum, CardNumber);
                    setFormEditVal(form, editCardExpDate, CardExpiration);
                }

            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
    public void AddPayment(SAPbouiCOM.Form form, double total)
    {
        try
        {
            SAPbobsCOM.Documents oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            int sNum = int.Parse(getItemValue(form, fidRecordID));
            oDoc.GetByKey(sNum);
            if (!oDoc.GetByKey(sNum))
            {
                SBO_Application.MessageBox(oCompany.GetLastErrorDescription());
                return;
            }
            SAPbobsCOM.Payments oPmt = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            oPmt.Invoices.DocEntry = sNum;
            oPmt.CardCode = oDoc.CardCode;
            oPmt.DocDate = DateTime.Now;
            oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
            string acct = GetConfig("SAPAccountForIncomingPayment");
            if (acct != "")
                oPmt.CashAccount = acct;
            oPmt.CashSum = total;
            oPmt.JournalRemarks = GetRemark(form);
            int r = oPmt.Add();
            if (r != 0)
                SBO_Application.MessageBox(oCompany.GetLastErrorDescription());
          
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    
    public bool createInvoice(SAPbouiCOM.Form form, bool bUseActiveForm = false)
    {
        try
        {
            SAPbobsCOM.Documents oInvDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            string s;
            int sNum;
            s = getFormEditVal(form, "8");
            sNum = int.Parse(s);
            if (oInvDoc.GetByKey(sNum))
                return true;
            oInvDoc.DocNum = sNum;
            if(bUseActiveForm)
            {
                int key = int.Parse(getItemValue(SBO_Application.Forms.ActiveForm, fidRecordID));
                oInvDoc.DocNum = key;
            }
            s = getFormEditVal(form, "4");
            oInvDoc.CardCode = s;
            s = getFormEditVal(form, "54");
            oInvDoc.CardName = s;

            oInvDoc.DocDueDate = DateTime.Now.AddDays(5);


            oInvDoc.DocDate = DateTime.Now;

            //oInvDoc.JournalMemo = GetRemark(form);
            try
            {
                s = getComboBoxVal(form, "20");
                oInvDoc.SalesPersonCode = int.Parse(s);
            }
            catch (Exception)
            { }
            SAPbouiCOM.Matrix oMatrix = form.Items.Item(fidMatrix).Specific;
            if (oMatrix.RowCount >= 2)
            {/*
                string str = "";
                for (int c = 1; c < oMatrix.Columns.Count; c++)
                {
                    str = str + "col=" + c + ",val=" + getMatrixItem(oMatrix, c.ToString(), 1) + "\r\n";
                }
                errorLog(str);
              */
                for (int i = 1; i < oMatrix.RowCount; i++)
                {

                    s = getMatrixItem(oMatrix, "1", i);
                    oInvDoc.Lines.ItemCode = s;

                    s = getMatrixItem(oMatrix, "3", i);
                    oInvDoc.Lines.ItemDescription = s;

                    s = getMatrixItem(oMatrix, "11", i);
                    int q = (int)double.Parse(s);
                    oInvDoc.Lines.Quantity = q;

                    s = getMatrixItem(oMatrix, "14", i).Replace("$", "").Replace(",", "").Trim();
                    oInvDoc.Lines.Price = double.Parse(s);

                    s = getMatrixItem(oMatrix, "160", i);
                    oInvDoc.Lines.TaxCode = s;

                    s = getMatrixItem(oMatrix, "21", i).Replace("$", "").Replace(",", "").Trim();
                    oInvDoc.Lines.LineTotal = double.Parse(s);

                    s = getMatrixItem(oMatrix, "15", i);
                    oInvDoc.Lines.DiscountPercent = double.Parse(s);
                    oInvDoc.Lines.Add();
                }

            }

            int r = oInvDoc.Add();
            if (r != 0)
            {
                errorLog(oCompany.GetLastErrorDescription());
                bAuto = false;
                SBO_Application.MessageBox(oCompany.GetLastErrorDescription());
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    /*
    public bool createInvoice(SAPbobsCOM.Documents oSODoc)
    {
        try
        {
           SAPbobsCOM.Documents oInvDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
           int key = int.Parse(getItemValue(SBO_Application.Forms.ActiveForm, fidRecordID));
            oInvDoc.DocNum = key;

            oInvDoc.CardCode = oSODoc.CardCode;
            oInvDoc.CardName = oSODoc.CardName;

            oInvDoc.DocDueDate = DateTime.Now.AddDays(5);


            oInvDoc.DocDate = DateTime.Now;

            try
            {
                oInvDoc.SalesPersonCode = oSODoc.SalesPersonCode;
            }catch(Exception)
            { }
            for (int i = 0; i < oSODoc.Lines.Count; i++)
            {
                try
                {
                    oInvDoc.Lines.BaseEntry = oSODoc.DocNum;
                    oInvDoc.Lines.BaseLine = i;
                    oInvDoc.Lines.BaseType = (int)SAPbobsCOM.BoObjectTypes.oOrders;

                    oDoc.Lines.ItemCode = s;

                    s = getMatrixItem(oMatrix, "3", i);
                    oDoc.Lines.ItemDescription = s;

                    s = getMatrixItem(oMatrix, "11", i);
                    int q = (int)double.Parse(s);
                    oDoc.Lines.Quantity = q;

                    s = getMatrixItem(oMatrix, "14", i).Replace("$", "").Replace(",", "").Trim();
                    oDoc.Lines.Price = double.Parse(s);

                    s = getMatrixItem(oMatrix, "160", i);
                    oDoc.Lines.TaxCode = s;

                    s = getMatrixItem(oMatrix, "21", i).Replace("$", "").Replace(",", "").Trim();
                    oDoc.Lines.LineTotal = double.Parse(s);

                    s = getMatrixItem(oMatrix, "15", i);
                    oDoc.Lines.DiscountPercent = double.Parse(s);
                    oInvDoc.Lines.Add();
                }catch(Exception ex)
                {
                    errorLog(ex);
                }

            }
                      
            int r = oInvDoc.Add();
            if (r != 0)
            {
                errorLog(oCompany.GetLastErrorDescription());
                SBO_Application.MessageBox(oCompany.GetLastErrorDescription());
                return false;
            }
          
            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
     */
    public bool createSO(SAPbouiCOM.Form form, ref SAPbobsCOM.Documents oDoc)
    {
        try
        {

           
                   
            string s = getFormEditVal(form, "8");
            int sNum = int.Parse(s);
            if (oDoc.GetByKey(sNum))
                return true;

            oDoc.DocNum = int.Parse(s);

            s = getFormEditVal(form, "4");
            oDoc.CardCode = s;
            s = getFormEditVal(form, "54");
            oDoc.CardName = s;
           
            /*
            s = getFormEditVal(form, "46");
            if(s.Length==8)
                s = s.Substring(4, 2) + "/" + s.Substring(6, 2) + "/" + s.Substring(0, 4);
            oDoc.DocDate = DateTime.Parse(s);
            oDoc.DocDueDate = oDoc.DocDate;
            */
            oDoc.DocDueDate = DateTime.Now.AddDays(5);
            oDoc.DocDate = DateTime.Now;
           // oDoc.JournalMemo = GetRemark(form);
            try
            {
                s = getComboBoxVal(form, "20");
                oDoc.SalesPersonCode = int.Parse(s);
            }
            catch (Exception)
            { }
            SAPbouiCOM.Matrix oMatrix = form.Items.Item(fidMatrix).Specific;
            if (oMatrix.RowCount >= 2)
            {/*
                string str = "";
                for (int c = 1; c < oMatrix.Columns.Count; c++)
                {
                    str = str + "col=" + c + ",val=" + getMatrixItem(oMatrix, c.ToString(), 1) + "\r\n";
                }
                errorLog(str);
              */
                for (int i = 1; i < oMatrix.RowCount; i++)
                {
                    s = getMatrixItem(oMatrix, "1", i);
                    oDoc.Lines.ItemCode = s;

                    s = getMatrixItem(oMatrix, "3", i);
                    oDoc.Lines.ItemDescription = s;

                    s = getMatrixItem(oMatrix, "11", i);
                    int q = (int)double.Parse(s);
                    oDoc.Lines.Quantity = q;

                    s = getMatrixItem(oMatrix, "14", i).Replace("$", "").Replace(",", "").Trim();
                    oDoc.Lines.Price = double.Parse(s);

                    s = getMatrixItem(oMatrix, "160", i);
                    oDoc.Lines.TaxCode = s;

                    s = getMatrixItem(oMatrix, "21", i).Replace("$", "").Replace(",", "").Trim();
                    oDoc.Lines.LineTotal = double.Parse(s);

                    s = getMatrixItem(oMatrix, "15", i);
                    oDoc.Lines.DiscountPercent = double.Parse(s);
                    oDoc.Lines.Add();
                }

            }

            int r = oDoc.Add();
            if (r != 0)
            {
                errorLog(oCompany.GetLastErrorDescription());
                SBO_Application.MessageBox(oCompany.GetLastErrorDescription());
                return false;
            }
           
            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    public bool createCreditMemo(SAPbouiCOM.Form form)
    {
        try
        {

            SAPbobsCOM.Documents oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);

            string s = getFormEditVal(form, "8");
            int sNum = int.Parse(s);
            if (oDoc.GetByKey(sNum))
                return true;

            oDoc.DocNum = int.Parse(s);

            s = getFormEditVal(form, "4");
            oDoc.CardCode = s;
            s = getFormEditVal(form, "54");
            oDoc.CardName = s;
            /*
            s = getFormEditVal(form, "46");
            s = s.Substring(4, 2) + "/" + s.Substring(6, 2) + "/" + s.Substring(0, 4);
            oDoc.DocDate = DateTime.Parse(s);
            oDoc.DocDueDate = oDoc.DocDate;
            */
            oDoc.DocDueDate = DateTime.Now.AddDays(5);
            oDoc.DocDate = DateTime.Now;
           */

            // oDoc.JournalMemo = GetRemark(form);
            try
            {
                s = getComboBoxVal(form, "20");
                oDoc.SalesPersonCode = int.Parse(s);
            }
            catch (Exception)
            { }
            SAPbouiCOM.Matrix oMatrix = form.Items.Item(fidMatrix).Specific;
            if (oMatrix.RowCount >= 2)
            {/*
                string str = "";
                for (int c = 1; c < oMatrix.Columns.Count; c++)
                {
                    str = str + "col=" + c + ",val=" + getMatrixItem(oMatrix, c.ToString(), 1) + "\r\n";
                }
                errorLog(str);
              */
                for (int i = 1; i < oMatrix.RowCount; i++)
                {
                    s = getMatrixItem(oMatrix, "1", i);
                    oDoc.Lines.ItemCode = s;

                    s = getMatrixItem(oMatrix, "3", i);
                    oDoc.Lines.ItemDescription = s;

                    s = getMatrixItem(oMatrix, "11", i);
                    int q = (int)double.Parse(s);
                    oDoc.Lines.Quantity = q;

                    s = getMatrixItem(oMatrix, "14", i).Replace("$", "").Replace(",", "").Trim();
                    oDoc.Lines.Price = double.Parse(s);

                    s = getMatrixItem(oMatrix, "160", i);
                    oDoc.Lines.TaxCode = s;

                    s = getMatrixItem(oMatrix, "21", i).Replace("$", "").Replace(",", "").Trim();
                    oDoc.Lines.LineTotal = double.Parse(s);

                    s = getMatrixItem(oMatrix, "15", i);
                    oDoc.Lines.DiscountPercent = double.Parse(s);
                    oDoc.Lines.Add();
                }

            }

            int r = oDoc.Add();
            if (r != 0)
            {
                errorLog(oCompany.GetLastErrorDescription());
                SBO_Application.MessageBox(oCompany.GetLastErrorDescription());
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    public void AddOutgoingPayment(SAPbouiCOM.Form form, double total)
    {
        try
        {
            SAPbobsCOM.Documents oDoc = null;
            oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
            int sNum = int.Parse(getItemValue(form, fidRecordID));
            if (!oDoc.GetByKey(sNum))
            {
                SBO_Application.MessageBox(oCompany.GetLastErrorDescription());
                return;
            }
            SAPbobsCOM.Payments oPmt = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            oPmt.Invoices.DocEntry = sNum;
            oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote;
            oPmt.CashAccount = GetConfig("SAPAccountForOutgoingPayment");
            oPmt.CardCode = oDoc.CardCode;
            oPmt.DocDate = DateTime.Now;
            oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
            oPmt.CashSum = total;
            oPmt.JournalRemarks = GetRemark(form);
            int r = oPmt.Add();
            if (r != 0)
                SBO_Application.MessageBox(oCompany.GetLastErrorDescription());
            else
                SBO_Application.MessageBox("Outgoing payment record added");
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
} 

