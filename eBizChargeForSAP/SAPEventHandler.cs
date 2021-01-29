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
using Outlook = Microsoft.Office.Interop.Outlook;
using eBizChargeForSAP.ServiceReference1;
using System.Reflection;
using System.Xml;

partial class SAP
{

    private void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
    {
        BubbleEvent = true;
        try
        {


        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    const string menuIDLastRecord = "1291";
    const string menuCAcctValidate = "eBizCardAccountValidateMenu";

    bool bValidated = false;
    bool bUsePaymentMean = false;
    bool beBizChargeClicked = false;
    private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;
        try
        {

            if (pVal.BeforeAction == true)
            {

                switch (pVal.MenuUID)
                {
                    case formExportBP:
                        CreateExportBp();
                        break;

                    case formExportItemMaster:
                        CreateExportItem();
                        break;
                    case formExportSO:
                        CreateExportSO();
                        break;
                    case formEmailPaymentSO:
                        CreateMenuCreatePaymentForSO();
                        break;
                    case formSyncUpload:
                        CreateSyncUpload();
                        break;
                    case formEmailEstimates:
                        CreateSendEmailSO();
                        break;
                    case formEmailPaymentPending:
                        CreatePendEmail();
                        break;
                    case formEmailPaymentReceived:
                        CreateRecPay();
                        break;
                    case formSyncDownload:
                        CreateSyncSO();
                        break;
                        /*
                    case menuCAcctValidate:
                        if (SBO_Application.MessageBox("Validate credit card account?", 1, "Yes", "No") == 1)
                        {
                            tProcess = new Thread(ValidateCardAccount);
                            tProcess.Start();
                            return;
                        }
                        break;
                    case menuTerminal:
                        TermAmount = "";
                        TermFunc = "";
                        TermGroup = "";
                        CreateTerminalForm();
                        break;
                    case menuCustImp:
                        CreateCustImpForm();
                        break;
                    case menuOrderImp:
                        CreateOrderImpForm();
                        break;
                    case menueBizConnect:
                        CreateeBizConnectForm();
                        break;
                    case menuePaymentForm:
                        CreateePaymentFormForm();
                        break;

                    case menuCfg:
                        CreateCfgForm();
                        break;
                    case menuBatchInv:
                        CreateBatchInvForm();
                        break;
                    case menuRBilling:
                        CreateRBillingForm();
                        break;
                    case "5892":  //Payment means

                        try
                        {
                            try
                            {
                                SAPbouiCOM.Form act = SBO_Application.Forms.ActiveForm;
                                if (act != null)
                                    theActiveForm = act;
                            }
                            catch (Exception ex)
                            {
                                errorLog(ex);
                            }
                            if (beBizChargeClicked)
                            {
                                beBizChargeClicked = false;
                                bUsePaymentMean = false;
                            }
                            else
                            {
                                if (theActiveForm.TypeEx == FORMINCOMINGPAYMENT && getFormItemVal(theActiveForm, "1") != "Add")
                                {
                                    bUsePaymentMean = true;
                                }
                                if (theActiveForm.TypeEx != FORMINCOMINGPAYMENT)
                                {
                                    bUsePaymentMean = true;
                                }
                                else
                                {
                                    bUsePaymentMean = false;
                                }
                            }
                        }
                        catch (Exception)
                        { }

                        break;
                    case "1284":  //Cancel
                        try
                        {
                            try
                            {
                                SAPbouiCOM.Form act = SBO_Application.Forms.ActiveForm;
                                if (act != null)
                                {
                                    if (act.TypeEx != FORMINCOMINGPAYMENT && act.TypeEx != FORMSALESORDER && act.TypeEx != FORMSALESQUOTE)
                                        return;
                                }
                            }
                            catch (Exception ex)
                            {
                                errorLog("menu cancel excception: " + ex.Message);
                            }
                            if (theActiveForm == null)
                                return;
                            if (theActiveForm.TypeEx != FORMINCOMINGPAYMENT && theActiveForm.TypeEx != FORMSALESORDER && theActiveForm.TypeEx != FORMSALESQUOTE)
                                return;
                            if (theActiveForm.TypeEx == FORMSALESORDER || theActiveForm.TypeEx == FORMSALESQUOTE)
                                formPayment = null;
                            string refNum = "";
                            string invoicenum = "";
                            double amount = 0;
                            double creditsum = 0;
                            if (formPayment != null)
                            {
                                string docID = getFormItemVal(formPayment, "3");
                                if (formPayment.TypeEx == "170")
                                {
                                    bAuto = false;
                                    /*
                                    if (SBO_Application.MessageBox("Do you want to force cancel this incoming payment?", 1, "Yes", "No") == 1)
                                    {
                                        paymentForceCancel(docID);
                                        if (SBO_Application.MessageBox("Payment document cancelled. Do you want to continue?", 1, "Yes", "No") != 1)
                                        {
                                            bAuto = false;
                                            BubbleEvent = false;
                                            return;
                                        }
                                    }
                                    */
                        /*
                        SAPbobsCOM.Recordset oCustomerRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oCustomerRS.DoQuery("select \"ConfNum\", \"CreditSum\" from \"RCT3\" where \"DocNum\"=" + docID);
                        if (!oCustomerRS.EoF)
                        {
                            refNum = GetFieldVal(oCustomerRS, 0);
                            creditsum = (double)oCustomerRS.Fields.Item(1).Value;
                            oCustomerRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string sql = string.Format("select \"U_refNum\",\"U_InvoiceID\", \"U_amount\", \"U_DownPaymentInvID\" from \"@CCTRANS\" where \"U_refNum\" = '{0}'", refNum);

                            oCustomerRS.DoQuery(sql);
                            if (!oCustomerRS.EoF)
                            {
                                string invID = GetFieldVal(oCustomerRS, 1);
                                amount = getDoubleValue(GetFieldVal(oCustomerRS, 2));

                                string DownPaymentInvoiceID = GetFieldVal(oCustomerRS, 3);
                                if (DownPaymentInvoiceID != "" && DownPaymentInvoiceID != null && invID != "" && invID != null)
                                {
                                    invoicenum = getDocNum("OINV", invID).ToString();
                                    if (SBO_Application.MessageBox("Cannot cancel a document with down payment that was drawn to an invoice. Do you want to continue?", 1, "Yes", "No") != 1)
                                    {
                                        bAuto = false;
                                        BubbleEvent = false;
                                        return;
                                    }
                                }
                            }

                        }
                    }
                    if (formPayment.TypeEx == "426")  //Outgoing
                    {
                        SAPbobsCOM.Recordset oCustomerRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oCustomerRS.DoQuery("select \"ref2\" from OVPM where \"DocNum\"=" + docID);
                        if (!oCustomerRS.EoF)
                            refNum = GetFieldVal(oCustomerRS, 0);

                    }
                }
                else
                {
                    try
                    {
                        SAPbouiCOM.Form form = SBO_Application.Forms.ActiveForm;
                        if (form != null)
                        {
                            if (form.TypeEx == FORMSALESORDER)  //Sales Order
                            {

                                SODocEntry = form.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);
                                SAPbobsCOM.Recordset oCustomerRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                oCustomerRS.DoQuery(string.Format("select \"U_refNum\" from \"@CCTRANS\" where \"U_OrderID\" = '{0}' and \"U_command\"='cc:authonly' and not exists(select 1 from \"@CCTRANS\" where \"U_OrderID\"='{0}' and \"U_command\"='cc:void')", SODocEntry));
                                if (!oCustomerRS.EoF)
                                {
                                    refNum = GetFieldVal(oCustomerRS, 0);
                                }
                            }
                            if (form.TypeEx == FORMSALESQUOTE)  //Sales quote
                            {

                                string qid = form.DataSources.DBDataSources.Item("OQUT").GetValue("DocEntry", 0);
                                SAPbobsCOM.Recordset oCustomerRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                oCustomerRS.DoQuery(string.Format("select \"U_refNum\" from \"@CCTRANS\" where \"U_QuoteID\" = '{0}' and \"U_command\"='cc:authonly' and not exists(select 1 from \"@CCTRANS\" where \"U_QuoteID\"='{0}' and \"U_command\"='cc:void')", qid));
                                if (!oCustomerRS.EoF)
                                    refNum = GetFieldVal(oCustomerRS, 0);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        errorLog("Menu Sales order Cancel exception. " + ex.Message);
                    }
                }

                if (refNum != "")
                {
                    bAuto = false;
                    if (SBO_Application.MessageBox("Void transaction " + refNum + "?", 1, "Yes", "No") == 1)
                    {
                        if (!voidCustomer(refNum, false))
                        {
                            BubbleEvent = true;
                            bAuto = false;
                        }
                        else
                        {
                            BubbleEvent = true;
                            bAuto = true;
                            bVoidCloseForm = true;
                            string subject = string.Format("Transaction Voided. Invoice: {0} amount: {1}", invoicenum, amount);
                            string body = string.Format("Invoice: {0}\r\n\r\nAmount: {1}\r\n\r\nTransaction ID: {2}", invoicenum, amount, refNum);
                            bool bRes = sendEmailViaOutlook(userEmail,
       cfgmerchantEmail, null,
       subject,
       body,
       null,
       null);

                        }
                    }
                    else
                    {
                        BubbleEvent = true;
                        bAuto = false;
                    };
                }
            }
            catch (Exception ex)
            {
                errorLog("Menu Cancel exception. " + ex.Message);
            }
            break;

    }
}
else
{
    switch (pVal.MenuUID)
    {
        case "1288":
        case "1289":
        case "1290":
        case menuIDLastRecord:
            try
            {
                if (theActiveForm != null)
                {
                    if (theActiveForm.TypeEx == FORMBPARTNER)
                    {
                        if (theActiveForm.PaneLevel == panePM)
                            theActiveForm.Items.Item(tbeBizCharge).Click();

                    }
                    else if (theActiveForm.TypeEx == FORMCREDITMEMO
                             || theActiveForm.TypeEx == FORMDOWNPAYMENT
                             || theActiveForm.TypeEx == FORMINVOICE
                             || theActiveForm.TypeEx == FORMSALESQUOTE
                             || theActiveForm.TypeEx == FORMSALESORDER)
                    {
                        if (theActiveForm.PaneLevel == paneCCLog)
                            theActiveForm.Items.Item(tbCreditCard).Click();

                    }
                }
            }
            catch (Exception) { }
            break;
        case "1284":  //cancel
            {

                if (bVoidCloseForm)
                {
                    bVoidCloseForm = false;
                    bAuto = true;
                    tProcess = new Thread(VoidCloseFormProc);
                    tProcess.Start();
                }

            }
            break;
            */
                }
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    bool bAuto = false;
    private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;
        try
        {
            SAPbouiCOM.Form form;
            if (pVal.FormTypeEx == "0" && !pVal.Before_Action && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            {
                if (bAuto)
                {
                    bAuto = false;
                    form = SBO_Application.Forms.Item(FormUID);
                    form.Items.Item("1").Click();
                }
            }

            if (pVal.FormType.Equals(FORMBPARTNER) && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD && pVal.Before_Action == true)
            {
                form = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                {
                    if (pVal.FormType.Equals(FORMBPARTNER))
                    {
                        UpdateXMLForm("Presentation_Layer.BP.xml", pVal.FormUID);
                        BusinessPartnerFormHandler2(form, pVal);
                    }
                }
            }

            
            #region TEK

            switch (pVal.FormTypeEx)
            {
                case formExportBP:
                    form = SBO_Application.Forms.Item(FormUID);
                    ExportBPFormHandler(form, pVal);
                    break;
                case formExportItemMaster:
                    form = SBO_Application.Forms.Item(FormUID);
                    ExportItemFormHandler(form, pVal);
                    break;
                case formSyncUpload:
                    form = SBO_Application.Forms.Item(FormUID);
                    SyncUploadFormHandler(form, pVal);
                    break;
                case formEmailPaymentSO:
                    form = SBO_Application.Forms.Item(FormUID);
                    SyncEmailPaymentSOFormHandler(form, pVal);
                    break;
                case formExportSO:
                    form = SBO_Application.Forms.Item(FormUID);
                    ExportSOFormHandler(form, pVal);
                    break;
                case formEmailEstimates:
                    form = SBO_Application.Forms.Item(FormUID);
                    SendEmailSOFormHandler(form, pVal);
                    break;
                case formEmailPaymentPending:
                    form = SBO_Application.Forms.Item(FormUID);
                    PendEmailFormHandler(form, pVal);
                    break;
                case formEmailPaymentReceived:
                    form = SBO_Application.Forms.Item(FormUID);
                    RecPayFormHandler(form, pVal);
                    break;
                case formSyncDownload:
                    form = SBO_Application.Forms.Item(FormUID);
                    SyncDownloadFormHandler(form, pVal);
                    break;
                case FORMINVOICE: //Invoice
                    form = SBO_Application.Forms.Item(FormUID);
                    theActiveForm = form;
                    InvoiceFormHandler(form, pVal, out BubbleEvent);
                    break;
                    /*
                case formTerminal:
                    form = SBO_Application.Forms.Item(FormUID);
                    TerminalFormHandler(form, pVal);
                    break;
                case FORMBPARTNER:  //Business Partner
                    form = SBO_Application.Forms.Item(FormUID);
                    BusinessPartnerFormHandler(form, pVal);
                    //BusinessPartnerFormHandler21(form, pVal);
                    break;
                case "146": //Deposit On Order & Paymenht Means
                    form = SBO_Application.Forms.Item(FormUID);
                    if (!bUsePaymentMean)
                        PaymentMeansFormHandler(form, pVal);
                    else
                        PaymentMeansImportHandler(form, pVal);

                    break;
                case "424":  //Credit Card Form
                    form = SBO_Application.Forms.Item(FormUID);
                    CreditCardFormHandler(form, pVal);
                    break;
                case "953":
                    form = SBO_Application.Forms.Item(FormUID);
                    DocWizardFormHandler(form, pVal);
                    break;
                case FORMSALESORDER: //Sales Order
                    form = SBO_Application.Forms.Item(FormUID);
                    theActiveForm = form;
                    SalesOrderFormHandler(form, pVal, out BubbleEvent);
                    break;
                case FORMSALESQUOTE: //Sales Order
                    form = SBO_Application.Forms.Item(FormUID);
                    theActiveForm = form;
                    SalesQuoteFormHandler(form, pVal, out BubbleEvent);
                    break;
                case FORMINVOICE: //Invoice
                    form = SBO_Application.Forms.Item(FormUID);
                    theActiveForm = form;
                    InvoiceFormHandler(form, pVal, out BubbleEvent);
                    break;
                case "179": //Credit Memo
                    form = SBO_Application.Forms.Item(FormUID);
                    theActiveForm = form;
                    CreditMemoFormHandler(form, pVal);
                    break;
                case "170":
                    //FORMTYPE.INCOMING_PAYMENT;
                    formPayment = SBO_Application.Forms.Item(FormUID);
                    theActiveForm = formPayment;
                    break;
                case "426":
                    //FORMTYPE.OUTGOING_PAYMENT;
                    formPayment = SBO_Application.Forms.Item(FormUID);
                    break;
                case "60090":  //Invoice + Payment
                    form = SBO_Application.Forms.Item(FormUID);
                    theActiveForm = form;
                    InvoiceFormHandler(form, pVal, out BubbleEvent);
                    break;
                case "60091":  //A/R Reserve Invoice
                    form = SBO_Application.Forms.Item(FormUID);
                    theActiveForm = form;
                    InvoiceFormHandler(form, pVal, out BubbleEvent);
                    break;
                case FORMDOWNPAYMENT: //Down Payment Invoice
                    form = SBO_Application.Forms.Item(FormUID);
                    theActiveForm = form;
                    InvoiceFormHandler(form, pVal, out BubbleEvent);
                    break;

                case formCredit:
                    form = SBO_Application.Forms.Item(FormUID);
                    CreditFormHandler(form, pVal, out BubbleEvent);
                    break;

                //case formSOImport:
                //    form = SBO_Application.Forms.Item(FormUID);
                //    SOImportFormHandler(pVal);
                //    break;

                case formBatchInv:
                    form = SBO_Application.Forms.Item(FormUID);
                    BatchInvFormHandler(form, pVal);
                    break;
                case formCustImp:
                    form = SBO_Application.Forms.Item(FormUID);
                    CustImpFormHandler(form, pVal);
                    break;
                case formOrderImp:
                    form = SBO_Application.Forms.Item(FormUID);
                    OrderImpFormHandler(form, pVal);
                    break;
                case formRBilling:
                    form = SBO_Application.Forms.Item(FormUID);
                    RBillingFormHandler(form, pVal);
                    break;
                case formCfg:
                    form = SBO_Application.Forms.Item(FormUID);
                    CFGFormHandler(form, pVal, out BubbleEvent);
                    break;
                case formJob:
                    form = SBO_Application.Forms.Item(FormUID);
                    JobFormHandler(form, pVal, out BubbleEvent);
                    break;
                case formeBizConnect:
                    form = SBO_Application.Forms.Item(FormUID);
                    eBizConnectFormHandler(form, pVal);
                    break;
                case formePaymentForm:
                    form = SBO_Application.Forms.Item(FormUID);
                    ePaymentFormFormHandler(form, pVal);
                    break;
                case FORMPICKMANAGER:
                    form = SBO_Application.Forms.Item(FormUID);
                    PickManagerFormHandler(form, pVal);
                    break;
                case "60508":
                    formPaymentList = SBO_Application.Forms.Item(FormUID);
                    break;
                    */
            }

            #endregion

            
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void CreditCardFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {

        try
        {
            if (!pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        form.Title = "eBizCharge Account Setup";
                        break;

                }
            }
            else
            {
                switch (pVal.EventType)
                {

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1")
                        {
                            createAcctList();
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
    public bool sendEmailViaOutlook(string sFromAddress, string sToAddress, string sCc, string sSubject, string sBody, List<string> arrAttachments = null, string sBcc = null)
    {
        //Send email via Office Outlook 2010
        //'sFromAddress' = email address sending from (ex: "me@somewhere.com") -- this account must exist in Outlook. Only one email address is allowed!
        //'sToAddress' = email address sending to. Can be multiple. In that case separate with semicolons or commas. (ex: "recipient@gmail.com", or "recipient1@gmail.com; recipient2@gmail.com")
        //'sCc' = email address sending to as Carbon Copy option. Can be multiple. In that case separate with semicolons or commas. (ex: "recipient@gmail.com", or "recipient1@gmail.com; recipient2@gmail.com")
        //'sSubject' = email subject as plain text
        //'sBody' = email body. Type of data depends on 'bodyType'
        //'bodyType' = type of text in 'sBody': plain text, HTML or RTF
        //'arrAttachments' = if not null, must be a list of absolute file paths to attach to the email
        //'sBcc' = single email address to use as a Blind Carbon Copy, or null not to use
        //RETURN:
        //      = true if success
        if (sFromAddress == "" || sToAddress == "" || cfgUseUserEmail != "Y")
        {
            trace(string.Format("sFromAddress={0}, sToAddress={1}, Void Notify={2}", sFromAddress, sToAddress, cfgUseUserEmail));
            return false;
        }
        bool bRes = false;

        try
        {
            //Get Outlook COM objects
            Outlook.Application app = new Outlook.Application();
            Outlook.MailItem newMail = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);

            //Parse 'sToAddress'
            if (!string.IsNullOrWhiteSpace(sToAddress))
            {
                string[] arrAddTos = sToAddress.Split(new char[] { ';', ',' });
                foreach (string strAddr in arrAddTos)
                {
                    if (!string.IsNullOrWhiteSpace(strAddr) &&
                        strAddr.IndexOf('@') != -1)
                    {
                        newMail.Recipients.Add(strAddr.Trim());
                    }
                    else
                        throw new Exception("Bad to-address: " + sToAddress);
                }
            }
            else
                throw new Exception("Must specify to-address");

            //Parse 'sCc'
            if (!string.IsNullOrWhiteSpace(sCc))
            {
                string[] arrAddTos = sCc.Split(new char[] { ';', ',' });
                foreach (string strAddr in arrAddTos)
                {
                    if (!string.IsNullOrWhiteSpace(strAddr) &&
                        strAddr.IndexOf('@') != -1)
                    {
                        newMail.Recipients.Add(strAddr.Trim());
                    }
                    else
                        throw new Exception("Bad CC-address: " + sCc);
                }
            }

            //Is BCC empty?
            if (!string.IsNullOrWhiteSpace(sBcc))
            {
                newMail.BCC = sBcc.Trim();
            }

            //Resolve all recepients
            if (!newMail.Recipients.ResolveAll())
            {
                throw new Exception("Failed to resolve all recipients: " + sToAddress + ";" + sCc);
            }


            newMail.Body = sBody;

            if (arrAttachments != null)
            {
                //Add attachments
                foreach (string strPath in arrAttachments)
                {
                    if (File.Exists(strPath))
                    {
                        newMail.Attachments.Add(strPath);
                    }
                    else
                        throw new Exception("Attachment file is not found: \"" + strPath + "\"");
                }
            }

            //Add subject
            if (!string.IsNullOrWhiteSpace(sSubject))
                newMail.Subject = sSubject;

            Outlook.Accounts accounts = app.Session.Accounts;
            Outlook.Account acc = null;

            //Look for our account in the Outlook
            foreach (Outlook.Account account in accounts)
            {
                if (account.SmtpAddress.Equals(sFromAddress, StringComparison.CurrentCultureIgnoreCase))
                {
                    //Use it
                    acc = account;
                    break;
                }
            }

            //Did we get the account
            if (acc != null)
            {
                //Use this account to send the e-mail. 
                newMail.SendUsingAccount = acc;

                //And send it
                ((Outlook._MailItem)newMail).Send();

                //Done
                bRes = true;
            }
            else
            {
                throw new Exception("Account does not exist in Outlook: " + sFromAddress);
            }
        }
        catch (Exception ex)
        {
            errorLog("ERROR: Failed to send mail: " + ex.Message);
        }

        return bRes;
    }
    public void ValidateCardAccount()
    {
        try
        {
            List<CCCUST> list = getCCUST();
            foreach (CCCUST c in list)
            {
                bool bValid = true;
                SBO_Application.SetStatusBarMessage("Validating credit card " + c.methodDescription + ".  Customer ID:" + c.CustomerID, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                try
                {
                    string currency = "";
                    string group = getGroupName(c.CustomerID);
                    string acctid = "";
                    string cardname = getCardName(group, c.cardType, currency, ref acctid, getBranchID());

                    SecurityToken token = getToken(acctid);
                    Customer cobj = ebiz.GetCustomer(token, c.CustomerID, "");
                }
                catch (Exception ex)
                {
                    if (ex.Message.ToLower().IndexOf("customer not found") >= 0)
                    {
                        string msg = string.Format("Error in finding Card {0} of customer {1} in gateway.\r\ncard code: {2}.\r\nError: {3} \r\n\r\nRemove card?", c.methodDescription, c.CustomerID, c.CCAccountID, ex.Message);
                        if (SBO_Application.MessageBox(msg, 1, "Yes", "No") == 1)
                        {
                            deleteCCCust(c.recID);
                        }
                        else
                        {
                            if (SBO_Application.MessageBox("Exit validation?", 1, "Yes", "No") == 1)
                            {
                                SBO_Application.SetStatusBarMessage("Card validation completed.", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                return;
                            }
                        }
                    }
                }
                if (bValid)
                    SBO_Application.SetStatusBarMessage(c.methodDescription + " validated.  Customer ID:" + c.CustomerID, SAPbouiCOM.BoMessageTime.bmt_Medium, false);

            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        SBO_Application.SetStatusBarMessage("Card validation completed.", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

    }

    public void UpdateXMLForm(string pathstr, string sFormUID)
    {
        try
        {
            System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();
            string[] abc = Assembly.GetExecutingAssembly().GetManifestResourceNames();

            System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("eBizChargeForSAP." + pathstr);
            System.IO.StreamReader streamreader = new System.IO.StreamReader(stream, true);
            xmldoc.LoadXml(streamreader.ReadToEnd());
            streamreader.Close();
            var xml = xmldoc.InnerXml;
            //    SBO_Application.LoadBatchActions(xmldoc.InnerXml);


            string sXPath = "Application//forms//action//form//@uid";
            XmlDocument xDoc = new XmlDocument();
            xDoc.LoadXml(xml);
            //xDoc.Load(sPath + sFileName);
            XmlNode xNode = xDoc.SelectSingleNode(sXPath);
            xNode.InnerText = sFormUID;
            string sXML = xDoc.InnerXml.ToString();
            SBO_Application.LoadBatchActions(ref sXML);


            //EventHandler.oApplication.LoadBatchActions(xmldoc.InnerXml);

        }
        catch (Exception ex)
        {
            String error = ex.Message;
            //EventHandler.oApplication.StatusBar.SetText("AddXML Method Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }
        finally
        {
        }
    }
    public void AddXMLForm(string pathstr)
    {
        try
        {
            System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();
            string[] abc = Assembly.GetExecutingAssembly().GetManifestResourceNames();

            System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("eBizChargeForSAP." + pathstr);
            System.IO.StreamReader streamreader = new System.IO.StreamReader(stream, true);
            xmldoc.LoadXml(streamreader.ReadToEnd());
            streamreader.Close();
            var xml = xmldoc.InnerXml;
            SBO_Application.LoadBatchActions(ref xml);

            //EventHandler.oApplication.LoadBatchActions(xmldoc.InnerXml);

        }
        catch (Exception ex)
        {
            String error = ex.Message;
            //EventHandler.oApplication.StatusBar.SetText("AddXML Method Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }
        finally
        {
        }
    }

}

