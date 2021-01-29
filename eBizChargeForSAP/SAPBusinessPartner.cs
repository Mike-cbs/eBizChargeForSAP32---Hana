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
using System.Reflection;

public class SAPCust
{
    public CCCUST cccust { get; set; }
    public Customer custObj { get; set; }
    public string key { get; set; }
    public string currency { get; set; }
    public string cardHolder { get; set; }
    public string cardNum { get; set; }
    public DateTime paymentDate { get; set; }
    //  public string CCAccount { get; set; }
    public string SaveCard { get; set; }
    public bool Changed()
    {

        try
        {
            if (custObj.PaymentMethodProfiles[0].MethodType == "cc")
            {
                if (custObj.PaymentMethodProfiles[0].CardNumber.IndexOf("XXX") == -1)
                {
                    //SAP.errorLog("CardNumber is new:  " + custObj.PaymentMethodProfiles[0].CardNumber);
                    return true;
                }
                if (custObj.PaymentMethodProfiles[0].CardNumber.IndexOf(cccust.cardNum) == -1)
                {
                    //SAP.errorLog("CardNumber is different:  custObj:" + custObj.PaymentMethodProfiles[0].CardNumber +",cccust: " + cccust.cardNum);
                    return true;
                }
                if (custObj.PaymentMethodProfiles[0].CardExpiration.Length == 4)
                {
                    if (custObj.PaymentMethodProfiles[0].CardExpiration != cccust.expDate)
                    {
                        //SAP.errorLog("CardExp is different:  custObj:" + custObj.PaymentMethodProfiles[0].CardExpiration + ",cccust: " + cccust.expDate);
                        return true;
                    }
                }

            }
            if (custObj.PaymentMethodProfiles[0].MethodType == "check")
            {
                if (custObj.PaymentMethodProfiles[0].Account.IndexOf("XXX") == -1
                    &&
                    custObj.PaymentMethodProfiles[0].Routing.IndexOf("XXX") == -1
                    )
                    return true;
            }
        }
        catch (Exception)
        {

        }
        return false;
    }
}

public class CB_JOB
{
    public CCJOB job { get; set; }
    public string key { get; set; }

}

partial class SAP
{

    const int paneCCLog = 127;
    int paneRPMT = 128;
    List<SAPCust> SAPCustList = new List<SAPCust>();
    List<CB_JOB> JOBList = new List<CB_JOB>();
    SAPCust selSAPCust = null;
    const string btnAdd = "Add";
    const string btnDelete = "Delete";
    const string btnUpdate = "Update";
    const string lbDeclined = "lbDecl";
    const string tbeBizCharge = "tbeBiz";
    const string tbRecursivePmt = "tbRPMT";
    const string tbPaymentMethod = "tbPayMe";
    const string NEWCARD = "New Card";
    const string NEWCHECK = "New eCheck";
    const string LB_CANCEL = "Cancel";
    const string LB_ACTIVATE = "Activate";
    bool bcontrolCreated = false;
    SAPbouiCOM.Form theBPForm;
    string selectedPM = NEWCARD;
    SAPbouiCOM.Form formSelTarget;
    bool bManualPopulate = false;
    SAPbouiCOM.Button Btn_paymentmethod1 = null;
    SAPbouiCOM.Button btn_TransactionLog1 = null;
    SAPbouiCOM.Button Btn_paymentmethod2 = null;
    SAPbouiCOM.Button btn_TransactionLog2 = null;
    SAPbouiCOM.Folder TabEbiz = null;

    public void selectCard()
    {
        SAPbouiCOM.ComboBox oCB = null;
        try
        {
            Thread.Sleep(500);
            oCB = (SAPbouiCOM.ComboBox)formSelTarget.Items.Item(cbPaymentMethod).Specific;
            oCB.Select(selectedPM);
            if (bManualPopulate)
                HandleSelect(formSelTarget);

        }
        catch (Exception ex)
        {
            try
            {
                oCB.Select(NEWCARD);
            }
            catch (Exception)
            {

            }

        }
        finally
        {
            if (theActiveForm != null)
            {
                try
                {
                    if (theActiveForm.TypeEx == FORMINCOMINGPAYMENT)
                    {
                        formSelTarget.Items.Item(editCAmount).Enabled = false;
                    }
                }
                catch (Exception) { }
            }
        }
    }
    public void initPB()
    {
        try
        {
            Thread.Sleep(2000);
            FormFreeze(formSelTarget, true);


            createControl(formSelTarget);

        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
        finally
        {
            FormFreeze(formSelTarget, false);
        }
    }
    public void createControl(SAPbouiCOM.Form form)
    {
        try
        {

            if (!bcontrolCreated)
            {
                bcontrolCreated = true;
                //showStatus("Loading customer data please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                form.Items.Item(tbeBizCharge).Visible = true;

                // AddTab(form, panePM, "Recurvsive Payment",  tbPaymentMethod, tbRecursivePmt, panePM, false, 50, -100);


                SAPbouiCOM.Item oItemRef = form.Items.Item(tbPaymentMethod);
                /*
                SAPbouiCOM.Item oItem = form.Items.Add("rect", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
                oItem.Top = oItemRef.Top + oItemRef.Height + 5;
                oItem.Left = 5;
                oItem.Width = form.ClientWidth - 20;
                oItem.Height = 1;
                oItem.FromPane = panePM;
                oItem.ToPane = paneCCLog;
                 */


                AddRecursivePaymentItem(form, paneCCLog, oItemRef.Top + oItemRef.Height + 10);
                AddCCLogMatrix(form, paneCCLog, oItemRef.Height + 10);
                AddCCFormField(form, panePM);


                form.Items.Item(tbeBizCharge).Visible = true;

                //loadCustData(form);
                // SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbRecursivePmt).Specific;
                // oCB.Select(ALL_TRANS);
                //  oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbPaymentMethod).Specific;

                //showStatus("Customer data loaded", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void BusinessPartnerFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            theBPForm = form;
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:

                        break;
                }
            }
            else
            {
                switch (pVal.EventType)
                {

                    case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
                        currentForm = form;
                        //tProcess = new Thread(BPMoveThreadProc);
                        //tProcess.Start();
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:

                        try
                        {
                            formSelTarget = form;

                            //form.Items.Item(tbeBizCharge).Visible = false;
                            bcontrolCreated = false;
                            tProcess = new Thread(BPCreateTabThreadProc);
                            tProcess.Start();
                        }
                        catch (Exception ex)
                        {
                            errorLog(ex);
                        }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        try
                        {
                            if (!bcontrolCreated)
                                return;
                            if (pVal.ItemChanged == true && pVal.ItemUID == editCardNum)
                            {
                                setCardName(form, selSAPCust);
                            }
                        }
                        catch (Exception)
                        { }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        {
                            switch (pVal.ItemUID)
                            {
                                case cbPaymentMethod:
                                    HandleSelect(form);
                                    break;
                                case cbActive:

                                    HandleActiveSelect(form);


                                    break;
                                case cbDefault:

                                    HandleDefaultSelect(form);


                                    break;
                                case cbRecursivePmt:
                                    string k = getFormItemVal(form, cbRecursivePmt);
                                    foreach (CB_JOB job in JOBList)
                                    {
                                        if (job.key == k)
                                        {
                                            setLabelCaption(form, stRPmtInvoice, job.job.InvoiceID.ToString());
                                            strWhere = string.Format(" Where \"U_jobID\"='{0}' ", job.job.jobID);
                                            populateLog(form);
                                            if (job.job.Cancelled == 'Y')
                                            {
                                                setLabelCaption(form, stRPmtStatus, "Cancelled on " + job.job.CancelledDate);
                                                showFormItem(form, btnRPmtCancel, true);
                                                setButtonCaption(form, btnRPmtCancel, LB_ACTIVATE);
                                            }
                                            else
                                            {
                                                setLabelCaption(form, stRPmtStatus, "");
                                                showFormItem(form, btnRPmtCancel, true);
                                                setButtonCaption(form, btnRPmtCancel, LB_CANCEL);
                                            }

                                        }
                                    }
                                    if (k == ALL_TRANS)
                                    {
                                        strWhere = string.Format(" Where \"U_customerID\"='{0}' ", getFormItemVal(form, "5"));
                                        populateLog(form);
                                        setLabelCaption(form, stRPmtStatus, "");
                                        showFormItem(form, btnRPmtCancel, false);
                                    }
                                    /*
                                    if(k == NEW_RECUR_BILLING)
                                    {
                                        JobCustID = getFormItemVal(form, "5");
                                        CreateJobForm();
                                    }
                                    */
                                    break;
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == mxtCCLog)
                        {
                            HandleCCLogSelect(form);
                        }
                        else if(pVal.ItemUID == tbCreditCard || pVal.ItemUID == tbPaymentMethod)
                        {
                            handleItemPress(form, pVal);
                        }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        handleItemPress(form, pVal);
                        break;
                }

            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void BusinessPartnerFormHandler2(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            theBPForm = form;
            Btn_paymentmethod1 = ((SAPbouiCOM.Button)(form.Items.Item("btnPM1").Specific));
            btn_TransactionLog1 = ((SAPbouiCOM.Button)(form.Items.Item(tbCreditCard).Specific));
            Btn_paymentmethod2 = ((SAPbouiCOM.Button)(form.Items.Item(tbPaymentMethod).Specific));
            btn_TransactionLog2 = ((SAPbouiCOM.Button)(form.Items.Item("btnTrn2").Specific));
            TabEbiz = ((SAPbouiCOM.Folder)(form.Items.Item(tbeBizCharge).Specific));
            SAPbouiCOM.Folder F0 = ((SAPbouiCOM.Folder)(form.Items.Item(tbeBizCharge).Specific));
            F0.GroupWith("156");
          

            Btn_paymentmethod1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Btn_paymentmethod_ClickAfter);
            btn_TransactionLog1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btn_TransactionLog_ClickAfter);
            Btn_paymentmethod2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Btn_paymentmethod_ClickAfter);
            btn_TransactionLog2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btn_TransactionLog_ClickAfter);
            TabEbiz.ClickAfter += new SAPbouiCOM._IFolderEvents_ClickAfterEventHandler(this.Tab_ebizCharge_ClickAfter);

            //setLabelCaption(form, stRecordID, getFormItemVal(form, "5"));

            //string curcode = getCurrencyCodeFormCard(cust.cccust.CCAccountID);
            //errorLog("PopulateCCInfo Currency Code: " + curcode + ", Card Code: " + cust.cccust.CCAccountID);
            //setComboValue(form, cbCurrencyCode, curcode);

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
   
    private void Btn_paymentmethod_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
    {
        if (pVal.ActionSuccess == true)
        {
            theBPForm.PaneLevel = 125;
        }
    }
    private void btn_TransactionLog_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
    {
        if (pVal.ActionSuccess == true)
        {
            theBPForm.PaneLevel = 127;
        }
    }
    private void Tab_ebizCharge_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
    {

        if (pVal.ActionSuccess == true)
        {
            theBPForm.PaneLevel = 125;
        }
    }

    private void loadCustData(SAPbouiCOM.Form form)
    {
        try
        {

            if (getFormItemVal(form, "40") == "C")
            {
                if (getFormItemVal(form, "5") != "")
                {
                    form.Items.Item(tbeBizCharge).Visible = true;
                    if (getFormItemVal(form, stRecordID) != getFormItemVal(form, "5"))
                    {
                        setLabelCaption(form, stRecordID, getFormItemVal(form, "5"));
                        bManualPopulate = true;
                        reload(form, ref SAPCustList, ref selSAPCust, true);
                        // SBO_Application.SetStatusBarMessage("Credit card for customer: " + getFormItemVal(form, stRecordID) + " loaded");

                    }
                }
                else
                    HideCC(form);
            }
            else
            {
                HideCC(form);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void HideCC(SAPbouiCOM.Form form)
    {
        try
        {

            clearForm(form);
            if (form.Items.Item(stRecordID) != null)
                setLabelCaption(form, stRecordID, "");

            if (form.Items.Item(tbCreditCard) != null)
                form.Items.Item(tbCreditCard).Visible = false;

            form.PaneLevel = 0;
            form.Items.Item("3").Click();

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }

    private void handleItemPress(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            switch (pVal.ItemUID)
            {

                case tbeBizCharge:
                    try
                    {
                        //Azeem
                        //form.Items.Item(tbeBizCharge).Click();
                        /*
                           FormFreeze(form, true);
                           if (getFormItemVal(form, "40") != "C" || getFormItemVal(form, "5") == "" || form.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                           {
                               if (form.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                   showStatus("Please add customer first.", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                               form.Items.Item("3").Click();
                               return;
                           }
                           createControl(form);
                           form.PaneLevel = panePM;


                           SAPbouiCOM.Folder f = (SAPbouiCOM.Folder)form.Items.Item(tbPaymentMethod).Specific;
                           if (!f.Selected)
                               f.Select();
                           f = (SAPbouiCOM.Folder)form.Items.Item(tbeBizCharge).Specific;
                           if (!f.Selected)
                               f.Select();
                           selectedPM = NEWCARD;
                           formSelTarget = form;
                           form.Items.Item(tbPaymentMethod).Click();
                        */
                    }
                    catch (Exception)
                    {

                    }
                    finally
                    {
                        FormFreeze(form, false);
                    }
                    break;

                case tbCreditCard:

                    try
                    {
                        FormFreeze(form, true);
                        form.PaneLevel = paneCCLog;
                        LoadJob(form);
                        strWhere = string.Format(" Where \"U_customerID\"='{0}' ", getFormItemVal(form, "5"));
                        populateLog(form);
                    }
                    catch (Exception)
                    { }
                    finally
                    {
                        FormFreeze(form, false);
                    }
                    break;
                case tbPaymentMethod:
                    try
                    {
                        FormFreeze(form, true);
                        BPMoveControl(form);
                        loadCustData(form);
                        form.PaneLevel = panePM;
                    }
                    catch (Exception)
                    { }
                    finally
                    {
                        FormFreeze(form, false);
                    }
                    break;
                case "18":
                    loadCustData(form);
                    var cardcode = getFormItemVal(form, "5");
                    strWhere = string.Format(" Where \"U_customerID\"='{0}' ", getFormItemVal(form, "5"));
                    populateLog(form);
                    break;
                case btnRPmtCancel:
                    string k = getFormItemVal(form, cbRecursivePmt);
                    CB_JOB selJob = null;
                    foreach (CB_JOB j in JOBList)
                    {
                        if (j.key == k)
                        {
                            selJob = j;
                        }

                    }
                    if (selJob != null)
                    {
                        bAuto = false;
                        if (getFormItemVal(form, btnRPmtCancel) == LB_CANCEL)
                        {

                            if (SBO_Application.MessageBox("Are you sure you want to cancel this payment schedule?", 1, "Yes", "No") == 1)
                            {
                                CancelJob(selJob.job.jobID.ToString());

                                LoadJob(form);
                                SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbRecursivePmt).Specific;
                                oCB.Select(selJob.key);
                            }
                        }
                        else
                        {
                            if (SBO_Application.MessageBox("Activate this payment schedule?", 1, "Yes", "No") == 1)
                            {
                                ActivateJob(selJob.job.jobID.ToString());

                                LoadJob(form);
                                SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbRecursivePmt).Specific;
                                oCB.Select(selJob.key);
                            }
                        }
                    }
                    break;

                case btnAdd:
                    if (IsWalkIn(getCustomerID(form)))
                    {
                        string msg = string.Format("Cannot add payment method.  Customer: {0} is a walk-in customer.", getCustomerID(form));
                        showMessage(msg);
                        return;
                    }
                    if (SBO_Application.MessageBox(string.Format("Add card {0} to customer {1}?", getFormItemVal(form, editCardNum), getCustomerID(form)), 1, "Yes", "No") != 1)
                        return;
                    showStatus("Add customer.  Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                    bAuto = false;
                    SAPCust sapcust = new SAPCust();
                    if (!loadFormData(form, ref sapcust, false))
                        return;
                    sapcust.cccust.CustomerID = getCustomerID(form);
                    sapcust.custObj.CustomerId = getCustomerID(form);
                    if (Validate(form, sapcust.custObj))
                    {
                        try
                        {

                            if (getFormItemVal(form, cbActive) == "Y")
                                sapcust.cccust.active = 'Y';
                            else
                                sapcust.cccust.active = 'N';
                            if (getFormItemVal(form, cbDefault) == "Y")
                                sapcust.cccust.@default = 'Y';
                            else
                                sapcust.cccust.@default = 'N';
                            if (AddCustomer(ref sapcust))
                            {
                                showMessage("Payment method added.");
                                bManualPopulate = false;
                                reload(form, ref SAPCustList, ref selSAPCust, true);
                            }
                        }
                        catch (Exception ex)
                        {
                            showMessage("Failed to add your payment method: " + ex.Message);
                        }

                    }
                    break;
                case btnDelete:
                    bAuto = false;
                    if (selSAPCust == null)
                    {
                        showMessage("Please select a payment method.");
                    }
                    else
                    {
                        bAuto = false;
                        if (SBO_Application.MessageBox("Are you sure you want to remove payment method " + selSAPCust.key + "?", 1, "Yes", "No") == 1)
                        {
                            showStatus("Delete customer.  Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                            //deleteCustomer(selsapcust.custObj.CustomerToken);
                            deleteCCCust(selSAPCust.cccust.recID);
                            bManualPopulate = false;
                            reload(form, ref SAPCustList, ref selSAPCust, true);
                            showMessage("Payment method removed.");
                        }
                    }
                    break;
                case btnUpdate:
                    {
                        bAuto = false;
                        if (selSAPCust == null)
                        {
                            showMessage("Please select a payment method.");
                            return;
                        }
                        if (!loadFormData(form, ref selSAPCust, false))
                            return;

                        if (Validate(form, selSAPCust.custObj))
                        {
                            if (SBO_Application.MessageBox("Update payment method " + selSAPCust.key + "?", 1, "Yes", "No") == 1)
                            {
                                showStatus("Update customer.  Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                                selSAPCust.custObj.CustomerId = selSAPCust.cccust.CustomerID;
                                selSAPCust.custObj.BillingAddress.CompanyName = selSAPCust.cccust.GroupName;
                                if (updateCustomer(ref selSAPCust))
                                {
                                    bManualPopulate = false;
                                    reload(form, ref SAPCustList, ref selSAPCust, true);
                                    showMessage("Payment method updated.");
                                }


                            }
                        }
                    }
                    break;
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        finally
        {
            //Call this function without set it to true first causes error in 8.82
            //form.Freeze(false);
        }

    }

    private bool loadFormData(SAPbouiCOM.Form form, ref SAPCust cust, bool bPaymentForm)
    {
        try
        {
            string pm = getFormItemVal(form, cbPaymentMethod);
            if (pm == NEWCARD || pm == NEWCHECK)
            {
                cust = new SAPCust();
                cust.cccust = new CCCUST();
                cust.custObj = new Customer();
                cust.custObj.PaymentMethodProfiles = new PaymentMethodProfile[1];
                cust.custObj.PaymentMethodProfiles[0] = new PaymentMethodProfile();
                cust.custObj.BillingAddress = new Address();
            }
            handleSwipe(form);
            if (cust == null)
                cust = new SAPCust();
            if (cust.cccust == null)
            {
                cust.cccust = new CCCUST();
                cust.cccust.CustNum = "";
            }
            if (cust.custObj == null)
            {
                cust.custObj = createCustObj();
            }
            if (cust.cccust.CustomerID == "" || cust.cccust.CustomerID == null)
            {
                if (form.TypeEx == FORMBPARTNER)
                {
                    cust.cccust.CustomerID = getFormItemVal(form, "5");
                }
                else
                {
                    cust.cccust.CustomerID = getCustomerID(form);
                }
                cust.currency = getCurrency(form);
            }
            cust.cccust.methodName = getFormItemVal(form, editMethodName);
            cust.custObj.Email = getFormItemVal(form, editHolderEmail);
            cust.custObj.BillingAddress.FirstName = getFirstName(getFormItemVal(form, editHolderName));
            cust.custObj.BillingAddress.LastName = getLastName(getFormItemVal(form, editHolderName));
            cust.custObj.FirstName = getFirstName(getFormItemVal(form, editHolderName));
            cust.custObj.LastName = getLastName(getFormItemVal(form, editHolderName));
            cust.custObj.BillingAddress.Address1 = getFormItemVal(form, editHolderAddr);
            cust.custObj.BillingAddress.City = getFormItemVal(form, editHolderCity);
            cust.custObj.BillingAddress.State = getFormItemVal(form, editHolderState);
            cust.custObj.BillingAddress.ZipCode = getFormItemVal(form, editHolderZip);
            //cust.custObj.SendReceipt = cfgsendCustomerEmail == "Y" ? true : false;
            cust.custObj.PaymentMethodProfiles[0].MethodName = getFormItemVal(form, editMethodName);
            if (getFormItemVal(form, editCardNum) != "")
            {
                cust.custObj.PaymentMethodProfiles[0].MethodType = "cc";
                cust.custObj.PaymentMethodProfiles[0].CardNumber = getFormItemVal(form, editCardNum);
                cust.custObj.PaymentMethodProfiles[0].CardExpiration = getFormItemVal(form, editCardExpDate);
                cust.cccust.expDate = getFormItemVal(form, editCardExpDate);
                cust.custObj.PaymentMethodProfiles[0].CardCode = getFormItemVal(form, editCardCode);
                cust.custObj.PaymentMethodProfiles[0].Account = "";
                cust.custObj.PaymentMethodProfiles[0].Routing = "";
                if (cust.custObj.PaymentMethodProfiles[0].CardNumber.IndexOf("X") == -1)
                {
                    cust.custObj.PaymentMethodProfiles[0].CardType = getCardType(cust.custObj.PaymentMethodProfiles[0].CardNumber);

                }
                cust.custObj.PaymentMethodProfiles[0].AccountHolderName = getFormItemVal(form, editHolderName);
            }
            if (getFormItemVal(form, editCheckAccount) != "")
            {
                cust.custObj.PaymentMethodProfiles[0].MethodType = "check";
                cust.custObj.PaymentMethodProfiles[0].CardNumber = "";
                cust.custObj.PaymentMethodProfiles[0].CardExpiration = "";
                cust.custObj.PaymentMethodProfiles[0].CardCode = "";

                cust.custObj.PaymentMethodProfiles[0].Account = getFormItemVal(form, editCheckAccount);
                cust.custObj.PaymentMethodProfiles[0].Routing = getFormItemVal(form, editCheckRouting);
                cust.custObj.PaymentMethodProfiles[0].AccountHolderName = getFormItemVal(form, editHolderName);
            }
            if (bPaymentForm)
            {

                if (cust.custObj.PaymentMethodProfiles[0].MethodType == "cc")
                {
                    if (cust.custObj.PaymentMethodProfiles[0].CardNumber[0] == 'X')
                    {
                        cust.custObj.PaymentMethodProfiles[0].CardNumber = getCardHeaderCode(cust.custObj.PaymentMethodProfiles[0].CardType) + cust.custObj.PaymentMethodProfiles[0].CardNumber;
                    }
                }
            }
            cust.cccust.CardName = getFormItemVal(form, stCardName);
            if (cust.cccust.CardName.IndexOf(",") >= 0)
            {
                string currency = getCurrency(form);
                cust.cccust.GroupName = getGroupName(cust.cccust.CustomerID);
                if ((cust.cccust.cardType == "" || cust.cccust.cardType == null) && cust.custObj.PaymentMethodProfiles[0].Account != "" && cust.custObj.PaymentMethodProfiles[0].Account != null)
                    cust.cccust.cardType = "eCheck";
                string acctid = "";
                string cardname = getCardName(cust.cccust.GroupName, cust.cccust.cardType, currency, ref acctid, getBranchID());
                if (cust.cccust.CardName != cardname || cust.cccust.CCAccountID != acctid)
                {
                    cust.cccust.CardName = cardname;
                    cust.cccust.CCAccountID = acctid;
                    update(cust.cccust);
                }
            }
            cust.cccust.cardCode = getCardCode(cust.cccust.CardName);
            cust.cccust.CCAccountID = cust.cccust.cardCode;
            if (form.TypeEx == FORMBPARTNER)
            {
                if (getFormItemVal(form, cbActive) == "Y")
                    cust.cccust.active = 'Y';
                else
                    cust.cccust.active = 'N';
                if (getFormItemVal(form, cbDefault) == "Y")
                    cust.cccust.@default = 'Y';
                else
                    cust.cccust.@default = 'N';
            }
            try
            {
                if (form.TypeEx == FORMPAYMENTMEANS || form.TypeEx == formCredit)
                {
                    string dt = getFormItemVal(form, editPaymentDate);
                    if (dt != "")
                        cust.paymentDate = DateTime.Parse(dt);
                    else
                        cust.paymentDate = DateTime.Today;
                }
            }
            catch (Exception)
            {
                cust.paymentDate = DateTime.Today;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    private string getCardHeaderCode(string type)
    {
        string head = "";
        switch (type.ToLower())
        {
            case "v":
                head = "4";
                break;
            case "m":
                head = "5";
                break;
            case "a":
                head = "34";
                break;
            case "ds":
                head = "6011";
                break;
        }
        return head;
    }
    private string getCardType(string CardNum)
    {
        string head = "";
        switch (CardNum.Substring(0, 1))
        {
            case "4":
                head = "v";
                break;
            case "5":
                head = "m";
                break;
            case "3":
                head = "a";
                break;
            case "6":
                head = "ds";
                break;
        }
        return head;
    }
    private SAPbouiCOM.ValidValue ComboAddItem(SAPbouiCOM.ComboBox oCB, string val)
    {
        try
        {
            return oCB.ValidValues.Add(val, val);

        }
        catch (Exception)
        {

        }
        return null;
    }
    private void reload(SAPbouiCOM.Form form, ref List<SAPCust> list, ref SAPCust sel, bool bAll = false)
    {
        SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbPaymentMethod).Specific;
        try
        {
            if (form.Items.Item(cbPaymentMethod) == null)
                return;
            string customerID = getCustomerID(form);
            bool hasDefault = false;

            sel = new SAPCust();
            list = new List<SAPCust>();
            try
            {
                while (oCB.ValidValues.Count > 0)
                    oCB.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception)
            { }
            ComboAddItem(oCB, NEWCARD);
            ComboAddItem(oCB, NEWCHECK);
            List<CCCUST> q = getCCUST(customerID);
            /*
            db.Refresh(System.Data.Linq.RefreshMode.OverwriteCurrentValues, db.CCCUSTs);
            var q = from a in db.CCCUSTs
                    where a.CustomerID == customerID
                    select a;
             */
            if (q.Count() == 0)
            {
                int n = oCB.ValidValues.Count - 2;
                showStatus("Number of record in combo= " + n.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                oCB.Select(NEWCARD);
                populateCustFromDB(form);
                showStatus("Credit card for customer: " + getFormItemVal(form, stRecordID) + " loaded. No credit card record found", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                return;
            }
            if (q.Count() > 0)
            {

                foreach (CCCUST cust in q)
                {
                    if (cust.active == 'Y' || bAll)
                    {
                        SAPCust sapcust = new SAPCust();
                        sapcust.cccust = cust;
                        sapcust.key = cust.methodDescription;
                        sapcust.custObj = new Customer();
                        sapcust.custObj.CustomerToken = cust.CustNum;
                        sapcust.custObj.BillingAddress = new Address();
                        sapcust.custObj.Email = cust.email;
                        sapcust.custObj.BillingAddress.FirstName = cust.firstName;
                        sapcust.custObj.BillingAddress.LastName = cust.lastName;
                        sapcust.custObj.BillingAddress.State = cust.state;
                        sapcust.custObj.BillingAddress.City = cust.city;
                        sapcust.custObj.BillingAddress.ZipCode = cust.zip;
                        sapcust.custObj.BillingAddress.Address1 = cust.street;
                        sapcust.custObj.PaymentMethodProfiles = new PaymentMethodProfile[1];
                        sapcust.custObj.PaymentMethodProfiles[0] = new PaymentMethodProfile();
                        sapcust.custObj.PaymentMethodProfiles[0].CardNumber = cust.cardNum;
                        sapcust.custObj.PaymentMethodProfiles[0].CardType = cust.cardType;
                        sapcust.custObj.PaymentMethodProfiles[0].CardExpiration = cust.expDate;
                        sapcust.custObj.PaymentMethodProfiles[0].CardCode = cust.cardCode;
                        sapcust.custObj.PaymentMethodProfiles[0].MethodID = cust.MethodID;
                        sapcust.custObj.PaymentMethodProfiles[0].Routing = cust.routingNumber;
                        sapcust.custObj.PaymentMethodProfiles[0].Account = cust.checkingAccount;
                        if (cust.checkingAccount != null)
                            sapcust.custObj.PaymentMethodProfiles[0].MethodType = "check";
                        else
                            sapcust.custObj.PaymentMethodProfiles[0].MethodType = "cc";
                        if (cust.Declined == 'Y')
                            sapcust.key = sapcust.key + ", Declined";
                        if (sapcust.key != "" && sapcust.key != null)
                            ComboAddItem(oCB, sapcust.key);
                        list.Add(sapcust);
                        if (cust.@default == 'Y')
                        {
                            sel = sapcust;
                            hasDefault = true;
                            selectedPM = sapcust.key;
                        }


                    }
                }
            }
            form.Items.Item(tbeBizCharge).Click();
            //populateCCInfo(form, sel);
            showStatus("Credit card for customer: " + CustomerID + " loaded. " + (oCB.ValidValues.Count - 2).ToString() + " credit card record(s) found", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            try
            {
                if (theActiveForm != null)
                {
                    if (theActiveForm.TypeEx == FORMINVOICE || theActiveForm.TypeEx == FORMINCOMINGPAYMENT)
                    {
                        if (preauthCCTRAN.CCTRANS.Count > 0)
                        {
                            selectedPM = preauthCCTRAN.paymentKey;
                            hasDefault = true;
                        }
                    }
                }
            }
            catch (Exception) { }
            formSelTarget = form;
            try
            {
                if (!hasDefault)
                {
                    oCB.Select(NEWCARD);
                    if (form.TypeEx != FORMBPARTNER)
                        HandleSelect(form);
                }
                else
                {
                    trace("Has default: " + selectedPM);
                    tProcess = new Thread(selectCard);
                    tProcess.Start();
                }
            }
            catch (Exception ex)
            {
                errorLog(ex);
            }
            //errorLog("reload");
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        finally
        {
            if (theActiveForm != null)
            {
                try
                {
                    if (theActiveForm.TypeEx == FORMINCOMINGPAYMENT)
                    {
                        form.Items.Item(editCAmount).Enabled = false;
                    }
                }
                catch (Exception) { }
            }
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCB);
                oCB = null;
            }
            catch (Exception) { }

        }

    }
    private void HandleSelect(SAPbouiCOM.Form form)
    {
        //errorLog("Handle select");
        try
        {
            if (form == null)
                return;
            FormFreeze(form, true);
            switch (form.TypeEx)
            {
                case FORMBPARTNER:

                    string sel = getFormItemVal(form, cbPaymentMethod);
                    if (sel == NEWCARD || sel == NEWCHECK)
                    {
                        selSAPCust = null;
                        populateCustFromDB(form);

                    }
                    else
                    {
                        selSAPCust = getCCInfoSelect(form, SAPCustList);
                        if (selSAPCust != null)
                            populateCCInfo(form, selSAPCust);

                    }
                    setPane(form, selSAPCust);
                    break;
                case FORMPAYMENTMEANS:
                    PMSAPCust = getCCInfoSelect(form, PMCustList);
                    if (PMSAPCust != null)
                        populateCCInfo(form, PMSAPCust);

                    break;
                case formCredit:
                    CFSAPCust = getCCInfoSelect(form, CFCustList);
                    if (CFSAPCust != null)
                        populateCCInfo(form, CFSAPCust);
                    break;
            }

            form.Items.Item(editCardNum).Click();



        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        finally
        {
            FormFreeze(form, false);

        }
    }
    private void HandleActiveSelect(SAPbouiCOM.Form form)
    {
        try
        {
            if (selSAPCust == null)
                return;
            string active = getFormItemVal(form, cbActive);
            if (active != selSAPCust.cccust.active.ToString())
            {
                if (updateCustState(selSAPCust.cccust.recID, active))
                {
                    if (active == "Y")
                        showMessage("Card activated.");
                    else
                        showMessage("Card deactivated.");
                    reload(form, ref SAPCustList, ref selSAPCust, true);
                    /*
                    selSAPCust.cccust.active = active[0];
                    foreach(SAPCust cust in SAPCustList)
                    {
                        if(cust.cccust.recID == selSAPCust.cccust.recID)
                            cust.cccust.active = active[0];
                    }
                    */
                }
                else
                {
                    showMessage("Failed to update card status.");
                }
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    private void HandleDefaultSelect(SAPbouiCOM.Form form)
    {
        try
        {
            selSAPCust = getCCInfoSelect(form, SAPCustList);
            if (selSAPCust == null)
                return;
            string def = getFormItemVal(form, cbDefault);
            if (def == selSAPCust.cccust.@default.ToString())
                return;
            string msg = "Select card as default?";
            if (def == "N")
                msg = "Unselect card as default?";
            if (SBO_Application.MessageBox(msg, 1, "Yes", "No") == 1)
            {

                if (updateCustDefaultState(selSAPCust.cccust.recID, def, getCustomerID()))
                {
                    if (def == "Y")
                    {
                        updateMethodToDefault(selSAPCust);
                        showMessage("Card selected as default.");
                    }
                    else
                        showMessage("Card unselected as default.");
                    reload(form, ref SAPCustList, ref selSAPCust, true);

                }
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    public bool updateMethodToDefault(SAPCust sapcust)
    {
        bool bResult = true;
        try
        {
            SecurityToken token = getToken(sapcust.cccust.CCAccountID);
            bResult = ebiz.SetDefaultCustomerPaymentMethodProfile(token, sapcust.custObj.CustomerToken, sapcust.cccust.MethodID);
            trace("SetDefaultCustomerPaymentMethodProfile on " + sapcust.cccust.CustomerID + "/" + sapcust.custObj.CustomerToken + "," + sapcust.cccust.MethodID + ", return=" + bResult);
        }
        catch (Exception ex)
        {
            trace("Payment method update failed.\r\n" + ex.Message);
            bResult = false;
        }
        return bResult;
    }
    private string populateCustFromDB(SAPbouiCOM.Form form)
    {
        string sql = "";
        string cardcode = getCustomerID(form);
        try
        {

            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            sql = string.Format("Select a.\"CardName\",b.\"Street\",b.\"City\",b.\"State\",b.\"ZipCode\",a.\"E_Mail\",a.\"CardCode\" from OCRD a, CRD1 b where a.\"CardCode\" = b.\"CardCode\" and a.\"BillToDef\" = b.\"Address\" and a.\"CardCode\"='{0}'", cardcode);

            oRS.DoQuery(sql);

            if (!oRS.EoF)
            {
                trace(string.Format("Get customer info from db. ID={6}, {0}, {1}, {2}, {3}, {4}, {5}", GetFieldVal(oRS, 0), GetFieldVal(oRS, 1), GetFieldVal(oRS, 2), GetFieldVal(oRS, 3), GetFieldVal(oRS, 4), GetFieldVal(oRS, 5), cardcode));
                setFormEditVal(form, editHolderName, GetFieldVal(oRS, 0), panePM, panePM + 1);
                setFormEditVal(form, editHolderAddr, GetFieldVal(oRS, 1), panePM, panePM + 1);
                setFormEditVal(form, editHolderCity, GetFieldVal(oRS, 2), panePM, panePM + 1);
                setFormEditVal(form, editHolderState, GetFieldVal(oRS, 3), panePM, panePM + 1);
                setFormEditVal(form, editHolderZip, GetFieldVal(oRS, 4), panePM, panePM + 1);
                setFormEditVal(form, editHolderEmail, GetFieldVal(oRS, 5), panePM, panePM + 1);
                var CardCode= GetFieldVal(oRS, 6);
                setFormEditVal(form, lblstRecID, CardCode, panePM, panePM);
            }
            else
            {
                trace("Get customer info from db." + cardcode + " has no billing record");
                setFormEditVal(form, editHolderName, "", panePM, panePM + 1);
                setFormEditVal(form, editHolderAddr, "", panePM, panePM + 1);
                setFormEditVal(form, editHolderCity, "", panePM, panePM + 1);
                setFormEditVal(form, editHolderState, "", panePM, panePM + 1);
                setFormEditVal(form, editHolderZip, "", panePM, panePM + 1);
                setFormEditVal(form, editHolderEmail, "", panePM, panePM + 1);
                setFormEditVal(form, lblstRecID, cardcode, panePM, panePM);
            }
            
            setFormEditVal(form, editCardNum, "", panePM, panePM);
            setFormEditVal(form, editCardExpDate, "", panePM, panePM);
            setFormEditVal(form, editCardCode, "", panePM, panePM);
            setFormEditVal(form, editCheckAccount, "", panePM + 1, panePM + 1);
            setFormEditVal(form, editCheckRouting, "", panePM + 1, panePM + 1);
            setFormEditVal(form, editMethodName, "", panePM, panePM + 1);
            if (form.TypeEx == FORMBPARTNER)
            {
                setComboValue(form, cbDefault, "N");
                setComboValue(form, cbActive, "Y");
            }
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        return cardcode;
    }

    private void populateCCInfo(SAPbouiCOM.Form form, SAPCust cust)
    {
        // errorLog("PopulateCCInfo");
        try
        {
            if (form == null)
                return;
            if (form.TypeEx != FORMBPARTNER && form.TypeEx != FORMPAYMENTMEANS && form.TypeEx != formCredit)
                return;
            string cardcode = "";
            if (cust == null)
            {
                cardcode = populateCustFromDB(form);

            }
            else
            {
                if (cust.custObj != null)
                {
                    cust.cardHolder = cust.custObj.BillingAddress.FirstName + " " + cust.custObj.BillingAddress.LastName;
                    cust.cardNum = cust.custObj.PaymentMethodProfiles[0].CardNumber;
                    setFormEditVal(form, editMethodName, cust.cccust.methodName, panePM, panePM + 1);
                    setFormEditVal(form, editHolderName, cust.custObj.BillingAddress.FirstName + " " + cust.custObj.BillingAddress.LastName, panePM, panePM + 1);
                    setFormEditVal(form, editHolderAddr, cust.custObj.BillingAddress.Address1 + " " + cust.custObj.BillingAddress.Address2, panePM, panePM + 1);
                    setFormEditVal(form, editHolderCity, cust.custObj.BillingAddress.City, panePM, panePM + 1);
                    setFormEditVal(form, editHolderState, cust.custObj.BillingAddress.State, panePM, panePM + 1);
                    setFormEditVal(form, editHolderZip, cust.custObj.BillingAddress.ZipCode, panePM, panePM + 1);
                    setFormEditVal(form, editHolderEmail, cust.custObj.Email, panePM, panePM + 1);
                    setFormEditVal(form, editCardNum, cust.custObj.PaymentMethodProfiles[0].CardNumber, panePM, panePM);
                    setFormEditVal(form, editCardExpDate, cust.cccust.expDate, panePM, panePM);
                    setFormEditVal(form, editCardCode, "", panePM, panePM);
                    setFormEditVal(form, editCheckAccount, cust.custObj.PaymentMethodProfiles[0].Account, panePM + 1, panePM + 1);
                    setFormEditVal(form, editCheckRouting, cust.custObj.PaymentMethodProfiles[0].Routing, panePM + 1, panePM + 1);

                    if (form.TypeEx == FORMBPARTNER)
                    {
                        setComboValue(form, cbActive, cust.cccust.active.ToString());
                        setComboValue(form, cbDefault, cust.cccust.@default.ToString());

                    }
                    string curcode = getCurrencyCodeFormCard(cust.cccust.CCAccountID);
                    errorLog("PopulateCCInfo Currency Code: " + curcode + ", Card Code: " + cust.cccust.CCAccountID);
                    setComboValue(form, cbCurrencyCode, curcode);
                    if (cust.cccust.Declined == 'Y')
                        setLabelCaption(form, lbDeclined, "Declined");
                    else
                        setLabelCaption(form, lbDeclined, "");


                }
            }
            if (cust != null)
            {
                if (cust.cccust != null)
                {
                    string currency = getCurrency(form);
                    cust.cccust.GroupName = getGroupName(cust.cccust.CustomerID);
                    if ((cust.cccust.cardType == "" || cust.cccust.cardType == null) && cust.cccust.checkingAccount != "" && cust.cccust.checkingAccount != null)
                        cust.cccust.cardType = "eCheck";
                    string acctid = "";
                    string cardname = getCardName(cust.cccust.GroupName, cust.cccust.cardType, currency, ref acctid);
                    if (cust.cccust.CardName != cardname)
                    {
                        cust.cccust.CardName = cardname;
                        cust.cccust.CCAccountID = acctid;
                        update(cust.cccust);
                    }

                    setLabelCaption(form, stCardName, cust.cccust.CardName);

                }
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    private void clearForm(SAPbouiCOM.Form form)
    {
        try
        {
            if (form.Items.Item(cbPaymentMethod) == null)
                return;
            SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbPaymentMethod).Specific;
            try
            {
                while (oCB.ValidValues.Count > 0)
                    oCB.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);

            }
            catch (Exception)
            {

            }
            oCB.ValidValues.Add("New Card", "New Card");
            setFormEditVal(form, editMethodName, "", panePM, panePM + 1);
            setFormEditVal(form, editHolderName, "", panePM, panePM + 1);
            setFormEditVal(form, editHolderAddr, "", panePM, panePM + 1);
            setFormEditVal(form, editHolderCity, "", panePM, panePM + 1);
            setFormEditVal(form, editHolderState, "", panePM, panePM + 1);
            setFormEditVal(form, editHolderZip, "", panePM, panePM + 1);
            setFormEditVal(form, editHolderEmail, "", panePM, panePM + 1);
            setFormEditVal(form, editCardExpDate, "", panePM, panePM);
            setFormEditVal(form, editCardNum, "", panePM, panePM);
            setComboValue(form, cbActive, "Y");
            setComboValue(form, cbDefault, "N");
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private SAPCust getCCInfoSelect(SAPbouiCOM.Form form, List<SAPCust> l)
    {
        try
        {
            string sel = getFormItemVal(form, cbPaymentMethod);
            if (sel == NEWCHECK)
                form.PaneLevel = panePM + 1;
            else if (sel == NEWCARD)
                form.PaneLevel = panePM;

            foreach (SAPCust c in l)
            {
                if (c.key == sel)
                {
                    if (c.custObj.PaymentMethodProfiles[0].MethodType == "cc")
                        form.PaneLevel = panePM;
                    else
                        form.PaneLevel = panePM + 1;
                    return c;
                }
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return null;
    }
    private void AddCCFormField(SAPbouiCOM.Form form, int pane)
    {
        try
        {
            //     AddXML("Presentation_Layer.BPtab1.srf");
            int edL = 110 + 15; //oItm.Left + oItm.Width;
            int edW = 250;
            int edT = 30;
            int edH = 15;
            int nGap = 16;
            SAPbouiCOM.Item oItm = null;
            SAPbouiCOM.Item oRefItm = null;
            edT = form.Items.Item(tbCreditCard).Top + form.Items.Item(tbCreditCard).Height * 2 + edT;
            oRefItm = form.Items.Item("7");
            edH = oRefItm.Height;
            nGap = edH + 2;
            edW = oRefItm.Width;
            edL = oRefItm.Width + 15;

            oItm = addPaneItem(form, stRecordID, edL, edT, edW, edH, "Customer ID", SAPbouiCOM.BoFormItemTypes.it_STATIC, pane, 0);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, cbCurrencyCode, edL, edT, edW, edH, "Currency:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, pane, 1987);
            string cstr = getCurrencyString();
            SAPbouiCOM.ComboBox oCBC = (SAPbouiCOM.ComboBox)form.Items.Item(cbCurrencyCode).Specific;

            oCBC.ValidValues.Add("", "");
            string[] cArr = cstr.Split(',');
            foreach (string cs in cArr)
            {
                oCBC.ValidValues.Add(cs, cs);
            }
            try
            {
                string curcode = getFormItemVal(form, "11");
                setComboValue(form, cbCurrencyCode, curcode);
            }
            catch (Exception) { }
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, cbPaymentMethod, edL, edT, edW, edH, "Payment Method:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, pane, 1);
            edT = oItm.Top + nGap;
            oItm.AffectsFormMode = false;
            oItm = addPaneItem(form, editCardNum, edL, edT, edW, edH, CC_CARDNUM, SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 2);
            oItm = form.Items.Add(lbDeclined, SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItm.FromPane = pane;
            oItm.ToPane = pane;
            oItm.Left = edW * 3;
            oItm.Top = edT;
            oItm.Width = edW / 3;
            oItm.Height = edH;
            edT = oItm.Top + nGap;
            oItm.AffectsFormMode = false;
            oItm = addPaneItem(form, editCardExpDate, edL, edT, edW, edH, CC_EXPDATE, SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 3);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCardCode, edL, edT, edW, edH, CC_CARDCODE, SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 4);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderName, edL, edT, edW, edH, "*Card/Acct. Holder:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 6);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCheckRouting, edL, edT, edW, edH, EC_ROUTING, SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 101);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCheckAccount, edL, edT, edW, edH, EC_ACCOUNT, SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 100);
            edT = oItm.Top + nGap + 10;
            oItm = addPaneItem(form, editHolderAddr, edL, edT, edW, edH, "Address:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 7);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderCity, edL, edT, edW, edH, "City:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 8);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderState, edL, edT, edW, edH, "State:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 9);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderZip, edL, edT, edW, edH, "Zip:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 10);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderEmail, edL, edT, edW, edH, "Email:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 11);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, cbActive, edL, edT, 60, edH, "Active:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, pane, 12, edW);
            SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbActive).Specific;
            oCB.ValidValues.Add("Y", "Y");
            oCB.ValidValues.Add("N", "N");
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, cbDefault, edL, edT, 60, edH, "Default:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, pane, 18, edW);
            oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbDefault).Specific;
            oCB.ValidValues.Add("Y", "Y");
            oCB.ValidValues.Add("N", "N");
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editMethodName, edL, edT, edW, edH, "Method Name:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 22);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, stCardName, edL, edT, edW, edH, "Card Name:", SAPbouiCOM.BoFormItemTypes.it_STATIC, pane, 20);


            edT = oItm.Top + 70;
            edL = oItm.Left - 20;
            edH = 20;
            int btnW = 60;
            nGap = 65;
            oItm = addPaneButton(form, btnAdd, edL, edT, btnW, edH, "Add", pane);
            oItm.ToPane = pane + 1;
            edL = oItm.Left + nGap;
            oItm = addPaneButton(form, btnDelete, edL, edT, btnW, edH, "Delete", pane);
            oItm.ToPane = pane + 1;
            edL = oItm.Left + nGap;
            oItm = addPaneButton(form, btnUpdate, edL, edT, btnW, edH, "Update", pane);
            oItm.ToPane = pane + 1;
            MoveBPCCControl(form);
        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
    }


    private void MoveBPCCControl(SAPbouiCOM.Form form)
    {
        try
        {

            int edL = 110 + 15; //oItm.Left + oItm.Width;
            int edW = 250;
            int edT = 30;
            int edH = 15;
            int nGap = 16;
            int l = 0;
            SAPbouiCOM.Item oItm = null;
            SAPbouiCOM.Item oRefItm = null;
            edT = form.Items.Item(tbCreditCard).Top + form.Items.Item(tbCreditCard).Height * 2 + edT;
            oRefItm = form.Items.Item("7");
            edH = oRefItm.Height;
            nGap = edH + 2;
            edW = oRefItm.Width;
            edL = oRefItm.Width + 15;


            oItm = form.Items.Item(editCardExpDate);
            nGap = oItm.Height + 2;
            edT = oItm.Top + nGap + 2;

            oItm = form.Items.Item(editCardCode);
            oItm.Top = edT;
            l = oItm.Width + oItm.Left;
            oItm = form.Items.Item("LB4");
            oItm.Top = edT;
            edT = oItm.Top + nGap + 2;

            oItm = form.Items.Item(editHolderName);
            oItm.Top = edT;
            l = oItm.Width + oItm.Left;
            oItm = form.Items.Item("LB6");
            oItm.Top = edT;
            edT = oItm.Top + nGap + 2;

            oItm = form.Items.Item(editCheckRouting);
            oItm.Top = edT;
            l = oItm.Width + oItm.Left;
            oItm = form.Items.Item("LB100");
            oItm.Top = edT;
            edT = oItm.Top + nGap + 2;

            oItm = form.Items.Item(editCheckAccount);
            oItm.Top = edT;
            l = oItm.Width + oItm.Left;
            oItm = form.Items.Item("LB101");
            oItm.Top = edT;
            edT = oItm.Top + nGap + 2;

            oItm = form.Items.Item(editHolderAddr);
            oItm.Top = edT;
            l = oItm.Width + oItm.Left;
            oItm = form.Items.Item("LB7");
            oItm.Top = edT;
            edT = oItm.Top + nGap + 2;

            oItm = form.Items.Item(editHolderCity);
            oItm.Top = edT;
            l = oItm.Width + oItm.Left;
            oItm = form.Items.Item("LB8");
            oItm.Top = edT;
            edT = oItm.Top + nGap + 2;

            oItm = form.Items.Item(editHolderState);
            oItm.Top = edT;
            l = oItm.Width + oItm.Left;
            oItm = form.Items.Item("LB9");
            oItm.Top = edT;
            edT = oItm.Top + nGap + 2;


            oItm = form.Items.Item(editHolderZip);
            oItm.Top = edT;
            l = oItm.Width + oItm.Left;
            oItm = form.Items.Item("LB10");
            oItm.Top = edT;
            edT = oItm.Top + nGap + 2;


            oItm = form.Items.Item(editHolderEmail);
            oItm.Top = edT;
            l = oItm.Width + oItm.Left;
            oItm = form.Items.Item("LB11");
            oItm.Top = edT;
            edT = oItm.Top + nGap + 2;

            oItm = form.Items.Item(cbActive);
            oItm.Top = edT;
            l = oItm.Width + oItm.Left;
            oItm = form.Items.Item("LB12");
            oItm.Top = edT;
            edT = oItm.Top + nGap + 2;

            oItm = form.Items.Item(cbDefault);
            oItm.Top = edT;
            l = oItm.Width + oItm.Left;
            oItm = form.Items.Item("LB18");
            oItm.Top = edT;
            edT = oItm.Top + nGap + 2;

            oItm = form.Items.Item(editMethodName);
            oItm.Top = edT;
            l = oItm.Width + oItm.Left;
            oItm = form.Items.Item("LB22");
            oItm.Top = edT;
            edT = oItm.Top + nGap + 2;

            oItm = form.Items.Item(stCardName);
            oItm.Top = edT;
            l = oItm.Width + oItm.Left;
            oItm = form.Items.Item("LB20");
            oItm.Top = edT;
            edT = oItm.Top + nGap + 12;

            oItm = form.Items.Item(btnAdd);
            oItm.Top = edT;
            oItm = form.Items.Item(btnDelete);
            oItm.Top = edT;
            oItm = form.Items.Item(btnUpdate);
            oItm.Top = edT;

        }
        catch (Exception)
        { }
    }

    private Customer createCustObj()
    {
        Customer CustomerData = new Customer();
        CustomerData.CustomerToken = "";
        CustomerData.BillingAddress = new Address();
        CustomerData.PaymentMethodProfiles = new PaymentMethodProfile[1];
        CustomerData.PaymentMethodProfiles[0] = new PaymentMethodProfile();
        return CustomerData;
    }
    private bool Validate(SAPbouiCOM.Form form, Customer cusObj)
    {
        string err = "";
        try
        {
            if (cusObj.PaymentMethodProfiles[0].MethodType == "cc" || cusObj.PaymentMethodProfiles[0].MethodType == "" || cusObj.PaymentMethodProfiles[0].MethodType == null)
            {
                if (cusObj.PaymentMethodProfiles[0].CardNumber == "")
                {
                    err += "Credit card number is required.\r\n";

                }
                if (cusObj.PaymentMethodProfiles[0].CardExpiration == "")
                {
                    err += "Card expiration date is required.\r\n";
                }
                if (cusObj.BillingAddress.FirstName == "")
                {
                    err += "Card holder name is required.\r\n";
                }

            }
            if (cusObj.PaymentMethodProfiles[0].MethodType == "check")
            {
                if (cusObj.PaymentMethodProfiles[0].Account == "")
                {
                    err += "Checking account number is required.\r\n";

                }
                if (cusObj.PaymentMethodProfiles[0].Routing == "")
                {
                    err += "Routing number is required.\r\n";
                }
                if (cusObj.BillingAddress.FirstName == "")
                {
                    err += "Account holder name is required.\r\n";
                }
            }
            if (cfgaddressRequired == "Y")
            {
                if (cusObj.BillingAddress.Address1 == "" || cusObj.BillingAddress.State == "" || cusObj.BillingAddress.ZipCode == "" || cusObj.BillingAddress.City == "")
                    err += "Billing address is required.";

            }

            if (err == "")
                return true;
            else
                showMessage(err);

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }

    const string cbRecursivePmt = "cbRPMT";
    const string stRPmtStatus = "stStatus";
    const string stRPmtInvoice = "stInv";
    const string btnRPmtCancel = "btnCanRP";
    const string ALL_TRANS = "All Transactions";
    const string NEW_RECUR_BILLING = "Add recurring billing on account";
    private void AddRecursivePaymentItem(SAPbouiCOM.Form form, int pane, int top)
    {
        try
        {
            int edL = 110 + 15; //oItm.Left + oItm.Width;
            int edW = 250;
            int edT = 178;
            int edH = 15;
            int nGap = 17;
            int btnW = 60;
            SAPbouiCOM.Item oItm = null;
            SAPbouiCOM.Item oRefItm = null;
            edT = top;
            oRefItm = form.Items.Item("7");
            edH = oRefItm.Height;
            nGap = edH + 2;
            edW = oRefItm.Width;
            edL = oRefItm.Width + 15;

            oItm = addPaneItem(form, cbRecursivePmt, edL, edT, edW, edH, "Transaction Type:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, pane, 800);
            //LoadJob(form);

            /*
            oItm = form.Items.Add(stRPmtInvoice, SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItm.FromPane = pane;
            oItm.ToPane = pane;
            oItm.Left = edL;
            oItm.Top = edT;
            oItm.Width = 200;
            oItm.Height = 0;

            oItm = form.Items.Add("LBRPMT", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oItm.FromPane = pane;
            oItm.ToPane = pane;
            oItm.Left = edL - 30;
            oItm.Top = edT;
            oItm.Width = 30;
            oItm.Height = edH;
            oItm.LinkTo = stRPmtInvoice;
            SAPbouiCOM.LinkedButton btn = (SAPbouiCOM.LinkedButton)oItm.Specific;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_Invoice).ToString();
            */
            edL = edL + edW;


            oItm = addPaneButton(form, btnRPmtCancel, edL, edT, btnW, edH, LB_CANCEL, pane);
            oItm.Visible = false;
            oItm = form.Items.Add(stRPmtStatus, SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItm.FromPane = pane;
            oItm.ToPane = pane;
            oItm.Left = edL + btnW;
            oItm.Top = edT;
            oItm.Width = 160;
            oItm.Height = edH;

            showFormItem(form, btnRPmtCancel, false);
            //populateLog(form);
        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
    }
    public void LoadJob(SAPbouiCOM.Form form)
    {
        try
        {
            string cid = getFormItemVal(form, "5");
            if (cid == "")
                return;
            SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbRecursivePmt).Specific;
            try
            {
                JOBList = new List<CB_JOB>();
                while (oCB.ValidValues.Count > 0)
                    oCB.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception)
            { }
            oCB.ValidValues.Add(ALL_TRANS, ALL_TRANS);
            // oCB.ValidValues.Add(NEW_RECUR_BILLING, NEW_RECUR_BILLING);
            List<CCJOB> q = findCCJOB(string.Format(" where \"U_CustomerID\"='{0}' ", cid));

            foreach (CCJOB j in q)
            {
                CB_JOB cbjob = new CB_JOB();
                cbjob.job = j;
                cbjob.key = string.Format("{0} {1}_{2}", j.Frequency, j.Description, j.jobID);
                JOBList.Add(cbjob);
                oCB.ValidValues.Add(cbjob.key, cbjob.key);
            }

        }
        catch (Exception)
        {

        }
    }

    public void setCardName(SAPbouiCOM.Form form, SAPCust sapcust)
    {
        try
        {
            FormFreeze(form, true);
            if (sapcust != null)
            {
                if (sapcust.cccust != null)
                {
                    if (sapcust.cccust.CardName != null && sapcust.cccust.CardName != "")
                    {
                        setLabelCaption(form, stCardName, sapcust.cccust.CardName);
                        return;
                    }
                }
            }
            handleSwipe(form);
            string cardNum = getFormItemVal(form, editCardNum);
            if (cardNum.IndexOf("X") >= 0 || cardNum == "")
            {
                setLabelCaption(form, stCardName, sapcust.cccust.CardName);
                return;
            }
            string custid = getCustomerID();
            if (form.TypeEx == FORMBPARTNER)
                custid = getCustomerID(form);
            string cardType = "";
            string currency = getCurrency(form);
            string group = getGroupName(custid);
            if (cardNum != "")
            {
                cardType = getCardType(cardNum);
            }
            string id = "";
            string cardName = getCardName(group, cardType, currency, ref id, getBranchID());
            setLabelCaption(form, stCardName, cardName);
        }
        catch (Exception)
        { }
        finally
        {
            FormFreeze(form, false);
        }

    }
    public void BPMoveThreadProc()
    {
        try
        {
            FormFreeze(currentForm, true);
            BPMoveControl(currentForm);
        }
        catch (Exception ex)
        {
        }
        finally
        {
            FormFreeze(currentForm, false);
        }
    }
    public void BPCreateTabThreadProc()
    {
        try
        {
            if (formSelTarget == null)
            {
                if (theBPForm != null)
                    formSelTarget = theBPForm;
                else
                    return;
            }
            if (formSelTarget.TypeEx != FORMBPARTNER)
                return;

            //AddTab(formSelTarget, panePM, "eBizCharge", "156", tbeBizCharge);
            //AddTab(formSelTarget, panePM, "Payment Method", tbeBizCharge, tbPaymentMethod, panePM, false, 50, -100);
            //AddTab(formSelTarget, panePM, "Transaction Log", tbPaymentMethod, tbCreditCard, panePM, false, 50, -100);

        }
        catch (Exception ex)
        {
        }

    }
    private void BPMoveControl(SAPbouiCOM.Form form)
    {
        try
        {
            SAPbouiCOM.Item oItm = form.Items.Item("LB800");
            int nGap = oItm.Height + 4;
            int edT = oItm.Top + nGap;
            int l = oItm.Left;

            oItm = form.Items.Item("LB19877");
            oItm.Top = edT;
            oItm.Left = l;

            oItm = form.Items.Item(editEmail);
            oItm.Top = edT;
            oItm.Left = l + oItm.Width;
            edT = oItm.Top + nGap;
            oItm = form.Items.Item(mxtCCLog);
            oItm.Top = edT;
            oItm.Left = l;
            MoveBPCCControl(form);
        }
        catch (Exception)
        { }
    }
}

