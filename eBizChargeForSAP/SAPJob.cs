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


partial class SAP
{
    const string formJob = "eBJobForm";
    
    SAPbouiCOM.Form oJobForm;
    string JobID = "";
    string JobCustID = "";
    string JobOrderID = "";
    string JobPaymentID = "";
    string JobAmount = "";
    CCJOB ccjob = new CCJOB();
    const string cbInvType = "cbInvType";

    public void selectPayment()
    {
        try
        {
            Thread.Sleep(500);
            if (ccjob.PaymentID != null && ccjob.PaymentID != "")
            {
                SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)oJobForm.Items.Item(cbPaymentMethod).Specific;
                oCB.Select(ccjob.PaymentID);
            }
            if (JobPaymentID != "")
            {
                SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)oJobForm.Items.Item(cbPaymentMethod).Specific;
                oCB.Select(JobPaymentID);
            }
            if(JobAmount != "")
            {
                setFormEditVal(oJobForm, editCAmount, JobAmount);
              
            }
            if (ccjob.InvoiceID != 0 && ccjob.InvoiceID != null)
            {
                int id = getDocNum("OINV", ccjob.InvoiceID.ToString());
                setFormEditVal(oJobForm, editInvNo, id.ToString()); ;
            }
            if(ccjob.Frequency != "" && ccjob.Frequency != null)
            {
                SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)oJobForm.Items.Item(cbFrequency).Specific;
                oCB.Select(ccjob.Frequency);
            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
    private void populateData()
    {
        try
        {
          
           
            SAPbouiCOM.Button oButton = (SAPbouiCOM.Button)oJobForm.Items.Item(btnUpdate).Specific;
           
            if (JobID != "")
                oButton.Caption = "Update";
            else
                oButton.Caption = "Add";
          
            if (JobID != "")
            {

                List<CCJOB> q = findCCJOB(string.Format(" where \"DocEntry\"='{0}'", JobID));
                foreach (CCJOB job in q)
                    ccjob = job;
                JobCustID = ccjob.CustomerID;
                setFormEditVal(oJobForm, editDescription, ccjob.Description);
                setFormEditVal(oJobForm, editCAmount, ccjob.Amount);
                setFormEditVal(oJobForm, editCustomerID, ccjob.CustomerID);
                setFormEditVal(oJobForm, editStartDate, ((DateTime)ccjob.StartDate).ToString("MM/dd/yyyy"));
                if (ccjob.EndDate != null)
                    setFormEditVal(oJobForm, editEndDate, ((DateTime)ccjob.EndDate).ToString("MM/dd/yyyy"));
                else
                    setFormEditVal(oJobForm, editEndDate, "");

                if (ccjob.CancelledDate != null)
                {
                    setLabelCaption(oJobForm, stCancel, LB_CANCELLED);
                    setButtonCaption(oJobForm, btnUpdate, "Activate");
                }
                setFormEditVal(oJobForm, editDescription, ccjob.Description);
                setComboValue(oJobForm, cbFrequency, ccjob.Frequency);
              
            }
            setFormEditVal(oJobForm, editCustomerID, JobCustID);
            populatePaymentList(oJobForm);
        
        }catch(Exception)
        {

        }
        finally
        {
            tProcess = new Thread(selectPayment);
            tProcess.Start();
        }
    }
    private void CreateJobForm()
    {
        try
        {
            try {
                oJobForm = SBO_Application.Forms.Item(formJob);
                oJobForm.Close();
            }catch(Exception)
            { }
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Button oButton = null;

            // add a new form
            SAPbouiCOM.FormCreationParams oCreationParams = null;

            oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.FormType = formJob;
            oCreationParams.UniqueID = formJob;
            try
            {
                oJobForm = SBO_Application.Forms.AddEx(oCreationParams);
            }
            catch (Exception ex)
            {
                oJobForm = SBO_Application.Forms.Item(formJob);
                errorLog(ex);
            
            }
            // set the form properties
            try
            {
            oJobForm.Title = "eBizCharge Recurring Billing Setup Form";
            oJobForm.Left = 400;
            oJobForm.Top = 100;
            oJobForm.ClientHeight = 300;
            oJobForm.ClientWidth = 500;
            int buttonTop = 270;
            int buttonLeft = 60;
            int buttonWidth = 65;
            int buttonHeight = 19;

            oItem = oJobForm.Items.Add(btnUpdate, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = buttonLeft;
            oItem.Width = buttonWidth;
            oItem.Top = buttonTop;
            oItem.Height = buttonHeight;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));
            if (JobID != "")
                oButton.Caption = "Update";
            else
                oButton.Caption = "Add";

            //************************
            // Adding a Cancel button
            //***********************
           
                oItem = oJobForm.Items.Add(btnClose, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = buttonLeft + buttonWidth + 5;
                oItem.Width = buttonWidth;
                oItem.Top = buttonTop;
                oItem.Height = buttonHeight;

                oButton = ((SAPbouiCOM.Button)(oItem.Specific));

                oButton.Caption = "Cancel";
            }
            catch (Exception) { }
            AddJobFormField(oJobForm);

            oJobForm.Visible = true;
            setFormEditVal(oJobForm, editStartDate, DateTime.Today.ToString("MM/dd/yyyy"));
            setFormEditVal(oJobForm, editEndDate, DateTime.Today.AddYears(1).ToString("MM/dd/yyyy"));
            setLabelCaption(oJobForm, stOrderID, JobOrderID);
            setFormEditVal(oJobForm, editDescription, strJobDesc);
           
           
        }
        catch (Exception ex)
        {

            errorLog(ex);
        }finally
        {
            populateData();
     
        }

    }
    const string cbFrequency = "cbFreq";
    const string F_Onetime = "Onetime";
    const string F_Daily = "Daily";
    const string F_Weekly = "Weekly";
    const string F_Biweekly = "Biweekly";
    const string F_Monthly = "Monthly";
    const string F_Bimonthly = "Bimonthly";
    const string F_Quarterly = "Quarterly";
    const string F_Semiannually = "Semiannually";
    const string F_Annually = "Annually";
    const string editEndDate = "edEndD";
    const string editDescription = "edDesc";
    const string editCustomerID = "edCID";
    const string cbCustomerID = "cbCID";
    const string cbGroup = "cbGrp";
    const string editInvNo = "edInvNo";
    const string stOrderID = "stOID";
    const string stCancel = "stCancel";
    const string LB_CANCELLED = "Cancelled";
    string strJobDesc = "";
    /*
     *  if(freq != F_Onetime && oDoc.GetByKey(int.Parse(INVDocEntry)))
                        {
                           
                        }
     */
    private void AddJobFormField(SAPbouiCOM.Form form)
    {
        try
        {
            int edL = 140; //oItm.Left + oItm.Width;
            int edW = 250;
            int edT = 20;
            int edH = 15;
            int nGap = 20;

            SAPbouiCOM.Item oItm = null;
            if (SODocEntry == "")
            {
                oItm = addPaneItem(form, editCustomerID, edL, edT, edW, edH, "Customer ID:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1100, 100);
                edT = oItm.Top + nGap;
                oItm = addPaneItem(form, cbPaymentMethod, edL, edT, edW, edH, "Payment Method:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 1101, 100);
                edT = oItm.Top + nGap;
                oItm = addPaneItem(form, editInvNo, edL, edT, edW, edH, "Invoice #:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1102, 100);
                edT = oItm.Top + nGap;
                oItm = addPaneItem(form, editCAmount, edL, edT, edW, edH, "Amount:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1103, 100);
                edT = oItm.Top + nGap;
                setFormEditVal(form, editCustomerID, JobCustID);
            }
            oItm = addPaneItem(form, cbFrequency, edL, edT, edW, edH, "Frequency:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 1, 100);
            SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbFrequency).Specific;
            oCB.ValidValues.Add(F_Daily, F_Daily);
            oCB.ValidValues.Add(F_Weekly, F_Weekly);
            oCB.ValidValues.Add(F_Biweekly, F_Biweekly);
            oCB.ValidValues.Add(F_Monthly, F_Monthly);
            oCB.ValidValues.Add(F_Bimonthly, F_Bimonthly);
            oCB.ValidValues.Add(F_Quarterly, F_Quarterly);
            oCB.ValidValues.Add(F_Semiannually, F_Semiannually);
            oCB.ValidValues.Add(F_Annually, F_Annually);
            oCB.Select(F_Monthly);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editDescription, edL, edT, edW, edH, "Description:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 3, 100);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editStartDate, edL, edT, edW, edH, "Start Date:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 4, 100);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editEndDate, edL, edT, edW, edH, "End Date:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 110, 100);
            edT = oItm.Top + nGap;
            if (SODocEntry != "")
            {
                oItm = addPaneItem(form, stOrderID, edL, edT, edW, edH, "Base On Order ID:", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 109, 100);
                edT = oItm.Top + nGap;
            }
            else
            {
                oItm = addPaneItem(form, stOrderID, edL, edT, edW, edH, "", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 109, 100);
                edT = oItm.Top + nGap;

            }
            oItm = addPaneItem(form, stCancel, edL, edT, edW, edH, "", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 111, 100);



        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
    }
    private void JobFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        oJobForm = form;
        BubbleEvent = true;
        try
        {
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        break;
                }
            }
            else
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                       // populateData(form);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                       
                        break;
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:

                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:

                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        HandleJobPress(form, pVal);
                        break;
                }

            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    private void populatePaymentList(SAPbouiCOM.Form form)
    {
        try
        {
            SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbPaymentMethod).Specific;
            try
            {
                while (oCB.ValidValues.Count > 0)
                    oCB.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception)
            { }


            List<CCCUST> q = findCCCust(string.Format(" where \"U_CustomerID\"='{0}' ", JobCustID));
            foreach (CCCUST cust in q)
            {
                try
                {
                    oCB.ValidValues.Add(cust.methodDescription, cust.methodDescription);
                }
                catch (Exception)
                {

                }
               
            }
        }catch(Exception ex)
        {
            errorLog(ex);

        }
    }
 
    private bool ValidateJob(SAPbouiCOM.Form form)
    {
        try
        {
            string errStr = "";
            DateTime dt;
            string sd = getFormItemVal(form, editStartDate);
            if (sd == "")
                errStr += "Please enter a start date.\r\n";
            else
            {
                if (!DateTime.TryParse(sd, out dt))
                    errStr += "Please enter a valid start date.\r\n";
            }
            sd = getFormItemVal(form, editEndDate);
            if (sd != "")
            {
                if (!DateTime.TryParse(sd, out dt))
                    errStr += "Please enter a valid end date or leave it blank for no end date.\r\n";
            }


            if (errStr != "")
            {
                SBO_Application.MessageBox(errStr);
                return false;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    private void HandleJobPress(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {

            switch (pVal.ItemUID)
            {
                case btnUpdate:
                    if (!ValidateJob(form))
                    {
                        return;
                    }
                    HandleUpdate(form);
                    break;
                case btnClose:
                    if (JobID != "" && getFormItemVal(oJobForm, stCancel) != LB_CANCELLED)
                    {
                        if (SBO_Application.MessageBox("Cancel this job?", 1, "Yes", "No") == 1)
                        {
                            CancelJob(JobID);
                        }
                    }
                    form.Close();
                    break;
                

            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void resetGlobal()
    {
     
        JobCustID = "";
        JobID = "";
        confirmNum = "";
        SODocEntry = "";
    }


    private void HandleUpdate(SAPbouiCOM.Form form)
    {
        try
        {
            string msg = "Add eBizCharge recurring billing for this order?";
            string invoiceNum = "";
            string custid = "";
            string amt = "";
            string paymentID = "";
            string oid = getFormItemVal(form, stOrderID);
            custid = getFormItemVal(form, editCustomerID);
            invoiceNum = getFormItemVal(form, editInvNo);
            amt = getFormItemVal(form, editCAmount);
            paymentID = getFormItemVal(form, cbPaymentMethod);
            
            if (JobID == "")  //This is a Add
            {
                if (invoiceNum != "")
                    msg = string.Format("Add term payment on Invoice:{0}?", invoiceNum);
                else
                    msg = string.Format("Add recurring payment on customer account?");
                if (SBO_Application.MessageBox(msg, 1, "Yes", "No") == 1)
                {
                    CCJOB job = new CCJOB();
                    job.CustomerID = custid;
                    if (oid != "")
                    {
                        job.OrderID = int.Parse(oid);
                        
                    }
                   
                    if (invoiceNum == "")
                        job.InvoiceID = null;
                    else
                        job.InvoiceID = int.Parse(invoiceNum);
                   
                    job.PaymentID = paymentID;
                    job.Amount = amt;
                    job.StartDate = DateTime.Parse(getFormItemVal(form, editStartDate));
                    if (getFormItemVal(form, editEndDate) != "")
                    {
                        job.EndDate = DateTime.Parse(getFormItemVal(form, editEndDate));
                    }
                    job.Frequency = getFormItemVal(form, cbFrequency);

                    job.Description = getFormItemVal(form, editDescription);
                    job.NextRunDate = GetNextRunDate(job);
                    

                    insert(job);
           
                    if (confirmNum != "")
                        CCTRANAdjustJobId(confirmNum, job.jobID.ToString());
                 
                    SBO_Application.MessageBox("eBizCharge job added.");
                    form.Close();
                    
                }
            }
            else
            {
                if (SBO_Application.MessageBox("Update job?", 1, "Yes", "No") == 1)
                {   
                    ccjob.StartDate = DateTime.Parse(getFormItemVal(form, editStartDate));
                    if (getFormItemVal(form, editEndDate) != "")
                    {
                        ccjob.EndDate = DateTime.Parse(getFormItemVal(form, editEndDate));
                    }
                    else
                        ccjob.EndDate = null;
                    ccjob.Frequency = getFormItemVal(form, cbFrequency);
                    ccjob.Description = getFormItemVal(form, editDescription);

                    if (invoiceNum == "")
                        ccjob.InvoiceID = null;
                    else
                        ccjob.InvoiceID = int.Parse(invoiceNum);
                    ccjob.Cancelled = 'N';
                    ccjob.CancelledDate = null;
                    ccjob.PaymentID = paymentID;
                    ccjob.Amount = amt;
                    ccjob.NextRunDate = GetNextRunDate(ccjob);
                    update(ccjob);
                    SBO_Application.MessageBox("Recurring billing job updated.");
                    form.Close();
                        
                }
            }
            populateRBillingMatrix();
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    public void ProcessJob()
    {
        try
        {
             setJobCancelledToNULL();
                   
            List<CCJOB> q = findCCJOB(" where \"U_NextRunDate\" <= CURRENT_DATE and (\"U_EndDate\" is NULL OR \"U_EndDate\" > CURRENT_DATE) and \"U_Cancelled\" is NULL "); 
            foreach(CCJOB job in q)
            {
               
                if (job.CustomerID == "")
                    job.CustomerID = getJobCustomerIDFromTrans(job.jobID);
                if (job.OrderID != null && job.OrderID != 0)
                {
                    handleRBillingOrder(job);
                }
                else if (job.InvoiceID != null && job.InvoiceID != 0)
                {
                    handleRBillingInvoice(job);
                }
                else
                    handleRBillingOnAccount(job);

            }
            if(q.Count == 0)
            {
                showMessage("No job to run");
            }
            setJobCancelledToNULL();
        }
        catch(Exception ex)
        {
            errorLog(ex);
        }
    }
     public void handleRBillingOnAccount(CCJOB job)
    {
        string CustNum = "";
        string CCAccountID = "";
        try
        {
 
            SAPCust sapcust = job_getCustomer(job);
            if (sapcust != null)
            {
                string amt = job.Amount;

                if (job_runCustomerTrans(sapcust, null, job, amt))
                {
                    if (job_AddPayment(job, null, sapcust, amt, job.CustomerID))
                    {
                        job_setNextRunDate(job.jobID, GetNextRunDate(job));
                    }
                }
            }
 
        }catch(Exception ex)
        {
            errorLog("Job failed on handleRBillingOnAccount. \r\n" +
               "JobID:" + job.jobID + "\r\n" +
                "CustNum:" + CustNum + "\r\n" +
                "CCAcountID:" + CCAccountID + "\r\n" +
               ex.Message);
        }
    }
    public void handleRBillingInvoice(CCJOB job)
    {
        string CustNum = "";
        string CCAccountID = "";
        try
        {
            int invID = getDocEntry("13", job.InvoiceID.ToString(), job.CustomerID);
            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            if (!oDoc.GetByKey(invID))
            {
                string err = string.Format("Job {0} failed. Invoice ID: {1} not found", job.jobID, job.InvoiceID);
                job.Result = err;
                job.LastRunDate = DateTime.Today;
                update(job);
                errorLog(err);
                return;
            }
            SAPCust sapcust = job_getCustomer(job);
            if (sapcust != null)
            {
                string amt = job.Amount;
                double balance = getBalance(invID.ToString());
                if (balance == 0)
                {
                    job.EndDate = DateTime.Today;
                    job.LastRunDate = DateTime.Today;
                    job.Result = "Invoice closed";
                    update(job);
                    amt = "";
                }
                else if (balance < getDoubleValue(amt))
                    amt = balance.ToString();
                if (amt != "")
                {
                    if (job_runCustomerTrans(sapcust, oDoc, job, amt))
                    {
                        if (job_AddPayment(job, oDoc, sapcust, amt))
                        {
                            job_setNextRunDate(job.jobID, GetNextRunDate(job));
                        }
                    }
                }
            }

        }catch(Exception ex)
        {
            errorLog("Job failed on handleRBillingInvoice. \r\n" +
                "JobID:" + job.jobID + "\r\n" +
                "InvoiceID:" + job.InvoiceID + "\r\n" +
                 "CustNum:" + CustNum + "\r\n" +
                "CCAcountID:" + CCAccountID + "\r\n" +
               ex.Message);
            job.Result = "Failed. " + ex.Message;
            job.EndDate = DateTime.Today;
            update(job);
        }
    }
    public void handleRBillingOrder(CCJOB job)
    {
        string refNum = "";
        int CCAccountID = 0;
        try {
            SAPbobsCOM.Documents oRefDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
            if (!oRefDoc.GetByKey((int)job.OrderID))
            {
                string err = string.Format("Job {0} failed. Order ID: {1} not found", job.jobID, job.OrderID);
                job.Result = err;
                job.LastRunDate = DateTime.Today;
                update(job);
                errorLog(err);
                return;
            }
            int invID = job_createInvoice(oRefDoc, job);
            if (invID != 0)
            {
                SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                oDoc.GetByKey(invID);
               
                job_getInfoformTrans(job, ref refNum);
                SAPCust sapcust = job_getCustomer(job);

                if (sapcust != null)
                {
                    if (job_runCustomerTrans(sapcust, oDoc, job))
                    {
                        if (job_AddPayment(job, oDoc, sapcust))
                        {

                            job_setNextRunDate(job.jobID, GetNextRunDate(job));
                        }
                    }
                }
            }
            
        }catch(Exception ex)
        {
            errorLog("Job failed on handleRBillingOrder. \r\n" +
                "JobID:" + job.jobID + "\r\n" +
                 "OrderID:" + job.OrderID + "\r\n" +
                  "CustNum:" + CustNum + "\r\n" +
                 "CCAcountID:" + CCAccountID + "\r\n" +
                ex.Message);
            job.Result = "Failed. " + ex.Message;
            job.LastRunDate = DateTime.Today;
            update(job);
        }
    }
    public bool job_AddPayment( CCJOB job , SAPbobsCOM.Documents oDoc, SAPCust cust, string amt = "", string customerID = "")
    {
        try
        {
            double dtotal = 0;
            if(oDoc != null)
                dtotal = oDoc.DocTotal;
            if (amt != "")
                dtotal = getDoubleValue(amt);
            SAPbobsCOM.Payments oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            int bplid = getBranchIdFromINV(oDoc.DocEntry.ToString());
            if (bplid != 0)
                oPmt.BPLID = bplid;    
            if (oDoc != null)
            {
                oPmt.Invoices.DocEntry = oDoc.DocEntry;
                oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice;
                oPmt.CardCode = oDoc.CardCode;
                oPmt.Invoices.SumApplied = dtotal;
            }
            else
                oPmt.CardCode = customerID;
            oPmt.DocDate = DateTime.Now;
            oPmt.Remarks = "eBizCharge Payment On Account Recurring Billing " + confirmNum;
            //oPmt.TransferDate = DateTime.Now;
            oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
            AddCreditCardPayment(oPmt, cust, dtotal, true);
            int r = oPmt.Add();
            if (r != 0)
            {
                errorLog("Job Failed on Add Payment. ID:" + job.jobID + "\r\n" +
                    "OrderID:" + job.OrderID + "\r\n" +
                    "CCAcountID:" + cust.cccust.CCAccountID + "\r\n" +
                    getErrorString());
                voidCustomer(confirmNum);
                job.Result = "Failed. " + getErrorString();
                job.LastRunDate = DateTime.Today;
                update(job);
                return false;
            }
            CCTRANUpdateIncomingPaymentID(confirmNum);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            job.Result = "Failed. " +ex.Message;
            job.LastRunDate = DateTime.Today;
            update(job);
            voidCustomer(confirmNum);
            return false;
            
        }
        return true;
    }
    public void job_getInfoformTrans(CCJOB job, ref string refNum)
    {

        string sql = string.Format("select \"U_refNum\" from \"@CCTRANS\" where \"U_jobID\"='{0}' and \"U_OrderID\"='{1}'", job.jobID, job.OrderID);
        try
        {

            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                refNum = (string)oRS.Fields.Item(0).Value;
             
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
  
    }
    public void job_setNextRunDate(int jobid, DateTime nextRunDate)
    {

        string sql = string.Format("update \"@CCJOB\" set \"U_NextRunDate\"={1}, \"U_LastRunDate\" = CURRENT_DATE where \"DocEntry\"='{0}'", jobid, toDateString(nextRunDate));
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
    public bool job_runCustomerTrans(SAPCust sapcust, SAPbobsCOM.Documents oDoc, CCJOB job, string amt ="")
    {
        try
        {
            string cmd = "cc:sale";
            if (sapcust.custObj.PaymentMethodProfiles[0].MethodType == "check")
                cmd = "check";
            confirmNum = "";
           
            CustomerTransactionRequest req = job_createCustomRequest(job, oDoc, amt);
            req.Command = cmd;
            trace("Job_runCustomRequest: amount = " + req.Details.Amount + ", tax=" + req.Details.Tax + ",Subtotal=" + req.Details.Subtotal + ",shipping=" + req.Details.Shipping + ",Discount=" + req.Details.Discount +
             string.Format("runCustomerTrans, token={0}, method={1}:{2}", sapcust.custObj.CustomerToken, sapcust.custObj.PaymentMethodProfiles[0].MethodID, sapcust.custObj.PaymentMethodProfiles[0].CardNumber));
            req.CustReceipt = cfgsendCustomerEmail == "Y" ? true : false;
            req.CustReceiptName = "vterm_customer";
            req.CustReceiptEmail = sapcust.custObj.Email;
            if (req.Command == "cc:authonly" && cfgPreAuthEmail != "Y")
                req.CustReceipt = false;
            if (req.CustReceiptEmail == "")
                req.CustReceipt = false;
            SecurityToken token = getToken(sapcust.cccust.CCAccountID);
            TransactionResponse resp = new TransactionResponse();
            resp = ebiz.runCustomerTransaction(token, sapcust.custObj.CustomerToken, sapcust.custObj.PaymentMethodProfiles[0].MethodID, req);
            confirmNum = resp.RefNum;
            sendReceipt(token, sapcust, resp);
            int invoiceID = 0;
            if (oDoc != null)
                invoiceID = oDoc.DocEntry;
            job_logCustomTransaction(req, resp, sapcust, job, invoiceID.ToString());
            if (resp.Error.ToLower() == "approved")
            {
                job.Result = "Success";
                job.LastRunDate = DateTime.Today;
                update(job);
                return true;
            }
            else
            {
                if (resp.ResultCode == "D")
                {
                    if (req.CustReceiptEmail != "" && req.CustReceiptEmail != null && cfgCustomerReceipt != "")
                    {

                        try
                        {
                            EmailReceiptResponse emresp = ebiz.EmailReceipt(token, resp.RefNum, resp.RefNum, cfgCustomerReceipt, req.CustReceiptEmail);

                            if (emresp.ErrorCode == 0)
                            {
                                showStatus("Declined receipt sent.", SAPbouiCOM.BoMessageTime.bmt_Medium,false);
                            }
                            else
                                showStatus("Failed to send declined receipt.\r\n" + resp.Error, SAPbouiCOM.BoMessageTime.bmt_Medium,true);

                        }
                        catch (Exception ex2)
                        {
                            showStatus("Failed to send receipt.\r\n" + ex2.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                        }

                    }
                }
                string err = "Job Failed. ID:" + job.jobID + "\r\n" + resp.Error;
                job.Result = err;
                job.LastRunDate = DateTime.Today;
                update(job);
                errorLog(err);
                return false;
            }

        }
        catch (Exception ex)
        {
            string err = "Job Failed. ID:" + job.jobID + "\r\n" + ex.Message;
            job.Result = err;
            job.LastRunDate = DateTime.Today;
            update(job);
            errorLog(ex);
           
            return false;
        }

    }
    public int job_createInvoice(SAPbobsCOM.Documents oRefDoc, CCJOB job)
    {
        try
        {
    
            
            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            oDoc.DocNum = getNextTableNum("OINV");
            oDoc.CardCode = oRefDoc.CardCode;
            oDoc.CardName = oRefDoc.CardName;
            oDoc.DocDueDate = DateTime.Now.AddDays(5);
            oDoc.DocDate = DateTime.Now;
            oDoc.Comments = "Created by eBizCharge Job base on order entry: " + oRefDoc.DocEntry;
            SAPbobsCOM.Document_Lines olines = oRefDoc.Lines;
            double total = 0;
            oDoc.DocType = oRefDoc.DocType;
            for (int i = 0; i < olines.Count; i++)
            {
                
                olines.SetCurrentLine(i);
                if (oRefDoc.DocType == SAPbobsCOM.BoDocumentTypes.dDocument_Items)
                {
                    oDoc.Lines.ItemCode = olines.ItemCode;
                    oDoc.Lines.SerialNum = olines.SerialNum;
                    oDoc.Lines.Quantity = olines.Quantity;
                    oDoc.Lines.Price = olines.Price;
                    oDoc.Lines.DiscountPercent = olines.DiscountPercent;

                }
                else
                {
                    oDoc.Lines.AccountCode = olines.AccountCode;

                }
                oDoc.Lines.ItemDescription = olines.ItemDescription;
                oDoc.Lines.TaxCode = olines.TaxCode;
                oDoc.Lines.LineTotal = olines.LineTotal;
                oDoc.Lines.Add();
                total += olines.LineTotal;
            }
            int r = oDoc.Add();
            if (r != 0)
            {
                string err = string.Format("Job create invoice failed. JobID:{0};OrderID:{1}\r\n", job.jobID, job.OrderID) + getErrorString();
                errorLog(err);
                job.Result = err;
                job.LastRunDate = DateTime.Today;
                update(job);
                return 0;
            }

            return getNextTableID("OINV") - 1;
        }
        catch (Exception ex)
        {
            errorLog(ex);
           
        }
        return 0;

    }
    private CustomerTransactionRequest job_createCustomRequest(CCJOB job, SAPbobsCOM.Documents oDoc, string amt ="")
    {


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
            req.Details.AllowPartialAuth = true;
            
            if(amt == "" && oDoc != null)
                req.Details.Amount = oDoc.DocTotal;
            else
                req.Details.Amount = getDoubleValue(amt);
  
            
            req.Details.Discount = 0;
           
            req.Details.Shipping = 0;
            
            req.Details.Tax = 0;
          
            string invID = "RB000";
            if(oDoc!= null)
                invID = oDoc.DocNum.ToString();
            req.Details.Invoice = invID;
            req.Details.OrderID = invID;
            req.Details.PONum = invID;
            req.Details.Description =job.Description;
            req.Details.Subtotal = 0; // req.Details.Amount;
           
            if (oDoc != null)
            {
                SAPbobsCOM.Document_Lines olines = oDoc.Lines;
                req.LineItems = new LineItem[olines.Count];
                for (int i = 0; i < olines.Count; i++)
                {
                    olines.SetCurrentLine(i);
                    req.LineItems[i] = new LineItem();
                    req.LineItems[i].CommodityCode = olines.ItemCode;
                    req.LineItems[i].SKU = olines.ItemCode;
                    req.LineItems[i].Qty = olines.Quantity.ToString();
                    req.LineItems[i].UnitPrice = olines.UnitPrice.ToString();
                    req.LineItems[i].TaxRate = olines.TaxCode;
                    req.LineItems[i].ProductName = olines.ItemDescription;
                    req.LineItems[i].DiscountRate = "";
                    req.LineItems[i].UnitOfMeasure = "EACH";
                }
            }else
            {
                req.LineItems = new LineItem[1];
                req.LineItems[0] = new LineItem();
                req.LineItems[0].CommodityCode = "000";
                req.LineItems[0].Qty = "1";
                req.LineItems[0].UnitPrice = amt;
                req.LineItems[0].TaxRate = "";
                req.LineItems[0].ProductName = "eBizCharge SAP Recurring Billing On Account";
                req.LineItems[0].DiscountRate = "";
                req.LineItems[0].UnitOfMeasure = "EACH";
                req.LineItems[0].SKU = "SKU1";
                
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return req;
    }
    public DateTime GetNextRunDate(CCJOB job)
    {
        try
        {
            if (job.LastRunDate == null)
                job.NextRunDate = job.StartDate;
            else
                job.NextRunDate = job.LastRunDate;
            if (job.NextRunDate == null)
                job.NextRunDate = DateTime.Today;
            switch (job.Frequency)
            {
                case F_Daily:
                    return ((DateTime)(job.NextRunDate)).AddDays(1);
                    
                case F_Weekly:
                    return ((DateTime)(job.NextRunDate)).AddDays(7);
                    
                case F_Biweekly:
                    return ((DateTime)(job.NextRunDate)).AddDays(14);
                    
                case F_Monthly:
                    return ((DateTime)(job.NextRunDate)).AddMonths(1);
                    
                case F_Bimonthly:
                    return ((DateTime)(job.NextRunDate)).AddMonths(2);
                    
                case F_Quarterly:
                    return ((DateTime)(job.NextRunDate)).AddMonths(3);
                    
                case F_Semiannually:
                    return ((DateTime)(job.NextRunDate)).AddMonths(6);
                    
                case F_Annually:
                    return ((DateTime)(job.NextRunDate)).AddYears(1);
                    
                
            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return DateTime.Today;
    }
    private void job_logCustomTransaction(CustomerTransactionRequest req, TransactionResponse resp, SAPCust sapcust, CCJOB job, string InvoiceID = "")
    {
        CCTRAN cctran = new CCTRAN();
         
        try
        {
            if (job.OrderID == null)
                job.OrderID = 0;
            if (job.InvoiceID == null)
                job.InvoiceID = 0;
       
            cctran.customerID = sapcust.cccust.CustomerID;
            cctran.CCAccountID = 0;
            cctran.jobID = (int)job.jobID;
            cctran.OrderID = job.OrderID.ToString();
            cctran.InvoiceID = InvoiceID;
            cctran.CreditMemoID = "";
            cctran.DownPaymentInvoiceID = "";
            if (sapcust != null)
            {
                try
                {
                    cctran.CCAccountID = int.Parse(sapcust.cccust.CCAccountID);
                    cctran.customerID = sapcust.cccust.CustomerID;
                    cctran.MethodID = sapcust.custObj.PaymentMethodProfiles[0].MethodID;
                    cctran.custNum = sapcust.cccust.CustNum;
                    cctran.crCardNum = sapcust.custObj.PaymentMethodProfiles[0].CardNumber;
                    cctran.CardHolder = sapcust.custObj.BillingAddress.FirstName + " " + sapcust.custObj.BillingAddress.LastName;
                }catch(Exception ex)
                {
                    errorLog(ex);
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
    public SAPCust job_getCustomer(CCJOB job)
    {
        
        SAPCust sapcust = new SAPCust();
        try
        {
                 CCCUST c = null;
                List<CCCUST> q = findCCCust(string.Format(" where \"U_CustomerID\"='{0}' AND \"U_methodDescription\"='{1}'",job.CustomerID, job.PaymentID ));
                if(q.Count() == 0)
                {
                    List<CCCUST> q2 = findCCCust(string.Format(" where \"U_CustomerID\"='{0}'", job.CustomerID));
                    foreach (CCCUST c1 in q2)
                        c = c1;

                }
                else
                {
                    foreach (CCCUST c1 in q)
                        c = c1;
                }
                if(c == null)
                {
                    string err = "No payment method found";
                    job.Result = err;
                    job.LastRunDate = DateTime.Today;
                    update(job);
                    return null;
                }
                sapcust.cccust = c;
            string id = "";
                string cardcode = getCardAcctByCustNum(c.CustNum, ref id);
                SecurityToken token = getToken(cardcode);
                sapcust.custObj = ebiz.GetCustomer(token, c.CustomerID, null);
                return sapcust;
            

        }
        catch (Exception ex)
        {
            errorLog("Failed to Create Customer Obj." +
                "\r\nJobID:" + job.jobID +
                "\r\n" + ex.Message);
            job.Result = "Failed on get payment method";
            job.LastRunDate = DateTime.Today;
            update(job);
            
        }
        return null;
    }
}


