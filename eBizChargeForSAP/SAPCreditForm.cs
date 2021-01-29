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


partial class SAP
{
    const string formCredit = "eBizCForm";
    const string formCfg = "eBCFGForm";
    SAPbouiCOM.Form oCreditForm;
    SAPbouiCOM.Form oCfgForm;
    string CFCustID;
    List<SAPCust> CFCustList = new List<SAPCust>();
    SAPCust CFSAPCust;
    bool bProcessing = false;
    bool bCreated = false;
    bool bLoaded = false;
    SAPbouiCOM.Form CFTargetedForm;
    public void ThreadReload()
    {
        try
        {
            oCreditForm.Freeze(true);
            bManualPopulate = false;
            reload(oCreditForm,  ref CFCustList, ref CFSAPCust);
            setFormEditVal(oCreditForm, "11", getFormItemVal(theActiveForm, "29"));
            double discRate = getCashDiscountRate(CFCustID);
            setFormEditVal(oCreditForm, "111", discRate.ToString());

            double bal = getItemDoubleValue(theActiveForm, "33");
            double disc = bal * discRate/100;
            setFormEditVal(oCreditForm, "110", string.Format("{0:C}", disc));

            double amount = bal - disc;
         
            if (getFormItemVal(oCreditForm, cbTransactionType) == "Credit")
            {
                setFormEditVal(oCreditForm, "14", getFormItemVal(theActiveForm, "33"));
                setFormEditVal(oCreditForm, editCAmount, getFormItemVal(theActiveForm, "33"));
            }else
            {
                setFormEditVal(oCreditForm, "14", string.Format("{0:C}", amount));
                setFormEditVal(oCreditForm, editCAmount, string.Format("{0:C}", amount));
            }


            /*
            string MethodID = getMethodID();
                            
                       
            if (MethodID != null)
            {
                SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)oCreditForm.Items.Item(cbPaymentMethod).Specific;
                foreach (SAPCust sapcust in CFCustList)
                {
                    if (sapcust.cccust.MethodID == MethodID)
                    {
                        oCB.Select(sapcust.key);
                        CFSAPCust = sapcust;
                        populateCCInfo(oCreditForm, CFSAPCust);
                    }
                }
            }
            */
            setBalVal(oCreditForm);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }finally
        {
            try
            {
                oCreditForm.Freeze(false);
            }catch(Exception)
            { }
        }
    }
    private void CreateCreditCardForm()
    {
        try
        {
            beBizChargeClicked = false;
            CFCustID = getCustomerID();
            if (CFCustID == "")
            {
                SBO_Application.MessageBox("Please select a customer first");
                return;
            }
            try
            {
                oCreditForm = SBO_Application.Forms.Item(formCredit);
                oCreditForm.Close();
            }catch(Exception)
            { }

            // add a new form
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Button oButton = null;

            // add a new form
            SAPbouiCOM.FormCreationParams oCreationParams = null;

            oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.FormType = formCredit;
            oCreationParams.UniqueID = formCredit;
            try
            {
                oCreditForm = SBO_Application.Forms.AddEx(oCreationParams);
            }
            catch (Exception ex)
            {
                errorLog(ex);
                oCreditForm = SBO_Application.Forms.Item(formCredit);

            }
            SAPbouiCOM.Item oRefItm = null;
            SAPbouiCOM.Item oRefBtnItm = null;
            if (theActiveForm.TypeEx != FORMINCOMINGPAYMENT)
                oRefItm = theActiveForm.Items.Item("4");
            else
                oRefItm = theActiveForm.Items.Item("5");
            oRefBtnItm = theActiveForm.Items.Item("1");
            // set the form properties
            oCreditForm.Title = "eBizCharge Credit Card Form";
            oCreditForm.Left = 400;
            oCreditForm.Top = 100;
            oCreditForm.ClientHeight = oRefItm.Height * 35;
            oCreditForm.ClientWidth = oRefItm.Width * 4;

            int buttonTop = oCreditForm.ClientHeight - oRefItm.Height * 2;
            oItem = oCreditForm.Items.Add(btnProcess, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 6;
            oItem.Width = oRefBtnItm.Width;
            oItem.Top = buttonTop;
            oItem.Height = oRefBtnItm.Height;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Credit";

            //************************
            // Adding a Cancel button
            //***********************

            oItem = oCreditForm.Items.Add(btnClose, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 6 + oRefBtnItm.Width + 2;
            oItem.Width = oRefBtnItm.Width;
            oItem.Top = buttonTop;
            oItem.Height = oRefBtnItm.Height;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";
            if (cfgHasTerminal == "Y")
            {
                oItem = oCreditForm.Items.Add(btnTerminal, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = oCreditForm.Items.Item(btnClose).Left + oCreditForm.Items.Item(btnClose).Width + 10;
                oItem.Width = oCreditForm.Items.Item(btnClose).Width;
                oItem.Top = oCreditForm.Items.Item(btnClose).Top;
                oItem.Height = oCreditForm.Items.Item(btnClose).Height;
                setButtonCaption(oCreditForm, btnTerminal, "Terminal");
            }
            AddCCPMFormField(oCreditForm, panePM, true);
            setLabelCaption(oCreditForm, stRecordID, getFormItemVal(theActiveForm, fidRecordID));
            oCreditForm.Visible = true;

        }
        catch (Exception ex)
        {

            errorLog(ex);
        }

    }
    private void CreditFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
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
                    case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
                       
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        CFTargetedForm = theActiveForm;
                        form.PaneLevel = panePM;
                        bProcessing = false;
                        bCreated = true;
                        bAuto = false;
                        bLoaded = false;
                        //ThreadReload();
                       // tProcess = new Thread(ThreadReload);
                       // tProcess.Start();
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE:
                        if (!bLoaded)
                        {
                            bLoaded = true;
                            tProcess = new Thread(ThreadReload);
                            tProcess.Start();
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (pVal.ItemChanged == true)
                        {
                            switch (pVal.ItemUID)
                            {
                                case editCardNum:
                                    setCardName(form, CFSAPCust);
                                    break;
                                case "110":
                                    setFormEditVal(form, editCAmount, getFormItemVal(form, "14"));
                                    break;
                                case "111":
                                    try
                                    {
                                        string oamt = getFormItemVal(form, "11");
                                        string disc = getFormItemVal(form, "111");
                                        double bal = Math.Round(getDoubleValue(oamt) * (100 - getDoubleValue(disc)) / 100, 2);
                                        double discamt = Math.Round(getDoubleValue(oamt) * getDoubleValue(disc) / 100, 2);
                                        setFormEditVal(form, "110", discamt.ToString());
                                        setFormEditVal(form, "14", bal.ToString());
                                        setFormEditVal(form, editCAmount, bal.ToString());
                                    }
                                    catch (Exception ex) { errorLog(ex); }
                                    break;

                            }

                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:

                        if ((pVal.CharPressed == 13 || pVal.CharPressed == 10 || pVal.CharPressed == 9) && pVal.ItemUID == editCardNum)
                        {
                            try
                            {
                                handleSwipe(form);
                            }
                            catch (Exception e)
                            {
                                errorLog(e);
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                    {
                        switch (pVal.ItemUID)
                        {
                            case cbPaymentMethod:
                                /*
                                CFSAPCust = getCCInfoSelect(form, CFCustList);
                                populateCCInfo(form, CFSAPCust);
                                */
                                HandleSelect(form);
                                break;
                            case cbTransactionType:
                                setButtonCaption(form,btnProcess,getFormItemVal(form,cbTransactionType));
                                    
                                  break;
                   
                        }
                    }
                    break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        theActiveForm = CFTargetedForm;
                        bProcessing = true;
                        HandleCFPress(form, pVal);
                        bProcessing = false;
                        break;
                }

            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    SAPbouiCOM.Form formCM;
    private void HandleCFPress(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {

            switch (pVal.ItemUID)
            {
                case btnTerminal:
                    TermAmount = getFormItemVal(form, editCAmount);
                    TermFunc = "Charge";
                    if (getFormItemVal(form, btnProcess) == "Credit")
                        TermFunc = "Credit";
                    TermGroup = getGroupName(getCustomerID());
                    TermInvoice = getInvoiceNum();
                    TermCustomer = getCustomerID();
                    oTerminalLaunchForm = form;
                    CreateTerminalForm();
                    break;
                
                case btnProcess:
                    formCM = form;
                    if (!ValidateChargeForm(form))
                        return;
                    tProcess = new Thread(ProcessCardCM);
                    tProcess.Start();
                    break;
                   case btnClose:
                            oCreditForm.Close();
                            break;
                    }

            
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void ProcessCardCM()
    {
        SAPbouiCOM.Form form = formCM;
        SAPbobsCOM.Documents oSODoc = null;
        try
        {
            try { form.Items.Item(btnProcess).Enabled = false; }
            catch (Exception) { }
            switch (getFormItemVal(form, cbTransactionType))
            {
                case "No Trans":
                    {


                        if (!loadFormData(form, ref CFSAPCust, true))
                            return;
                        string s = "Create payment record for " + getFormItemVal(oCreditForm, editCAmount) + "on " + CFSAPCust.paymentDate.ToString("MM/dd/yyyy") + "? (Transaction will not run)";
                        if (SBO_Application.MessageBox(s, 1, "Yes", "No") == 1)
                        {
                            if (theActiveForm.TypeEx == FORMCREDITMEMO)
                            {
                                if (createCreditMemo(theActiveForm, CFSAPCust))
                                {
                                    double amt = getItemDoubleValue(form, editCAmount);

                                    if (AddOutgoingPayment(theActiveForm, amt, CFSAPCust))
                                    {
                                        form.Close();
                                        bAuto = true;
                                        tProcess = new Thread(RecordRefreshThreadProc);
                                        tProcess.Start();
                                    }
                                    else
                                    {
                                        form.Close();
                                    }

                                }
                            }
                            else
                            {
                                handleCharge(form, CFSAPCust, true);
                                form.Close();
                            }
                        }
                    }
                    break;
                case "Credit":
                    {
                        if (theActiveForm.TypeEx != "179")
                        {
                            errorLog("Credit error. Target form not a credit memo. form type:" + theActiveForm.TypeEx);
                            form.Close();
                            return;
                        }


                        if (!loadFormData(form, ref CFSAPCust, true))
                            return;
                        string s = "Credit " + getFormItemVal(oCreditForm, editCAmount) + " to card " + getFormItemVal(oCreditForm, editCardNum) + "?";
                        if (CFSAPCust.custObj.PaymentMethodProfiles[0].MethodType == "check")
                            s = "Credit " + getFormItemVal(form, editCAmount) + " to checking account " + getFormItemVal(form, editCheckAccount);

                        if (SBO_Application.MessageBox(s, 1, "Yes", "No") == 1)
                        {
                            if (CreditCustomer(oCreditForm, CFSAPCust))
                            {
                                if (createCreditMemo(theActiveForm, CFSAPCust))
                                {
                                    double amt = getItemDoubleValue(form, editCAmount);

                                    if (AddOutgoingPayment(theActiveForm, amt, CFSAPCust))
                                    {
                                        updateCCTRAN(SAPbobsCOM.BoObjectTypes.oCreditNotes, CFSAPCust.cccust.CCAccountID);
                                        form.Close();
                                        bAuto = true;
                                        tProcess = new Thread(RecordRefreshThreadProc);
                                        tProcess.Start();
                                    }
                                    else
                                    {

                                        form.Close();
                                    }

                                }
                                else
                                {
                                    voidCustomer(confirmNum);
                                    form.Close();
                                }


                            }
                        }
                    }

                    break;
                case "Charge":
                    {

                        bAuto = false;
                        if (!loadFormData(form, ref CFSAPCust, true))
                            return;
                        if (CFSAPCust.cccust.CustomerID == null)
                        {
                            CFSAPCust.cccust.CustomerID = CFCustID;
                            CFSAPCust.custObj.CustomerId = CFCustID;

                        }

                        if (!Validate(form, CFSAPCust.custObj))
                            return;
                        string s = "Charge " + getFormItemVal(form, editCAmount) + " to card " + getFormItemVal(form, editCardNum) + "?";
                        if (CFSAPCust.custObj.PaymentMethodProfiles[0].MethodType == "check")
                            s = "Charge " + getFormItemVal(form, editCAmount) + " to checking account " + getFormItemVal(form, editCheckAccount);
                        double amt = getItemDoubleValue(form, editCAmount);
                        if (amt < 0)
                            s = s.Replace("Charge ", "Refund ");
                        if (SBO_Application.MessageBox(s, 1, "Yes", "No") == 1)
                        {
                            showStatus("Processing charge.  Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                            handleCharge(form, CFSAPCust);
                        }
                    }
                    break;
                case "Pre-auth":
                    {
                        if (theActiveForm.TypeEx == FORMSALESORDER)
                            SODocEntry = theActiveForm.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);

                        bAuto = false;
                        if (!loadFormData(form, ref CFSAPCust, true))
                            return;
                        if (CFSAPCust.custObj.CustomerId == null || CFSAPCust.cccust.CustomerID == null)
                        {
                            if (CFCustID == null)
                                CFCustID = getCustomerID();
                            CFSAPCust.custObj.CustomerId = CFCustID;
                            CFSAPCust.cccust.CustomerID = CFCustID;
                        }
                        if (!Validate(form, CFSAPCust.custObj))
                            return;



                        if (CFSAPCust.custObj.PaymentMethodProfiles[0].MethodType == "check")
                        {
                            SBO_Application.MessageBox("eCheck does not support pre-auth.  Please select other payment method.");
                            return;
                        }

                        strTransDesc = "Pre-auth " + getFormItemVal(form, editCAmount) + " to card " + getFormItemVal(form, editCardNum);


                        if (SBO_Application.MessageBox(strTransDesc + "?", 1, "Yes", "No") == 1)
                        {

                            if (PreAuthCustomer(form, CFSAPCust))
                            {
                                if (theActiveForm.TypeEx == FORMSALESORDER)
                                {
                                    oSODoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                                    if (createSO(theActiveForm, CFSAPCust, ref oSODoc))
                                    {
                                        dChargeAmount = getItemDoubleValue(form, editCAmount);
                                        if (getFormItemVal(form, cbRecurringBilling) == "Y")
                                            bRecurringBilling = true;
                                        else
                                            bRecurringBilling = false;
                                        updateCCTRAN(SAPbobsCOM.BoObjectTypes.oOrders, CFSAPCust.cccust.CCAccountID);
                                        bAuto = true;
                                        form.Close();
                                        
                                    }
                                }
                                if (theActiveForm.TypeEx == FORMSALESQUOTE)
                                {
                                    oSQDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations);
                                    SQDocEntry = theActiveForm.DataSources.DBDataSources.Item("OQUT").GetValue("DocEntry", 0);
                                    if (oSQDoc.GetByKey(int.Parse(SQDocEntry)))
                                    {
                                        dChargeAmount = getItemDoubleValue(form, editCAmount);
                                        CCTRANAdjustWithSQId(confirmNum, SQDocEntry, CFSAPCust.cccust.CustomerID, dChargeAmount.ToString());
                                        bAuto = true;
                                        form.Close();
                                        tProcess = new Thread(RecordRefreshThreadProc);
                                        tProcess.Start();
                                    }
                                }
                            }
                        }
                    }
                    break;
                case "Capture":
                    HandleCapture(form, ref CFSAPCust);
                    break;
             }//end switch
        }
        catch(Exception )
        {

        }
        finally
        {
            try {
                if(oSODoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSODoc);
                    oSODoc = null;
                }
                form.Items.Item(btnProcess).Enabled = true; }
            catch (Exception) { }
        }
    }
    public void setBalVal(SAPbouiCOM.Form form)
    {
        try
        {
            if (theActiveForm.TypeEx == FORMSALESORDER && SODocEntry != "")
            {
                SODocEntry = theActiveForm.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);
                double transamt = getTransAmount(SODocEntry);
                double bal = getItemDoubleValue(form, editCAmount) - transamt;
                setFormEditVal(form, editCAmount, bal.ToString());
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }

}


