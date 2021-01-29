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
using System.Collections.Generic;
using eBizChargeForSAP.ServiceReference1;


partial class SAP
{
    
    public void AddDocThreadProc()
    {
        try
        {

            bAuto = true;
            string custID = getCustomerID();
            trace("CustomerID=" + custID);
            string docID = getFormItemVal(theActiveForm, fidRecordID);
            theActiveForm.Items.Item("1").Click();
          
            for (int i = 0; i <= 3; i++)
            {
                if (getCustomerID() == "")
                {
                    docID = getDocNumAdded(custID, theActiveForm.TypeEx);
                    SBO_Application.Menus.Item("1281").Activate();
                    setFormEditVal(theActiveForm, fidRecordID, docID);
                    theActiveForm.Items.Item("1").Click();
                    return;
                }
                Thread.Sleep(100);
            }

        }
        catch (Exception)
        {
            //errorLog(ex);
        }finally
        {
            oFormEvent.Set();
        }
    }

    private void SalesOrderFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;
        
        try
        {
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        AddTab(form, paneCCLog, "Credit Card Log", "114");
                        AddCCLogMatrix(form, paneCCLog);
                        FormSetup(form);
                        theActiveForm = form;
                        break;
                  
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                       
                        break;
                }
            }
            else
            {
                //errorLog(pVal.EventType.ToString());
                switch (pVal.EventType)
                {

                    case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        /*
                        SAPbouiCOM.Folder f1 = (SAPbouiCOM.Folder)form.Items.Item(tbCreditCard).Specific;
                        if (f1.Selected)
                        {
                            SODocEntry = form.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);
                            if (SODocEntry != "")
                                strWhere = string.Format(" Where CustomerID='{0}' and OrderID={1} ", getFormItemVal(form, fidCustID), SODocEntry);
                            else
                                strWhere = string.Format(" Where CustomerID='{0}'", getFormItemVal(form, fidCustID));

                            populateLog(form);
                        }
                        */
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        SODocEntry = form.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);
                       // handleDataChangeInit(form);

                         break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                         if (pVal.ItemUID == mxtCCLog)
                        {
                            HandleCCLogSelect(form);
                        }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        HandleSOPress(form, pVal);
                        break;
                }

            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }

    }

    private void HandleCCLogSelect(SAPbouiCOM.Form form)
    {
        try
        {
            string CustomerID = getCustomerID(form);

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item(mxtCCLog).Specific;
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                if (oMatrix.IsRowSelected(i))
                {
                    string id = getMatrixItem(oMatrix, "ccTRANSID", i);
                    string cmd =  getMatrixItem(oMatrix, "command", i);
                    string refnum = getMatrixItem(oMatrix, "RefNum", i);
                    string SOID = getMatrixItem(oMatrix, "OrderID", i);
                    string custNum = getMatrixItem(oMatrix, "custNum", i);

                    if (cmd == "cc:authonly" && form.TypeEx == FORMSALESORDER)
                    {
                        bAuto = false;
                       if (SBO_Application.MessageBox("Void preauth transaction " + refnum + "?", 1, "Yes", "No") == 1)
                       {       
                            if (voidTransaction(refnum))
                            {
                                voidUpdateTrans(id);
                                populateLog(form);
                            }
                        }
                    }

                    if (cmd == "cc:authonly" || cmd == "cc:sale" || cmd == "check")
                    {
                        bAuto = false;
                        SAPCust sapcust = getCustomerByID(CustomerID, refnum);
                        string email = getFormItemVal(form, editEmail);
                        if (email == "")
                            email = getCustEmail(CustomerID);
                        if (email != "" && email != null)
                        {
                            if (SBO_Application.MessageBox("Send email receipt to " + email + "?", 1, "Yes", "No") == 1)
                            {
                                try
                                {
                                    SecurityToken token = getToken(sapcust.cccust.CCAccountID);
                                    EmailReceiptResponse resp = ebiz.EmailReceipt(token, refnum,refnum, "vterm_customer", email);

                                    if (resp.ErrorCode == 0)
                                    {
                                        SBO_Application.MessageBox("Receipt sent.");
                                    }
                                    else
                                        SBO_Application.MessageBox("Failed to send receipt.\r\n" + resp.Error);
                                }
                                catch (Exception ex2)
                                {
                                    SBO_Application.MessageBox("Failed to send receipt.\r\n" + ex2.Message);
                                }
                            }
                        }
                    }
                }
            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
    
    const string BTN_EBIZCHARGE = "eBizCharge";
    private void FormSetup(SAPbouiCOM.Form form)
    {
        try
        {
            if (form.Items.Item("1").Enabled == true)
            {
                SAPbouiCOM.Item oRef = form.Items.Item("2");
                form.Items.Add(btnCharge, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                SAPbouiCOM.Item oItm = form.Items.Item(btnCharge);

                oItm.Left = oRef.Left + int.Parse(cfgBtnPos);
                oItm.Top = oRef.Top;
                oItm.Width = oRef.Width + 20;
                oItm.Height = oRef.Height;
                oItm.AffectsFormMode = false;

                SAPbouiCOM.Button oBt = (SAPbouiCOM.Button)form.Items.Item(btnCharge).Specific;
                oBt.Caption = BTN_EBIZCHARGE;
            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
    private void HandleSOPress(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {

            switch (pVal.ItemUID)
            {

                case "1":
                    CCTRANASetOrderIDwithQuoteID();
                    break;
                  
         
                case tbCreditCard:
                    
                    form.PaneLevel = paneCCLog;
                    string OID = form.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);
                    string CustomerID = getCustomerID(form);
                    updateTransCustomerID(OID, CustomerID);
                    if (OID != "")
                        strWhere = string.Format(" Where \"U_customerID\"='{0}' and \"U_OrderID\"='{1}' ",CustomerID, OID);
                    else
                        strWhere = " Where 1=0 ";
      
                    populateLog(form);
                    break;
                case btnCharge:
                    string title = form.Title;
                  
                    if (title.ToLower().IndexOf("draft") == -1)
                    {
                        SODocEntry = form.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);
                        if (SODocEntry != "")
                        {
                            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                            oDoc.GetByKey(int.Parse(SODocEntry));
                            if (oDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Close)
                            {
                                SBO_Application.MessageBox("Order is closed");
                                bAuto = true;
                                tProcess = new Thread(RecordRefreshThreadProc);
                                tProcess.Start();
                                return;
                            }
                        }
                    }
                    if(!ItemValidate(form))
                        return;
                    theActiveForm = form;
                  
                    if (cfgAutoPreauth == "Y")
                    {
                        string customerid = getCustomerID();
                        List<SAPCust> list = getCustomerListByID(customerid);
                        if (list.Count() == 0)
                        {
                            if (SBO_Application.Menus.Item("5892").Enabled == false)
                            {
                                CreateCreditCardForm();
                            }
                            else
                            {
                                //CreateCreditCardForm();
                                SBO_Application.Menus.Item("5892").Activate();
                            }
                        }
                        else
                        {
                            string docentry = form.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);
                            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                            if (!oDoc.GetByKey(int.Parse(docentry)))
                            {
                                showMessage("Cannot retrieve document. DocEntry=" + docentry);
                                return;
                            }
                            if (oDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Close)
                            {
                                showMessage("Order close.");
                                return;
                            }

                            if (isOrderPreauthed(docentry.ToString()))
                            {
                                showMessage("Order already pre-authed.");
                                return;
                            }
                            int i = 0;
                            foreach (SAPCust sapcust in list)
                            {
                                i++;
                                string s = string.Format("Preauth {0} on card {1}/{3}/{4} {5}?  Customer has {2} card(s) on file. {6} 0f {2}", oDoc.DocTotal, sapcust.cccust.cardNum, list.Count()
                                    , sapcust.cccust.expDate, sapcust.cccust.firstName, sapcust.cccust.lastName, i);
                                bAuto = false;
                                if (SBO_Application.MessageBox(s, 1, "Yes", "No") == 1)
                                {
                                    try
                                    {
                                        form.Freeze(true);
                                        BatchPreauth_PreAuthCustomer(oDoc, sapcust);
                                    }
                                    catch (Exception)
                                    {
                                    }
                                    finally
                                    {
                                        form.Freeze(false);
                                    }
                                    return;
                                }
                            }
                            if (SBO_Application.Menus.Item("5892").Enabled == false)
                            {
                                CreateCreditCardForm();
                            }
                            else
                            {
                                //CreateCreditCardForm();
                                SBO_Application.Menus.Item("5892").Activate();
                            }
                        }

                    }
                    else
                    {
                        if (SBO_Application.Menus.Item("5892").Enabled == false)
                        {
                            CreateCreditCardForm();
                        }
                        else
                        {
                            //CreateCreditCardForm();
                            SBO_Application.Menus.Item("5892").Activate();
                        }
                    }
                    break;
                  
            }

        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
 //   bool bProcessing = false;
    private bool ItemValidate(SAPbouiCOM.Form form)
    {
        beBizChargeClicked = true;
        bUsePaymentMean = false;
        bool bRet = true;
        string err = "";
        try
        {
            string cid = getCustomerID();
            updatePaymentTerm(cid);
        }
        catch (Exception)
        { }
        try
        {
           
            if (form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                return true;
           
            if(form.TypeEx == FORMSALESORDER)
            {
                if(getFormItemVal(form, "12") == "")
                {
                    err = err + "Delivery date is required\r\n";
                }
               // int id = getNextTableNum("ORDR");
               // setFormEditVal(form, fidRecordID, id.ToString());
            }

            if(err != "")
            {
                SBO_Application.MessageBox(err);
                return false;
            }
            tProcess = new Thread(AddDocThreadProc);
            tProcess.Start();
            oFormEvent.WaitOne();
            oFormEvent.Reset();
            if (theActiveForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                return false;
            else
                return true;
        }
        catch(Exception ex)
        {
            errorLog(ex);
        }
        return true;

    }

}

