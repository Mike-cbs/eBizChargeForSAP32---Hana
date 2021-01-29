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



partial class SAP
{
  
    string strWhere;
    //   const string strBaseOnSO = "Based On Sales Orders";
    //  const string strBaseOnDelivery = "Based On Deliveries";
    public void UpdateTransThreadProc()
    {
        try
        {
            if (cfgCaptureOnPrint == "Y")
            {
               captureOnAdd();
               // RecordRefreshThreadProc();              
            }
        }
        catch (Exception)
        {
            //errorLog(ex);
        }
        finally
        {
           
        }
    }
    private void InvoiceFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
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
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        /*
                        SAPbouiCOM.Folder f1 = (SAPbouiCOM.Folder)form.Items.Item(tbCreditCard).Specific;
                        if (f1.Selected)
                        {
                            string id = getFormItemVal(form, fidRecordID);
                            if (form.TypeEx == "65300")
                            {
                                DPDocEntry = form.DataSources.DBDataSources.Item("ODPI").GetValue("DocEntry", 0);
                                if (DPDocEntry != "")
                                    strWhere = string.Format(" Where CustomerID='{0}' and DownPaymentInvoiceID={1} ", getFormItemVal(form, fidCustID), DPDocEntry);
                                else
                                    strWhere = string.Format(" Where CustomerID='{0}' ", getFormItemVal(form, fidCustID));

                            }
                            else
                            {
                                
                                INVDocEntry = form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0);

                                if (getFormItemVal(form, "1").ToLower() != "add")
                                    strWhere = string.Format(" Where CustomerID='{0}' and InvoiceID={1} ", getFormItemVal(form, fidCustID), INVDocEntry);
                                else
                                    strWhere = string.Format(" Where CustomerID='{0}' ", getFormItemVal(form, fidCustID));

                            }
                            populateLog(form);
                        }
                        */
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        //SetInvoiceID(form);
                      
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
                        HandleInvPress(form, pVal);
                        break;
                }

            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }

    }
    private void HandleInvPress(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {

            switch (pVal.ItemUID)
            {
                case tbCreditCard:
                  
                    form.PaneLevel = paneCCLog;
                    string id = getFormItemVal(form, fidRecordID);
                    if (form.TypeEx == "65300")
                    {
                        DPDocEntry = form.DataSources.DBDataSources.Item("ODPI").GetValue("DocEntry", 0);
                        if (DPDocEntry != "")
                            strWhere = string.Format(" Where \"U_customerID\"='{0}' and \"U_DownPaymentInvID\"='{1}' ", getFormItemVal(form, fidCustID), DPDocEntry);
                        else
                            strWhere = " Where 0 = 1";
 
                    }
                    else
                    {
                   
                        string invID = form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0);
                        string customerID = getCustomerID(form);
                        if(customerID != "" && invID != "")
                            updateTransInvoiceID(invID, customerID);
                        if (getFormItemVal(form, "1").ToLower() != "add")
                            strWhere = string.Format(" Where \"U_customerID\"='{0}' and \"U_InvoiceID\"='{1}' ", customerID, invID);
                        else
                            strWhere = " Where 0 = 1";

                    }
                    populateLog(form);
                    break;
                case btnCharge:
                    try
                    {
                        string invID = form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0);
                        string customerID = getCustomerID(form);
                        if (customerID != "" && invID != "")
                            updateTransInvoiceID(invID, customerID);
                    }
                    catch (Exception) { }
                    if (!ItemValidate(form))
                        return;
                    if (SBO_Application.Menus.Item("5892").Enabled == false)
                    {
                        if (getFormItemVal(form, "33") != "")
                        {
                            theActiveForm = form;
                            //SetInvoiceID(form);
                            CreateCreditCardForm();
                        }
                    }
                    else
                    {
                        if (getFormItemVal(form, "33") != "" && getFormItemVal(form, "31") != "")
                        {
                            theActiveForm = form;
                            // SetInvoiceID(form);
                            CreateCreditCardForm();
                        }
                        else
                        {
                            bUseeBizCharge = true;
                            SBO_Application.Menus.Item("5892").Activate();
                        }
                    }
                    break;
                case "1":
                    if(form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        tProcess = new Thread(UpdateTransThreadProc);
                        tProcess.Start();
                    }
                   
                    break;
           }

        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }

}


