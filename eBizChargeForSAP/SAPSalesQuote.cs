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

partial class SAP
{
   

    private void SalesQuoteFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
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
                       
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                      
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
                        HandleSQPress(form, pVal);
                        break;
                }

            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }

   

    
    private void HandleSQPress(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {

            switch (pVal.ItemUID)
            {
                case tbCreditCard:
                    form.PaneLevel = paneCCLog;
                    SQDocEntry = form.DataSources.DBDataSources.Item("OQUT").GetValue("DocEntry", 0);
                    if (SQDocEntry != "")
                        strWhere = string.Format(" Where \"U_customerID\"='{0}' and \"U_QuoteID\"='{1}' ", getFormItemVal(form, fidCustID), SQDocEntry);
                    else
                        strWhere = string.Format(" Where 1=0 ", getFormItemVal(form, fidCustID));

                    populateLog(form);
                    break;
                case btnCharge:
                    SQDocEntry = form.DataSources.DBDataSources.Item("OQUT").GetValue("DocEntry", 0);
                    if (SQDocEntry != "")
                    {
                        SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations);
                        oDoc.GetByKey(int.Parse(SQDocEntry));
                        if (oDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Close)
                        {
                            showMessage("Quotation is closed");
                            bAuto = true;
                            tProcess = new Thread(RecordRefreshThreadProc);
                            tProcess.Start();
                            return;
                        }
                    }
                     if(!ItemValidate(form))
                        return;
                     theActiveForm = form;
                    CreateCreditCardForm();
                    break;

            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    //   bool bProcessing = false;
   

}

