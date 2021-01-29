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
    const string strBaseOnARInvoice = "Based On A/R Invoices";
    const string strBaseOnARDownPayment = "Based On A/R Down Payment";
    string CustNum = "";
    //string MethodID = "";
    string CardNum = "";
    string CardHolder = "";
    string CMDocNum = "";
    string CMDocDate = "";
    string CustomerID = "";
    string CCAccountID = "";
    string MethodType = "";
   
    private void CreditMemoFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        AddTab(form, paneCCLog, "Credit Card Log", "114");
                        AddCCLogMatrix(form, paneCCLog);
                        CustNum = "";
                       // MethodID = "";
                        CardNum = "";
                        CardHolder = "";
                        
                        CustomerID = "";
                        CCAccountID = "";
                        theActiveForm = form;
               
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if(pVal.ItemUID == "1")
                        {
                            CMDocNum = getFormItemVal(form,fidRecordID);
                            CMDocDate = getFormItemVal(form, "46");
                        }
                        break;
                }
            }
            else
            {
                switch (pVal.EventType)
                {
                    
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        /*
                        SAPbouiCOM.Folder f1 = (SAPbouiCOM.Folder)form.Items.Item(tbCreditCard).Specific;
                        if (f1.Selected)
                        {
                            CMDocEntry = form.DataSources.DBDataSources.Item("ORIN").GetValue("DocEntry", 0);
                            if (CMDocEntry != "")
                                strWhere = string.Format(" Where CustomerID='{0}' and CreditMemoID={1} ", getFormItemVal(form, fidCustID), CMDocEntry);
                            else
                                strWhere = string.Format(" Where CustomerID='{0}'", getFormItemVal(form, fidCustID));
      
                            populateLog(form);
                        }
                        */
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                       
                        if(CheckStatus(form))
                            FormSetup(form);
                        break;
                   
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == mxtCCLog)
                        {
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item(mxtCCLog).Specific;
                            for (int i = 1; i < oMatrix.RowCount; i++)
                            {
                                if (oMatrix.IsRowSelected(i))
                                {
                                    CustNum = getMatrixItem(oMatrix, "custNum", i);
                                   // MethodID = getMatrixItem(oMatrix, "MethodID", i);
                                    CardNum = getMatrixItem(oMatrix, "crCardNum", i);
                                    CardHolder = getMatrixItem(oMatrix, "CardHolder", i);
                                
                                    CustomerID = getFormItemVal(form, fidCustID);
                                    CCAccountID = getMatrixItem(oMatrix, "CCAID", i);
                                    MethodType = getMatrixItem(oMatrix, "command", i);
                                    if (MethodType != "check")
                                    {
                                        if (CardNum == "" || CustNum == "")
                                        {
                                            SBO_Application.MessageBox("Invalid transaction record for credit.  Please select another one.");
                                        }
                                    }
                                }
                            }
                        }
                        
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        
                        HandleCMPress(form, pVal);
                   
                        break;
                }

            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    private void HandleCMPress(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
           
            switch (pVal.ItemUID)
            {
                case tbCreditCard:
                    if (!ItemValidate(form))
                        return;
                    form.PaneLevel = paneCCLog;
                    string CRID = form.DataSources.DBDataSources.Item("ORIN").GetValue("DocEntry", 0);
                    updateTransCMCustomerID(CRID);
                    if (CMDocEntry!="")
                        strWhere = string.Format(" Where \"U_customerID\"='{0}' and \"U_CreditMemoID\"='{1}' ", getCustomerID(), CRID);
                    else
                        strWhere = string.Format(" Where \"U_customerID\"='{0}'", getCustomerID());  
      
                    populateLog(form);
                    break;
                case btnCharge:
                    if (!ItemValidate(form))
                        return;
                    theActiveForm = form;
                   
                    
                    CreateCreditCardForm();
                    break;
                case "1":
                    {
                        if (CardNum != "" || MethodType=="check" )
                        {

                            string acct = getCCAccountNameByID(CCAccountID);
                            if(acct == null)
                            {
                                SBO_Application.MessageBox("Credit card not found. Please select a different transaction.");
                                return;
                            }
                            CMDocEntry = getDocEntry("14", CMDocNum, getFormItemVal(form, fidCustID)).ToString();
                            if(CMDocEntry=="")
                            {
                                SBO_Application.MessageBox("DocEntry for Credit Memo not found.");
                                return;
                            }
                            double amt = getCreditMemoTotal(int.Parse(CMDocEntry));
                            string strTransDesc = string.Format("Credit {0} back to card: {1} {2}", amt, CardNum, CardHolder);
                            if(MethodType == "check")
                            {
                                strTransDesc = string.Format("Credit {0} back to checking account: {1} {2}", amt, CardNum, CardHolder);
                            }
                            if (SBO_Application.MessageBox(strTransDesc + "?", 1, "Yes", "No") == 1)
                            {

                                if (MethodType == "check")
                                {
                                    if (!runCustomerTrans("checkcredit", amt, CustNum, CardNum, CustomerID, CMDocEntry, CCAccountID))
                                        return;
                                }else
                                {
                                    if (!runCustomerTrans("cc:credit", amt, CustNum, CardNum, CustomerID, CMDocEntry, CCAccountID))
                                        return;
                       
                                }
                                if (AddOutgoingPayment(int.Parse(CMDocEntry), CustomerID, CardNum,MethodType, amt, CCAccountID))
                                {
                                    updateCCTRAN(SAPbobsCOM.BoObjectTypes.oCreditNotes, CCAccountID);
                                   
                                    
                                    SBO_Application.Menus.Item("1289").Activate();
             
                                }
                                else
                                {
                                    voidCustomer(confirmNum);
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
        
    }
    /*
    public string getBaseOnInvoiceID(SAPbouiCOM.Form form)
    {

        return getBaseOnID(form, strBaseOnARInvoice);
    }
    public string getBaseOnDownPaymentInvoiceID(SAPbouiCOM.Form form)
    {

        return getBaseOnID(form, strBaseOnARDownPayment);
    }
     * */
}


