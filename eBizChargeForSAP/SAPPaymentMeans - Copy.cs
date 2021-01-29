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
    int panePM = 15;
    string PMCustID;
    List<SAPCust> PMCustList = new List<SAPCust>();
    SAPCust PMSAPCust;
    CCTRAN preauthCCTRAN = null;
    const string stRecordID = "stRecID";
    private void PaymentMeansFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        PMCustID = getCustomerID();
                        if (PMCustID == "")
                        {
                            SBO_Application.MessageBox("Please select a customer first");
                            form.Visible = false;
                            return;
                        }
                        AddTab(form, panePM, "eBizCharge", "5");
                        AddCCPMFormField(form, paneBP);
                        form.Items.Item("5").Visible = false;
                        reload(form, PMCustID, ref PMCustList, ref  PMSAPCust);
                        break;
                }

            }
            else
            {

                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        currentForm = form;
                        tProcess = new Thread(ThreadProc);
                        tProcess.Start();

                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        {
                            switch (pVal.ItemUID)
                            {
                                case cbPaymentMethod:
                                    PMSAPCust = getCCInfoSelect(form, PMCustList);
                                    populateCCInfo(form, PMSAPCust);

                                    break;
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        handlePMItemPress(form, pVal);
                        break;
                }


            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
    private void AddCCPMFormField(SAPbouiCOM.Form form, int pane)
    {
        try
        {

                     
            int edL = 110 + 15; //oItm.Left + oItm.Width;
            int edW = 250;
            int edT = 54;
            int edH = 15;
            int nGap = 17;
            SAPbouiCOM.Item oItm = null;
            switch (theActiveForm.TypeEx)
            {

                case "139": //Sales Order
                    oItm = addPaneItem(form, stRecordID, edL, edT, edW, edH, "Sales Order ID", SAPbouiCOM.BoFormItemTypes.it_STATIC, pane, 0);
                    break;
                case "133": //Invoice
                    oItm = addPaneItem(form, stRecordID, edL, edT, edW, edH, "Invoice ID", SAPbouiCOM.BoFormItemTypes.it_STATIC, pane, 0);
                    break;
                case "179": //Credit Memo
                    oItm = addPaneItem(form, stRecordID, edL, edT, edW, edH, "Credit Memo ID", SAPbouiCOM.BoFormItemTypes.it_STATIC, pane, 0);
                    break;
                case "170":
                    //FORMTYPE.INCOMING_PAYMENT;
                    oItm = addPaneItem(form, stRecordID, edL, edT, edW, edH, "Incoming Payment", SAPbouiCOM.BoFormItemTypes.it_STATIC, pane, 0);
                    break;
                case "420": //FORMTYPE.OUTGOING_PAYMENT;
                    oItm = addPaneItem(form, stRecordID, edL, edT, edW, edH, "Outgoing Payment", SAPbouiCOM.BoFormItemTypes.it_STATIC, pane, 0);
                    break;
                case "60091":  //A/R Reserve Invoice
                    oItm = addPaneItem(form, stRecordID, edL, edT, edW, edH, "Invoice ID:", SAPbouiCOM.BoFormItemTypes.it_STATIC, pane, 0);
                    break;
                case "65300": //Down Payment Invoice
                    oItm = addPaneItem(form, stRecordID, edL, edT, edW, edH, "Invoice ID:", SAPbouiCOM.BoFormItemTypes.it_STATIC, pane, 0);
                    break;
                default:
                    oItm = addPaneItem(form, stRecordID, edL, edT, edW, edH, "Record ID:", SAPbouiCOM.BoFormItemTypes.it_STATIC, pane, 0);
                    break;
            }
            
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, cbPaymentMethod, edL, edT, edW, edH, "Payment Method:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, pane, 1);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCardNum, edL, edT, edW, edH, "Card Number:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 2);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCardExpDate, edL, edT, edW, edH, "Card Exp. Date:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 3);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCardCode, edL, edT, edW, edH, "Card Code:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 4);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderName, edL, edT, edW, edH, "Card Holder:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 6);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderAddr, edL, edT, edW, edH, "Address:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 7);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderCity, edL, edT, edW, edH, "City:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 8);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderState, edL, edT, edW, edH, "State:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 9);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderZip, edL, edT, edW, edH, "Zip:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 10);
            edT = oItm.Top + 2 * nGap;

            oItm = addPaneItem(form, editCAmount, edL, edT, 100, edH, "Amount:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 11);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, cbTransactionType, edL, edT,100 , edH, "Transaction Type:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, pane, 12);
            SAPbouiCOM.ComboBox oCB = form.Items.Item(cbTransactionType).Specific;
            switch(theActiveForm.TypeEx)
            {
                
                case "139": //Sales Order
                   oCB.ValidValues.Add("Pre-auth", "Pre-auth");
                   oCB.ValidValues.Add("Charge", "Charge");
                   oCB.Select("Pre-auth");
                   setLabelCaption(form, stRecordID, getItemValue(theActiveForm, fidRecordID));
                    break;
                case "133": //Invoice
                  
                    oCB.ValidValues.Add("Charge", "Charge");
                    oCB.Select("Charge");
                    if (CheckCapture(form))
                    {
                        oCB.ValidValues.Add("Capture", "Capture");
                        oCB.Select("Capture");
                    }
                    setLabelCaption(form, stRecordID, getItemValue(theActiveForm, fidRecordID));
                    break;
                case "179": //Credit Memo
                    oCB.ValidValues.Add("Credit", "Credit");
                    oCB.Select("Credit");
                    setLabelCaption(form, stRecordID, getItemValue(theActiveForm, fidRecordID));
                    break;
                case "170":
                    //FORMTYPE.INCOMING_PAYMENT;
                    oCB.ValidValues.Add("Charge", "Charge");
                    oCB.Select("Charge");
                    break;
                case "420": //FORMTYPE.OUTGOING_PAYMENT;
                    oCB.ValidValues.Add("Credit", "Credit");
                    oCB.Select("Credit");
                    break;
                case "60091":  //A/R Reserve Invoice
                    oCB.ValidValues.Add("Charge", "Charge");
                    oCB.Select("Charge");
                    setLabelCaption(form, stRecordID, getItemValue(theActiveForm, fidRecordID));
                    break;
                case "65300": //Down Payment Invoice
                    oCB.ValidValues.Add("Charge", "Charge");
                    oCB.Select("Charge");
                    setLabelCaption(form, stRecordID, getItemValue(theActiveForm, fidRecordID));
                    break;
            }
            setFormEditVal(form, editCAmount, getItemValue(form, "14"));
        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
    }
    private bool CheckCapture(SAPbouiCOM.Form form)
    {
        try
        {
            var q = from a in db.CCTRANs
                    where a.customerID == PMCustID
                    && a.command == "cc:authonly" 
                    && a.recDate > (DateTime.Today.AddDays(- 30))
                    && !db.CCTRANs.Any(a2 => (a2.refNum == a.refNum) && (a2.command == "cc:capture"))
                    select a;
            if (q.Count() == 0)
                return false;
            foreach(CCTRAN t in q)
            {
                if(double.Parse(t.amount) == getItemDoubleValue(theActiveForm, fidBalance))
                {
                    preauthCCTRAN = t;
                    return true;
                }
            }

        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
}
