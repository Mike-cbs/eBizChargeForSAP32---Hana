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

public class PREAUTHTRANS
{
    public string amount;
    public string refNum;
    public string paymentKey;
    public string CCAcountID;
    public List<CCTRAN> CCTRANS;
    public PREAUTHTRANS()
    {
        amount = "";
        refNum = "";
        CCTRANS = new List<CCTRAN>();
    }
}
partial class SAP
{
    int panePM = 125;
    string PMCustID;
    List<SAPCust> PMCustList = new List<SAPCust>();
    SAPCust PMSAPCust;
    Thread tProcess;
    SAPbouiCOM.Form currentForm;
    SAPbouiCOM.Form targetedForm;
    const string btnProcess = "btnProc";
    PREAUTHTRANS preauthCCTRAN = null;
    const string stRecordID = "stRecID";
    const string stCCAccount = "stCCAcct";

    SAPbobsCOM.Documents oSQDoc = null;
    double dChargeAmount = 0;
    string confirmNum;
    string theInvoiceID = "";
    string strTransDesc = "";
    bool bRecurringBilling;
    bool bPMControlCreated = false;
    bool bDeclined = false;

    public void DeclinedAction()
    {
        try
        {
            theActiveForm.Items.Item("114").Click();
            ((SAPbouiCOM.CheckBox)theActiveForm.Items.Item("136").Specific).Checked = false;
            theActiveForm.Items.Item("1").Click();
        }
        catch (Exception) { }
    }

    public void SaveInvoiceThreadProc()
    {
        try
        {

            bAuto = true;

            SBO_Application.Menus.Item("1289").Activate();
            SBO_Application.Menus.Item("1288").Activate();


        }
        catch (Exception)
        {
            //errorLog(ex);
        }
    }
    SAPCust DPSAPCust = null;
    public void SOChargeLastRecordThreadProc(double dAmount, SAPbouiCOM.Form formSales)
    {
        SAPbobsCOM.Documents oSODoc = null;
        try
        {
            showStatus("Creating down payment invoice.  Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

            // bAuto = true;
            // SBO_Application.Menus.Item("1289").Activate();
            // SBO_Application.Menus.Item("1288").Activate();
            string cardcode = getCustomerID();
            string docNum = getFormItemVal(formSales, "8");
            int id = getDocEntry("17", docNum, cardcode);
            oSODoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

            if (oSODoc.GetByKey(id))
            {
                List<int> list = getSOInvoice(id);
                foreach (int inv in list)
                {
                    double bal = getInvoiceBalance(inv.ToString());
                    if (bal > 0 && dAmount > 0)
                    {
                        INVDocEntry = inv.ToString();
                        double amt = dAmount > bal ? bal : dAmount;
                        if (AddPayment(amt, DPSAPCust, inv.ToString()))
                        {
                            dAmount = dAmount - amt;
                        }
                    }
                }
                if (dAmount <= 0)
                    return;
                //    SBO_Application.Menus.Item("2071").Activate();  // Down Payment Invoice
                if (createDownpaymentInvoice(formSales, DPSAPCust, dAmount, oSODoc))
                {

                    if (DPDocEntry == "")
                    {
                        voidCustomer(confirmNum);
                        return;
                    }
                    AddDownPayment(dAmount, int.Parse(DPDocEntry), DPSAPCust);
                    updateCCTRAN(SAPbobsCOM.BoObjectTypes.oDownPayments, DPSAPCust.cccust.CCAccountID);
                    //SBO_Application.Menus.Item("1289").Activate();
                    //SBO_Application.Menus.Item("1288").Activate();
                    //theActiveForm.Items.Item("1").Click();
                    if (bRecurringBilling)
                    {
                        bRecurringBilling = false;
                        if (DPSAPCust.custObj.PaymentMethodProfiles[0].MethodType == "cc")
                            strJobDesc = string.Format("Recursive charge of {0} on {1}({2})", dAmount, DPSAPCust.custObj.PaymentMethodProfiles[0].CardNumber, DPSAPCust.custObj.PaymentMethodProfiles[0].MethodName);
                        else
                            strJobDesc = string.Format("Recursive charge of {0} on {1}({2})", dAmount, DPSAPCust.custObj.PaymentMethodProfiles[0].Account, DPSAPCust.custObj.PaymentMethodProfiles[0].MethodName);

                        CreateJobForm();
                    }
                }

            }
            else
            {
                if (SBO_Application.MessageBox(string.Format("DocEntry for Sales Order not found.  DocNum: {0}, CardCode: {1}.  Do you want to void this transaction?", docNum, cardcode)) == 1)
                    voidCustomer(confirmNum, true);
                return;
            }


        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        finally
        {
            try
            {
                if (oSODoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSODoc);
                    oSODoc = null;
                }

            }
            catch (Exception) { }
        }

    }

    string savedCustomerID = "";

    public void RecordRefreshThreadProc()
    {
        SAPbobsCOM.Documents oSODoc = null;
        try
        {
            if (theActiveForm.TypeEx == FORMSALESORDER)
                return;
            if (theActiveForm.TypeEx == FORMINCOMINGPAYMENT)
            {
                string CustomerID = getCustomerID();
                string docID = getDocNumAdded(CustomerID, theActiveForm.TypeEx);
                SBO_Application.Menus.Item("1281").Activate();
                setFormEditVal(theActiveForm, "3", docID);
                theActiveForm.Items.Item("1").Click();
                return;
            }

            try
            {
                SBO_Application.Menus.Item("1304").Activate();  //8.8 don't have refresh

            }
            catch
            {

                if (SBO_Application.Menus.Item("1281").Enabled)  //Find enabled
                {
                    string docID = getFormItemVal(theActiveForm, fidRecordID);
                    trace("Refresh order Num: " + docID);
                    SBO_Application.Menus.Item("1281").Activate();
                    setFormEditVal(theActiveForm, fidRecordID, docID);
                    theActiveForm.Items.Item("1").Click();
                }
            }
            //errorLog("Pre/next Record");
            if (oSODoc != null)
            {
                if (bRecurringBilling)
                {
                    bRecurringBilling = false;
                    if (PMSAPCust.custObj.PaymentMethodProfiles[0].MethodType == "cc")
                        strJobDesc = string.Format("Recursive charge of {0} on {1}({2})", dChargeAmount, PMSAPCust.custObj.PaymentMethodProfiles[0].CardNumber, PMSAPCust.custObj.PaymentMethodProfiles[0].MethodName);
                    else
                        strJobDesc = string.Format("Recursive charge of {0} on {1}({2})", dChargeAmount, PMSAPCust.custObj.PaymentMethodProfiles[0].Account, PMSAPCust.custObj.PaymentMethodProfiles[0].MethodName);
                    JobCustID = oSODoc.CardCode;
                    JobOrderID = oSODoc.DocEntry.ToString();
                    JobPaymentID = PMSAPCust.cccust.methodDescription;
                    JobAmount = dChargeAmount.ToString();
                    resetGlobal();
                    CreateJobForm();
                }
            }

        }
        catch (Exception)
        {
            //errorLog(ex);
        }
        finally
        {
            try
            {
                if (oSODoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSODoc);
                    oSODoc = null;
                }

            }
            catch (Exception) { }
        }
    }

    public void ThreadProc()
    {
        try
        {
            trace("PM Load Process");
            string bal = getFormItemVal(currentForm, "14");
            if (bal == "")
            {
                if (currentForm.Items.Item("5") != null)
                    currentForm.Items.Item("5").Visible = true;
                //currentForm.Items.Item(tbCreditCard).Visible = false;

                return;
            }
            if (bUseeBizCharge)
            {

                trace("PM Load Process bUseeBizCharge");
                if (currentForm.Items.Item(tbCreditCard) != null)
                    currentForm.Items.Item(tbCreditCard).Click();
                else
                    trace("currentForm.Items.Item(tbCreditCard) is null");
            }

        }
        catch (Exception ex)
        {
            errorLog("Payment thread exception: " + ex.Message);
        }
        finally
        {

        }
    }
    private void setEnable(SAPbouiCOM.Form form, string fItem, string itemname, bool bEnable)
    {
        try
        {
            form.Items.Item(itemname).Enabled = bEnable;
        }
        catch (Exception)
        {
            form.Items.Item(fItem).Click();
            form.Items.Item(itemname).Enabled = bEnable;
        }
    }
    private void CreatePMControls(SAPbouiCOM.Form form)
    {
        try
        {
            trace("CreatePMControls");
            PMCustID = getCustomerID();
            if (PMCustID == "")
            {
                form.Items.Item("6").Click();
                return;
            }
            if (!bPMControlCreated)
            {
                trace("PMControls Not Created");
                AddCCPMFormField(form, panePM);
                SAPbouiCOM.Item oItem = form.Items.Add(btnProcess, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = form.Items.Item("1").Left;
                oItem.Width = form.Items.Item("1").Width;
                oItem.Top = form.Items.Item("1").Top;
                oItem.Height = form.Items.Item("1").Height;
                if (cfgHasTerminal == "Y")
                {
                    oItem = form.Items.Add(btnTerminal, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    oItem.Left = form.Items.Item("2").Left + form.Items.Item("2").Width + 10;
                    oItem.Width = form.Items.Item("2").Width;
                    oItem.Top = form.Items.Item("2").Top;
                    oItem.Height = form.Items.Item("2").Height;
                    setButtonCaption(form, btnTerminal, "Terminal");
                }
                bPMControlCreated = true;
            }
            bManualPopulate = true;
            trace("CreatePMControls reload");
            reload(form, ref PMCustList, ref PMSAPCust);
            /*
            string MethodID = getMethodID();

            if (MethodID != null)
            {
                SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)currentForm.Items.Item(cbPaymentMethod).Specific;
                foreach (SAPCust sapcust in PMCustList)
                {
                    if (sapcust.cccust.MethodID == MethodID)
                    {
                        oCB.Select(sapcust.key);
                        HandleSelect(form);
                    }
                }
            }
            */
            setButtonCaption(currentForm, btnProcess, getFormItemVal(currentForm, cbTransactionType));
            if (form.Items.Item("1") != null)
                form.Items.Item("1").Visible = false;
            if (form.Items.Item(btnProcess) != null)
                form.Items.Item(btnProcess).Visible = true;
            trace("CreatePMControls before setpane");
            setPane(form, PMSAPCust);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private string getInvoiceNum()
    {
        try
        {
            switch (theActiveForm.TypeEx)
            {
                case FORMCREDITMEMO:
                case FORMSALESORDER:
                case FORMINVOICE:
                    return getFormItemVal(theActiveForm, "8");
                case FORMINCOMINGPAYMENT:
                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)theActiveForm.Items.Item("20").Specific;
                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        if (getMatrixSelect(oMatrix, "10000127", i))
                        {
                            string type = getMatrixItem(oMatrix, "45", i);
                            string DocNum = getMatrixItem(oMatrix, "1", i);
                            string DocDate = getMatrixItem(oMatrix, "21", i);
                            string amount = getMatrixItem(oMatrix, "24", i);
                            if (type == "13")
                            {
                                return DocNum;
                            }

                        }

                    }

                    break;

            }

        }
        catch (Exception)
        {

        }
        return "";
    }
    public void MoveThreadProc()
    {
        try
        {
            currentForm.Freeze(true);
            MoveControl(currentForm);
        }
        catch (Exception ex)
        {
        }
        finally
        {
            currentForm.Freeze(false);
        }
    }
    private void MoveControl(SAPbouiCOM.Form form)
    {
        try
        {
            SAPbouiCOM.Item oItm = form.Items.Item(btnProcess);
            oItm.Left = form.Items.Item("1").Left;
            oItm.Width = form.Items.Item("1").Width;
            oItm.Top = form.Items.Item("1").Top;
            oItm.Height = form.Items.Item("1").Height;
            int nGap = 16;
            int edT = 50;
            int l = 0;
            oItm = form.Items.Item(stCardName);
            nGap = oItm.Height + 2;

            edT = oItm.Top + nGap;
            oItm = form.Items.Item(editHolderAddr);
            oItm.Top = edT;
            l = oItm.Width + oItm.Left;
            oItm = form.Items.Item("LB1027");
            oItm.Top = edT;

            edT = oItm.Top + nGap;
            oItm = form.Items.Item(editHolderCity);
            oItm.Top = edT;
            l = oItm.Width + oItm.Left;
            oItm = form.Items.Item("LB8");
            oItm.Top = edT;

            oItm = form.Items.Item("LB9");
            oItm.Top = edT;
            oItm.Left = l;
            l += oItm.Width;
            oItm = form.Items.Item(editHolderState);
            oItm.Top = edT;
            oItm.Left = l;
            l += oItm.Width;
            oItm = form.Items.Item("LB10");
            oItm.Top = edT;
            oItm.Left = l;
            l += oItm.Width;
            oItm = form.Items.Item(editHolderZip);
            oItm.Top = edT;
            oItm.Left = l;

            edT = oItm.Top + nGap;

            oItm = form.Items.Item(editHolderEmail);
            oItm.Top = edT;
            oItm = form.Items.Item("LB108");
            oItm.Top = edT;
            edT = oItm.Top + nGap;
            oItm = form.Items.Item(editMethodName);
            oItm.Top = edT;
            oItm = form.Items.Item("LB202");
            oItm.Top = edT;
            edT = oItm.Top + nGap;
            oItm = form.Items.Item(editCAmount);
            oItm.Top = edT;
            oItm = form.Items.Item("LB13");
            oItm.Top = edT;
            edT = oItm.Top + nGap;
            oItm = form.Items.Item(editPaymentDate);
            oItm.Top = edT;
            oItm = form.Items.Item("LB6677");
            oItm.Top = edT;
            edT = oItm.Top + nGap;

            oItm = form.Items.Item("LB14");
            oItm.Top = edT;
            oItm = form.Items.Item(cbTransactionType);
            oItm.Top = edT;
            l = oItm.Left + oItm.Width;
            oItm = form.Items.Item("LB134");
            oItm.Top = edT;
            oItm.Left = l;
            l += oItm.Width;
            oItm = form.Items.Item(cbRecurringBilling);
            oItm.Top = edT;
            oItm.Left = l;



        }
        catch (Exception)
        { }
    }
    private void PaymentMeansFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {

        try
        {
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        targetedForm = theActiveForm;
                        break;
                }

            }
            else
            {

                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:

                        currentForm = form;
                        tProcess = new Thread(MoveThreadProc);
                        tProcess.Start();
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:

                        if (bDeclined && theActiveForm.TypeEx == FORMSALESORDER && cfgDeclineUnapprove == "Y")
                        {
                            bDeclined = false;
                            tProcess = new Thread(DeclinedAction);
                            tProcess.Start();
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        try
                        {

                            bRecurringBilling = false;
                            bPMControlCreated = false;
                            form.Items.Item("5").Visible = false;
                            AddTab(form, panePM, "eBizCharge", "5");



                        }
                        catch (Exception)
                        {

                        }
                        currentForm = form;
                        tProcess = new Thread(ThreadProc);
                        tProcess.Start();
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
                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (!bPMControlCreated)
                            return;
                        if (pVal.ItemChanged == true)
                        {
                            switch (pVal.ItemUID)
                            {
                                case editCardNum:
                                    setCardName(form, PMSAPCust);
                                    break;
                                case "110":
                                case "111":
                                    setFormEditVal(form, editCAmount, getFormItemVal(form, "14"));
                                    break;

                            }

                        }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        {

                            switch (pVal.ItemUID)
                            {
                                case cbPaymentMethod:
                                    HandleSelect(form);

                                    break;
                                case cbTransactionType:
                                    if (form.Items.Item(btnProcess) != null)
                                        setButtonCaption(form, btnProcess, getFormItemVal(form, cbTransactionType));


                                    break;

                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                        theActiveForm = targetedForm;
                        handlePMItemPress(form, pVal);

                        break;
                }


            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    private void setPane(SAPbouiCOM.Form form, SAPCust sapcust)
    {
        try
        {
            if (form != null)
            {
                trace("setpan init pane to PMPane.");
                form.PaneLevel = panePM;
            }
            if (form == null || sapcust == null)
            {
                trace("form == null || sapcust == null");
                return;
            }

            string sel = getFormItemVal(form, cbPaymentMethod);
            if (sel == NEWCHECK)
                form.PaneLevel = panePM + 1;
            else if (sel == NEWCARD)
                form.PaneLevel = panePM;
            else
            {
                if (sapcust == null)
                    return;
                if (sapcust.custObj == null)
                {
                    trace("Setpane sapcust.custObj is null");
                    return;
                }
                if (sapcust.custObj.PaymentMethodProfiles[0] == null)
                {
                    trace("sapcust.custObj.PaymentMethodProfiles[0] is null");
                    return;
                }
                if (sapcust.custObj.PaymentMethodProfiles[0].MethodType == "check")
                    form.PaneLevel = panePM + 1;
                else
                    form.PaneLevel = panePM;

            }
            trace("setpan set pane to " + form.PaneLevel);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private bool ValidateChargeForm(SAPbouiCOM.Form form)
    {
        bool bRtn = true;
        try
        {
            if (getFormItemVal(form, cbTransactionType) == "No Trans")
                return true;
            string err = "";
            string pm = getFormItemVal(form, cbPaymentMethod);
            if (pm == NEWCARD)
            {
                if (getFormItemVal(form, editCardNum) == "")
                    err += "Credit card number is required.\r\n";
                if (getFormItemVal(form, editCardExpDate) == "")
                    err += "Credit card expiration date is required.\r\n";
                /*
                if (getFormItemVal(form, editCardCode) == "")
                    err += "Credit card code is required.\r\n";
                */
            }
            if (pm == NEWCHECK)
            {
                if (getFormItemVal(form, editCheckAccount) == "")
                    err += "Checking account number is required.\r\n";
                if (getFormItemVal(form, editCheckRouting) == "")
                    err += "Routing number is required.\r\n";

            }

            if (err != "")
            {
                showMessage(err);
                return false;
            }

        }
        catch (Exception)
        {

        }
        return bRtn;
    }
    public string getMethodID()
    {
        SAPbouiCOM.Form form = theActiveForm;
        string MethodID = null;
        try
        {
            if (form.TypeEx == FORMINCOMINGPAYMENT)
                return getMethodIDByCustomerID(getFormItemVal(form, "5"));
            MethodID = getMethodIDByCustomerID(getFormItemVal(form, fidCustID));
            switch (form.TypeEx)
            {

                case FORMSALESORDER: //Sales Order
                    SODocEntry = form.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);
                    if (SODocEntry != "")
                        return getMethodIDByOrderID(SODocEntry);

                    break;
                case FORMINVOICE: //Invoice
                    INVDocEntry = form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0);
                    if (INVDocEntry != "")
                        return getMethodID(INVDocEntry, "U_InvoiceID");
                    break;
                case "179": //Credit Memo
                    CMDocEntry = form.DataSources.DBDataSources.Item("ORIN").GetValue("DocEntry", 0);
                    if (CMDocEntry != "")
                        return getMethodID(CMDocEntry, "U_CreditMemoID");
                    break;

                case "60090":  //Invoice + Payment
                    INVDocEntry = form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0);
                    if (INVDocEntry != "")
                        return getMethodID(INVDocEntry, "U_InvoiceID");
                    break;
                case "60091":  //A/R Reserve Invoice
                    INVDocEntry = form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0);
                    if (INVDocEntry != "")
                        return getMethodID(INVDocEntry, "U_InvoiceID");
                    break;
                case "65300": //Down Payment Invoice
                    DPDocEntry = theActiveForm.DataSources.DBDataSources.Item("ODPI").GetValue("DocEntry", 0);
                    if (DPDocEntry != "")
                        return getMethodID(DPDocEntry, "U_DownPaymentInvID");
                    break;
            }
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)theActiveForm.Items.Item(fidMatrix).Specific;
            if (oMatrix.RowCount >= 2)
            {
                for (int i = 1; i < oMatrix.RowCount; i++)
                {
                    string s = getMatrixItem(oMatrix, "43", i);
                    if (s != "-1")
                    {
                        if (s == "17")
                        {
                            s = getMatrixItem(oMatrix, "45", i);
                            return getMethodIDByOrderID(s);
                        }
                        else if (s == "13")
                        {
                            s = getMatrixItem(oMatrix, "45", i);
                            return getMethodID(s, "U_InvoiceID");
                        }
                        else if (s == "15")  //Delivery
                        {
                            s = getMatrixItem(oMatrix, "45", i);
                            return getMethodID(getSOFromDelivery(s).ToString(), "U_OrderID");
                        }

                    }
                }
            }
            oMatrix = (SAPbouiCOM.Matrix)theActiveForm.Items.Item("39").Specific;
            if (oMatrix.RowCount >= 2)
            {
                for (int i = 1; i < oMatrix.RowCount; i++)
                {
                    string s = getMatrixItem(oMatrix, "23", i);
                    if (s != "-1")
                    {
                        if (s == "17")
                        {
                            s = getMatrixItem(oMatrix, "45", i);
                            return getMethodIDByOrderID(s);
                        }
                        if (s == "13")
                        {
                            s = getMatrixItem(oMatrix, "45", i);
                            return getMethodID(s, "U_InvoiceID");
                        }
                        if (s == "15")  //Delivery
                        {
                            s = getMatrixItem(oMatrix, "45", i);
                            return getMethodIDByOrderID(getSOFromDelivery(s).ToString());
                        }
                    }
                }
            }


        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return MethodID;
    }
    public bool bUseeBizCharge = true;
    SAPbouiCOM.Form formPM;
    private void handlePMItemPress(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {

        try
        {
            switch (pVal.ItemUID)
            {
                case btnTerminal:
                    TermAmount = getFormItemVal(form, editCAmount);
                    TermFunc = "Charge"; //getFormItemVal(form, btnProcess);
                    TermGroup = getGroupName(getCustomerID());
                    TermInvoice = getInvoiceNum();
                    TermCustomer = getCustomerID();
                    oTerminalLaunchForm = form;
                    CreateTerminalForm();
                    break;
                case tbCreditCard:
                    try
                    {
                        trace("tbCreditCard clicked");
                        bUseeBizCharge = true;
                        form.Freeze(true);
                        CreatePMControls(form);
                        form.Items.Item("1").Visible = false;
                        if (form.Items.Item(btnProcess) != null)
                            form.Items.Item(btnProcess).Visible = true;
                    }
                    catch (Exception)
                    {

                    }
                    finally
                    {
                        form.Freeze(false);
                    }
                    break;
                case "3":
                case "4":
                case "6":
                    bUseeBizCharge = false;
                    form.Items.Item("1").Visible = true;
                    if (form.Items.Item(btnProcess) != null)
                        form.Items.Item(btnProcess).Visible = false;
                    break;
                case btnProcess:
                    if (!ValidateChargeForm(form))
                        return;
                    formPM = form;
                    tProcess = new Thread(ProcessCard);
                    tProcess.Start();
                    break;
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    private void ProcessCard()
    {
        SAPbouiCOM.Form form = formPM;
        SAPbobsCOM.Documents oSODoc = null;
        try
        {

            try { form.Items.Item(btnProcess).Enabled = false; } catch (Exception) { }
            theInvoiceID = null;
            switch (getFormItemVal(form, cbTransactionType))
            {
                case "Pre-auth":
                    {
                        SODocEntry = theActiveForm.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);

                        bAuto = false;
                        if (!loadFormData(form, ref PMSAPCust, true))
                            return;
                        if (PMSAPCust.custObj.CustomerId == null || PMSAPCust.cccust.CustomerID == null)
                        {
                            if (PMCustID == null)
                                PMCustID = getCustomerID();
                            PMSAPCust.custObj.CustomerId = PMCustID;
                            PMSAPCust.cccust.CustomerID = PMCustID;
                        }
                        if (!Validate(form, PMSAPCust.custObj))
                            return;
                        if (PMSAPCust.custObj.PaymentMethodProfiles[0].MethodType == "check")
                            strTransDesc = "PreAuth " + getFormItemVal(form, editCAmount) + " to checking account " + getFormItemVal(form, editCheckAccount);
                        else
                            strTransDesc = "PreAuth " + getFormItemVal(form, editCAmount) + " to card " + getFormItemVal(form, editCardNum);

                        if (SBO_Application.MessageBox(strTransDesc + "?", 1, "Yes", "No") == 1)
                        {

                            if (PreAuthCustomer(form, PMSAPCust))
                            {
                                oSODoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                                if (createSO(theActiveForm, PMSAPCust, ref oSODoc))
                                {
                                    dChargeAmount = getItemDoubleValue(form, editCAmount);
                                    if (getFormItemVal(form, cbRecurringBilling) == "Y")
                                        bRecurringBilling = true;
                                    else
                                        bRecurringBilling = false;
                                    updateCCTRAN(SAPbobsCOM.BoObjectTypes.oOrders, PMSAPCust.cccust.CCAccountID);
                                    bAuto = true;
                                    form.Close();
                                    tProcess = new Thread(RecordRefreshThreadProc);
                                    tProcess.Start();
                                }


                            }
                        }
                    }
                    break;
                case "Capture":
                    HandleCapture(form, ref PMSAPCust);
                    break;
                case "No Trans":
                    loadFormData(form, ref PMSAPCust, true);
                    if (SBO_Application.MessageBox("Create payment record on " + PMSAPCust.paymentDate.ToString("MM/dd/yyyy") + "?", 1, "Yes", "No") == 1)
                    {
                        if (!loadFormData(form, ref PMSAPCust, true))
                            return;
                        if (PMSAPCust.custObj.CustomerId == null || PMSAPCust.cccust.CustomerID == null)
                        {
                            if (PMCustID == null)
                                PMCustID = getCustomerID();
                            PMSAPCust.custObj.CustomerId = PMCustID;
                            PMSAPCust.cccust.CustomerID = PMCustID;
                        }
                        confirmNum = "";
                        showStatus("Creating records.  Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                        handleCharge(form, PMSAPCust, true);

                    }
                    break;

                case "Charge":
                    {
                        bAuto = false;
                        if (!loadFormData(form, ref PMSAPCust, true))
                            return;
                        if (PMSAPCust.custObj.CustomerId == null || PMSAPCust.cccust.CustomerID == null)
                        {
                            if (PMCustID == null)
                                PMCustID = getCustomerID();
                            PMSAPCust.custObj.CustomerId = PMCustID;
                            PMSAPCust.cccust.CustomerID = PMCustID;
                        }
                        if (!Validate(form, PMSAPCust.custObj))
                            return;
                        if (PMSAPCust.custObj.PaymentMethodProfiles[0].MethodType == "check")
                            strTransDesc = "Charge " + getFormItemVal(form, editCAmount) + " to checking account " + getFormItemVal(form, editCheckAccount);
                        else
                            strTransDesc = "Charge " + getFormItemVal(form, editCAmount) + " to card " + getFormItemVal(form, editCardNum);
                        double amt = getItemDoubleValue(form, editCAmount);
                        if (amt < 0)
                            strTransDesc = strTransDesc.Replace("Charge ", "Refund ");
                        if (SBO_Application.MessageBox(strTransDesc + "?", 1, "Yes", "No") == 1)
                        {
                            showStatus("Processing charge.  Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                            handleCharge(form, PMSAPCust);



                        }
                    }
                    break;
            }
        }
        catch (Exception)
        {

        }
        finally
        {
            try
            {
                if (oSODoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSODoc);
                    oSODoc = null;
                }
                form.Items.Item(btnProcess).Enabled = true;
            }
            catch (Exception) { }
        }
    }
    public void HandleCapture(SAPbouiCOM.Form form, ref SAPCust sapcust)
    {
        try
        {
            {
                bAuto = false;
                if (!loadFormData(form, ref sapcust, true))
                    return;
                if (sapcust.custObj.CustomerId == null || sapcust.cccust.CustomerID == null)
                {
                    string CustomerID = getCustomerID();
                    sapcust.custObj.CustomerId = CustomerID;
                    sapcust.cccust.CustomerID = CustomerID;
                }
                if (!Validate(form, sapcust.custObj))
                    return;

                double Camt = getItemDoubleValue(form, editCAmount);
                if (Camt < getDoubleValue(preauthCCTRAN.amount))
                    preauthCCTRAN.amount = Camt.ToString();
                trace(string.Format("Preauth Capture {0},{2} cfgCaptureOnly={1}", Camt, cfgCaptureOnly, preauthCCTRAN.amount));
                if (Camt > getDoubleValue(preauthCCTRAN.amount) && cfgCaptureOnly != "Y")
                {
                    strTransDesc = "Preauth amount: " + preauthCCTRAN.amount + " is less then the invoice balance: " + Camt + ".  Run charge of " + Camt + " and void orignal Preauth transaction";

                }
                else
                {

                    strTransDesc = "Capture transaction ref number: " + preauthCCTRAN.refNum + ", amount: " + preauthCCTRAN.amount;
                }


                if (SBO_Application.MessageBox(strTransDesc + "?", 1, "Yes", "No") == 1)
                {
                    double amount = 0;
                    if (Camt > getDoubleValue(preauthCCTRAN.amount) && cfgCaptureOnly != "Y")
                    {
                        amount = Camt;
                        if (ChargeCustomer(form, sapcust))
                        {
                            if (CaptureCreatePayment(Camt, form, sapcust, preauthCCTRAN.CCTRANS[0]))
                            {
                                foreach (CCTRAN t in preauthCCTRAN.CCTRANS)
                                {
                                    voidCustomer(t.refNum, false);
                                }
                                bAuto = true;
                                form.Close();
                                tProcess = new Thread(RecordRefreshThreadProc);
                                tProcess.Start();
                            }
                        }
                    }
                    else
                    {
                        double remain = getDoubleValue(preauthCCTRAN.amount);
                        foreach (CCTRAN t in preauthCCTRAN.CCTRANS)
                        {
                            if (remain > 0 && getDoubleValue(t.amount) > 0)
                            {
                                sapcust.cccust.CCAccountID = preauthCCTRAN.CCAcountID;
                                amount = getDoubleValue(t.amount);
                                if (getDoubleValue(t.amount) > remain)
                                    amount = remain;
                                remain = remain - amount;
                                if (CaptureCustomer(form, sapcust, t, amount))
                                {
                                    CaptureCreatePayment(amount, form, sapcust, t);

                                }
                            }
                        }
                        bAuto = true;
                        form.Close();
                        tProcess = new Thread(RecordRefreshThreadProc);
                        tProcess.Start();
                    }
                }

            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    public bool CaptureCreatePayment(double amount, SAPbouiCOM.Form form, SAPCust sapcust, CCTRAN t)
    {
        bool ret = true;
        try
        {
            switch (theActiveForm.TypeEx)
            {
                case FORMINVOICE:
                    {
                        if (createInvoice(theActiveForm, sapcust))
                        {
                            AddPayment(amount, sapcust, t.InvoiceID, getFormItemVal(form, "111"));
                            updateCCTRAN(SAPbobsCOM.BoObjectTypes.oInvoices, PMSAPCust.cccust.CCAccountID);

                        }
                        else
                        {
                            voidCustomer(confirmNum);

                        }
                    }
                    break;
                case FORMINCOMINGPAYMENT:
                    try
                    {

                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)theActiveForm.Items.Item("20").Specific;
                        string CustomerID = getCustomerID();
                        int invID = 0;
                        int invCount = 0;
                        for (int i = 1; i <= oMatrix.RowCount; i++)
                        {
                            if (getMatrixSelect(oMatrix, "10000127", i))
                            {
                                string type = getMatrixItem(oMatrix, "45", i);
                                string DocNum = getMatrixItem(oMatrix, "1", i);
                                string DocDate = getMatrixItem(oMatrix, "21", i);
                                if (type == "13")
                                {
                                    invCount++;
                                    invID = getDocEntry("13", DocNum, CustomerID);

                                }

                            }

                        }
                        AddPayment(amount, sapcust, invID.ToString());
                        CCTRANAdjustWithINVId(confirmNum, invID.ToString());
                        //updateCCTRAN(SAPbobsCOM.BoObjectTypes.oInvoices, PMSAPCust.cccust.CCAccountID);
                    }
                    catch (Exception ex) { errorLog(ex); }
                    break;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

        return ret;
    }
    private void handleCharge(SAPbouiCOM.Form form, SAPCust sapcust, bool bNoTrans = false)
    {
        SAPbouiCOM.Form formSales = theActiveForm;
        SAPbobsCOM.Documents oSODoc = null;

        try
        {


            confirmNum = "";
            if (!IsBPActive(sapcust.cccust.CustomerID))
            {
                SBO_Application.MessageBox("Can not process your request.  Customer is inactive.");
                return;
            }
            if (!bNoTrans)
            {
                if (!ChargeCustomer(form, sapcust))
                    return;
            }
            dChargeAmount = getItemDoubleValue(form, editCAmount);
            //errorLog("dChargeAmount=" + dChargeAmount.ToString());
            //errorLog("sapcust=" + sapcust.cccust.CCAccountID.ToString());

            switch (formSales.TypeEx)
            {
                case FORMSALESORDER:  //Sales Order
                    {
                        oSODoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                        if (createSO(formSales, sapcust, ref oSODoc))
                        {
                            if (getFormItemVal(form, cbRecurringBilling) == "Y")
                                bRecurringBilling = true;
                            else
                                bRecurringBilling = false;
                            DPSAPCust = sapcust;
                            SOChargeLastRecordThreadProc(dChargeAmount, formSales);
                            form.Close();
                            //tProcess = new Thread(SOChargeLastRecordThreadProc);
                            // tProcess.Start();


                        }
                        else
                        {
                            voidCustomer(confirmNum);
                            form.Close();
                        }

                    }
                    break;
                case FORMINCOMINGPAYMENT:  //Incoming Payment
                    {
                        SAPbobsCOM.Payments oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                        try
                        {
                            string bplid = getFormItemVal(formSales, "1320002037");
                            if (bplid != "")
                            {
                                trace("Create Payment from incomming payment. Branch ID: " + bplid);
                                oPmt.BPLID = int.Parse(bplid);
                            }
                        }
                        catch (Exception) { }
                        oPmt.CardCode = getFormItemVal(formSales, "5");
                        oPmt.DocDate = DateTime.Now;
                        oPmt.Remarks = getFormItemVal(formSales, "26");
                        oPmt.JournalRemarks = getFormItemVal(formSales, "59");
                        trace("Incomming Payment Counter reference = " + getFormItemVal(formSales, "22"));
                        oPmt.CounterReference = getFormItemVal(formSales, "22");

                        oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
                        oPmt.JournalRemarks = getFormItemVal(formSales, "59");
                        if (sapcust.paymentDate == null)
                            oPmt.DocDate = DateTime.Now;
                        else
                            oPmt.DocDate = sapcust.paymentDate;
                        AddCreditCardPayment(oPmt, sapcust, dChargeAmount, true);

                        bool bCanClose = false;

                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)formSales.Items.Item("20").Specific;
                        List<String> ids = new List<string>();

                        for (int i = 1; i <= oMatrix.RowCount; i++)
                        {
                            if (getMatrixSelect(oMatrix, "10000127", i))
                            {
                                string type = getMatrixItem(oMatrix, "45", i);
                                string DocNum = getMatrixItem(oMatrix, "1", i);
                                string DocDate = getMatrixItem(oMatrix, "21", i);
                                string amount = getMatrixItem(oMatrix, "24", i);
                                int docEntry = getDocEntry(type, DocNum, sapcust.cccust.CustomerID);
                                string dcount = getMatrixItem(oMatrix, "55", i);
                                string installment = getMatrixItem(oMatrix, "71", i);
                                installment = installment.Split(' ')[0];

                                trace("Incoming payment type=" + type + ", DocNum=" + DocNum + ",docEntry = " + docEntry.ToString() + ", totalDiscount=" + dcount);
                                oPmt.Invoices.DocEntry = docEntry;
                                oPmt.Invoices.InstallmentId = int.Parse(installment);
                                bool bFC = false;

                                try
                                {
                                    SAPbobsCOM.Documents oInv = (SAPbobsCOM.Documents)oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)int.Parse(type));
                                    oInv.GetByKey(docEntry);
                                    if (oInv.DocTotalFc != 0)
                                        bFC = true;

                                }
                                catch (Exception)
                                { }


                                oPmt.Invoices.InvoiceType = (SAPbobsCOM.BoRcptInvTypes)int.Parse(type);
                                if (bFC)
                                    oPmt.Invoices.AppliedFC = getDoubleValue(getMoneyString(amount));

                                else
                                    oPmt.Invoices.SumApplied = getDoubleValue(getMoneyString(amount));
                                if (type == "24" || type == "30" || type == "-2")
                                {
                                    oPmt.Invoices.DocLine = getJELineID(docEntry.ToString(), sapcust.cccust.CustomerID);
                                }
                                //else
                                //oPmt.Invoices.DocLine = 0;
                                try
                                {
                                    if (bFC)
                                        oPmt.Invoices.TotalDiscountFC = getDoubleValue(dcount);
                                    else
                                        oPmt.Invoices.TotalDiscount = getDoubleValue(dcount);
                                }
                                catch (Exception)
                                {

                                }
                                oPmt.Invoices.Add();

                            }

                        }



                        int r = oPmt.Add();
                        if (r != 0)
                        {
                            errorLog(getErrorString());
                            bAuto = false;
                            SBO_Application.MessageBox(getErrorString());
                            voidCustomer(confirmNum);

                        }
                        else
                        {
                            bCanClose = true;
                            showStatus("eBizCharge: incoming payment added.  Amount: " + dChargeAmount, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                        }
                        if (bCanClose)
                        {
                            form.Close();
                            bAuto = true;
                            savedCustomerID = getFormItemVal(formSales, "5");
                            SBO_Application.Menus.Item(menuIDLastRecord).Activate();
                        }
                        //setFormEditVal(formSales, "12", "");
                        //bAuto = true;
                        //form.Close();
                    }
                    break;
                case "65300":  //downpayment invoice
                    {
                        dChargeAmount = getItemDoubleValue(form, editCAmount);

                        if (createDownpaymentInvoice(formSales, sapcust))
                        {
                            double amt = getItemDoubleValue(form, editCAmount);
                            if (DPDocEntry == "")
                            {
                                SBO_Application.MessageBox("DocEntry for down payment not found");
                                voidCustomer(confirmNum);
                                return;
                            }
                            AddDownPayment(dChargeAmount, int.Parse(DPDocEntry), sapcust);
                            updateCCTRAN(SAPbobsCOM.BoObjectTypes.oDownPayments, sapcust.cccust.CCAccountID);
                            bAuto = true;
                            form.Close();

                            tProcess = new Thread(RecordRefreshThreadProc);
                            tProcess.Start();
                        }

                    }
                    break;
                default:
                    {

                        if (createInvoice(formSales, sapcust))
                        {
                            double amt = getItemDoubleValue(form, editCAmount);
                            AddPayment(amt, sapcust, INVDocEntry, getFormItemVal(form, "111"));
                            updateCCTRAN(SAPbobsCOM.BoObjectTypes.oInvoices, sapcust.cccust.CCAccountID);

                            bAuto = true;
                            form.Close();

                            tProcess = new Thread(RecordRefreshThreadProc);
                            tProcess.Start();
                        }
                        else
                        {
                            voidCustomer(confirmNum);
                            form.Close();
                        }

                    }
                    break;
            }
        }
        catch (Exception)
        {

        }
        finally
        {
            try
            {
                if (oSODoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSODoc);
                    oSODoc = null;
                }
                form.Items.Item(btnProcess).Enabled = true;
            }
            catch (Exception) { }
        }
    }
    public void HandleeConnectInvoice(string invoiceID, string amt, string CardCode = "", string UploadedBalance = "", string RefNum = "")
    {
        double AmtPaid = 0;
        try
        {
            if (amt == "")
            {
                //  invoiceID = getDocEntry("13", invoiceID, CardCode).ToString();
                double paid = getInvoicePaidToDate(int.Parse(invoiceID));
                if (paid == 0)
                    return;
                amt = paid.ToString();
            }
            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            if (!oDoc.GetByKey(int.Parse(invoiceID)))
            {
                trace("HandleeConnectInvoice invoice not found: " + invoiceID);
                return;
            }
            string currency = oDoc.DocCurrency;
            string cardcode = "";
            string acctname = getAccountName(oDoc.CardCode);
            // string cardName = getCardName("", "", currency, ref cardcode, acctname);
            //SecurityToken securityToken = getToken(cardcode);
            SecurityToken securityToken = getToken();

            Invoice inv = ebiz.GetInvoice(securityToken, oDoc.CardCode, "", oDoc.DocNum.ToString(), "");
            if (inv.AmountDue == 0)
            {
                trace("HandleeConnectInvoice invoice: " + oDoc.DocNum + " already set to 0 balance.");
                updateCCConnectInvoice(oDoc.DocNum.ToString(), amt, inv.AmountDue.ToString());
                return;
            }
            decimal newBal = (decimal)(Math.Round(oDoc.DocTotal - getDoubleValue(amt), 2));
            if ((decimal)Math.Round((double)inv.AmountDue, 2) == newBal)
            {
                trace("HandleeConnectInvoice invoice: " + invoiceID + " balance is up to date. balance = " + inv.AmountDue + " new balance = " + newBal);
                return;
            }
            decimal PaidAmtDiff = Convert.ToDecimal(inv.AmountDue - newBal);

            //if (diff == 0)
            inv.InvoiceAmount = (decimal)oDoc.DocTotal;
            inv.NotifyCustomer = true;
            if (newBal == 0)
                inv.InvoiceDueDate = oDoc.DocDueDate.AddYears(1).ToString("MM/dd/yyyy");
            inv.AmountDue = newBal;
            inv.Software = Software;
            //InvoiceResponse resp = ebiz.UpdateInvoice(securityToken, inv, oDoc.CardCode, "", oDoc.DocNum.ToString(), "");
            try
            {
                Customer cust = ebiz.GetCustomer(securityToken, oDoc.CardCode, null);
                InvoicePaymentRequest req = new InvoicePaymentRequest();
                req.CustomerId = oDoc.CardCode;
                req.CustNum = cust.CustomerToken;
                req.InvoicePaymentDetails = new InvoicePaymentDetails[1];
                req.InvoicePaymentDetails[0] = new InvoicePaymentDetails();
                req.InvoicePaymentDetails[0].InvoiceInternalId = inv.InvoiceInternalId;
                req.InvoicePaymentDetails[0].PaidAmount = PaidAmtDiff;
                req.TotalPaidAmount = PaidAmtDiff;
                
                req.PaymentMethodType = "B1";
                req.PaymentMethodId = "";
                req.Software = Software;
                //var refNumFromConnect=ebiz.get
                req.RefNum = RefNum;//getSQLString(string.Format("select Cast([TransId[ as nvarchar(50)) from OINV where [DocEntry[ = '{0}'", oDoc.DocEntry).Replace("[","\""));
                InvoicePaymentResponse iresp = ebiz.AddInvoicePayment(securityToken, req);
                trace("add invoice payment request resp: " + iresp.Error + " Customer: " + oDoc.CardCode + "; Invoice: " + oDoc.DocNum + "; amount: " + amt + "; internalID: " + inv.InvoiceInternalId);

            }
            catch (Exception ex)
            {
                errorLog(ex);
            }
            updateCCConnectInvoice(oDoc.DocNum.ToString(), amt, inv.AmountDue.ToString());
            trace("Payment made in B1 for invoice: " + oDoc.DocNum + " in the amount of " + amt + ". Update eConnect invoicedue to " + inv.AmountDue);
        }
        catch (Exception ex)
        {
            if (ex.Message.ToLower() != "not found")
            {
                trace("HandleeConnectInvoice exception. " + ex.Message);
            }
            else
            {
                trace(invoiceID + " not uploaded to eConnect");
            }
        }
    }

    public void HandleeConnectSales(string OrderID, string amt, string CardCode = "", string UploadedBalance = "", string RefNum = "")
    {
        double AmtPaid = 0;
        try
        {
            if (amt == "")
            {
                double paid = getBalance(OrderID,"ORDR");
                if (paid == 0)
                    return;
                amt = paid.ToString();
            }
            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
            if (!oDoc.GetByKey(int.Parse(OrderID)))
            {
                trace("HandleeConnectSales order not found: " + OrderID);
                return;
            }
            string currency = oDoc.DocCurrency;
            string cardcode = "";
            string acctname = getAccountName(oDoc.CardCode);
            SecurityToken securityToken = getToken();

            SalesOrder oOrder = ebiz.GetSalesOrder(securityToken, oDoc.CardCode, "", oDoc.DocNum.ToString(), "");

            if (oOrder.AmountDue == 0)
            {
                trace("HandleeConnectSales order: " + oDoc.DocNum + " already set to 0 balance.");
                updateCCConnectSalesOrder(oDoc.DocNum.ToString(), amt, oOrder.AmountDue.ToString());
                return;
            }
            decimal newBal = (decimal)(Math.Round(oDoc.DocTotal - getDoubleValue(amt), 2));
            if ((decimal)Math.Round((double)oOrder.AmountDue, 2) == newBal)
            {
                trace("HandleeConnectSales order: " + OrderID + " balance is up to date. balance = " + oOrder.AmountDue + " new balance = " + newBal);
                return;
            }
            decimal PaidAmtDiff = Convert.ToDecimal(oOrder.AmountDue - newBal);

            //if (diff == 0)
            oOrder.Amount = (decimal)oDoc.DocTotal;
            oOrder.NotifyCustomer = true;
            if (newBal == 0)
                oOrder.Date = oDoc.DocDueDate.AddYears(1).ToString("MM/dd/yyyy");
            oOrder.AmountDue = newBal;
            oOrder.Software = Software;
            SalesOrderResponse resp = ebiz.UpdateSalesOrder(securityToken, oOrder, oDoc.CardCode, "", oDoc.DocNum.ToString(), "");
            try
            {
                oOrder.Amount = newBal;

                ApplicationTransactionRequest oAppTranReq = new ApplicationTransactionRequest();
                oAppTranReq.SoftwareId = Software;
                oAppTranReq.LinkedToTypeId = "SalesOrder";// oOrder.SalesOrderNumber;
                oAppTranReq.LinkedToInternalId = oOrder.SalesOrderInternalId;
                oAppTranReq.LinkedToExternalUniqueId = oOrder.SalesOrderNumber;

                Customer ocust = ebiz.GetCustomer(getToken(), oOrder.CustomerId, null);
                oAppTranReq.CustomerInternalId = ocust.CustomerInternalId;
                var TransactionType = ebiz.GetTransactionDetails(securityToken, RefNum).TransactionType;
                oAppTranReq.TransactionTypeId = TransactionType; 
                oAppTranReq.TransactionDate = DateTime.Now.ToShortDateString();
                oAppTranReq.TransactionId = RefNum;
                var respns = ebiz.AddApplicationTransaction(securityToken, oAppTranReq);
                ShowSystemLog(String.Format("Sales Order Payment {0} uploaded to connect successfully", oOrder.SalesOrderNumber));

                //trace("add invoice payment request resp: " + iresp.Error + " Customer: " + oDoc.CardCode + "; Invoice: " + oDoc.DocNum + "; amount: " + amt + "; internalID: " + oOrder.InvoiceInternalId);

            }
            catch (Exception ex)
            {
                errorLog(ex);
            }
            updateCCConnectSalesOrder(oDoc.DocNum.ToString(), amt, oOrder.AmountDue.ToString());
            trace("Payment made in B1 for sales: " + oDoc.DocNum + " in the amount of " + amt + ". Update eConnect Sales Due to " + oOrder.AmountDue);
        }
        catch (Exception ex)
        {
            if (ex.Message.ToLower() != "not found")
            {
                trace("HandleeConnectInvoice exception. " + ex.Message);
            }
            else
            {
                trace(OrderID + " not uploaded to eConnect");
            }
        }
    }
    private void RemoveRow(SAPbouiCOM.Matrix oMatrix, string id, string type)
    {
        for (int i = 1; i <= oMatrix.RowCount; i++)
        {
            if (getMatrixSelect(oMatrix, "10000127", i))
            {
                if (getMatrixItem(oMatrix, "1", i) == id && getMatrixItem(oMatrix, "45", i) == type)
                {
                    oMatrix.Columns.Item("10000127").Cells.Item(i).Click();
                    oMatrix.DeleteRow(i);
                    return;
                }

            }

        }
    }
    const string CC_CARDNUM = "*Card Number:";
    const string CC_EXPDATE = "*Exp. Date(MMYY):";
    const string CC_CARDCODE = "Card Code:";
    const string CC_CARDNAME = "Card Name:";
    const string EC_ACCOUNT = "*Checking Account:";
    const string EC_ROUTING = "*Routing Number:";
    const string cbRecurringBilling = "cbRBill";
    const string cbCurrencyCode = "cbCurr";

    private void AddCCPMFormField(SAPbouiCOM.Form form, int pane, bool bCreateExtra = false)
    {
        try
        {


            int edL = 110 + 15; //oItm.Left + oItm.Width;
            int edW = 250;
            int edT = 30;
            int edH = 15;
            int nGap = 16;
            SAPbouiCOM.Item oItm = null;
            SAPbouiCOM.Item oRefItm = null;
            if (form.TypeEx == FORMPAYMENTMEANS)
            {
                edT = form.Items.Item(tbCreditCard).Top + form.Items.Item(tbCreditCard).Height + edT;
            }
            if (theActiveForm.TypeEx != FORMINCOMINGPAYMENT)
                oRefItm = theActiveForm.Items.Item("4");
            else
                oRefItm = theActiveForm.Items.Item("5");
            edH = oRefItm.Height;
            nGap = edH + 1;
            edW = oRefItm.Width;
            edL = oRefItm.Width + 15;
            switch (theActiveForm.TypeEx)
            {

                case FORMSALESORDER: //Sales Order
                    oItm = addPaneItem(form, stRecordID, edL, edT, edW, edH, "Sales Order ID", SAPbouiCOM.BoFormItemTypes.it_STATIC, pane, 0);
                    break;
                case FORMINVOICE: //Invoice
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
                case "60090":
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
                if (form.TypeEx == FORMPAYMENTMEANS)
                {
                    string curcode = getFormItemVal(form, "8");
                    setComboValue(form, cbCurrencyCode, curcode);
                }
            }
            catch (Exception) { }


            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, cbPaymentMethod, edL, edT, edW, edH, "Payment Method:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, pane, 1);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCardNum, edL, edT, edW, edH, CC_CARDNUM, SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 2);
            oItm = form.Items.Add(lbDeclined, SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItm.FromPane = pane;
            oItm.ToPane = pane;
            oItm.Left = edW * 3;
            oItm.Top = edT;
            oItm.Width = 50;
            oItm.Height = edH;
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCardExpDate, edL, edT, edW, edH, CC_EXPDATE, SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 3);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCardCode, edL, edT, edW, edH, CC_CARDCODE, SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 4);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderName, edL, edT, edW, edH, "*Card/Acct. Holder:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 6);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, stCardName, edL, edT, edW, edH, CC_CARDNAME, SAPbouiCOM.BoFormItemTypes.it_STATIC, pane, 20);
            oItm = addPaneItem(form, editCheckRouting, edL, edT, edW, edH, EC_ROUTING, SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 101);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCheckAccount, edL, edT, edW, edH, EC_ACCOUNT, SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 100);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderAddr, edL, edT, edW, edH, "Address:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 1027);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderCity, edL, edT, edW, edH, "City:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 8);
            oItm = addPaneItem(form, editHolderState, edL + edW + edW / 3, edT, edW / 4, edH, "State:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 9, edW / 4);
            oItm = addPaneItem(form, editHolderZip, edL + 2 * edW, edT, edW / 4, edH, "Zip:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 10, edW / 4);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderEmail, edL, edT, edW, edH, "Email:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 108);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editMethodName, edL, edT, edW, edH, "Method Name:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 202);

            if (!bCreateExtra)
                edT = oItm.Top + nGap;
            else
                edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCAmount, edL, edT, edW, edH, "Amount:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 13);
            if (theActiveForm.TypeEx == FORMINCOMINGPAYMENT)
            {
                oItm.Enabled = false;
            }
            else
                oItm.Enabled = true;
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editPaymentDate, edL, edT, edW, edH, "Payment Date:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 6677);
            setFormEditVal(form, editPaymentDate, DateTime.Today.ToString("MM/dd/yyyy"));
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, cbTransactionType, edL, edT, edW, edH, "Transaction Type:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, pane, 14);
            SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbTransactionType).Specific;
            int fPane = 999;
            string fLB = "";
            switch (theActiveForm.TypeEx)
            {

                case FORMSALESORDER: //Sales Order
                    if (cfgProcessPreAuth == "Y")
                    {
                        oCB.ValidValues.Add("Pre-auth", "Pre-auth");
                        oCB.ValidValues.Add("Charge", "Charge");
                        oCB.Select("Pre-auth");
                    }
                    else
                    {
                        oCB.ValidValues.Add("Charge", "Charge");
                        oCB.ValidValues.Add("Pre-auth", "Pre-auth");
                        oCB.Select("Charge");
                    }
                    fPane = pane;
                    fLB = "Recurring Billing:";

                    setLabelCaption(form, stRecordID, getFormItemVal(theActiveForm, fidRecordID));
                    break;
                case FORMSALESQUOTE:
                    oCB.ValidValues.Add("Pre-auth", "Pre-auth");
                    oCB.Select("Pre-auth");
                    fPane = pane;
                    setLabelCaption(form, stRecordID, getFormItemVal(theActiveForm, fidRecordID));
                    break;
                case FORMINVOICE: //Invoice

                    oCB.ValidValues.Add("Charge", "Charge");
                    oCB.Select("Charge");

                    if (CheckCapture(form, ref theInvoiceID))
                    {
                        oCB.ValidValues.Add("Capture", "Capture");
                        oCB.Select("Capture");
                    }

                    break;
                case "179": //Credit Memo
                    oCB.ValidValues.Add("Credit", "Credit");

                    oCB.Select("Credit");
                    setLabelCaption(form, stRecordID, getFormItemVal(theActiveForm, fidRecordID));


                    break;
                case FORMINCOMINGPAYMENT:
                    //FORMTYPE.INCOMING_PAYMENT;
                    oCB.ValidValues.Add("Charge", "Charge");
                    oCB.ValidValues.Add("No Trans", "No Trans");
                    oCB.Select("Charge");
                    if (CheckCapture(form, ref theInvoiceID))
                    {
                        oCB.ValidValues.Add("Capture", "Capture");
                        oCB.Select("Capture");
                    }
                    break;
                case "420": //FORMTYPE.OUTGOING_PAYMENT;
                    oCB.ValidValues.Add("Credit", "Credit");
                    oCB.Select("Credit");
                    break;
                case "60090":
                case "60091":  //A/R Reserve Invoice
                    oCB.ValidValues.Add("Charge", "Charge");
                    oCB.Select("Charge");
                    setLabelCaption(form, stRecordID, getFormItemVal(theActiveForm, fidRecordID));
                    break;
                case "65300": //Down Payment Invoice
                    oCB.ValidValues.Add("Charge", "Charge");
                    oCB.Select("Charge");
                    setLabelCaption(form, stRecordID, getFormItemVal(theActiveForm, fidRecordID));
                    break;
            }
            if (cfgNoTransOnInv != "N" && theActiveForm.TypeEx != FORMINCOMINGPAYMENT && theActiveForm.TypeEx != FORMSALESQUOTE && theActiveForm.TypeEx != FORMSALESORDER)
                oCB.ValidValues.Add("No Trans", "No Trans");
            if (!bCreateExtra)
            {
                oItm = addPaneItem(form, cbRecurringBilling, edL + edW * 2, edT, edW / 4, edH, fLB, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, fPane, 134, edW / 2);
                if (fPane == pane)
                {
                    oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbRecurringBilling).Specific;
                    oCB.ValidValues.Add("Y", "Y");
                    oCB.ValidValues.Add("N", "N");
                    oCB.Select("N");
                }
            }

            edT = oItm.Top + nGap;

            if (bCreateExtra)
            {
                int margin = 10;
                SAPbouiCOM.Item oItem = form.Items.Add("Rect1", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
                oItem.Left = margin;
                oItem.Width = form.ClientWidth - 2 * margin;
                oItem.Top = margin;
                oItem.Height = edT + oRefItm.Height * 4 + margin;
                edT = oItem.Height + nGap;
                oItm = addPaneItem(form, "11", edL, edT, edW / 2, edH, "Overall Amount:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 15);
                oItm.Enabled = false;
                oItm = addPaneItem(form, "111", edW * 2, edT, edW / 3, edH, "", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 16);
                oItm.Enabled = true;
                oItm = addPaneItem(form, "110", edW * 3, edT, edW / 2, edH, "Discount %", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 17);
                oItm.Enabled = false;
                edT = oItm.Top + nGap;
                oItm = addPaneItem(form, "14", edL, edT, edW / 2, edH, "Balance Due:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 18);
                oItm.Enabled = false;
            }
            setFormEditVal(form, editCAmount, getFormItemVal(form, "14")); // getBalance());
            try
            {
                if (form.TypeEx == FORMPAYMENTMEANS)
                {
                    string fc = getFormItemVal(form, "48");
                    if (fc != "")
                        setFormEditVal(form, editCAmount, fc);
                }
            }
            catch (Exception) { }
            try
            {
                if (cfgAmountField != null && cfgAmountField != "")
                {
                    string fd = getFormItemVal(theActiveForm, cfgAmountField);
                    if (fd != "")
                        setFormEditVal(form, editCAmount, getMoneyString(fd));
                }
            }
            catch (Exception) { }
            setBalVal(form);
        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
    }
    private string getBalance()
    {
        try
        {
            switch (theActiveForm.TypeEx)
            {
                case FORMINCOMINGPAYMENT:
                    trace(theActiveForm.Title + " getBalance: " + getFormItemVal(theActiveForm, "12"));
                    return getMoneyString(getFormItemVal(theActiveForm, "12"));
                case FORMSALESQUOTE:
                case FORMSALESORDER:
                    trace(theActiveForm.Title + " getBalance: " + getFormItemVal(theActiveForm, "29"));
                    return getMoneyString(getFormItemVal(theActiveForm, "29"));
                default:
                    trace(theActiveForm.Title + " getBalance: " + getFormItemVal(theActiveForm, "29"));
                    return getMoneyString(getFormItemVal(theActiveForm, "29"));


            }
        }
        catch (Exception)
        { }
        return "";
    }
    private bool getInvoiceID(ref string invID)
    {
        invID = null;
        try
        {
            switch (theActiveForm.TypeEx)
            {
                case FORMINVOICE:
                    try
                    {
                        string id = theActiveForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0);
                        SAPbobsCOM.Documents oInvDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                        if (oInvDoc.GetByKey(int.Parse(id)))
                        {
                            invID = id;
                            return true;
                        }
                    }
                    catch (Exception)
                    {

                    }
                    break;
                case FORMINCOMINGPAYMENT:
                    try
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)theActiveForm.Items.Item("20").Specific;
                        string CustomerID = getCustomerID();

                        int invCount = 0;
                        for (int i = 1; i <= oMatrix.RowCount; i++)
                        {
                            if (getMatrixSelect(oMatrix, "10000127", i))
                            {
                                string type = getMatrixItem(oMatrix, "45", i);
                                string DocNum = getMatrixItem(oMatrix, "1", i);
                                string DocDate = getMatrixItem(oMatrix, "21", i);
                                string amount = getMatrixItem(oMatrix, "24", i);
                                if (type == "13")
                                {
                                    invCount++;
                                    invID = getDocEntry("13", DocNum, CustomerID).ToString();


                                }

                            }

                        }
                        if (invCount == 1 && invID != "")
                            return true;
                        else
                            return false;
                    }
                    catch (Exception ex) { errorLog(ex); }
                    break;

            }
        }
        catch (Exception)
        { }
        return true;
    }
    private bool CheckCapture(SAPbouiCOM.Form form, ref string invID)
    {
        string sql = "";

        try
        {


            preauthCCTRAN = new PREAUTHTRANS();
            if (!getInvoiceID(ref invID))
            {


                return false;
            }
            if (invID == null || invID == "")
                return false;
            string customerID = getCustomerID();
            if (customerID != "" && invID != "")
            {
                updateTransInvoiceID(invID, customerID);
            }
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string id = invID;
            double amt = 0;
            sql = string.Format("select \"U_recDate\", \"U_amount\", \"U_refNum\", \"U_CCAccountID\", \"U_MethodID\", \"U_customerID\", \"U_OrderID\"\r\n" +
            " from \"@CCTRANS\" a where \"U_command\" = 'cc:authonly' and \"U_InvoiceID\"='{0}' " +
            " order by a.\"U_refNum\", a.\"U_recDate\" ASC", id);
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                while (!oRS.EoF)
                {
                    amt += getDoubleValue((string)oRS.Fields.Item(1).Value);
                    preauthCCTRAN.amount = amt.ToString();
                    CCTRAN cctran = new CCTRAN();
                    cctran.amount = (string)oRS.Fields.Item(1).Value;
                    cctran.refNum = (string)oRS.Fields.Item(2).Value;
                    cctran.OrderID = (string)oRS.Fields.Item(6).Value;
                    cctran.InvoiceID = invID;
                    preauthCCTRAN.CCAcountID = (string)oRS.Fields.Item(3).Value;
                    preauthCCTRAN.refNum = (string)oRS.Fields.Item(2).Value;
                    preauthCCTRAN.paymentKey = getPaymentKey((string)oRS.Fields.Item(4).Value, (string)oRS.Fields.Item(5).Value);
                    if (!IsCaptured(preauthCCTRAN.refNum))
                    {
                        preauthCCTRAN.CCTRANS.Add(cctran);
                    }
                    oRS.MoveNext();
                }
            }
            else
            {
                trace("CheckCapture order has no preauth record");
                return false;
            }
            if (preauthCCTRAN.CCTRANS.Count() == 0)
            {
                trace("CheckCapture preauthCCTRAN.CCTRANS.Count() == 0");
                return false;
            }
            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);

        }
        return false;
    }

    private bool IsCaptured(string refNum)
    {
        string sql = "";
        try
        {

            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            sql = string.Format("select 1 from \"@CCTRANS\" a where \"U_command\" in ( 'cc:void', 'cc:capture' ) and \"U_refNum\"='{0}' ", refNum);
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                return true;
            }
            else
                return false;
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return false;
    }
    private void PaymentMeansImportHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {

        try
        {
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

                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        try
                        {
                            if (cfgImportCust == "Y")
                            {
                                SAPbouiCOM.Item oItem = form.Items.Add(btnProcess, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                                oItem.Left = form.Items.Item("2").Left + form.Items.Item("2").Width + 3;
                                oItem.Width = form.Items.Item("2").Width;
                                oItem.Top = form.Items.Item("2").Top;
                                oItem.Height = form.Items.Item("2").Height;
                                oItem.Enabled = true;
                                setButtonCaption(form, btnProcess, "Import");
                            }
                        }
                        catch (Exception)
                        {

                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        try
                        {
                            switch (pVal.ItemUID)
                            {
                                case btnProcess:
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item("112").Specific;

                                    string cardNum = getMatrixItem(oMatrix, "63", 1);
                                    string expdate = getMatrixItem(oMatrix, "64", 1);
                                    if (cardNum.IndexOf("X") >= 0 || cardNum.Length < 15)
                                    {
                                        showMessage("Invalid card number:" + cardNum);
                                        return;
                                    }
                                    if (SBO_Application.MessageBox("Import credit card " + cardNum + "; expdate: " + expdate + " to EBizCharge?", 1, "Yes", "No") == 1)
                                    {
                                        string CustomerID = getCustomerID();
                                        Customer c = AddCustomer(CustomerID);
                                        PaymentMethodProfile pm = new PaymentMethodProfile();
                                        pm.CardExpiration = expdate.Replace("/", "");
                                        pm.CardNumber = cardNum;
                                        SecurityToken token = getToken("");
                                        string mid = ebiz.AddCustomerPaymentMethodProfile(token, c.CustomerInternalId, pm);
                                        trace("add payment method returns: " + mid);
                                        c = ebiz.GetCustomer(token, CustomerID, "");
                                        foreach (PaymentMethodProfile p in c.PaymentMethodProfiles)
                                        {
                                            if (p.MethodID == mid)
                                            {
                                                pm = p;
                                                trace("Customer Payment profile: " + pm.CardNumber);
                                            }
                                        }

                                        CCCUST cccust = new CCCUST();
                                        cccust.CustomerID = CustomerID;
                                        cccust.CustNum = c.CustomerToken;
                                        cccust.MethodID = pm.MethodID;
                                        cccust.GroupName = getGroupName(CustomerID);
                                        string acctid = "";
                                        string bid = getBranchID();
                                        string cardname = getCardName(cccust.GroupName, pm.CardType, "", ref acctid, bid);
                                        cccust.CCAccountID = acctid;
                                        cccust.CardName = cardname;
                                        cccust.active = 'Y';
                                        cccust.email = c.Email;
                                        cccust.firstName = c.BillingAddress.FirstName;
                                        cccust.lastName = c.BillingAddress.LastName;
                                        cccust.street = c.BillingAddress.Address1;
                                        cccust.city = c.BillingAddress.City;
                                        cccust.state = c.BillingAddress.State;
                                        cccust.zip = c.BillingAddress.ZipCode;
                                        cccust.expDate = pm.CardExpiration;
                                        cccust.cardCode = pm.CardCode;
                                        int id = getNextTableID("@CCCUST");
                                        string desc = c.BillingAddress.FirstName + " " + c.BillingAddress.LastName;
                                        cccust.cardNum = pm.CardNumber;
                                        cccust.methodDescription = id.ToString() + "_" + pm.CardNumber + " " + pm.CardExpiration + "(" + desc + ")";
                                        cccust.cardType = pm.CardType;
                                        insert(cccust);
                                        showMessage("Card imported");
                                    }
                                    break;

                            }

                        }
                        catch (Exception ex)
                        {
                            errorLog(ex);
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

}
