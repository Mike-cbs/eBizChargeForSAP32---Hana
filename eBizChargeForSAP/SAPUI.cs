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
using eBizChargeForSAP.ServiceReference1;

partial class SAP
{

    //private string BtnRefID;
    string edFormRefID = "33";
    string stFormRefID = "34";

    string btnPreAuth = "btnPreAuth";
    const string btnCharge = "btnCharge";
    const string btnCredit = "btnCredit";
    const string cbActive = "cbActive";
    const string FORMBPARTNER = "134";
    const string FORMCREDITMEMO = "179";
    const string FORMPAYMENTMEANS = "146";
    const string FORMPICKMANAGER = "81";
    const string FORMSALESQUOTE = "149";
    const string FORMSALESORDER = "139";
    const string FORMDOWNPAYMENT = "65300";
    const string FORMINCOMINGPAYMENT = "170";
    const string FORMINVOICE = "133";
    const string formExportBP = "ExportBP";
    const string formSyncUpload = "SyncUp";
    const string formExportItemMaster = "ExportItm";

    const string formEmailPaymentSO = "EmailPaySO";
    const string formExportSO = "ExportSO";
    const string formEmailEstimates = "EmailEsti";
    const string formEmailPaymentPending = "EmailPaymtPend";
    const string formEmailPaymentReceived = "EmailPaymtRec";
    const string formSyncDownload = "SyncD";
    //string btnPayM= "btnPayM";
    const string editCheckAccount = "ChkAcct";
    const string editCheckRouting = "ChkRout";


    const string editCardNum = "CardNum";
    const string editCardExpDate = "ExpDt";
    const string editCardCode = "CardCode";
    const string editCAmount = "CAmount";
    const string editPaymentDate = "PDate";
    //string lbCardNum = "lbCardNum";
    const string cbTransactionType = "cbTType";
    const string tbCreditCard = "tbCCLog";
    string mxtCCLog = "mtxCCLog";
    string dtCCTransLog = "CCTRANSLOG";
    string dtCCImport = "CCIMPORT";
    string strCCTransTableName = "CCTRANS";
    public string fidCustID = "4";
    public string fidRecordID = "8";
    public string fidCustName = "54";
    public string fidBalance = "33";
    public string fidSubTotal = "22";
    public string fidTax = "27";
    public string fidDiscount = "42";
    public string fidFreight = "89";
    public string fidRemark = "16";
    public string fidRemarkStatic = "17";
    public string fidDesc = "18";
    public string fidMatrix = "38";
    public string fidStatus = "81";
    public SAPbouiCOM.SboGuiApi SboGuiApi = null;
    string strHeaderApproved = "**Approved. ";
    string strHeaderCredited = "**Credited. ";
    string strHeaderCaptured = "**Captured. ";
    string strHeaderCharged = "**Charged. ";
    const string cbPaymentMethod = "cbPayM";
    const string editHolderName = "edHName";
    const string editHolderAddr = "edHAddr";
    const string editHolderCity = "edHCity";
    const string editHolderState = "edHState";
    const string editHolderZip = "edHZip";
    const string editHolderEmail = "ccemail";
    const string editMethodName = "edMName";
    const string cbCCAccount = "cbCCAcct";
    const string stCardName = "stCName";
    const string lblstRecID = "stRecID";

    //const string cbCCSave = "cbCCSave";

    Customer theSAPCustomer;
    SAPbouiCOM.Form theActiveForm;
    bool bVoidCloseForm = false;
    SAPbouiCOM.Form formPayment;
    SAPbouiCOM.Form formPaymentList;
    private bool PMItemValidate(SAPbouiCOM.Form form)
    {
        bool bRet = true;
        string err = "";
        try
        {

            if (form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                return true;

            if (form.TypeEx == FORMSALESORDER)
            {
                if (getFormItemVal(form, "12") == "")
                {
                    err = err + "Delivery date is required\r\n";
                }
            }

            if (err != "")
            {
                showMessage(err);
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
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;

    }


    public void VoidCloseFormProc()
    {
        try
        {

            formPayment.Close();
            formPaymentList.Close();
            SBO_Application.Menus.Item("1289").Activate();
            SBO_Application.Menus.Item("1288").Activate();

        }
        catch (Exception)
        {
            //errorLog(ex);
        }
        finally
        {
            voidCCTRANS(confirmNum);
        }
    }


    private bool CheckStatus(SAPbouiCOM.Form form)
    {
        if (getFormItemVal(form, fidStatus) == "1" || getFormItemVal(form, fidStatus) == "2")
            return true;
        return false;
    }
    private void AddTab(SAPbouiCOM.Form form, int pane, string title, string refItemID, string tabID = tbCreditCard, int tabPane = 0, bool bVisible = true, int width = 0, int left = 0)
    {
        try
        {
            SAPbouiCOM.Item oItem;

            SAPbouiCOM.Item oItemRef;

            oItemRef = form.Items.Item(refItemID);
            oItem = form.Items.Add(tabID, SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            oItem.Top = oItemRef.Top;
            oItem.Height = oItemRef.Height;

            oItem.Left = oItemRef.Left + oItemRef.Width + left;
            oItem.Width = oItemRef.Width + width;
            oItem.AffectsFormMode = false;
            try
            {

                if (refItemID == tbeBizCharge)
                {
                    oItem.Left = form.Items.Item("3").Left + 8;
                    oItem.Top = oItemRef.Top + 28;
                }
            }
            catch (Exception ex)
            {

            }
            oItem.FromPane = tabPane;
            oItem.ToPane = tabPane + 3;
            oItem.Visible = bVisible;
            SAPbouiCOM.Folder oFolder = (SAPbouiCOM.Folder)oItem.Specific;
            oFolder.Caption = title;
            if (refItemID != tbeBizCharge)
                oFolder.GroupWith(oItemRef.UniqueID);
            oFolder.Pane = pane;

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    const string H_CUSTOMERID = "hCustID";
    const string H_RECORDID = "hRecID";
    const string editEmail = "edEmail";
    private void AddCCLogMatrix(SAPbouiCOM.Form form, int pane, int top = 0)
    {
        try
        {
            SAPbouiCOM.Item oItem = form.Items.Add(H_CUSTOMERID, SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.FromPane = 1000;
            oItem.ToPane = 1000;
            oItem.AffectsFormMode = false;
            oItem = form.Items.Add(H_RECORDID, SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.FromPane = 1000;
            oItem.ToPane = 1000;
            oItem.AffectsFormMode = false;
            oItem = form.Items.Item("4");
            if (form.TypeEx == FORMBPARTNER)
                oItem = form.Items.Item("7");
            SAPbouiCOM.Item oItemRef = form.Items.Item(tbCreditCard);
            int edW = oItem.Width;
            int edT = oItemRef.Top + top + oItemRef.Height + 10;
            int edL = 20 + oItem.Width;
            oItem = addPaneItem(form, editEmail, edL, edT, edW, oItem.Height, "Receipt email:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 19877);
            edT = edT + oItem.Height + 2;
            oItem = form.Items.Add(mxtCCLog, SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oItem.FromPane = pane;
            oItem.ToPane = pane;
            oItem.AffectsFormMode = false;


            oItem.Top = edT;
            oItem.Left = 20;
            oItem.Height = form.ClientHeight / 3;
            oItem.Width = form.ClientWidth - 60;

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item(mxtCCLog).Specific;
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Add("RefNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ref. No.";
            oColumn.Width = 50;
            oColumn.AffectsFormMode = false;

            oColumn = oMatrix.Columns.Add("PaymentID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "";
            SAPbouiCOM.LinkedButton btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_JournalPosting).ToString();
            oColumn.Width = 0;
            oColumn.AffectsFormMode = false;


            oColumn = oMatrix.Columns.Add("OrderID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Order";
            btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_Order).ToString();
            oColumn.Width = 50;
            oColumn.AffectsFormMode = false;
            oColumn = oMatrix.Columns.Add("OrderNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Order No.";
            oColumn.Width = 50;
            oColumn = oMatrix.Columns.Add("InvoiceID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Invoice";
            btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_Invoice).ToString();
            oColumn.Width = 50;
            oColumn.AffectsFormMode = false;
            oColumn = oMatrix.Columns.Add("InvoiceNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Invoice No.";
            oColumn.Width = 50;
            oColumn = oMatrix.Columns.Add("CMID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Credit Memo";
            btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo).ToString();
            oColumn.Width = 50;
            oColumn.AffectsFormMode = false;

            oColumn = oMatrix.Columns.Add("command", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Command";
            oColumn.Width = 50;
            oColumn.AffectsFormMode = false;

            oColumn = oMatrix.Columns.Add("Amount", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Amount";
            oColumn.Width = 50;
            oColumn.AffectsFormMode = false;

            oColumn = oMatrix.Columns.Add("Desc", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Description";
            oColumn.Width = 50;
            oColumn.AffectsFormMode = false;

            oColumn = oMatrix.Columns.Add("CardHolder", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Card Holder";
            oColumn.Width = 50;
            oColumn.AffectsFormMode = false;


            oColumn = oMatrix.Columns.Add("crCardNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Card No.";
            oColumn.Width = 50;
            oColumn.AffectsFormMode = false;

            /*
            oColumn = oMatrix.Columns.Add("result", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Result";
            oColumn.Width = 50;
            */
            oColumn = oMatrix.Columns.Add("error", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Result";
            oColumn.Width = 50;
            oColumn.AffectsFormMode = false;

            oColumn = oMatrix.Columns.Add("avs", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "AVS";
            oColumn.Width = 150;
            oColumn.AffectsFormMode = false;

            oColumn = oMatrix.Columns.Add("CardCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Card Code";
            oColumn.Width = 150;
            oColumn.AffectsFormMode = false;

            oColumn = oMatrix.Columns.Add("CardLevel", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Card Level";
            oColumn.Width = 150;
            oColumn.AffectsFormMode = false;

            oColumn = oMatrix.Columns.Add("custNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Customer Number";
            oColumn.Width = 50;
            oColumn.AffectsFormMode = false;

            oColumn = oMatrix.Columns.Add("CCAID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Credit Card";
            oColumn.Width = 0;
            oColumn.AffectsFormMode = false;

            oColumn = oMatrix.Columns.Add("recdate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Process Date";
            oColumn.Width = 100;
            oColumn.AffectsFormMode = false;

            oColumn = oMatrix.Columns.Add("ccTRANSID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "ID";
            oColumn.Width = 10;
            oColumn.AffectsFormMode = false;


            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;






            //populateLog(form);
        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
    }

    private void populateLog(SAPbouiCOM.Form form)
    {
        string sql = "";
        try
        {

            //SAPbouiCOM.EditText customer = form.Items.Item(fidCustID).Specific;
            if (form.Items.Item(mxtCCLog) == null)
                return;
            form.Items.Item(mxtCCLog).AffectsFormMode = false;
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item(mxtCCLog).Specific;
            oMatrix.Clear();

            //oMatrix.AutoResizeColumns();
            SAPbouiCOM.DataTable oDts = getDataTable(form);

            sql = string.Format("select " +
            " (select CAST(\"DocEntry\" AS VARCHAR) from OINV where CAST(\"DocEntry\" AS VARCHAR) = \"U_InvoiceID\") as InvoiceID," +
            " (select CAST(\"DocNum\" AS VARCHAR) from OINV where CAST(\"DocEntry\" AS VARCHAR) = \"U_InvoiceID\") as InvoiceNum,  " +
            " (select CAST(\"DocNum\" AS VARCHAR) from ORDR where CAST(\"DocEntry\" AS VARCHAR) = \"U_OrderID\") as OrderNum, " +
            " (select CAST(\"DocEntry\" AS VARCHAR) from ORDR where CAST(\"DocEntry\" AS VARCHAR) = \"U_OrderID\") as OrderID, " +
            " \"U_CCAccountID\" as CCAID, " +
            "\"U_PaymentID\", \"U_DownPaymentInvID\", \"U_CreditMemoID\" as CMID, \"U_custNum\", \"U_customerID\",	" +
            " \"U_customerName\",\"U_CardHolder\",\"U_crCardNum\",\"U_Description\",\"U_recID\",\"U_acsUrl\",\"U_authAmount\", " +
            " \"U_authCode\",\"U_avsResult\",\"U_avsResultCode\",\"U_batchNum\",\"U_batchRef\",\"U_cardCodeResult\"," +
            "\"U_cardCodeRCode\",\"U_cardLevelResult\", \"U_cardLevelRCode\",\"U_conversionRate\"," +
            " \"U_convertedAmount\",\"U_convertedCurrency\",\"U_error\",\"U_errorCode\",\"U_isDuplicate\", " +
            " \"U_payload\",\"U_profilerScore\",\"U_profilerResponse\"," +
            "\"U_profilerReason\",\"U_refNum\",\"U_remainingBalance\",\"U_result\",\"U_resultCode\",\"U_status\"," +
            "\"U_statusCode\",	\"U_vpasResultCode\", \"U_recDate\", \"U_command\", \"U_amount\", \"DocEntry\" " +
            " from \"@CCTRANS\" {0} order by \"DocEntry\" desc", strWhere);
            // trace(sql);
            oDts.ExecuteQuery(sql);
            BindMatrix(oMatrix, "RefNum", "U_refNum");
            BindMatrix(oMatrix, "PaymentID", "U_PaymentID");
            BindMatrix(oMatrix, "OrderID", "OrderID");
            BindMatrix(oMatrix, "InvoiceID", "InvoiceID");
            BindMatrix(oMatrix, "OrderNum", "OrderNum");
            BindMatrix(oMatrix, "InvoiceNum", "InvoiceNum");
            BindMatrix(oMatrix, "CMID", "CMID");
            BindMatrix(oMatrix, "CCAID", "CCAID");
            BindMatrix(oMatrix, "custNum", "U_custNum");
            BindMatrix(oMatrix, "command", "U_command");
            BindMatrix(oMatrix, "Amount", "U_Amount");
            BindMatrix(oMatrix, "Desc", "U_Description");
            BindMatrix(oMatrix, "CardHolder", "U_CardHolder");
            BindMatrix(oMatrix, "crCardNum", "U_crCardNum");

            BindMatrix(oMatrix, "error", "U_error");
            BindMatrix(oMatrix, "avs", "U_avsResult");
            BindMatrix(oMatrix, "CardCode", "U_cardCodeResult");
            BindMatrix(oMatrix, "CardLevel", "U_cardLevelResult");
            BindMatrix(oMatrix, "recdate", "U_recDate");
            BindMatrix(oMatrix, "ccTRANSID", "DocEntry");
            oMatrix.LoadFromDataSource();
            string cardcode = getCustomerID();
            setFormEditVal(form, editEmail, getCustEmail(cardcode));
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void BindMatrix(SAPbouiCOM.Matrix oMatrix, string mtCol, string dtCol, string dt = null, bool editable = false)
    {
        try
        {
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item(mtCol);
            oColumn.Editable = editable;
            if (dt == null)
            {
                oColumn.DataBind.Bind(dtCCTransLog, dtCol);
            }
            else
                oColumn.DataBind.Bind(dt, dtCol);


        }
        catch (Exception ex)
        {
            errorLog("Can not bind " + mtCol + " to " + dtCol + ".  error: " + ex.Message);
        }

    }
    private SAPbouiCOM.DataTable getDataTable(SAPbouiCOM.Form form)
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = form.DataSources.DataTables.Item(dtCCTransLog);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
            oDT = form.DataSources.DataTables.Add(dtCCTransLog);

        }

        return oDT;
    }
    private string getRSValue(SAPbobsCOM.Recordset oRS, string id)
    {
        string ret = "";
        try
        {
            ret = oRS.Fields.Item(id).Value.ToString();

        }
        catch (Exception ex)
        {
            errorLog("'" + id + "' not found in record set.");
            errorLog(ex);
        }
        return ret;
    }
    private SAPbouiCOM.Item addPaneButton(SAPbouiCOM.Form form, string id, int l, int t, int w, int h, string label, int pane)
    {
        SAPbouiCOM.Item oItem = null;
        try
        {
            oItem = form.Items.Add(id, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            SAPbouiCOM.Item oItm = form.Items.Item(id);
            oItm.FromPane = pane;
            oItm.ToPane = pane;
            oItm.Left = l;
            oItm.Top = t;
            oItm.Width = w;
            oItm.Height = h;
            setButtonCaption(form, id, label);
        }
        catch (Exception)
        {

        }
        finally
        {
            /*
            try
            {
                if(oItem != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                oItem = null;
            }
            catch (Exception) { }
            */
        }
        return oItem;
    }
    private SAPbouiCOM.Item addPaneItem(SAPbouiCOM.Form form, string id, int l, int t, int w, int h, string label, SAPbouiCOM.BoFormItemTypes type, int pane, int order, int lbw = 0)
    {
        SAPbouiCOM.Item oItem = null;
        try
        {
            if (lbw == 0)
                lbw = w;
            int toPane = pane + 1;
            int fromPane = pane;
            if (label == CC_CARDNUM || label == CC_CARDCODE || label == CC_EXPDATE || label == CC_CARDNAME)
            {
                toPane = pane;
                fromPane = pane;

            }
            else if (label == EC_ROUTING || label == EC_ACCOUNT)
            {
                toPane = pane + 1;
                fromPane = pane + 1;
            }
            oItem = form.Items.Add(id, type);
            oItem.AffectsFormMode = false;

            SAPbouiCOM.Item oItm = form.Items.Item(id);
            oItm.FromPane = fromPane;
            oItem.ToPane = toPane;
            oItm.Left = l;
            oItm.Top = t;
            oItm.Width = w;
            oItm.Height = h;
            if (type == SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            {
                SAPbouiCOM.Button btn = (SAPbouiCOM.Button)form.Items.Item(id).Specific;
                btn.Caption = label;
            }
            else
            {
                if (label != "")
                {
                    string lb = "LB" + order.ToString();
                    oItm = form.Items.Add(lb, SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oItm.ToPane = toPane;
                    oItm.FromPane = fromPane;
                    oItm.Left = l - lbw;
                    oItm.Top = t;
                    oItm.Width = lbw;
                    oItm.Height = h;
                    setLabelCaption(form, lb, label);
                }
            }
            if (type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            {
                SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)form.Items.Item(id).Specific;
                oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return oItem;
    }
    private void addLabel(SAPbouiCOM.Form form, string id, int l, int t, int w, int h, string label)
    {
        try
        {
            form.Items.Add(id, SAPbouiCOM.BoFormItemTypes.it_STATIC);
            SAPbouiCOM.Item oItm = form.Items.Item(id);
            oItm.Left = l;
            oItm.Top = t;
            oItm.Width = w;
            oItm.Height = h;
            setLabelCaption(form, id, label);
        }
        catch (Exception)
        {

        }
    }

    private void setMatrixItem(SAPbouiCOM.Matrix oMatrix, string Col, int Row, string val)
    {
        string sValue = "";
        try
        {
            if (oMatrix.Columns.Item(Col) == null)
            {
                trace("setMatrixItem Cannot found col: " + Col);
                return;
            }
            if (oMatrix.Columns.Item(Col).Cells.Item(Row) == null)
            {
                trace("setMatrixItem Cannot found col: " + Col + ", row: " + Row);
                return;
            }
            SAPbouiCOM.Cell cell = oMatrix.Columns.Item(Col).Cells.Item(Row);
            if (cell.Specific is SAPbouiCOM.EditText)
            {
                SAPbouiCOM.EditText oEt = (SAPbouiCOM.EditText)oMatrix.Columns.Item(Col).Cells.Item(Row).Specific;
                oEt.Value = val;
                return;
            }
            if (cell.Specific is SAPbouiCOM.StaticText)
            {
                SAPbouiCOM.StaticText oStatic = (SAPbouiCOM.StaticText)oMatrix.Columns.Item(Col).Cells.Item(Row).Specific;
                oStatic.Caption = val;
                return;
            }
            if (cell.Specific is SAPbouiCOM.ComboBox)
            {
                SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(Col).Cells.Item(Row).Specific;
                oCB.Select(val);
                return;
            }
        }
        catch (Exception ex)
        {
            errorLog("Can't find row:" + Row.ToString() + ",col:" + Col + "\r\n" + ex.Message);
        }


    }

    private string getMatrixItem(SAPbouiCOM.Matrix oMatrix, string Col, int Row)
    {
        string sValue = "";
        try
        {
            if (oMatrix.Columns.Item(Col) == null)
            {
                trace("getMatrixItem Cannot found col: " + Col);
                return "";
            }
            if (oMatrix.Columns.Item(Col).Cells.Item(Row) == null)
            {
                trace("getMatrixItem Cannot found col: " + Col + ", row: " + Row);
                return "";
            }
            SAPbouiCOM.Cell cell = oMatrix.Columns.Item(Col).Cells.Item(Row);
            if (cell.Specific is SAPbouiCOM.EditText)
            {
                SAPbouiCOM.EditText oEt = (SAPbouiCOM.EditText)oMatrix.Columns.Item(Col).Cells.Item(Row).Specific;
                return oEt.Value;
            }
            if (cell.Specific is SAPbouiCOM.StaticText)
            {
                SAPbouiCOM.StaticText oStatic = (SAPbouiCOM.StaticText)oMatrix.Columns.Item(Col).Cells.Item(Row).Specific;
                return oStatic.Caption;
            }
            if (cell.Specific is SAPbouiCOM.ComboBox)
            {
                SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(Col).Cells.Item(Row).Specific;
                return oCB.Value;
            }
        }
        catch (Exception ex)
        {
            errorLog("Can't find row:" + Row.ToString() + ",col:" + Col + "\r\n" + ex.Message);
        }
        return sValue;

    }

    private bool getMatrixSelect(SAPbouiCOM.Matrix oMatrix, string Col, int Row)
    {
        try
        {
            SAPbouiCOM.CheckBox oCB = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item(Col).Cells.Item(Row).Specific; //Cast the Cell of the matrix to the respective type , in this case EditText
            return oCB.Checked; //Get the value form the EditText
        }
        catch (Exception ex)
        {
            errorLog("Can't find row:" + Row.ToString() + ",col:" + Col + "\r\n" + ex.Message);
        }
        return false;

    }
    private double getDoubleValue(string s)
    {
        double ret = 0;
        try
        {
            string str = getMoneyString(s);
            string uiSep = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            str = str.Replace(".", uiSep);

            ret = double.Parse(str);

        }
        catch (Exception ex)
        {
            errorLog("getDoubleValue string= " + s + ". error: " + ex.Message);
        }
        return ret;
    }
    private double getItemDoubleValue(SAPbouiCOM.Form form, string id)
    {
        double ret = 0;
        try
        {
            if (form.Items.Item(id) == null)
            {
                trace("Cann't found double value: " + form.Title + "," + id);
                return 0;
            }
            string s = getFormItemVal(form, id);
            if (s != "")
            {

                ret = getDoubleValue(getMoneyString(s));
            }


        }
        catch (Exception ex)
        {
            errorLog("Item ID: " + id + " error. " + ex.Message);
        }
        return ret;
    }

    public string getMoneyString(string CurrencyString)
    {
        try
        {
            CurrencyString = CurrencyString.Replace("(", "-");
            StringBuilder TempStr = new System.Text.StringBuilder();
            for (int i = 0; i < CurrencyString.Length; i++)
            {
                if (CurrencyString[i] == '-' || CurrencyString[i] == '.' || Char.IsDigit(CurrencyString[i]))
                {
                    TempStr.Append(CurrencyString[i]);
                }

            }
            return TempStr.ToString();
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return "";
    }
    /*
    private string getFormItemVal(SAPbouiCOM.Form form, string id)
    {
        string ret = "";
        try
        {
            SAPbouiCOM.EditText oET = (SAPbouiCOM.EditText)form.Items.Item(id).Specific;
            ret = oET.Value;
        }
        catch (Exception)
        {
        }
        return ret;
    }*/
    private void showMessage(string msg)
    {
        try
        {
            if (!bRunJob)
            {
                if (!bAuto)
                {
                    //    SBO_Application.MessageBox(msg);
                }
                else
                {
                    showStatus(msg, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                }
            }
            else
                trace(msg);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    public void showStatus(string msg, SAPbouiCOM.BoMessageTime t = SAPbouiCOM.BoMessageTime.bmt_Medium, bool error = true)
    {
        try
        {
            if (!bRunJob)
            {
                SBO_Application.SetStatusBarMessage(msg, t, error);
            }
            trace(msg);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }

    private void setComboValue(SAPbouiCOM.Form form, string id, string val)
    {
        string ret = "";
        SAPbouiCOM.ComboBox oItem = null;
        try
        {
            oItem = (SAPbouiCOM.ComboBox)form.Items.Item(id).Specific;
            oItem.Select(val);
        }
        catch (Exception)
        {
        }
        finally
        {
            /*
            try
            {
                if (oItem != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                oItem = null;
            }
            catch (Exception) { }
            */
        }

    }


    private void hideItem(SAPbouiCOM.Form form, string id)
    {
        try
        {
            form.Items.Item(id).Visible = false;
        }
        catch (Exception)
        {
        }
    }
    private void setButtonCaption(SAPbouiCOM.Form form, string id, string cap)
    {
        SAPbouiCOM.Button oItem = null;
        try
        {
            oItem = (SAPbouiCOM.Button)form.Items.Item(id).Specific;
            oItem.Caption = cap;
        }
        catch (Exception)
        {
        }
        finally
        {
            /*
            try
            {
                if (oItem != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                oItem = null;
            }
            catch (Exception) { }
            */
        }

    }

    private void setLabelCaption(SAPbouiCOM.Form form, string id, string cap)
    {

        SAPbouiCOM.StaticText oItem = null;
        try
        {
            oItem = (SAPbouiCOM.StaticText)form.Items.Item(id).Specific;
            oItem.Caption = cap;
        }
        catch (Exception)
        {
        }
        finally
        {
            /*
            try
            {
                if (oItem != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                oItem = null;
            }
            catch (Exception) { }
            */
        }

    }

    private bool GetTotalAmount(SAPbouiCOM.Form form, ref string amt)
    {
        try
        {
            string a = getFormItemVal(form, editCAmount);
            if (a == "")
            {
                showMessage("Please enter charge amount.");
                return false;
            }


            try
            {

                if (getItemDoubleValue(form, "14") < getDoubleValue(getMoneyString(a)))
                {
                    showMessage("The payment amount cannot exceed the current balance due..");
                    return false;
                }
                amt = a;
            }
            catch (Exception)
            {
                showMessage("Invalid amount.  Please enter numerical value only.");
                return false;
            }


            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }
        return false;
    }


    public string GetFieldVal(SAPbobsCOM.Recordset rs, int index)
    {
        try
        {
            return (string)rs.Fields.Item(index).Value;

        }
        catch (Exception ex)
        {
            errorLog("getFieldVal exceptionon Index " + index + ", error=" + ex.Message);

        }
        return "";
    }
    public static bool bHasFileAccess = true;
    public static void trace(string msg)
    {
        try
        {

            errorLog("Trace: " + msg);

        }
        catch (Exception)
        {

        }
    }
    public static void errorLog(Exception ex)
    {
        try
        {
            errorLog(ex.Message + "\r\n\r\n" + ex.StackTrace);
        }
        catch (Exception)
        {
        }
    }
    public static void errorLog(string msg)
    {
        try
        {
            try
            {
                if (GetConfig("trace").ToLower() == "off")
                    return;
                eBizChargeForSAP.B1InstallService.B1InstallServiceSoapClient install = new eBizChargeForSAP.B1InstallService.B1InstallServiceSoapClient();
                eBizChargeForSAP.B1InstallService.SAPError error = new eBizChargeForSAP.B1InstallService.SAPError();
                error.Company = oCompany.CompanyName + "-" + ebizVersion;
                if (bRunJob)
                    error.Company = oCompany.CompanyName + "-" + ebizVersion + "_Job";

                error.ErrorMsg = msg;
                error.DBName = oCompany.CompanyDB;
                error.AppUser = oCompany.UserName;
                install.LogError(error);
            }
            catch (Exception) { }
            /*
            
            if (B1Info.SubnetMask == "traceoff")
                return;
            if (!bHasFileAccess)
                return;
                */
            string logFileName = @"c:\CBS\eBizChargeForSAP_Hana.log";
            if (bRunJob)
                logFileName = @"c:\CBS\eBizChargeForSAP_Job.log";


            msg = string.Format("{0}:\r\n\t{1}\r\n\r\n", DateTime.Now, msg);
            if (!Directory.Exists(@"c:\CBS"))
                Directory.CreateDirectory(@"c:\CBS");
            if (File.Exists(logFileName))
            {
                FileInfo fi = new FileInfo(logFileName);
                int nSize = 1000000;

                if (fi.Length > nSize)
                {
                    string fn = logFileName.Replace(".log", DateTime.Now.ToString("yyyyMMddhhmmss") + ".log");
                    File.Copy(logFileName, fn);
                    File.Delete(logFileName);
                }
            }
            FileStream fs = new FileStream(logFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamReader sr = new StreamReader(fs);

            fs.Seek(0, SeekOrigin.End);
            StreamWriter sw = new StreamWriter(fs);
            sw.Write(msg);
            sw.Close();
            fs.Close();
        }
        catch (Exception)
        {
            bHasFileAccess = false;
        }
    }



    static public string GetConfig(string key)
    {

        try
        {
            if (key == "SAPDBUserPasswd" || key == "SAPSiteUserPasswd")
                return Base64Decode(System.Configuration.ConfigurationManager.AppSettings[key].ToString());
            else
                return System.Configuration.ConfigurationManager.AppSettings[key].ToString();
        }
        catch (Exception ex)
        {
            errorLog("GetConfig Error. Key = " + key);
        }
        return "";
    }

    public static string Base64Decode(string base64EncodedData)
    {
        try
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }
        catch (Exception)
        {

        }
        return base64EncodedData;
    }
    public string getFormItemVal(SAPbouiCOM.Form form, string id)
    {
        SAPbouiCOM.EditText oEt = null;
        SAPbouiCOM.ComboBox oCB = null;
        SAPbouiCOM.StaticText oStatic = null;
        SAPbouiCOM.Button oBtn = null;
        try
        {
            if (form == null)
            {
                errorLog("getFormItemVal form is null. ");
                return "";
            }
            if (form.Items.Item(id) == null)
            {
                errorLog("getFormItemVal ID " + id + " not exist. in " + form.Title);
                return "";
            }
            switch (form.Items.Item(id).Type)
            {
                case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                    oEt = (SAPbouiCOM.EditText)form.Items.Item(id).Specific;
                    return oEt.Value;
                case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                    oCB = (SAPbouiCOM.ComboBox)form.Items.Item(id).Specific;
                    return oCB.Value;
                case SAPbouiCOM.BoFormItemTypes.it_STATIC:
                    oStatic = (SAPbouiCOM.StaticText)form.Items.Item(id).Specific;
                    return oStatic.Caption;
                case SAPbouiCOM.BoFormItemTypes.it_BUTTON:
                    oBtn = (SAPbouiCOM.Button)form.Items.Item(id).Specific;
                    return oBtn.Caption;
            }


        }
        catch (Exception ex)
        {
            errorLog("getFormItemVal exceptionon ID " + id + ", error=" + ex.Message);
        }
        finally
        {
            try
            {
                if (oEt != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEt);
                if (oCB != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCB);
                if (oStatic != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oStatic);
                if (oBtn != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBtn);
                oEt = null;
                oCB = null;
                oStatic = null;
                oBtn = null;
            }
            catch (Exception) { }
        }
        return "";
    }

    public void showFormItem(SAPbouiCOM.Form form, string id, bool bShow)
    {
        SAPbouiCOM.Item item = null;
        try
        {
            item = form.Items.Item(id);
            item.Visible = bShow;
            if (bShow)
                item.Height = 20;
            else
                item.Height = 0;
        }
        catch (Exception)
        {

        }
        finally
        {
            try
            {
                if (item != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                item = null;
            }
            catch (Exception) { }
        }
    }
    public void setFormEditVal(SAPbouiCOM.Form form, string id, string val, int paneFrom = 0, int paneTo = 0)
    {
        SAPbouiCOM.EditText oEt = null;
        try
        {
            oEt = (SAPbouiCOM.EditText)form.Items.Item(id).Specific;
            oEt.Value = val;
            if (paneFrom != 0)
            {
                form.Items.Item(id).FromPane = paneFrom;
                form.Items.Item(id).ToPane = paneTo;
            }

        }
        catch (Exception ex)
        {
            errorLog("setFormEditVal exceptionon ID " + id + ", val=" + val + ", error=" + ex.Message);
        }
        finally
        {
            try
            {
                if (oEt != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEt);
                oEt = null;
            }
            catch (Exception) { }
        }
    }

    private void FormFreeze(SAPbouiCOM.Form form, bool bFreeze)
    {

        try
        {
            if (form != null)
                form.Freeze(bFreeze);

        }
        catch (Exception) { }
    }

    

}

