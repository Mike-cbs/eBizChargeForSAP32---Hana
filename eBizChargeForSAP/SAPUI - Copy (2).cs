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


partial class SAP  {

    //private string BtnRefID;
    string edFormRefID = "33";
    string stFormRefID = "34";

    string btnPreAuth = "btnPreAuth";
    const string btnCharge = "btnCharge";
    const string btnCredit = "btnCredit";
    const string cbActive = "cbActive";
    const string FORMBPARTNER = "134";
    //string btnPayM= "btnPayM";
    const string editCheckAccount = "ChkAcct";
    const string editCheckRouting = "ChkRout";
   

    const string editCardNum = "CardNum";
    const string editCardExpDate = "ExpDt";
    const string editCardCode = "CardCode";
    const string editCAmount = "CAmount";
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
    string editHolderName = "edHName";
    string editHolderAddr = "edHAddr";
    string editHolderCity = "edHCity";
    string editHolderState = "edHState";
    const string editHolderZip = "edHZip";
    const string cbCCAccount = "cbCCAcct";
    const string cbCCSave = "cbCCSave";
  
    CustomerObject theSAPCustomer;
    SAPbouiCOM.Form theActiveForm;
    string theSAPCustomerID = "";
    string selectedPaymentMethodID = "";
    bool bVoidCloseForm = false;
    SAPbouiCOM.Form formIncomingPayment;
    SAPbouiCOM.Form formPaymentList;

    public void BPReloadThreadProc()
    {
        try
        {
            loadCustData(SBO_Application.Forms.ActiveForm);

        }
        catch (Exception)
        {
            //errorLog(ex);
        }
    }
    
    public void VoidCloseFormProc()
    {
        try
        {

            formIncomingPayment.Close();
            formPaymentList.Close();
            SBO_Application.Menus.Item("1289").Activate();
            SBO_Application.Menus.Item("1288").Activate();
  
        }
        catch (Exception)
        {
            //errorLog(ex);
        }finally
        {
            voidCCTRANS(confirmNum);
        }
    }
    private void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
    {
        BubbleEvent = true;
        try
        {
           
           
        }
        catch (Exception ex)
        {
            errorLog(ex);        }

    }
    private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;
        try
        {

            if (pVal.BeforeAction == true)
            {
                switch(pVal.MenuUID)
                {
                    case menuSOImport:
                        CreateSOImportForm();
                        break;
                    case menuCfg:
                        CreateCfgForm();
                        break;
             
                    case  "5892":  //Payment means
                        theActiveForm = SBO_Application.Forms.ActiveForm;
                        if (!ItemValidate(theActiveForm))
                            BubbleEvent = false;
                       
                        break;
                    case "1284":  //Cancel
                        string refNum = "";
                        SAPbouiCOM.Form oform = SBO_Application.Forms.ActiveForm;
                        if (oform.TypeEx == "170")
                        {
                            string docID = getFormEditVal(oform, "3");
                             SAPbobsCOM.Recordset oCustomerRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                             oCustomerRS.DoQuery("select ConfNum from RCT3 where DocNum=" + docID);
                             refNum = GetFieldVal(oCustomerRS, 0);
                        }
                        if (oform.TypeEx == "426")  //Outgoing
                        {
                            string docID = getFormEditVal(oform, "3");
                            SAPbobsCOM.Recordset oCustomerRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            oCustomerRS.DoQuery("select ref2 from OVPM where DocNum=" + docID);
                            refNum = GetFieldVal(oCustomerRS, 0);
                        }
                        if (refNum != "")
                        {
                            bAuto = false;
                            if (SBO_Application.MessageBox("Void transaction " + refNum + "?", 1, "Yes", "No") == 1)
                            {       
                                if (!voidCustomer(refNum))
                                {
                                    BubbleEvent = false;
                                }
                                else
                                {
                                    bAuto = true;
                                    bVoidCloseForm = true;
                                }
                            }else
                                BubbleEvent = false;
                        }
                        break;
                    
                }
            }
            else
            {
                switch (pVal.MenuUID)
                {
                    case "1288":
                    case "1289":
                    case "1290":
                    case "1291":
                        if (SBO_Application.Forms.ActiveForm.TypeEx=="134")
                        {
                            bAuto = true;
                            tProcess = new Thread(BPReloadThreadProc);
                            tProcess.Start();
                        }
                    
                    break;
                    case "1284":  //cancel
                    {
                        if(bVoidCloseForm)
                        {
                            bVoidCloseForm = false;
                            bAuto = true;
                            tProcess = new Thread(VoidCloseFormProc);
                            tProcess.Start();
                        }
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
    bool bAuto = false;   
    private void SBO_Application_ItemEvent( string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent ) { 
        
        // *************************************************************************
        //  BubbleEvent sets the behavior of SAP Business One.
        //  False means that the application will not continue processing this event.
        // *************************************************************************
        BubbleEvent = true;
      
        try
        {
            SAPbouiCOM.Form form;
            if (pVal.FormTypeEx == "0" && !pVal.Before_Action && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            {
                if(bAuto)
                {
                    bAuto = false;
                    form = SBO_Application.Forms.Item(FormUID);
                    form.Items.Item("1").Click();
                }
            }
            
            // string msg = "FormID=" + pVal.FormTypeEx + "," + form.Title;
            // errorLog(msg);
            switch (pVal.FormTypeEx)
            {
                case FORMBPARTNER:  //Business Partner
                    form = SBO_Application.Forms.Item(FormUID);
                    BusinessPartnerFormHandler(form, pVal);
                    break;
                case "424":  //Credit Card Form
                    form = SBO_Application.Forms.Item(FormUID);
                    CreditCardFormHandler(form, pVal);
                    break;
                case "146": //Deposit On Order & Paymenht Means
                    form = SBO_Application.Forms.Item(FormUID);
                    PaymentMeansFormHandler(form, pVal);
                    break;
                case "139": //Sales Order
                    form = SBO_Application.Forms.Item(FormUID);
                 
                    SalesOrderFormHandler(form, pVal,out BubbleEvent);
                    break;
                case "133": //Invoice
                    form = SBO_Application.Forms.Item(FormUID);
                    InvoiceFormHandler(form, pVal, out BubbleEvent);
                    break;
                case "179": //Credit Memo
                    form = SBO_Application.Forms.Item(FormUID);
                    CreditMemoFormHandler(form, pVal);
                    break;
                case "170":
                    //FORMTYPE.INCOMING_PAYMENT;
                     formIncomingPayment = SBO_Application.Forms.Item(FormUID);
                   
                    break;
                case "420":
                    //FORMTYPE.OUTGOING_PAYMENT;
                    break;
                case "60090":  //Invoice + Payment
                    form = SBO_Application.Forms.Item(FormUID);
                    InvoiceFormHandler(form, pVal, out BubbleEvent);
                    break;
                case "60091":  //A/R Reserve Invoice
                    form = SBO_Application.Forms.Item(FormUID);
                    InvoiceFormHandler(form, pVal, out BubbleEvent);
                    break;
                case "65300": //Down Payment Invoice
                    form = SBO_Application.Forms.Item(FormUID);
                    InvoiceFormHandler(form, pVal, out BubbleEvent);
                    break;

                case formCredit:
                    form = SBO_Application.Forms.Item(FormUID);
                    CreditFormHandler(form, pVal, out BubbleEvent);
                    break;

                case formSOImport:
                    form = SBO_Application.Forms.Item(FormUID);
                    SOImportFormHandler(pVal); 
                    break;
                case formCfg:
                    form = SBO_Application.Forms.Item(FormUID);
                    CFGFormHandler(form,pVal, out BubbleEvent); 
                    break;
                   
                case "60508":
                    formPaymentList = SBO_Application.Forms.Item(FormUID);
                    break;
            }
        }
        catch (Exception)
        {
            //errorLog(ex);
        }
    }
    private void CreditCardFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        
        try
        {
            if (!pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        form.Title = "eBizCharge Account Setup";
                        /*
                        SAPbouiCOM.Matrix oMatrix = form.Items.Item("3").Specific;
                        SAPbouiCOM.Item oItemRef = form.Items.Item("3");
                        SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("CardName");
                        SAPbouiCOM.Item oItem = form.Items.Add("LBTit", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        oItem.Top = oItemRef.Top + 2;
                        oItem.Height = 18;
                        oItem.Left = 32 + oItemRef.Left;
                        oItem.Enabled = false; 
                        oItem.Width = 100;
                        SAPbouiCOM.StaticText s = oItem.Specific;
                        s.Caption = "Account Name";
                         */
                        //SAPbouiCOM.Cell item = oColumn.Cells.Item("2"); //Cast the Cell of the matrix to the respective type , in this case EditText
                        //item. = "Test";
                        break;
                   
                }
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    private bool CheckStatus(SAPbouiCOM.Form form )
    {
        if (getComboValue(form , fidStatus) == "1")
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
            oItem.FromPane = tabPane;
            oItem.ToPane = tabPane + 3;
            oItem.Visible = bVisible;
            SAPbouiCOM.Folder oFolder = oItem.Specific;
            oFolder.Caption = title;
            oFolder.GroupWith(oItemRef.UniqueID);
            oFolder.Pane = pane;
        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
    private void AddCCLogMatrix(SAPbouiCOM.Form form, int pane, string refID = "39", int height = 0, int top = 0)
    {
        try{

            SAPbouiCOM.Item oItem = form.Items.Add(mxtCCLog, SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oItem.FromPane = pane;
            oItem.ToPane = pane;
            SAPbouiCOM.Item oItemRef = form.Items.Item(refID);
            oItem.Top = oItemRef.Top + top;
            oItem.Left = oItemRef.Left;
            oItem.Width = oItemRef.Width;
            oItem.Height = oItemRef.Height + height;

            SAPbouiCOM.Matrix oMatrix = form.Items.Item(mxtCCLog).Specific;
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Add("RefNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ref. No.";
            oColumn.Width = 50;
           
           
            oColumn = oMatrix.Columns.Add("PaymentID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Journal Entry";
            SAPbouiCOM.LinkedButton btn = oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_JournalPosting).ToString();
            oColumn.Width = 50;
            

            oColumn = oMatrix.Columns.Add("OrderID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Order";
            btn = oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_Order).ToString();
            oColumn.Width = 50;

            oColumn = oMatrix.Columns.Add("InvoiceID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Invoice";
            btn = oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_Invoice).ToString();
            oColumn.Width = 50;

            oColumn = oMatrix.Columns.Add("CMID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Credit Memo";
            btn = oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo).ToString();
            oColumn.Width = 50;

            oColumn = oMatrix.Columns.Add("command", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Command";
            oColumn.Width = 50;
            oColumn = oMatrix.Columns.Add("Amount", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Amount";
            oColumn.Width = 50;
            oColumn = oMatrix.Columns.Add("Desc", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Description";
            oColumn.Width = 50;

            oColumn = oMatrix.Columns.Add("CardHolder", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Card Holder";
            oColumn.Width = 50;
            
            
            oColumn = oMatrix.Columns.Add("crCardNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Card No.";
            oColumn.Width = 50;
            /*
            oColumn = oMatrix.Columns.Add("result", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Result";
            oColumn.Width = 50;
            */
            oColumn = oMatrix.Columns.Add("error", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Error";
            oColumn.Width = 50;
            oColumn = oMatrix.Columns.Add("avs", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "AVS";
            oColumn.Width = 150;
            oColumn = oMatrix.Columns.Add("CardCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Card Code";
            oColumn.Width = 150;
            oColumn = oMatrix.Columns.Add("CardLevel", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Card Level";
            oColumn.Width = 150;
            oColumn = oMatrix.Columns.Add("custNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Customer Number";
            oColumn.Width = 50;
            oColumn = oMatrix.Columns.Add("CCAID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Credit Card";
            oColumn.Width = 50;
            oColumn = oMatrix.Columns.Add("recdate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Process Date";
            oColumn.Width = 100;

            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;




 

            //populateLog(form);
        }
        catch(Exception ex)
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
            SAPbouiCOM.Matrix oMatrix = form.Items.Item(mxtCCLog).Specific;
            oMatrix.Clear();
            //oMatrix.AutoResizeColumns();
            SAPbouiCOM.DataTable oDts = getDataTable(form);

            sql = string.Format("select top " + GetConfig("maxSelectRow") + " InvoiceID, CCAccountID as CCAID, OrderID, PaymentID, DownPaymentInvoiceID, CreditMemoID as CMID, custNum, customerID,	customerName,CardHolder,crCardNum,Description,recID,acsUrl,authAmount,authCode,avsResult,avsResultCode,	batchNum,batchRef,cardCodeResult," +
            "cardCodeResultCode,cardLevelResult, cardLevelResultCode,conversionRate,convertedAmount,convertedAmountCurrency,error,errorCode,isDuplicate,payload,profilerScore,profilerResponse," +
            "profilerReason,refNum,remainingBalance,result,resultCode,	status,	statusCode,	vpasResultCode, recDate, Command, Amount from CCTRANS {0} order by recDate desc", strWhere);
           
            oDts.ExecuteQuery(sql);
            BindMatrix(oMatrix, "RefNum", "refNum");
            BindMatrix(oMatrix, "PaymentID", "PaymentID");
            BindMatrix(oMatrix, "OrderID", "OrderID");
            BindMatrix(oMatrix, "InvoiceID", "InvoiceID");
            BindMatrix(oMatrix, "CMID", "CMID");
            BindMatrix(oMatrix, "CCAID", "CCAID");
            BindMatrix(oMatrix, "custNum", "custNum");
            BindMatrix(oMatrix, "command", "Command");
            BindMatrix(oMatrix, "Amount", "Amount");
            BindMatrix(oMatrix, "Desc", "Description");
            BindMatrix(oMatrix, "CardHolder", "CardHolder");
            BindMatrix(oMatrix, "crCardNum", "crCardNum");
         
            BindMatrix(oMatrix, "error", "error");
            BindMatrix(oMatrix, "avs", "avsResult");
            BindMatrix(oMatrix, "CardCode", "cardCodeResult");
            BindMatrix(oMatrix, "CardLevel", "cardLevelResult");
            BindMatrix(oMatrix, "recdate", "recDate");

            oMatrix.LoadFromDataSource();

        }catch(Exception ex)
        { 
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void BindMatrix(SAPbouiCOM.Matrix oMatrix, string mtCol, string dtCol, string dt = null)
    {
        try
        {
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item(mtCol);
            if (dt == null)
            {
                oColumn.DataBind.Bind(dtCCTransLog, dtCol);
            }else
                oColumn.DataBind.Bind(dt, dtCol);
  

        }catch(Exception ex)
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
        }catch(Exception)
        {
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

        }catch(Exception ex)
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
        return oItem;
    }
    private SAPbouiCOM.Item addPaneItem(SAPbouiCOM.Form form, string id, int l, int t, int w, int h, string label, SAPbouiCOM.BoFormItemTypes type, int pane, int order, int lbw = 100)
    {
        SAPbouiCOM.Item oItem = null;
        try
        {
            int toPane = pane + 1;
            int fromPane = pane;
            if(label==CC_CARDNUM || label==CC_CARDCODE || label==CC_EXPDATE)
            {
                toPane = pane;
                fromPane = pane;
         
            }else if(label==EC_ROUTING || label == EC_ACCOUNT)
            {
                toPane = pane + 1;
                fromPane = pane + 1;
            }
            oItem = form.Items.Add(id, type);
            SAPbouiCOM.Item oItm = form.Items.Item(id);
            oItm.FromPane = fromPane;
            oItem.ToPane = toPane;
            oItm.Left = l;
            oItm.Top = t;
            oItm.Width = w;
            oItm.Height = h;
            
            string lb =  "LB" + order.ToString();
            oItm = form.Items.Add(lb, SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.ToPane = toPane;
            oItem.FromPane = fromPane;
            oItm.Left = l - lbw;
            oItm.Top = t;
            oItm.Width = lbw;
            oItm.Height = h;
            setLabelCaption(form, lb, label);
            if (type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            {
                SAPbouiCOM.ComboBox oComboBox = form.Items.Item(id).Specific;
                oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            }
        }catch(Exception ex)
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
        }catch(Exception)
        {

        }
    }
    
   
    private void setMatrixItem(SAPbouiCOM.Matrix oMatrix, string Col, int Row, string val)
    {
        try
        {
            //oMatrix.SetCellWithoutValidation(Row, Col, val);
            
            SAPbouiCOM.EditText oEdit = oMatrix.Columns.Item(Col).Cells.Item(Row).Specific; //Cast the Cell of the matrix to the respective type , in this case EditText
            oEdit.Value = val; //Get the value form the EditText
            
        }
        catch (Exception ex)
        {
            errorLog("Can't find row:" + Row.ToString() + ",col:" + Col + "\r\n" + ex.Message);
        }
       
    }
    private string getMatrixItem(SAPbouiCOM.Matrix oMatrix, string Col, int Row )
    {
        string sValue = "";
        try
        {
            SAPbouiCOM.EditText oEdit = oMatrix.Columns.Item(Col).Cells.Item(Row).Specific; //Cast the Cell of the matrix to the respective type , in this case EditText
            sValue = oEdit.Value; //Get the value form the EditText
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
            SAPbouiCOM.CheckBox oCB = oMatrix.Columns.Item(Col).Cells.Item(Row).Specific; //Cast the Cell of the matrix to the respective type , in this case EditText
            return oCB.Checked; //Get the value form the EditText
        }
        catch (Exception ex)
        {
            errorLog("Can't find row:" + Row.ToString() + ",col:" + Col + "\r\n" + ex.Message);
        }
        return false;

    }
    private double getItemDoubleValue(SAPbouiCOM.Form form, string id)
    {
        double ret = 0;
        try
        {
            string s = getFormEditVal(form, id);
            if(s != "")
                ret = double.Parse(s.Replace("$","").Replace(",","").Trim());
        }
        catch (Exception ex)
        {
            errorLog("Item ID: " + id + " error. " + ex.Message);
        }
        return ret;
    }
    
    private string getItemValue(SAPbouiCOM.Form form, string id)
    {
        string ret = "";
        try
        {
            SAPbouiCOM.EditText oET = form.Items.Item(id).Specific;
            ret = oET.Value;
        }
        catch (Exception)
        {
        }
        return ret;
    }
    
    private string getComboValue(SAPbouiCOM.Form form, string id)
    {
        string ret = "";
        try
        {
            SAPbouiCOM.ComboBox cb = form.Items.Item(id).Specific;
            ret = cb.Value;
        }
        catch (Exception)
        {
        }
        return ret;
    }
    private string setComboValue(SAPbouiCOM.Form form, string id, string val)
    {
        string ret = "";
        try
        {
            SAPbouiCOM.ComboBox cb = form.Items.Item(id).Specific;
            cb.Select(val);
        }
        catch (Exception)
        {
        }
        return ret;
    }
  
    private string getButtonCaption(SAPbouiCOM.Form form, string id)
    {
        string ret = "";
        try
        {
            SAPbouiCOM.Button oButton = form.Items.Item(id).Specific;
            ret = oButton.Caption;
        }
        catch (Exception)
        {
        }
        return ret;
    }
    private void hideItem(SAPbouiCOM.Form form, string id)
    {
        try
        {
            form.Items.Item(id).Visible=false;
        }
        catch (Exception)
        {
        }
    }
    private void setButtonCaption(SAPbouiCOM.Form form, string id, string cap)
    {
       
        try
        {
            SAPbouiCOM.Button oButton = form.Items.Item(id).Specific;
            oButton.Caption = cap;
        }
        catch (Exception)
        {
        }
        
    }
    private string getLabelCaption(SAPbouiCOM.Form form, string id)
    {
        string ret = "";
        try
        {
            SAPbouiCOM.StaticText oStatic = form.Items.Item(id).Specific;
            ret = oStatic.Caption;
        }
        catch (Exception)
        {
        }
        return ret;
    }
    private void setLabelCaption(SAPbouiCOM.Form form, string id, string cap)
    {

        try
        {
            SAPbouiCOM.StaticText oStatic = form.Items.Item(id).Specific;
            oStatic.Caption = cap;
        }
        catch (Exception)
        {
        }

    }
    private string getComboBoxVal(SAPbouiCOM.Form form, string id)
    {

        try
        {
            SAPbouiCOM.ComboBox oCB = form.Items.Item(id).Specific;
            return oCB.Value;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return "";

    }
   
    private bool ClearInput(SAPbouiCOM.Form form)
    {
        try
        {
            setFormEditVal( form,editCardNum, "");
            setFormEditVal( form,editCardExpDate,"");
            setFormEditVal( form,editCardCode, "");
            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    private string getCardNumber(SAPbouiCOM.Form form)
    {
        try
        {
           
            SAPbouiCOM.EditText cardNumber = form.Items.Item(editCardNum).Specific;
            return cardNumber.Value;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return "";
    }
    private bool VerifyInput(SAPbouiCOM.Form form, ref string amt,ref string cardNo, ref string exp, ref string code )
    {
        try
        {
            amt = getFormEditVal(form,editCAmount);
            if (amt == "")
            {
                SBO_Application.MessageBox("Please enter charge amount.", 1, "Ok", "", "");
                return false;
            }

           
            if (!GetTotalAmount(form, ref amt))
                return false;
            try
            {
                double a = double.Parse(amt);
                if(a <= 0.01)
                {
                    SBO_Application.MessageBox("Invalid amount.  Please add item.", 1, "Ok", "", "");
                    return false;
                }
            }
            catch (Exception)
            {
                SBO_Application.MessageBox("Invalid amount.  Please enter numerical value only.", 1, "Ok", "", "");
                return false;
            }

            cardNo = getFormEditVal(form,editCardNum);
            if (cardNo == "")
            {
                SBO_Application.MessageBox("Please enter credit card number.", 1, "Ok", "", "");
                return false;
            }
            exp = getFormEditVal(form,editCardExpDate);
            if (exp == "")
            {
                SBO_Application.MessageBox("Please enter card expiration date.", 1, "Ok", "", "");
                return false;
            }
            code = getFormEditVal(form,editCardCode);
            if (code == "")
            {
                SBO_Application.MessageBox("Please enter card code.", 1, "Ok", "", "");
                return false;
            }
            return true;
        }catch(Exception ex)
        { 
            errorLog(ex); 
        }
        return false;
    }
    private bool GetTotalAmount(SAPbouiCOM.Form form, ref string amt)
    {
        try
        {
            string a = getFormEditVal(form, editCAmount);
            if (a == "")
            {
                SBO_Application.MessageBox("Please enter charge amount.", 1, "Ok", "", "");
                return false;
            }

         
            try
            {
                a = a.Replace("$", "").Replace(",", "");
                if(double.Parse(a)<= 0.01)
                {
                    SBO_Application.MessageBox("Invalid amount.  Please add item.", 1, "Ok", "", "");
                    return false;
                }
                if (getItemDoubleValue(form, "14") < double.Parse(a))
                {
                    SBO_Application.MessageBox("The payment amount cannot exceed the current balance due..", 1, "Ok", "", "");
                    return false;
                }
                amt = a;
            }
            catch (Exception)
            {
                SBO_Application.MessageBox("Invalid amount.  Please enter numerical value only.", 1, "Ok", "", "");
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
            return rs.Fields.Item(index).Value;

        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return "";
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
            string logFileName = @"c:\CBS\CBSForSAP.log";
            msg = string.Format("{0}:\r\n\t{1}\r\n\r\n", DateTime.Now, msg);
            if (!Directory.Exists(@"c:\CBS"))
                Directory.CreateDirectory(@"c:\CBS");
            if (File.Exists(logFileName))
            {
                FileInfo fi = new FileInfo(logFileName);
                if (fi.Length > 100000)
                {
                    File.Copy(logFileName, logFileName + DateTime.Now.ToString("yyyyMMddhhmmss"));
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
        }
    }
    private bool testItem(SAPbouiCOM.Form form , string id)
    {
        try
        {
            if (form.Items.Item(id).UniqueID == id)
                return true;
        }catch(Exception)
        {
            return false;
        }
        return true;

    }
    private string ConverToDollar(string total, string rate)
    {
        string[] dl = total.Split(' ');
        string ret = dl[0];
        ret =  ret.Replace(",", "");
        try
        {
            

        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return ret;
    }
   
    private string GetRemark(SAPbouiCOM.Form form)
    {
        try
        {
            SAPbouiCOM.EditText remark = form.Items.Item(fidRemark).Specific;
            return remark.Value;
        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return "";
    }
    static public string GetConfig(string key)
    {
        try
        {
            return System.Configuration.ConfigurationManager.AppSettings[key].ToString();
        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return "";
    }
    private void SetRemark(SAPbouiCOM.Form form, string r)
    {
        try
        {
            setFormEditVal(form, fidRemark, r);
            /*
            SAPbouiCOM.EditText soid = form.Items.Item(fidSaleOrderID).Specific;
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(string.Format("update ordr set comments='{1}' where DocEntry = {0}", soid.Value, r));
             */
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        
    }
 
    private bool GetRefNoAndAmountFromRemark(SAPbouiCOM.Form form, ref string refNo, ref string amount)
    {
        try
        {
            string s = getFormEditVal(form, fidRemark);
            if (s == "")
                return false;
            string t = "ref. No.:";
            int nStart = s.IndexOf(t) + t.Length;
            if (nStart == -1)
                return false;
            int nEnd = s.IndexOf(" ", nStart);
            if (nEnd == -1)
                return false;
            refNo = s.Substring(nStart, nEnd - nStart).Trim();
            t = "Amount:";
            nStart = s.IndexOf(t,nEnd) + t.Length;
            if (nStart == -1)
                return false;
            nEnd = s.IndexOf(" ", nStart);
            if (nEnd == -1)
                return false;
            amount = s.Substring(nStart, nEnd - nStart).Trim();
            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    public string getDocDate(SAPbouiCOM.Form form)
    {
       string s = "";
        try
        {
            s = getItemValue(form, "46");
            if(s.Length==8)
            {
                s=string.Format("{0}/{1}/{2}",s.Substring(4,2),s.Substring(6,2),s.Substring(0,4));
                return s;
            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return s;
    }
    public string getFormEditVal(SAPbouiCOM.Form form, string id)
    {
        try
        {
            SAPbouiCOM.EditText oEt = form.Items.Item(id).Specific;
            return oEt.Value;

        }catch(Exception)
        {

        }
        return "";
    }
    public void setFormEditVal(SAPbouiCOM.Form form, string id, string val)
    {
        try
        {
            SAPbouiCOM.EditText oEt = form.Items.Item(id).Specific;
            oEt.Value = val;

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
   
} 

