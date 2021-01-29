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


partial class SAP  {

    private string BtnRefID;
    string edFormRefID = "33";
    string stFormRefID = "34";

    string btnPreAuth = "btnPreAuth";
    string btnCharge = "btnCharge";
    string btnCapture = "btnCapture";
    string btnCredit = "btnCredit";
    string cbSave = "cbSave";
    string btnPayM= "btnPayM";
    string lnkPayM = "lnkPayM";
    string editCardNum = "CardNum";
    string editCardExpDate = "ExpDt";
    string editCardCode = "CardCode";
    string editCAmount = "CAmount";
    string lbCardNum = "lbCardNum";
    string lbCardExpDate = "lbExpDt";
    string lbCardCode = "lbCardCode";
    string lbCAmount = "lbCAmount";
    string tbCreditCard = "tbCCLog";
    string mxtCCLog = "mtxCCLog";
    string CCTabTitle = "Credit Card Log";
    string dtCCTransLog = "CCTRANSLOG";
    string strCCTransTableName = "CCTRANS";
    public string fidCustID = "4";
    public string fidRecordID = "8";
    public string fidCustName = "54";
    public string fidTotalAmount = "29";
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
    string strTotalAmount = "Total Amount";
    string cbPaymentMethod = "cbPayM";
    string editHolderName = "edHName";
    string editHolderAddr = "edHAddr";
    string editHolderCity = "edHCity";
    string editHolderState = "edHState";
    string editHolderZip = "edHZip";

    string strFormTitleSelectPayment = "Select Payment Method";
    string strFormTitleCardHolderInfo = "Card Holder Info";
    bool bCreditCardProcess = false;
    bool bCardHolderInfoProcess = false;
    bool bDefaultSet = false;
    string bSaveCard = "";
   
    Thread tProcess;
    SAPbouiCOM.Form theSAPForm;
    CustomerObject theSAPCustomer;
    string theSAPCustomerID = "";
    string selectedPaymentMethodID = "";
    CCTYPE CCType;
    enum CCTYPE { CC_PREAUTH = 0, CC_CHARGE = 1, CC_CREDIT=2 };
    public void ShowCreditCardSelectForm()
    {
        try
        {
            SBO_Application.MessageBox("  ", 1, "Cancel", "Remove", "Select");
           
        }
        catch (Exception)
        {
            //errorLog(ex);
        }
    }
    public void selectDefault()
    {
        try
        {
            
        }
        catch (Exception)
        {
            //errorLog(ex);
        }
    }
    public void ShowCardHolderInfoForm()
    {
        try
        {
            SBO_Application.MessageBox("\r\n\r\n\r\n\r\n\r\n\r\n\r\n", 1, "Cancel", "OK");
            
        }
        catch (Exception)
        {
            //errorLog(ex);
        }
    }
    private void SBO_Application_ItemEvent( string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent ) { 
        
        // *************************************************************************
        //  BubbleEvent sets the behavior of SAP Business One.
        //  False means that the application will not continue processing this event.
        // *************************************************************************
        BubbleEvent = true;
        try
        {
            if(FormUID ==formSOImport)
            {
                SOImportFormHandler(pVal);
                return;
            }
            if (pVal.FormTypeEx == "0" && !pVal.Before_Action && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            {
                if(bCardHolderInfoProcess)
                {
                    SAPbouiCOM.Form form = SBO_Application.Forms.Item(FormUID);
                    //SAPbouiCOM.CheckBox oCB = form.Items.Item(cbSave).Specific;
                    SAPbouiCOM.ComboBox oCB = form.Items.Item(cbSave).Specific;
                    oCB.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    //oCB.Checked = true;
                }
            }
            if (pVal.FormTypeEx == "0" && pVal.Before_Action)
            {
                if (!bCreditCardProcess)
                {
                    SAPbouiCOM.Form form = SBO_Application.Forms.Item(FormUID);
                    hideItem(form, cbPaymentMethod);
                    hideItem(form, lbCardNum);
                  
                }
                if (!bCardHolderInfoProcess)
                {
                    SAPbouiCOM.Form form = SBO_Application.Forms.Item(FormUID);
                    hideItem(form, editHolderName);
                    hideItem(form, "LB1");
                    hideItem(form, editHolderAddr);
                    hideItem(form, "LB2");
                    hideItem(form, editHolderCity);
                    hideItem(form, "LB3");
                    hideItem(form, editHolderState);
                    hideItem(form, "LB4");
                    hideItem(form, editHolderZip);
                    hideItem(form, "LB5");
                }
                if(bCreditCardProcess)
                {
                    if (pVal.ItemUID == "3") //select
                    {
                        bCreditCardProcess = false;
                        SAPbouiCOM.Form form = SBO_Application.Forms.Item(FormUID);
                        SAPbouiCOM.ComboBox oCB = form.Items.Item(cbPaymentMethod).Specific;
                        selectedPaymentMethodID = oCB.Value;
                        form.Close();
                        setPaymentMethod(selectedPaymentMethodID);
                    }
                    else if (pVal.ItemUID == "2")  //remove
                    {
                        bCreditCardProcess = false;
                        SAPbouiCOM.Form form = SBO_Application.Forms.Item(FormUID);
                        SAPbouiCOM.ComboBox oCB = form.Items.Item(cbPaymentMethod).Specific;
                        string sel = oCB.Value;
                        form.Close();
                        if(sel =="")
                        {
                            SBO_Application.MessageBox("Please select a credit card to remove.");
                        }else
                        {
                            if (SBO_Application.MessageBox("Are you sure you want to remove this card?", 1, "Cancel", "OK") == 2)
                            {
                                removePaymentMethod(sel);
                                if (selectedPaymentMethodID == sel)
                                {
                                    selectedPaymentMethodID = "";
                                    ClearInput(form);
                                }
                            }
                        }
                        return;
                    }
                    else if (pVal.ItemUID == "1")  //cancel
                    {
                        bCreditCardProcess = false;
                        return;
                    }
                    else if (pVal.ItemUID == "") //load
                    {
                        SAPbouiCOM.Form form = SBO_Application.Forms.Item(FormUID);
                        form.Title = strFormTitleSelectPayment;
                        addPaymentMethodSelect(form);
                    }
                }
                if(bCardHolderInfoProcess)  //card holder 
                {

                    if (pVal.ItemUID == "2")
                    {
                        bCardHolderInfoProcess = false;
                        SAPbouiCOM.Form form = SBO_Application.Forms.Item(FormUID);
                        string name = getFormEditVal(form, editHolderName);
                        if(name == "")
                        {
                            form.Close();
                            SBO_Application.MessageBox("Please enter card holder name.");
                            return;
                        }
                        
                        theSAPCustomer.BillingAddress.FirstName = getFirstName(name);
                        theSAPCustomer.BillingAddress.LastName = getLastName(name);
                        theSAPCustomer.BillingAddress.Street = getFormEditVal(form, editHolderAddr);
                        theSAPCustomer.BillingAddress.City = getFormEditVal(form, editHolderCity);
                        theSAPCustomer.BillingAddress.State = getFormEditVal(form, editHolderState);
                        theSAPCustomer.BillingAddress.Zip = getFormEditVal(form, editHolderZip);
                       // bSaveCard = getComboBoxVal(form, cbSave);
                        form.Close();
                        updateCustomer();
                        if (selectedPaymentMethodID != "" && getCardNumber(theSAPForm).IndexOf("XXX") > 0)
                        {
                            switch(CCType)
                            {
                                case CCTYPE.CC_PREAUTH:
                                    PreAuthCustomer(theSAPForm);
                                    break;
                                case CCTYPE.CC_CHARGE:
                                    ChargeCustomer(theSAPForm);
                                    break;
                                case CCTYPE.CC_CREDIT:
                                    CreditCustomer(theSAPForm);
                                    break;
                            }
                            
                        }
                        else
                        {
                            switch(CCType)
                            {
                                case CCTYPE.CC_PREAUTH:
                                    PreAuth(theSAPForm);
                                    break;
                                case CCTYPE.CC_CHARGE:
                                    Charge(theSAPForm);
                                    break;
                                case CCTYPE.CC_CREDIT:
                                    Credit(theSAPForm);
                                    break;
                            }
                        }
                        
                        
                    }
                    else if (pVal.ItemUID == "1")
                    {
                        bCardHolderInfoProcess = false;
                    }
                    else if (pVal.ItemUID == cbSave && !bDefaultSet)
                    {
                        bDefaultSet = true;
                        SAPbouiCOM.Form form = SBO_Application.Forms.Item(FormUID);
                        SAPbouiCOM.CheckBox oCB = form.Items.Item(cbSave).Specific;
                        //oCB.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        oCB.Checked = true;
                    }
                    else if (pVal.ItemUID == "")
                    {
                        handleSwipe(theSAPForm);
                        SAPbouiCOM.Form form = SBO_Application.Forms.Item(FormUID);
                        form.Title = strFormTitleCardHolderInfo;
                        addCardHolderInfo(form);
                        bDefaultSet = false;
                    }
                }
            }

            if (pVal.FormType != 0)
            {
                if (pVal.ItemUID == tbCreditCard && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.Before_Action == false)
                {
                   SAPbouiCOM.Form form = SBO_Application.Forms.Item(FormUID);
                   form.PaneLevel = 5;
                   populateLog(form);
                   
                }
                //get payment method btn click
                if (pVal.ItemUID == btnPayM && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.Before_Action == false)
                {
                    theSAPForm = SBO_Application.Forms.Item(FormUID);
                    string id = getSAPCustomerID(theSAPForm);
                    
                    if (id == "")
                    {
                        SBO_Application.MessageBox("Please select a customer.");
                        return;
                    }
                    else if (id != theSAPCustomerID)
                    {
                        theSAPCustomerID = id;
                        theSAPCustomer = getCustomer(theSAPCustomerID);
                    }
                    if (theSAPCustomer.PaymentMethods == null)
                    {
                        SBO_Application.MessageBox("No payment method found.", 1, "Ok", "", "");
                        return;
                    }
                    bCreditCardProcess = true;
                    tProcess = new Thread(ShowCreditCardSelectForm);
                    tProcess.Start();
                    BubbleEvent = false;
                    return;
                }
                //Pre-auth btn click
                if (pVal.ItemUID == btnPreAuth && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.Before_Action == false)
                {
                    string id = getSAPCustomerID(theSAPForm);
                    if (id == "")
                    {
                        SBO_Application.MessageBox("Please select a customer.");
                        return;
                    }
                    else if (id != theSAPCustomerID)
                    {
                        theSAPCustomerID = id;
                        theSAPCustomer = getCustomer(theSAPCustomerID);
                    }
                    BubbleEvent = false;
                    theSAPForm = SBO_Application.Forms.Item(FormUID);
                    bCardHolderInfoProcess = true;
                    CCType = CCTYPE.CC_PREAUTH;
                    tProcess = new Thread(ShowCardHolderInfoForm);
                    tProcess.Start();
                  
                    return;
                }
                //Charge button click
                if (pVal.ItemUID == btnCharge && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.Before_Action == false)
                {
                    BubbleEvent = false;
                    string id = getSAPCustomerID(theSAPForm);
                    if (id == "")
                    {
                        SBO_Application.MessageBox("Please select a customer.");
                        return;
                    }
                    else if (id != theSAPCustomerID)
                    {
                        theSAPCustomerID = id;
                        theSAPCustomer = getCustomer(theSAPCustomerID);
                    }
                    theSAPForm = SBO_Application.Forms.Item(FormUID);
                    bCardHolderInfoProcess = true;
                    CCType = CCTYPE.CC_CHARGE;
                 
                    tProcess = new Thread(ShowCardHolderInfoForm);
                    tProcess.Start();
                    return;
                }
                //Capture button click
                if (pVal.ItemUID == btnCapture && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.Before_Action == false)
                {
                    BubbleEvent = false;
                    theSAPForm = SBO_Application.Forms.Item(FormUID);
                   
                    Capture(theSAPForm);
                    return;
                }
                //Credit button click
                if (pVal.ItemUID == btnCredit && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.Before_Action == false)
                {
                    BubbleEvent = false;
                    string id = getSAPCustomerID(theSAPForm);
                    if (id == "")
                    {
                        SBO_Application.MessageBox("Please select a customer.");
                        return;
                    }
                    else if (id != theSAPCustomerID)
                    {
                        theSAPCustomerID = id;
                        theSAPCustomer = getCustomer(theSAPCustomerID);
                    }
                    theSAPForm = SBO_Application.Forms.Item(FormUID);
                    bCardHolderInfoProcess = true;
                    CCType = CCTYPE.CC_CREDIT;
                 
                    tProcess = new Thread(ShowCardHolderInfoForm);
                    tProcess.Start();
                    return;
                } 
                
                SAPbouiCOM.BoEventTypes EventEnum = 0;
                EventEnum = pVal.EventType;

               /*
                msg = "pVal.FormTypeEx: " + pVal.FormTypeEx + ", form title: " + SBO_Application.Forms.Item(FormUID).Title ;
                errorLog(msg);
                */
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action)
                {
                    if (pVal.FormTypeEx == GetConfig("SAPARInvoiceFormID") || pVal.FormTypeEx == GetConfig("SAPInvoicePaymentFormID") || pVal.FormTypeEx == GetConfig("SAPReserveInvoiceFormID") || pVal.FormTypeEx == GetConfig("SAPARDownPaymentFormID"))
                    {
                        SAPbouiCOM.Form form = SBO_Application.Forms.Item(FormUID);
                        if (pVal.Before_Action == true)
                        {

                            if (!testItem(form, btnCharge))
                            {
                                SAPbouiCOM.Item oRef = form.Items.Item(BtnRefID);
                                form.Items.Add(btnCharge, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                                SAPbouiCOM.Item oItm = form.Items.Item(btnCharge);

                                oItm.Left = oRef.Left + 68;
                                oItm.Top = oRef.Top;
                                oItm.Width = oRef.Width;
                                oItm.Height = oRef.Height;
                                SAPbouiCOM.Button oBt = form.Items.Item(btnCharge).Specific;
                                oBt.Caption = "Charge";


                                form.Items.Add(btnCapture, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                                oItm = form.Items.Item(btnCapture);
                                oRef = form.Items.Item(btnCharge);
                                oItm.Left = oRef.Left + 68;
                                oItm.Top = oRef.Top;
                                oItm.Width = oRef.Width;
                                oItm.Height = oRef.Height;
                                oBt = form.Items.Item(btnCapture).Specific;
                                oBt.Caption = "Capture";
                                addCreditCardForm(form);
                                if (pVal.FormTypeEx == GetConfig("SAPARDownPaymentFormID"))
                                {
                                    oItm = form.Items.Item(editCAmount);
                                    SAPbouiCOM.Item oRemarkStatic = form.Items.Item(fidRemarkStatic);
                                    SAPbouiCOM.Item oRemark = form.Items.Item(fidRemark);
                                    oRemarkStatic.Top = oItm.Top + 16;
                                    oRemark.Top = oItm.Top + 16;
                                    oRemark.Height = oRemark.Height - 14;
                                }
                            }
                            else
                            {
                                CheckButton(form, btnCharge, strHeaderCharged);
                                CheckButtonForCapture(form);
                            }

                        }
                    }
                }
                      
                        /*
                           foreach (SAPbouiCOM.Item i in form.Items)
                     {
                         if (i.Type == SAPbouiCOM.BoFormItemTypes.it_EDIT)
                         {
                             SAPbouiCOM.EditText oEd = form.Items.Item(i.UniqueID).Specific;
                             msg = "Edit UniqueID=" + i.UniqueID + ",value=" + oEd.Value + ", i.Top=" + i.Top.ToString() + ",i.Left=" + i.Left.ToString() + ",i.Height=" + i.Height.ToString() + ", type=" + i.Type.ToString();
                             errorLog(msg);
                         }
                         else if (i.Type == SAPbouiCOM.BoFormItemTypes.it_STATIC)
                         {
                             SAPbouiCOM.StaticText ocb = form.Items.Item(i.UniqueID).Specific;
                             msg = "Label UniqueID=" + i.UniqueID + ",value=" + ocb.Caption + ", i.Top=" + i.Top.ToString() + ",i.Left=" + i.Left.ToString() + ",i.Height=" + i.Height.ToString() + ", type=" + i.Type.ToString();
                             errorLog(msg);
                         }
                         else
                         {
                             msg = "Unkown UniqueID=" + i.UniqueID + ", i.Top=" + i.Top.ToString() + ",i.Left=" + i.Left.ToString() + ",i.Height=" + i.Height.ToString() + ", type=" + i.Type.ToString() + ", Desc=" + i.Description ;
                             errorLog(msg);
                         }
                     }
                         
                         */
               
                if (pVal.FormTypeEx == GetConfig("SAPCreditFormID"))
                {
                    SAPbouiCOM.Form form = SBO_Application.Forms.Item(FormUID);
                    if (pVal.Before_Action == true)
                    {
                     
                        if (!testItem(form, btnCredit))
                        {
                          
                            form.Items.Add(btnCredit, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            SAPbouiCOM.Item oRef = form.Items.Item(BtnRefID);
                            SAPbouiCOM.Item oItm = form.Items.Item(btnCredit);
                            oItm.Left = oRef.Left + 68;
                            oItm.Top = oRef.Top;
                            oItm.Width = oRef.Width;
                            oItm.Height = oRef.Height; 
                            SAPbouiCOM.Button oBt = form.Items.Item(btnCredit).Specific;
                            oBt.Caption = "Credit";
                            addCreditCardForm(form);
                        }
                       
                        
                    }else
                        CheckButton(form, btnCredit, strHeaderCredited);
                }

                if (pVal.FormTypeEx == GetConfig("SAPSalesOrderFormID"))
                {
                    
                    if (pVal.Before_Action == true)
                    {

                        SAPbouiCOM.Form form = SBO_Application.Forms.Item(FormUID);

                        if (!testItem(form, btnPreAuth))
                        {
                            form.Items.Add(btnPreAuth, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            SAPbouiCOM.Item oItm = form.Items.Item(btnPreAuth);
                            oItm.LinkTo =  BtnRefID;
                            oItm.Left = form.Items.Item(BtnRefID).Left + 68;
                            oItm.Top = form.Items.Item(BtnRefID).Top;
                            oItm.Width = form.Items.Item(BtnRefID).Width;
                            oItm.Height = form.Items.Item(BtnRefID).Height;
                            SAPbouiCOM.Button oBt = form.Items.Item(btnPreAuth).Specific;
                            oBt.Caption = "Pre-Auth";

                            addCreditCardForm(form);

                        }
                        CheckButton(form, btnPreAuth, strHeaderApproved);
                        
                    }
                }
            }
         }
        catch(Exception ex)
         {
             errorLog(ex);
         }
    }
    private void CheckButtonForCapture(SAPbouiCOM.Form form)
    {
        try
        {
            SAPbouiCOM.Item oItm = form.Items.Item(btnCapture);
            string s = getButtonCaption(form, "1");
            if (s == "Add")
            {
                oItm.Enabled = false;
                return;
            }
            string remark = GetRemark(form);
            if (remark.IndexOf(strHeaderApproved) == -1)
            {
                oItm.Enabled = false;
            }
            else
            {
                if (remark.IndexOf(strHeaderCaptured) == -1)
                    oItm.Enabled = true;
                else
                    oItm.Enabled = false;
            }
            checkMethod(form);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    private void checkMethod(SAPbouiCOM.Form form)
    {
        try
        {
            string id = getSAPCustomerID(form);
            if (theSAPCustomer == null)
                theSAPCustomer = new CustomerObject();
            if (theSAPCustomer.CustomerID != id)
            {
                theSAPCustomerID = id;
                selectedPaymentMethodID = null;
                theSAPCustomer = getCustomer(id);
                ClearInput(form);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void CheckButton(SAPbouiCOM.Form form, string btnID, string strHeader)
    {
        try
        {
            SAPbouiCOM.Item oBPM = form.Items.Item(btnPayM); 
            SAPbouiCOM.Item oItm = form.Items.Item(btnID);
            string s = getButtonCaption(form, "1");
            if(s == "Add" && btnID!=btnPreAuth)
            {
                oBPM.Enabled = false;
                oItm.Enabled = false;
                return;
            }
            
            string remark = GetRemark(form);
            if (remark.IndexOf(strHeader) >= 0)
            {
                oBPM.Enabled = false;
                oItm.Enabled = false;
            }
            else
            {
                bool bShow = true;
                if (btnID == btnCredit || btnID == btnCharge)
                    bShow = CheckStatus(form);
                if (bShow)
                {
                    oBPM.Enabled = true;
                    oItm.Enabled = true;
                }else
                {
                    oBPM.Enabled = false;
                    oItm.Enabled = false;
                }
            }
            checkMethod(form);
            
        }catch(Exception ex)
        {
            errorLog(ex);
        }
        
    }
    private bool CheckStatus(SAPbouiCOM.Form form)
    {
        if (getComboValue(form , fidStatus) == "1")
            return true;
        return false;
    }
    private void AddTab(SAPbouiCOM.Form form)
    {
        try
        {
            SAPbouiCOM.Item oItem; 

            SAPbouiCOM.Item oItemRef;

            oItemRef = form.Items.Item("114");
            oItem = form.Items.Add(tbCreditCard, SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            oItem.Top = oItemRef.Top;
            oItem.Height = oItemRef.Height;
            oItem.Left = oItemRef.Left + oItemRef.Width;
            oItem.Width = oItemRef.Width;
            oItem.Visible = true;
            SAPbouiCOM.Folder oFolder = oItem.Specific;
            oFolder.Caption = CCTabTitle;
            oFolder.GroupWith(oItemRef.UniqueID);
            oFolder.Pane = 5;
            //Create a matrix on the folder
            oItem = form.Items.Add(mxtCCLog, SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oItem.FromPane = 5;
            oItem.ToPane = 5;
            oItemRef = form.Items.Item("39");
            oItem.Top = oItemRef.Top;
            oItem.Left = oItemRef.Left;
            oItem.Width = oItemRef.Width;
            oItem.Height = oItemRef.Height;

            SAPbouiCOM.Matrix oMatrix = form.Items.Item(mxtCCLog).Specific;
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Add("RefNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ref. No.";
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
            oColumn = oMatrix.Columns.Add("RecID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Rec. ID";
            oColumn.Width = 50;
            oColumn = oMatrix.Columns.Add("crCardNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Card No.";
            oColumn.Width = 50;
            
            oColumn = oMatrix.Columns.Add("result", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Result";
            oColumn.Width = 50;
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
            oColumn = oMatrix.Columns.Add("recdate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Process Date";
            oColumn.Width = 50;

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
       
            SAPbouiCOM.EditText customer = form.Items.Item(fidCustID).Specific;
            SAPbouiCOM.Matrix oMatrix = form.Items.Item(mxtCCLog).Specific;
            oMatrix.Clear();
            //oMatrix.AutoResizeColumns();
            SAPbouiCOM.DataTable oDts = getDataTable(form);
            
            sql = string.Format("select top " + GetConfig("maxSelectRow") +  " customerID,	customerName,crCardNum,Description,recID,acsUrl,authAmount,authCode,avsResult,avsResultCode,	batchNum,batchRef,cardCodeResult," +
            "cardCodeResultCode,cardLevelResult, cardLevelResultCode,conversionRate,convertedAmount,convertedAmountCurrency,custNum,error,errorCode,isDuplicate,payload,profilerScore,profilerResponse," +
            "profilerReason,refNum,remainingBalance,result,resultCode,	status,	statusCode,	vpasResultCode, recDate, Command, Amount from CCTRANS where CustomerID='{0}' order by recDate desc", customer.Value);
           
            oDts.ExecuteQuery(sql);
            BindMatrix(oMatrix, "RefNum", "refNum");
            BindMatrix(oMatrix, "command", "Command");
            BindMatrix(oMatrix, "Amount", "Amount");
            BindMatrix(oMatrix, "Desc", "Description");
            BindMatrix(oMatrix, "RecID", "recID");
            BindMatrix(oMatrix, "crCardNum", "crCardNum");
            BindMatrix(oMatrix, "result", "result");
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
    private void BindMatrix(SAPbouiCOM.Matrix oMatrix, string mtCol, string dtCol)
    {
        try
        {
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item(mtCol);
            oColumn.DataBind.Bind(dtCCTransLog, dtCol);

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
    private void addCreditCardForm(SAPbouiCOM.Form form)
    {
        try
        {
            theSAPForm = form;
            selectedPaymentMethodID = "";
            AddTab(form);
            string edID = edFormRefID;
            string stID = stFormRefID;
            int edL = form.Items.Item(edID).Left;
            int edW = form.Items.Item(edID).Width - 20;
            int edT = form.Items.Item(edID).Top;
            int edH = form.Items.Item(edID).Height;

            int stL = form.Items.Item(stID).Left;
            int stW = form.Items.Item(stID).Width;
            int stH = form.Items.Item(stID).Height;
            int nGap = 15;

            form.Items.Add(editCardNum, SAPbouiCOM.BoFormItemTypes.it_EDIT);
            SAPbouiCOM.Item oItm = form.Items.Item(editCardNum);
            oItm.Left = edL;
            edT = edT + nGap;
            oItm.Top = edT;
            oItm.Width = edW;
            oItm.Height = edH;
            
            form.Items.Add(lnkPayM, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oItm = form.Items.Item(lnkPayM);
            oItm.Left = edL - 30;
            oItm.Top = edT;
            oItm.Width = 20;
            oItm.Height = edH;
            SAPbouiCOM.LinkedButton lb = form.Items.Item(lnkPayM).Specific;
            lb.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
            lb.LinkedObjectType = "2";


            form.Items.Add(btnPayM, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItm = form.Items.Item(btnPayM);
            oItm.Left = edL + edW;
            oItm.Top = edT;
            oItm.Width = 20;
            oItm.Height = edH;
            setButtonCaption(form,btnPayM,"...");
            /*
            SAPbouiCOM.LinkedButton lb = form.Items.Item(lnkPayM).Specific;
            lb.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
            lb.LinkedObjectType = "2";
            */
            form.Items.Add(lbCardNum, SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItm = form.Items.Item(lbCardNum);
            oItm.Left = stL;
            oItm.Top = edT + 1;
            oItm.Width = stW - 40;
            oItm.Height = stH;
            setLabelCaption(form,lbCardNum,"Card Number:");

            form.Items.Add(editCardExpDate, SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItm = form.Items.Item(editCardExpDate);
            oItm.Left = edL;
            edT = edT + nGap;
            oItm.Top = edT;
            oItm.Width = edW;
            oItm.Height = edH;

            form.Items.Add(lbCardExpDate, SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItm = form.Items.Item(lbCardExpDate);
            oItm.Left = stL;
            oItm.Top = edT + 1;
            oItm.Width = stW;
            oItm.Height = stH;
            setLabelCaption(form,lbCardExpDate,"Card Exp. Date:");

            form.Items.Add(editCardCode, SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItm = form.Items.Item(editCardCode);
            oItm.Left = edL;
            edT = edT + nGap;
            oItm.Top = edT;
            oItm.Width = edW;
            oItm.Height = edH;

            form.Items.Add(lbCardCode, SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItm = form.Items.Item(lbCardCode);
            oItm.Left = stL;
            oItm.Top = edT + 1;
            oItm.Width = stW;
            oItm.Height = stH;
            setLabelCaption(form,lbCardCode,"Card Code:");

            form.Items.Add(editCAmount, SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
            oItm = form.Items.Item(editCAmount);
            oItm.Left = edL;
            edT = edT + nGap;
            oItm.Top = edT;
            oItm.Width = edW;
            oItm.Height = edH;

            form.Items.Add(lbCAmount, SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItm = form.Items.Item(lbCAmount);
            oItm.Left = stL;
            oItm.Top = edT + 1;
            oItm.Width = stW;
            oItm.Height = stH;
            setLabelCaption(form,lbCAmount,"Charge Amount:");

            ClearInput(form);
            
        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
    private void addCardHolderInfo(SAPbouiCOM.Form form)
    {
        try
        {
            if (testItem(form, editHolderName))
                return;
            int edL = 120;
            int edW = 150;
            int edT = 0;
            int edH = 15;

            int stL = 10;
            int stW = 100;
            int stH = 14;
            int nGap = 17;
            string cname = "";

            //name
            cname = editHolderName;
            form.Items.Add(cname, SAPbouiCOM.BoFormItemTypes.it_EDIT);
            SAPbouiCOM.Item oItm = form.Items.Item(cname);
            oItm.Left = edL;
            edT = edT + nGap;
            oItm.Top = edT;
            oItm.Width = edW;
            oItm.Height = edH;
            setFormEditVal(form, cname, theSAPCustomer.BillingAddress.FirstName + " " + theSAPCustomer.BillingAddress.LastName);
            addLabel(form, "LB1", stL, edT + 1, stW, stH, "Name:");
            //address
            cname = editHolderAddr;
            form.Items.Add(cname, SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItm = form.Items.Item(cname);
            oItm.Left = edL;
            edT = edT + nGap;
            oItm.Top = edT;
            oItm.Width = edW;
            oItm.Height = edH;
            setFormEditVal(form, cname, theSAPCustomer.BillingAddress.Street + " " + theSAPCustomer.BillingAddress.Street2);
            addLabel(form, "LB2", stL, edT + 1, stW, stH, "Address:");
            //city
            cname = editHolderCity;
            form.Items.Add(cname, SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItm = form.Items.Item(cname);
            oItm.Left = edL;
            edT = edT + nGap;
            oItm.Top = edT;
            oItm.Width = edW;
            oItm.Height = edH;
            setFormEditVal(form, cname, theSAPCustomer.BillingAddress.City);
            addLabel(form, "LB3", stL, edT + 1, stW, stH, "City:");
            //State
            cname = editHolderState;
            form.Items.Add(cname, SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItm = form.Items.Item(cname);
            oItm.Left = edL;
            edT = edT + nGap;
            oItm.Top = edT;
            oItm.Width = edW;
            oItm.Height = edH;
            setFormEditVal(form, cname, theSAPCustomer.BillingAddress.State);
            addLabel(form, "LB4", stL, edT + 1, stW, stH, "State:");
            //Zip
            cname = editHolderZip;
            form.Items.Add(cname, SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItm = form.Items.Item(cname);
            oItm.Left = edL;
            edT = edT + nGap;
            oItm.Top = edT;
            setFormEditVal(form, cname, theSAPCustomer.BillingAddress.Zip);
            addLabel(form, "LB5", stL, edT + 1, stW, stH, "Zip:");
            //Save
            form.Items.Add(cbSave, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            //form.Items.Add(cbSave, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            
            oItm = form.Items.Item(cbSave);
            oItm.Left = edL + 2;
            oItm.Width = 40;
            oItm.Height = edH;
            edT = edT + nGap;
            oItm.Top = edT;
            oItm.Enabled = true;
            
            SAPbouiCOM.ComboBox oCB = form.Items.Item(cbSave).Specific;
            oCB.ValidValues.Add("Y", "Yes");
            oCB.ValidValues.Add("N", "No");

            oItm.DisplayDesc = true;
             
            addLabel(form, "LB6", stL, edT + 1, stW, stH, "Save Card Info:");
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
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
    private void addPaymentMethodSelect(SAPbouiCOM.Form form)
    {
        try
        {
            if(testItem(form,cbPaymentMethod))
                return;
            int edL = 100;
            int edW = 150;
            int edT = 0;
            int edH = 15;

            int stL = 10;
            int stW = 80;
            int stH = 14;
            int nGap = 17;

            form.Items.Add(cbPaymentMethod, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            SAPbouiCOM.Item oItm = form.Items.Item(cbPaymentMethod);
            oItm.Left = edL;
            edT = edT + nGap;
            oItm.Top = edT;
            oItm.Width = edW;
            oItm.Height = edH;
            SAPbouiCOM.ComboBox oCB = form.Items.Item(cbPaymentMethod).Specific;
           
            foreach(PaymentMethod p in theSAPCustomer.PaymentMethods)
            {
                oCB.ValidValues.Add(p.MethodID,p.CardNumber);
            }
           
            form.Items.Add(lbCardNum, SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItm = form.Items.Item(lbCardNum);
            oItm.Left = stL;
            oItm.Top = edT + 1;
            oItm.Width = stW;
            oItm.Height = stH;
            setLabelCaption(form,lbCardNum, "Payment Method:");
        }
        catch (Exception ex)
        {
            errorLog(ex);
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
            setFormEditVal( form,editCAmount,strTotalAmount);
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

            if (amt == strTotalAmount)
            {
                string t = getFormEditVal(form,fidTotalAmount);    
                amt = t.Replace("$", "").Replace(",", "");

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

            if (a == strTotalAmount)
            {
                string t = getFormEditVal(form, fidTotalAmount);
                a = t.Replace("$", "").Replace(",", "");
                 setFormEditVal(form, editCAmount,a);
            }
            try
            {
                if(double.Parse(a)<= 0.01)
                {
                    SBO_Application.MessageBox("Invalid amount.  Please add item.", 1, "Ok", "", "");
                    return false;
                }
                amt = a;
            }
            catch (Exception)
            {
                SBO_Application.MessageBox("Invalid amount.  Please enter numerical value only.", 1, "Ok", "", "");
                return false;
            }

            updateMethod(form);
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
    private string GetConfig(string key)
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
        catch (Exception)
        {

        }
    }
} 

