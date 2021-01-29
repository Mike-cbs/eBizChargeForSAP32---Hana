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
using EBizEMV;
using Newtonsoft.Json;
using eBizChargeForSAP.ServiceReference1;

partial class SAP
{

    SAPbouiCOM.Form oTerminalForm;
    SAPbouiCOM.Form oTerminalLaunchForm;
    const string btnTerminal = "btnTerm";
    const string cbTerminal = "cbTerm";
    const string cbDefault = "cbDef";
    const string editURL = "edURL";
    const string editDeviceKey = "edDevK";
    const string editAPIKey = "edAPIK";
    const string editPIN = "edPIN";
    const string editName = "edName";
    const string stTransaction = "stTrans";
    const string stAmount = "stAmt";
    const string stDeviceName = "stDNAME";
    const string stDeviceStatus = "stDStat";
    const string stInvoice = "stInv";
    const string stTransResult = "stTRes";
    const string stTransRefnum = "stRefnum";
    const string stCreditCard = "stCCard";
    const string stCustomer = "stCust";
    const string menuTerminal = "eBizTerminalMenu";
    const string NEWDEVICE = "New Device";
    const string formTerminal = "eBizTerminalForm";
    

    string TermAmount = "";
    string TermFunc = "";
    string TermGroup = "";
    string TermInvoice = "";
    string TermCustomer = "";
    Thread tTermProcess;
    List<CCCTerminal> listTerminal = new List<CCCTerminal>();
    private void AddTerminalMenu()
    {
        try
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = SBO_Application.Menus.Item("2048"); // Sales A/R
            oMenus = oMenuItem.SubMenus;
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
            oCreationPackage.UniqueID = menuTerminal;
            oCreationPackage.String = "eBizCharge Terminal Setup";
            oCreationPackage.Enabled = true;
            oCreationPackage.Image = null;
            oCreationPackage.Position = 16;

            oMenus.AddEx(oCreationPackage);

        }
        catch (Exception ex)
        {
            if (ex.Message.IndexOf("66000-68") == -1)
                errorLog(ex);
        }

    }
    private void CreateTerminalForm()
    {
        try
        {
            
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Button oButton = null;
            //SAPbouiCOM.StaticText oStaticText = null;
            // SAPbouiCOM.EditText oEditText = null;
            //SAPbouiCOM.ComboBox oComboBox = null;

            // add a new form
            SAPbouiCOM.FormCreationParams oCreationParams = null;

            oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.FormType = formTerminal;

            oCreationParams.UniqueID = formTerminal;
            try
            {
                oTerminalForm = SBO_Application.Forms.AddEx(oCreationParams);
            }
            catch (Exception)
            {
                oTerminalForm = SBO_Application.Forms.Item(formTerminal);
              
                oTerminalForm.Visible = true;
              
            }
            trace("Create Terminal Form. ");

            // set the form properties
            oTerminalForm.Title = "eBizCharge Device Manager";
            oTerminalForm.Left = 400;
            oTerminalForm.Top = 100;
            oTerminalForm.ClientHeight = 460;
            oTerminalForm.ClientWidth = 450;

            //************************
            // Adding a Rectangle
            //***********************
            int margin = 5;
            oItem = oTerminalForm.Items.Add("Rect1", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = margin;
            oItem.Width = oTerminalForm.ClientWidth - 2 * margin;
            oItem.Top = margin;
            oItem.Height = oTerminalForm.ClientHeight - 40;

            int nTopGap = 25;
            int left = 6;
            int wBtn = 70;
            int hBtn = 19;
            int span = 80;

            oItem = oTerminalForm.Items.Add(btnProcess, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oTerminalForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));
            if(TermFunc == "")
                oButton.Caption = "Add";
            else
                oButton.Caption = "Process";
            left = oItem.Left + wBtn + 10;
            oItem = oTerminalForm.Items.Add(btnClose, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oTerminalForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";

            int edL = 110 + 15; //oItm.Left + oItm.Width;
            int edW = 250;
            int edT = 30;
            int edH = 15;
            int nGap = 22;
            int lbw = 100;
            oItem = addPaneItem(oTerminalForm, cbTerminal, edL, edT, edW, edH, "Terminal(s)", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 0, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oTerminalForm, editName, edL, edT, edW, edH, "Name:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oTerminalForm, editURL, edL, edT, edW, edH, "URL:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oTerminalForm, editDeviceKey, edL, edT, edW, edH, "Device Key:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 3, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oTerminalForm, editAPIKey, edL, edT, edW, edH, "API Key:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 4, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oTerminalForm, editPIN, edL, edT, edW, edH, "PIN:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 5, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oTerminalForm, cbDefault, edL, edT, 30, edH, "Default:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 6, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oTerminalForm, stTransaction, edL, edT, edW, edH, "Transaction:", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 8, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oTerminalForm, stAmount, edL, edT, edW, edH, "Amount:", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 9, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oTerminalForm, stCustomer, edL, edT, edW, edH, "Customer:", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 15, lbw);
            edT = oItem.Top + nGap;
            
            oItem = addPaneItem(oTerminalForm, stInvoice, edL, edT, edW, edH, "Invoice:", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 10, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oTerminalForm, stDeviceName, edL, edT, edW, edH, "Device Name:", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 11, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oTerminalForm, stDeviceStatus, edL, edT, edW, edH, "Device Status:", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 12, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oTerminalForm, stTransResult, edL, edT, edW, edH, "Transaction Result:", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 13, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oTerminalForm, stCreditCard, edL, edT, edW, edH, "Credit Card:", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 14, lbw);
            edT = oItem.Top + nGap;
            setLabelCaption(oTerminalForm, stTransaction, TermFunc);
                
            if (TermFunc == "")
            {
                setLabelCaption(oTerminalForm, "LB8", "");
                setLabelCaption(oTerminalForm, "LB9", "");
                setLabelCaption(oTerminalForm, "LB10", "");
                setLabelCaption(oTerminalForm, "LB13", "");
                setLabelCaption(oTerminalForm, "LB14", "");
                setLabelCaption(oTerminalForm, "LB15", "");

            }
            else
            {
                if (TermInvoice == "")
                {
                    TermInvoice = getInvoiceNum();
                }
                if (theActiveForm.TypeEx == FORMSALESORDER)
                    setLabelCaption(oTerminalForm, "LB10", "Sales Order:");
                if (theActiveForm.TypeEx == FORMCREDITMEMO)
                    setLabelCaption(oTerminalForm, "LB10", "Credit Memo:");
                if(getDoubleValue(TermAmount) < 0)
                {
                    TermFunc = "Credit";
                    TermAmount = TermAmount.Replace("-", "").Trim();
                }
                setLabelCaption(oTerminalForm, stAmount, TermAmount);
                setLabelCaption(oTerminalForm, stTransaction, TermFunc);
                setLabelCaption(oTerminalForm, stInvoice, TermInvoice);
                setLabelCaption(oTerminalForm, stCustomer, TermCustomer);
            }

            SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)oTerminalForm.Items.Item(cbDefault).Specific;
            oCB.ValidValues.Add("Y", "Y");
            oCB.ValidValues.Add("N", "N");

            loadTerminalInfo(oTerminalForm);

            oTerminalForm.Visible = true;
        }
        catch (Exception ex)
        {

            errorLog(ex);
        }

        oTerminalForm.Visible = true;
    }
    public void loadTerminalInfo(SAPbouiCOM.Form form, string newunit = "")
    {
        try
        {
            selectedTerminal = null;         

            SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbTerminal).Specific;
            try
            {
                while (oCB.ValidValues.Count > 0)
                    oCB.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception)
            { }

            oCB.ValidValues.Add(NEWDEVICE,NEWDEVICE);
            oCB.Select(NEWDEVICE);
            listTerminal = new List<CCCTerminal>();
            string sql = string.Format("select \"DocEntry\",\"U_name\",\"U_url\",\"U_deviceKey\",\"U_apiKey\",\"U_pin\",\"U_isDefault\" from \"@CCCTERMINAL\" Order by \"U_name\"");
          
            try
            {
                SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery(sql);
                while(!oRS.EoF)
                {
                    CCCTerminal t = new CCCTerminal();
                    t.recID = (int)oRS.Fields.Item(0).Value;
                    t.name = (string)oRS.Fields.Item(1).Value;
                    t.url = (string)oRS.Fields.Item(2).Value;
                    t.deviceKey = (string)oRS.Fields.Item(3).Value;
                    t.apiKey = (string)oRS.Fields.Item(4).Value;
                    t.pin = (string)oRS.Fields.Item(5).Value;
                    t.isDefault = (string)oRS.Fields.Item(6).Value;
                   
                    listTerminal.Add(t);
                    string key = string.Format("{0}_{1}", t.recID, t.name);
                    oCB.ValidValues.Add(key, key);
                    if(t.isDefault == "Y")
                    {
                        selectedTerminal = t;
                       
                    }
                    if (t.name.ToLower().IndexOf(oCompany.UserName.ToLower()) >= 0)
                    {
                        selectedTerminal = t;

                    }
                    if (t.name == newunit)
                    {
                        selectedTerminal = t;
                    }
                    oRS.MoveNext();

                }
                if(selectedTerminal != null)
                {
                    string key = selectedTerminal.recID.ToString() + "_" + selectedTerminal.name;
                    oCB.Select(key);
                    HandleTerminalSelect(form);
                }

            }
            catch (Exception ex)
            {
                errorLog(ex);
                errorLog(sql);
            }

        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
    private void TerminalFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            if (pVal.BeforeAction)
            {

            }
            else
            {
                switch (pVal.EventType)
                {
                    
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        {
                            switch (pVal.ItemUID)
                            {
                                case cbTerminal:
                                    HandleTerminalSelect(form);
                                    break;

                                case cbDefault:
                                    HandleTerminalDefaultSelect(form);
                                    break;
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        {
                            switch (pVal.ItemUID)
                            {
                                case btnClose:
                                    try
                                    {
                                        if (tTermProcess.IsAlive)
                                        {
                                            if (SBO_Application.MessageBox("Exit waiting thread?", 1, "Yes", "No") != 1)
                                                return;
                                            tTermProcess.Abort();
                                            if (etran != null)
                                            {
                                                if (SBO_Application.MessageBox("Transaction may occurred.  Do you want to create payment record?", 1, "Yes", "No") == 1)
                                                {
                                                    string acctid = "";

                                                    string cardname = getCardName("", "", "", ref acctid);
                                                    TAddPayment(TermCustomer, TermInvoice, acctid, getMoneyString(TermAmount), "", "");
                                                    etran = null;
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception) { }
                                    oTerminalForm.Close();
                                    try
                                    {
                                        if (oTerminalLaunchForm != null)
                                            oTerminalLaunchForm.Close();
                                        if(theActiveForm.TypeEx == FORMINVOICE)
                                        {
                                            tProcess = new Thread(RecordRefreshThreadProc);
                                            tProcess.Start();

                                        }else
                                        {
                                            SBO_Application.Menus.Item(menuIDLastRecord).Activate();
                                        }
                                    }catch(Exception){}
                                    break;
                                case btnProcess:
                                    HandleButtonPress(form);
                                    break;
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
    CCCTerminal selectedTerminal;
    private void HandleTerminalSelect(SAPbouiCOM.Form form)
    {
        //errorLog("Handle select");
        try
        {
            FormFreeze(form,true);
            string sel = getFormItemVal(form, cbTerminal);
            if (sel == NEWDEVICE)
            {
                setFormEditVal(form, editName, "");
                setFormEditVal(form, editURL, "");
                setFormEditVal(form, editDeviceKey, "");
                setFormEditVal(form, editAPIKey, "");
                setFormEditVal(form, editPIN, "");
                setComboValue(form, cbDefault, "N");
                if (TermFunc == "")
                    setButtonCaption(form, btnProcess, "Add");
            }
            else
            {
                string[] key = sel.Split('_');
                foreach (CCCTerminal t in listTerminal)
                {
                    if (t.name == key[1])
                    {
                        selectedTerminal = t;
                        setFormEditVal(form, editName, t.name);
                        setFormEditVal(form, editURL, t.url);
                        setFormEditVal(form, editDeviceKey, t.deviceKey);
                        setFormEditVal(form, editAPIKey, t.apiKey);
                        setFormEditVal(form, editPIN, t.pin);
                        setComboValue(form, cbDefault, t.isDefault);
                        if(TermFunc =="")
                            setButtonCaption(form, btnProcess, "Update");
                    }
                }

            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        finally
        {
            FormFreeze(form,false);
          
        }
    }
    private void HandleButtonPress(SAPbouiCOM.Form form)
    {
        //errorLog("Handle select");
        try
        {
            FormFreeze(form,true);
            string sel = getFormItemVal(form, btnProcess);
            if (sel == "Add")
            {
                CCCTerminal t = new CCCTerminal();
                t.name = getFormItemVal(form, editName);
                t.url = getFormItemVal(form, editURL);
                t.deviceKey = getFormItemVal(form, editDeviceKey);
                t.apiKey = getFormItemVal(form, editAPIKey);
                t.pin = getFormItemVal(form, editPIN);
                t.isDefault = getFormItemVal(form, cbDefault);
                if (!validateTerminal(form, t))
                    return;
                if(addDevice(t))
                {
                    showMessage("Device Added");
                }else
                    showMessage("Add device failed.");

                loadTerminalInfo(form, t.name);
                return;
            }
            if (sel == "Update")
            {
                CCCTerminal t = new CCCTerminal();
                t.recID = selectedTerminal.recID;
                t.name = getFormItemVal(form, editName);
                t.url = getFormItemVal(form, editURL);
                t.deviceKey = getFormItemVal(form, editDeviceKey);
                t.apiKey = getFormItemVal(form, editAPIKey);
                t.pin = getFormItemVal(form, editPIN);
                t.isDefault = getFormItemVal(form, cbDefault);
                if (!validateTerminal(form,t))
                    return;
                if (updateDevice(t))
                {
                    showMessage("Device updated");
                }
                else
                    showMessage("Update device failed.");

                loadTerminalInfo(form);
                SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbTerminal).Specific;
                string key = t.recID.ToString() + "_" + t.name;
                oCB.Select(key);
                HandleTerminalSelect(form);
                return;
            }
            if (sel == "Process")
            {
                if(selectedTerminal == null)
                {
                    showMessage("Please select a device.");
                    return;
                }
                if (!validateTerminal(form, selectedTerminal))
                    return;
                string status = getFormItemVal(form, stDeviceStatus);
                if(status !="connected")
                {
                    showMessage("Device is not ready to process transaction.  Please reset device.");
                    return;
                }
                if (tTermProcess != null && etran != null)
                {
                    if (tTermProcess.IsAlive)
                    {
                        string resp = EMVHelper.Payrequest(selectedTerminal.url, selectedTerminal.deviceKey, selectedTerminal.apiKey, selectedTerminal.pin, etran, etran.command);
                        resp = EMVHelper.GetPayrequestResult(selectedTerminal.url, selectedTerminal.deviceKey, selectedTerminal.apiKey, selectedTerminal.pin, resp);
                        var res = JsonConvert.DeserializeObject<TerminalResp>(resp);
                        if (res.status.ToLower() == "transaction complete")
                        {
                            handleComplete(res);
                            showMessage("Transaction completed.");
                        }
                        else
                        {
                            if (SBO_Application.MessageBox("Waiting for device response.  If you wish to continue please make sure transaction is cancelled. Stop waiting?", 1, "Yes", "No") == 1)
                            {
                                tTermProcess.Abort();
                                etran = null;
                            }
                            else
                                return;
                        }
                    }
                }
                if (SBO_Application.MessageBox("Send request to " + TermFunc.ToLower() + " customer " + TermAmount + "?", 1, "Yes", "No") == 1)
                {
                    oTerminalForm.Items.Item(btnProcess).Enabled = false;
                    tTermProcess = new Thread(TerminalThreadProc);
                    tTermProcess.Start();
                }
            }
        }
        catch (Exception ex)
        {
            showMessage(ex.Message);
            errorLog(ex);
        }
        finally
        {
            try
            {
                FormFreeze(form,false);
            }
            catch (Exception) { }
        }
    }
    public void handleComplete(TerminalResp res)
    {
        try
        {
            trace("Terminal Transaction completed: " + res.transaction.result + " " + res.transaction.error + " " + res.transaction.avs.result + " authorized code:" + res.transaction.refnum);
            setLabelCaption(oTerminalForm, stTransResult, res.transaction.result + " " + res.transaction.error + " " + res.transaction.avs.result + " authorized code:" + res.transaction.refnum);
            setLabelCaption(oTerminalForm, stCreditCard, res.transaction.creditcard.number);
            if (res.transaction.result.ToLower() == "approved")
            {
                if (IsTransactionExists(res.transaction.refnum, TermAmount))
                {
                    trace("terminal transaction refNum: " + res.transaction.refnum + ", exists. ");
                }
                else
                {

                    CCTRAN cctran = new CCTRAN();
                    cctran.amount = TermAmount;
                    cctran.refNum = res.transaction.refnum;
                    cctran.CardHolder = res.transaction.creditcard.cardholder;
                    cctran.avsResult = res.transaction.avs.result;
                    cctran.command = "cc:sale";
                    cctran.result = res.transaction.result;
                    addToCCTRAN(cctran);
                    printReceipt(res.transaction.refnum);
                    string currency = getCurrency("OINV", TermInvoice);
                    string group = getGroupName(TermCustomer);

                    string acctid = "";
                    string cardname = getCardName(group, res.transaction.creditcard.type, currency, ref acctid);
                    TAddPayment(TermCustomer, TermInvoice, acctid, getMoneyString(TermAmount), res.transaction.creditcard.number, res.transaction.refnum);
                }
                setButtonCaption(oTerminalForm, btnClose, "Close");
                etran = null;
            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
    public bool validateTerminal(SAPbouiCOM.Form form, CCCTerminal t)
    {
        try
        {
            string error = "";
            if (t.name == "")
            {
                error = "Name is required.\r\n";
            }
            if (t.url == "")
            {
                error = error + "URL is required.\r\n";
            }
            if (t.deviceKey == "")
            {
                error = error + "Device Key is required.\r\n";
            }
            if (t.apiKey == "")
            {
                error = error + "API Key is required.\r\n";
            }
            if (t.pin == "")
            {
                error = error + "PIN is required.\r\n";
            }
            if (error != "")
            {
                showMessage(error);
                return false;
            }
            string resp = EMVHelper.GetDevices(t.url, t.deviceKey, t.apiKey, t.pin);
            var res = JsonConvert.DeserializeObject<TerminalResp>(resp);
            setLabelCaption(form, stDeviceName, res.name);
            setLabelCaption(form, stDeviceStatus, res.status);
            return true;
        }
        catch (Exception ex)
        {
            if (ex.Message.ToLower().IndexOf("not found") >= 0)
            {
                showMessage("Device not found.");

            }
            else
                showMessage(ex.Message);
        }
        return true;
    }
    private void HandleTerminalDefaultSelect(SAPbouiCOM.Form form)
    {
        try
        {

            if (selSAPCust == null)
                return;
            string def = getFormItemVal(form, cbDefault);
            if (def != selectedTerminal.isDefault)
            {
                string msg = "Select device as default?";
                if (def == "N")
                    msg = "Unselect device as default?";
               if (SBO_Application.MessageBox(msg, 1, "Yes", "No") == 1)
               {
                    if (updateTerminalDefault(selectedTerminal.recID, def))
                    {
                        if (def == "Y")
                            showMessage(selectedTerminal.name + " selected as default.");
                        else
                            showMessage(selectedTerminal.name + " unselected as default.");

                        loadTerminalInfo(form);
                    }
                    else
                    {
                        showMessage("Failed to update default status.");
                    }
                }
             }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    public bool updateTerminalDefault(int id, string def)
    {
        string sql = string.Format("UPDATE CCCTerminal set [isDefault] = '{0}' where recID={1}", def, id);
        try
        {

            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);
            
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            return false;
        }
        return true;
    }
    public bool addDevice(CCCTerminal t)
    {
        int id = getNextTableID("@CCCTERMINAL");
        string sql = string.Format("insert into \"@CCCTERMINAL\" (\"DocEntry\",\"U_name\",\"U_url\",\"U_deviceKey\",\"U_apiKey\",\"U_pin\",\"U_isDefault\") values({6}, '{0}','{1}','{2}','{3}','{4}','{5}')"
            , t.name,t.url,t.deviceKey,t.apiKey,t.pin, t.isDefault, id);
        try
        {

            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            return false;
        }
        return true;
    }
    public bool updateDevice(CCCTerminal t)
    {
        string sql = string.Format("UPDATE \"@CCCTERMINAL\" set \"U_name\" = '{0}', \"U_url\" = '{1}', \"U_deviceKey\"='{2}', \"U_apiKey\"='{3}', \"U_pin\"='{4}',\"U_isDefault\"='{5}' where \"DocEntry\"={6}"
            , t.name, t.url, t.deviceKey, t.apiKey, t.pin, t.isDefault, t.recID);
        try
        {

            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            return false;
        }
        return true;
    }
    EBizEMV.TransactionRequest etran = null;
    public void TerminalThreadProc()
    {
        try
        {
            etran = new EBizEMV.TransactionRequest();
            etran.command = "cc:sale";
            if (TermFunc == "Credit")
                etran.command = "cc:credit";    
            etran.amount = TermAmount;
            etran.invoice = TermInvoice;
            etran.manual_key = false;
            int i = 0;
            try
            {
                SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                if (TermFunc == "Credit")
                    oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
                if (theActiveForm.TypeEx == FORMSALESORDER)
                    oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

                if (oDoc.GetByKey(int.Parse(TermInvoice)))
                {
                    SAPbobsCOM.Document_Lines olines = oDoc.Lines;
                    etran.lineitems = new EBizEMV.LineItem[olines.Count];
                    for (i = 0; i < olines.Count; i++)
                    {
                        olines.SetCurrentLine(i);
                        etran.lineitems[i] = new EBizEMV.LineItem();
                        etran.lineitems[i].sku = olines.ItemCode;
                        etran.lineitems[i].description = olines.ItemDescription;
                        etran.lineitems[i].qty = olines.Quantity.ToString();
                        etran.lineitems[i].cost = olines.UnitPrice.ToString();
                        etran.lineitems[i].taxamount = olines.TaxTotal.ToString();
                        etran.lineitems[i].discountrate = olines.DiscountPercent.ToString();
                        etran.lineitems[i].um = "EACH";
                        etran.lineitems[i].name = olines.ItemCode;
                    }
                }
                else
                {
                    etran.lineitems = new EBizEMV.LineItem[1];
                    etran.lineitems[0] = new EBizEMV.LineItem();
                    etran.lineitems[0].sku = "SKU1";
                    etran.lineitems[0].description = "B1 EMV";
                    etran.lineitems[0].qty = "1";
                    etran.lineitems[0].cost = TermAmount;
                    etran.lineitems[0].taxamount = "";
                    etran.lineitems[0].um = "EACH";
                    etran.lineitems[0].name = "B1 EMV";

                }

            }
            catch (Exception ex)
            {
                errorLog(ex);
            }
            trace("etran.command:" + etran.command + ",etran.amount:" + etran.amount);

            string resp = EMVHelper.Payrequest(selectedTerminal.url, selectedTerminal.deviceKey, selectedTerminal.apiKey, selectedTerminal.pin, etran, etran.command);
            string wait =  ". Please wait.";
            i = 0;
            int max = 20;
            while (i < max)
            {
                i++;
                resp = EMVHelper.GetPayrequestResult(selectedTerminal.url, selectedTerminal.deviceKey, selectedTerminal.apiKey, selectedTerminal.pin, resp);
               
                var res = JsonConvert.DeserializeObject<TerminalResp>(resp);
                setLabelCaption(oTerminalForm, stDeviceStatus, res.status);
                switch (res.status.ToLower())
                {
                    case "transaction canceled":
                        bAuto = true;
                        i = max + 1;
                        oTerminalForm.Items.Item(btnProcess).Enabled = true;
                        etran = null;
                        break;
                    case "transaction complete":
                        trace("Terminal Transaction completed: " + res.transaction.result + " " + res.transaction.error + " " + res.transaction.avs.result + " authorized code:" + res.transaction.refnum);
                        setLabelCaption(oTerminalForm, stTransResult, res.transaction.result + " " + res.transaction.error + " " + res.transaction.avs.result + " authorized code:" + res.transaction.refnum);
                        setLabelCaption(oTerminalForm, stCreditCard, res.transaction.creditcard.number);
                        if (res.transaction.result.ToLower() == "approved")
                        {
                            handleComplete(res);
                        }
                        

                        i = max + 1;
                        break;
                    default:
                        wait = wait + ".";
                        setLabelCaption(oTerminalForm, stDeviceStatus, res.status + wait);
                        System.Threading.Thread.Sleep(1000);
                        if (i == max - 1)
                        {
                            if (SBO_Application.MessageBox("Device not responding continue to wait?", 1, "Yes", "No") == 1)
                            {
                                i = 0;

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
    public void addToCCTRAN(CCTRAN cctran)
    {
        try
        {
            cctran.customerID = TermCustomer;
            switch (theActiveForm.TypeEx)
            {
                case FORMINVOICE:
                    cctran.InvoiceID = theActiveForm.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0);
                    break;
                case FORMSALESORDER:
                    cctran.OrderID = theActiveForm.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);
                    break;
                case FORMINCOMINGPAYMENT:
                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)theActiveForm.Items.Item("20").Specific;
                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        if (getMatrixSelect(oMatrix, "10000127", i))
                        {
                            string type = getMatrixItem(oMatrix, "45", i);
                            string DocNum = getMatrixItem(oMatrix, "1", i);
                            if (type == "13")
                                cctran.InvoiceID = getDocEntry(type, DocNum, TermCustomer).ToString();
                        }
                    }
                    break;
            }
            insert(cctran);

        }
        catch (Exception ex)
        {

        }
    }
    public bool TAddPayment(string Custid, string InvoiceID, string cardID, string amount, string cardNum, string refNum)
    {
        try
        {
            trace(string.Format("Custid = {0}, InvoiceID={1}, cardID={2}, amount={3}, cardNum={4}, refNum={5}", Custid, InvoiceID, cardID, amount, cardNum, refNum));
            int bplid = int.Parse(getBranchID());
            switch (theActiveForm.TypeEx)
            {
                case FORMSALESORDER:
                    {
                        showStatus("eBizCharge: adding incoming payment to downpayment invoice.  Sales Order ID: " + InvoiceID + ".  Amount: " + amount, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                        DPSAPCust = getCustomerByID(Custid, "");
                        SOChargeLastRecordThreadProc(getDoubleValue(amount), theActiveForm);
                        showStatus("eBizCharge: Incoming payment of " + amount + " added to downpayment invoice.  Sales Order ID: " + InvoiceID, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    }
                    break;
                case FORMINVOICE:
                    {
                        SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                        int invID = getDocEntry("13", InvoiceID, Custid);
                        if (!oDoc.GetByKey(invID))
                        {
                            string err = string.Format("Cannot find the invoice document for customer: {0}, Invoice Number: {1}, DocEntry: {2}\r\n{3}", Custid, InvoiceID, invID, oCompany.GetLastErrorDescription());
                            errorLog(err);
                            showMessage(err);

                            return false;
                        }
                        SAPbobsCOM.Payments oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                        
                        trace("Invoice Branch ID: " + bplid);
                        if (bplid != 0)
                            oPmt.BPLID = bplid;
                        oPmt.Invoices.DocEntry = invID;
                        oPmt.CardCode = oDoc.CardCode;
                        oPmt.DocDate = DateTime.Now;
                        oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
                        TAddCreditCardPayment(oPmt, Custid, cardNum, getDoubleValue(amount), int.Parse(cardID), refNum);
                        string sql = "select \"InstlmntID\", \"InsTotal\" , \"InsTotalFC\", \"Paid\"  from \"INV6\" where \"Status\" = 'O' and \"DocEntry\" = " + invID;
                        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRS.DoQuery(sql);
                        double remain = getDoubleValue(amount);
                        while (!oRS.EoF)
                        {
                            double insTotal = (double)oRS.Fields.Item(1).Value;
                            double insTotalFC = (double)oRS.Fields.Item(2).Value;
                            double paid = (double)oRS.Fields.Item(3).Value;

                            double bal = (insTotal == 0 ? insTotalFC : insTotal) - paid;
                            if (remain >= bal)
                            {
                                oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice;
                                oPmt.Invoices.SumApplied = bal;
                                oPmt.Invoices.InstallmentId = (int)oRS.Fields.Item(0).Value;
                                oPmt.Invoices.DocEntry = invID;
                                oPmt.Invoices.Add();
                            }
                            else if (remain > 0)
                            {
                                oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice;
                                oPmt.Invoices.SumApplied = remain;
                                oPmt.Invoices.InstallmentId = (int)oRS.Fields.Item(0).Value;
                                oPmt.Invoices.DocEntry = invID;
                                oPmt.Invoices.Add();
                            }
                            remain = Math.Round(remain - bal, 2);
                            oRS.MoveNext();
                        }
                   
                        oPmt.CardCode = oDoc.CardCode;
                        int r = oPmt.Add();
                        if (r != 0)
                        {
                            errorLog(oCompany.GetLastErrorDescription());
                            showMessage("Can not create payment for invoice: " + InvoiceID + "\r\n" + oCompany.GetLastErrorDescription());
                            voidCustomer(confirmNum);
                            return false;
                        }
                    }
                    break;
                case FORMINCOMINGPAYMENT:
                    {
                        SAPbobsCOM.Payments oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                        if (bplid != 0)
                            oPmt.BPLID = bplid;
                        if (TermFunc == "Credit")
                        {
                            amount = "-" + getDoubleValue(amount);
                        }
                        oPmt.CardCode = getFormItemVal(theActiveForm, "5");
                        oPmt.DocDate = DateTime.Now;
                        oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
                        oPmt.Remarks = getFormItemVal(theActiveForm, "26");
                        oPmt.JournalRemarks = getFormItemVal(theActiveForm, "59");
                        trace("reference = " + getFormItemVal(theActiveForm, "22"));
                        oPmt.CounterReference = getFormItemVal(theActiveForm, "22");
                        oPmt.DocDate = DateTime.Now;
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)theActiveForm.Items.Item("20").Specific;
                        List<String> ids = new List<string>();

                        for (int i = 1; i <= oMatrix.RowCount; i++)
                        {
                            if (getMatrixSelect(oMatrix, "10000127", i))
                            {
                                string type = getMatrixItem(oMatrix, "45", i);
                                string DocNum = getMatrixItem(oMatrix, "1", i);
                                string DocDate = getMatrixItem(oMatrix, "21", i);
                                string amt = getMatrixItem(oMatrix, "24", i);
                                int docEntry = getDocEntry(type, DocNum, Custid);

                                string dcount = getMatrixItem(oMatrix, "29", i);
                                string installment = getMatrixItem(oMatrix, "71", i);
                                installment = installment.Split(' ')[0];

                                trace("Incoming payment type=" + type + ", DocNum=" + DocNum + ",docEntry = " + docEntry.ToString());
                                oPmt.Invoices.DocEntry = docEntry;
                                try
                                {
                                    if (type == "13")
                                    {
                                        try
                                        {
                                            SAPbobsCOM.Documents oInv = (SAPbobsCOM.Documents)oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)int.Parse(type));
                                            oInv.GetByKey(docEntry);
                                            trace("Payment Block: " + oInv.PaymentBlock);
                                            oInv.PaymentBlock = SAPbobsCOM.BoYesNoEnum.tNO;
                                            oInv.Update();
                                            oPmt.Invoices.InstallmentId = int.Parse(installment);
                                        }
                                        catch (Exception ex)
                                        {
                                            trace("Invoice update error. " + ex.Message);
                                        }

                                    }
                                }
                                catch (Exception) { }

                                oPmt.Invoices.InvoiceType = (SAPbobsCOM.BoRcptInvTypes)int.Parse(type);

                                oPmt.Invoices.SumApplied = getDoubleValue(getMoneyString(amt));
                                if (type == "24" || type == "30" || type == "-2")
                                {
                                    oPmt.Invoices.DocLine = getJELineID(docEntry.ToString(), Custid);
                                }
                                //else
                                //oPmt.Invoices.DocLine = 0;
                                try
                                {
                                    oPmt.Invoices.DiscountPercent = getDoubleValue(dcount);
                                }
                                catch (Exception)
                                {

                                }
                                oPmt.Invoices.Add();

                            }
                        }

                        TAddCreditCardPayment(oPmt, Custid, cardNum, getDoubleValue(amount), int.Parse(cardID), refNum);
                        int r = oPmt.Add();
                        if (r != 0)
                        {
                            errorLog(oCompany.GetLastErrorDescription());
                            bAuto = false;
                            showMessage("Cannot add incomming payment.\r\n" + oCompany.GetLastErrorDescription());
                            voidCustomer(confirmNum);
                            return false;
                        }
                        return true;
                    }
                    break;
                case FORMCREDITMEMO:
                    {
                        SAPbobsCOM.Documents oDoc = null;
                        oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
                        int cmID = getDocEntry("14", InvoiceID, Custid);
                        if (!oDoc.GetByKey(cmID))
                        {
                            bAuto = false;
                            showMessage("Cannot get credit memo. docNum=" + InvoiceID + "\r\n" + oCompany.GetLastErrorDescription());
                            return false;
                        }
                        SAPbobsCOM.Payments oPmt;
                        double total = getDoubleValue(amount);
                        if (cfgNegativeIncomingPayment == "Y")
                        {
                            oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                            total = -total;
                        }
                        else
                        {
                            oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);

                        }
                        trace("Credit Memo Branch ID: " + bplid);

                        if (bplid != 0)
                            oPmt.BPLID = bplid;
                        oPmt.Invoices.DocEntry = cmID;
                        oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote;
                        oPmt.Invoices.SumApplied = total;
                        //oPmt.CashAccount = GetConfig("SAPAccountForOutgoingPayment");
                        oPmt.CardCode = oDoc.CardCode;
                        //oPmt.Reference2 = confirmNum;
                        oPmt.DocDate = DateTime.Now;
                        oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
                        //oPmt.CashSum = total;
                        //oPmt.Remarks = "eBizCharge " + confirmNum;
                        // TAddCreditCardPayment(oPmt, Custid, cardNum, getDoubleValue(amount), int.Parse(cardID), refNum);
                        confirmNum = refNum;
                        AddCreditCardPayment(oPmt, Custid, CardNum, "cc", total, cardID, false);
                        int r = oPmt.Add();
                        if (r != 0)
                        {
                            errorLog(oCompany.GetLastErrorDescription());
                            bAuto = false;
                            showMessage(oCompany.GetLastErrorDescription());
                            voidCustomer(confirmNum);
                            return false;
                        }
                    }
                    break;
            }





        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    public void TAddCreditCardPayment(SAPbobsCOM.Payments vPay, string customerID, string cardNumber, double total, int creditcardID, string refNum)
    {
        try
        {
            vPay.CreditCards.AdditionalPaymentSum = 0;

            vPay.CreditCards.CardValidUntil = DateTime.Today.AddYears(2);

            vPay.CreditCards.CreditCard = creditcardID;
         
            vPay.CreditCards.CreditCardNumber = cardNumber;
            //vPay.CreditCards.CreditCur = "EUR"
            //vPay.CreditCards.CreditRate = 0
            vPay.CreditCards.CreditSum = total;
            //vPay.CreditCards.CreditType = 1;
            vPay.CreditCards.FirstPaymentDue = DateTime.Now;
            vPay.CreditCards.FirstPaymentSum = total;
            vPay.CreditCards.NumOfCreditPayments = 1;
            vPay.CreditCards.NumOfPayments = 1;
            vPay.CreditCards.VoucherNum = "eBizCharge";
            vPay.CreditCards.ConfirmationNum = refNum;
            vPay.CreditCards.OwnerIdNum = customerID;
            // vPay.CreditCards.OwnerPhone = "383838888"
            vPay.CreditCards.PaymentMethodCode = getPaymentTypeCode(creditcardID);
            vPay.CreditCards.Add();
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void printReceipt(string refNum)
    {
        try
        {
            if (cfgSaveReceipt != "Y")
                return;
            string folder = @"C:\CBS\";
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            folder = @"C:\CBS\Receipts\";
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            string filename = folder + refNum + ".html";
            SecurityToken token = getToken("");
            string s = ebiz.RenderReceipt(token, refNum, null, "vterm", "HTML");
            if (s != "")
            {

                byte[] todecode = Convert.FromBase64String(s);
                s = string.Concat(System.Text.Encoding.UTF8.GetString(todecode));
                s = s.Replace("<tr><td align=right><b>AMOUNT:</b></td>", "<tr style='display:none'><td align=right><b>AMOUNT:</b></td>");
                s = s.Replace("<tr><td align=right><b>TAX:</b></td>", "<tr style='display:none'><td align=right><b>TAX:</b></td>");

                FileStream fs = new FileStream(filename, FileMode.Create, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs);
                sw.Write(s);
                sw.Close();
                System.Diagnostics.Process.Start(filename);
            }
            try
            {
                if (userEmail != "")
                    ebiz.EmailReceipt(token, refNum, null, "vterm", userEmail);

            }
            catch (Exception ex)
            {
                errorLog(ex);
            }

        }
        catch (Exception ex2)
        {
            errorLog(ex2);
        }
    }
}

public class TerminalResp
{
    public string status { get; set; }
    public string name { get; set; }
    public string type { get; set; }
    public string key { get; set; }
    public terminalTransaction transaction { get; set; }
}
public class terminalTransaction
{
    public string type { get; set; }
    public string key { get; set; }
    public string refnum { get; set; }
    public string result { get; set; }
    public string error { get; set; }
    public AVS avs { get; set; }
    public CVC cvc { get; set; }
    public terminalCreditCard creditcard { get; set; }

}
public class terminalCreditCard
{
    public string number { get; set; }
    public string type { get; set; }
    public string cardholder { get; set; }

}
public class AVS
{
    public string result_code {get;set;}
    public string result { get; set; }
}
public class CVC
{
    public string result_code { get; set; }
    public string result { get; set; }
}