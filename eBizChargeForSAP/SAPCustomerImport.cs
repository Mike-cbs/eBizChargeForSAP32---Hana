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
using System.IO;
using eBizChargeForSAP.ServiceReference1;
public class ImportData
{
    public string CustomerID { get; set; }
    public string CompanyName { get; set; }
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public string Street { get; set; }
    public string Street2 { get; set; }
    public string Zip { get; set; }
    public string City { get; set; }
    public string State { get; set; }
    public string Phone { get; set; }
    public string Fax { get; set; }
    public string Email { get; set; }
    public string URL { get; set; }
    public string CardNumber { get; set; }
    public string CardExpiration { get; set; }
    public string MethodName { get; set; }
    public string defaultCard { get; set; }
    public string CustNum { get; set; }
}
public class CustomerData
{
    public CustomerData()
    {
        list = new List<ImportData>();
        total = 0;
    }
    public string code { get; set; }
    public string checkNum { get; set; }
    public string checkDate { get; set; }
    public double total { get; set; }
    public List<ImportData> list { get; set; }
}
partial class SAP
{
    const string menuCustImp = "eBizCustImpMenu";
    const string formCustImp = "eBizCustImpForm";
    const string matrixCustImp = "mBInv";
    const string btnImp = "btnImp";
    const string btnExpFImp = "btnEFImp";
    const string btnFImp = "btnFImp";
    const string editCustNum = "edCustNum";
    const string editFileName = "edXFN";
    const string editExpFileName = "edEXFN";
    const string editImportLog = "edxLog";
    //  const string cbUsePaymentName = "cbUPM";
     Customer custObjForImport = null;


    SAPbouiCOM.Form oCustImpForm;

    private void AddCustImpMenu()
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
            oCreationPackage.UniqueID = menuCustImp;
            oCreationPackage.String = "eBizCharge Payment Method Import";
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
    private void CreateCustImpForm()
    {
        try
        {
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Button oButton = null;

            // add a new form
            SAPbouiCOM.FormCreationParams oCreationParams = null;

            oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.FormType = formCustImp;

            oCreationParams.UniqueID = formCustImp;
            try
            {
                oCustImpForm = SBO_Application.Forms.AddEx(oCreationParams);
            }
            catch (Exception)
            {
                oCustImpForm = SBO_Application.Forms.Item(formCustImp);

            }

            // set the form properties
            oCustImpForm.Title = "eBizCharge Payment Method Import";
            oCustImpForm.Left = 400;
            oCustImpForm.Top = 100;
            oCustImpForm.ClientHeight = 320;
            oCustImpForm.ClientWidth = 550;






            int nTopGap = 25;
            int left = 6;
            int wBtn = 70;
            int hBtn = 19;
            int span = 80;

            oItem = oCustImpForm.Items.Add(btnFind, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oCustImpForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;
            oItem.Visible = false;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Find";
            left += span;
            oItem = oCustImpForm.Items.Add(btnImp, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oCustImpForm.ClientHeight - nTopGap;
            oItem.Visible = false;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Import";
            left += span;

            oItem = oCustImpForm.Items.Add(btnClose, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oCustImpForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";

            int margin = 8;
            int top = 15;
            int edL = 200; //oItm.Left + oItm.Width;
            int edW = 200;
            int edH = 15;
            int nGap = 26;

            /*
                List<String> list = FindCustomer("");
           
                    SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)oCustImpForm.Items.Item(cbCustomerID).Specific;
                    try
                    {
                        while (oCB.ValidValues.Count > 0)
                            oCB.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    catch (Exception)
                    { }

                    foreach (string c in list)
                    {
                        ComboAddItem(oCB, c);
                    }
               */

            oItem = addPaneItem(oCustImpForm, editFileName, edL, top, edW, edH, "Customer Import File(tab delimited):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2035, 200);
            oItem.Width = 240;
            int l = oItem.Left + oItem.Width + 10;
            oItem = oCustImpForm.Items.Add(btnFImp, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = l;
            oItem.Width = wBtn;
            oItem.Top = top;
            oItem.Height = hBtn - 2;
            oButton = ((SAPbouiCOM.Button)(oItem.Specific));
            oButton.Caption = "File Import";
            top = oItem.Top + nGap;
            oItem = addPaneItem(oCustImpForm, editExpFileName, edL, top, edW, edH, "Exported Customer File(tab delimited):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2038, 200);
            oItem.Width = 240;
            oItem = oCustImpForm.Items.Add(btnExpFImp, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = l;
            oItem.Width = wBtn;
            oItem.Top = top;
            oItem.Height = hBtn - 2;
            oButton = ((SAPbouiCOM.Button)(oItem.Specific));
            oButton.Caption = "File Import";
            top = oItem.Top + nGap;
            oItem = addPaneItem(oCustImpForm, editImportLog, edL, top, edW, edH, "import Log:", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT, 0, 2036, 150);
            oItem.Height = 80;
            oItem.Width = 240;
            top = oItem.Top + 120;
            oItem = addPaneItem(oCustImpForm, editCustNum, edL, top, edW, edH, "Gateway cardholder ID:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 100, 150);
            oItem.Visible = false;
            oCustImpForm.Items.Item("LB100").Visible = false;

            top = oItem.Top + nGap;

            oItem = addPaneItem(oCustImpForm, cbPaymentMethod, edL, top, edW, edH, "Payment method(s) found:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 200, 150);
            oItem.Visible = false;
            oCustImpForm.Items.Item("LB200").Visible = false;


            top = oItem.Top + nGap;

            //   oItem = addPaneItem(oCustImpForm, cbUsePaymentName, edL, top, edW, edH, "Use Payment Name(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2030);
            //   ((SAPbouiCOM.EditText)(oItem.Specific)).Value = "Y";
            //    top = oItem.Top + nGap;

            oItem = addPaneItem(oCustImpForm, editCustomerID, edL, top, edW, edH, "Customer Id import to:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 204, 150);
            oItem.Visible = false;
            oCustImpForm.Items.Item("LB204").Visible = false;


            //   ((SAPbouiCOM.EditText)(oItem.Specific)).Value = "Y";
            //    top = oItem.Top + nGap;
        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
        oCustImpForm.Visible = true;

    }
    Thread tImpProcess = null;
    private void CustImpFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
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
                    case SAPbouiCOM.BoEventTypes.et_CLICK:

                        break;

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        {
                            switch (pVal.ItemUID)
                            {
                                case btnClose:
                                    try
                                    {
                                        if (tImpProcess != null)
                                            tImpProcess.Abort();
                                    }
                                    catch (Exception) { }
                                    oCustImpForm.Close();
                                    break;
                                case btnImp:
                                    try
                                    {
                                        if (custObjForImport == null)
                                        {
                                            form.Items.Item(btnFind).Click();
                                        }
                                        else
                                        {
                                            string cardcode = getFormItemVal(form, editCustomerID);
                                            if (cardcode == "")
                                            {
                                                showMessage("Enter customer ID.");
                                            }
                                            else
                                            {
                                                if (custObjForImport == null)
                                                    form.Items.Item(btnFind).Click();
                                                else
                                                {
                                                    string code = FindCustomerByID(cardcode);
                                                    if (code == "")
                                                    {
                                                        showMessage("Customer not found");
                                                        return;
                                                    }

                                                    if (SBO_Application.MessageBox("Import payment method(s) to " + cardcode + "?", 1, "Yes", "No") == 1)
                                                    {


                                                        foreach (PaymentMethodProfile pm in custObjForImport.PaymentMethodProfiles)
                                                        {
                                                            if (!isPaymentMethodExists(cardcode, pm.CardNumber, pm.CardExpiration))
                                                            {

                                                                CCCUST cccust = new CCCUST();
                                                                cccust.CustomerID = cardcode;
                                                                cccust.active = 'Y';
                                                                cccust.email = custObjForImport.Email;
                                                                cccust.firstName = custObjForImport.BillingAddress.FirstName;
                                                                cccust.lastName = custObjForImport.BillingAddress.LastName;
                                                                cccust.street = custObjForImport.BillingAddress.Address1;
                                                                cccust.city = custObjForImport.BillingAddress.City;
                                                                cccust.state = custObjForImport.BillingAddress.State;
                                                                cccust.zip = custObjForImport.BillingAddress.ZipCode;
                                                                cccust.expDate = pm.CardExpiration;
                                                                cccust.cardCode = pm.CardCode;
                                                                try
                                                                {
                                                                    if (pm.MethodName != null && pm.MethodName != "")
                                                                    {
                                                                        if (pm.MethodName.IndexOf(" ") >= 0)
                                                                        {
                                                                            string name = pm.MethodName;
                                                                            string[] s = name.Split(' ');

                                                                            cccust.firstName = s[0];
                                                                            cccust.lastName = s[s.Length - 1];
                                                                        }
                                                                    }
                                                                }
                                                                catch (Exception) { }
                                                                cccust.routingNumber = pm.Routing;
                                                                cccust.checkingAccount = pm.Account;
                                                                cccust.GroupName = getGroupName(cardcode);
                                                                cccust.CustNum = custObjForImport.CustomerToken;
                                                                cccust.MethodID = pm.MethodID;
                                                                int id = getNextTableID("@CCCUST");
                                                                string ccname = cccust.firstName + " " + cccust.lastName;
                                                                string desc = ccname;
                                                                if (pm.MethodName != "" && pm.MethodName != "")
                                                                    desc = pm.MethodName;
                                                                if (pm.MethodType == "check")
                                                                {
                                                                    cccust.methodDescription = id.ToString() + "_" + pm.Routing + " " + pm.Account + "(" + desc + ")";
                                                                    cccust.checkingAccount = pm.Account;
                                                                    cccust.routingNumber = pm.Routing;
                                                                }
                                                                else
                                                                {
                                                                    cccust.cardNum = pm.CardNumber;
                                                                    cccust.methodDescription = id.ToString() + "_" + pm.CardNumber + " " + pm.CardExpiration + "(" + desc + ")";
                                                                    cccust.cardType = pm.CardType;
                                                                }
                                                                string cardid = "";
                                                                cccust.CardName = getCardName(cccust.GroupName, cccust.cardType, "", ref cardid);
                                                                cccust.CCAccountID = cardid;
                                                                insert(cccust);
                                                                SBO_Application.SetStatusBarMessage(string.Format("Import payment method {0} to {1}.", cccust.cardNum, cardcode), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                                                            }

                                                        }
                                                    }
                                                }

                                            }
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        SBO_Application.SetStatusBarMessage(string.Format("Import payment method error: {0}.", ex.Message), SAPbouiCOM.BoMessageTime.bmt_Medium, true);


                                    }
                                    break;
                                case btnFImp:
                                    string fn = getFormItemVal(form, editFileName);
                                    if (fn == "")
                                    {
                                        showMessage("Please enter file name.");
                                    }
                                    else
                                    {
                                        if (!File.Exists(fn))
                                        {
                                            showMessage("File: " + fn + " not found.");
                                        }
                                        else
                                        {
                                            if (SBO_Application.MessageBox("Import payment method(s) from file " + fn + "?", 1, "Yes", "No") == 1)
                                            {

                                                SetText("Import Begin.  Please wait...\r\n");
                                                tImpProcess = new Thread(importThread);
                                                tImpProcess.Start();
                                            }
                                        }
                                    }
                                    break;
                                case btnExpFImp:
                                    fn = getFormItemVal(form, editExpFileName);
                                    if (fn == "")
                                    {
                                        showMessage("Please enter exported customer file name.");
                                    }
                                    else
                                    {
                                        if (!File.Exists(fn))
                                        {
                                            showMessage("File: " + fn + " not found.");
                                        }
                                        else
                                        {
                                            if (SBO_Application.MessageBox("Import payment method(s) from file " + fn + "?", 1, "Yes", "No") == 1)
                                            {

                                                SetText("Import Begin.  Please wait...\r\n");
                                                tImpProcess = new Thread(ExpImportThread);
                                                tImpProcess.Start();
                                            }
                                        }
                                    }
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
    private void importThread()
    {
        try
        {
            oCustImpForm.Items.Item(btnFImp).Enabled = false;
            oCustImpForm.Items.Item(btnFind).Enabled = false;
            oCustImpForm.Items.Item(btnImp).Enabled = false;

            List<ImportData> dataList = new List<ImportData>();
            string fn = getFormItemVal(oCustImpForm, editFileName);

            var fileLines = File.ReadAllLines(fn).Skip(1).ToList();
            foreach (string line in fileLines)
            {
                string[] i = line.Split('\t');
                ImportData data = new ImportData();
                try
                {
                    data.CustomerID = i[0];
                    data.CompanyName = i[1];
                    data.FirstName = i[2];
                    data.LastName = i[3];
                    data.Street = i[4];
                    data.Street2 = i[5];
                    data.Zip = i[6];
                    data.City = i[7];
                    data.State = i[8];
                    data.Phone = i[9];
                    data.Fax = i[10];
                    data.Email = i[11];
                    data.URL = i[12];
                    data.CardNumber = i[13];
                    data.CardExpiration = i[14];
                    data.MethodName = i[15];
                    if (i.Length >= 16)
                        data.defaultCard = i[16];
                }
                catch (Exception)
                { }
                if (data.CustomerID != null && data.CustomerID != "")
                    dataList.Add(data);
            }
            foreach (ImportData imp in dataList)
            {
                SAPCust sapcust = new SAPCust();
                sapcust.cccust = new CCCUST();
                sapcust.cccust.city = imp.City;
                sapcust.cccust.email = imp.Email;
                sapcust.cccust.street = imp.Street;
                sapcust.cccust.state = imp.State;
                sapcust.cccust.zip = imp.Zip;
                sapcust.cccust.firstName = imp.FirstName;
                sapcust.cccust.lastName = imp.LastName;
                sapcust.cccust.CustomerID = imp.CustomerID;
                sapcust.cccust.GroupName = getGroupName(imp.CustomerID);
                sapcust.cccust.cardType = getCardType(imp.CardNumber);
                sapcust.cccust.methodName = imp.MethodName;
                try
                {
                    if (imp.defaultCard != null && imp.defaultCard != "")
                        sapcust.cccust.@default = imp.defaultCard[0];
                }
                catch (Exception) { }
                string code = "";
                sapcust.cccust.CardName = getCardName(sapcust.cccust.GroupName, sapcust.cccust.cardType, "", ref code);
                sapcust.cccust.CCAccountID = code;
                sapcust.cccust.cardNum = imp.CardNumber;
                sapcust.cccust.expDate = imp.CardExpiration;
                CCCUST cust = sapcust.cccust;
                sapcust.custObj = new  Customer();
                sapcust.custObj.BillingAddress = new Address();
                sapcust.custObj.Email = cust.email;
                sapcust.custObj.BillingAddress.FirstName = cust.firstName;
                sapcust.custObj.BillingAddress.LastName = cust.lastName;
                sapcust.custObj.BillingAddress.State = cust.state;
                sapcust.custObj.BillingAddress.City = cust.city;
                sapcust.custObj.BillingAddress.ZipCode = cust.zip;
                sapcust.custObj.BillingAddress.Address1 = cust.street;
                sapcust.custObj.PaymentMethodProfiles = new PaymentMethodProfile[1]; 
                sapcust.custObj.PaymentMethodProfiles[0] = new PaymentMethodProfile(); 
                sapcust.custObj.PaymentMethodProfiles[0].CardNumber = cust.cardNum;
                sapcust.custObj.PaymentMethodProfiles[0].CardType = cust.cardType;
                sapcust.custObj.PaymentMethodProfiles[0].CardExpiration = cust.expDate;
                sapcust.custObj.PaymentMethodProfiles[0].CardCode = cust.cardCode;
                sapcust.custObj.PaymentMethodProfiles[0].MethodName = imp.MethodName;
                sapcust.custObj.PaymentMethodProfiles[0].MethodType = "cc";
                if (AddCustomer(ref sapcust))
                    SetText("\r\n" + string.Format("Customer {0} card {1} added", sapcust.cccust.CustomerID, sapcust.cccust.cardNum));
                else
                    SetText("\r\n" + string.Format("Failed to add Customer {0} card {1}", sapcust.cccust.CustomerID, sapcust.cccust.cardNum));

            }
            SetText("\r\nImport Completed.");
        }
        catch (Exception ex)
        {
            showMessage(ex.Message);
        }
        finally
        {
            try
            {
                oCustImpForm.Items.Item(btnFImp).Enabled = true;
                oCustImpForm.Items.Item(btnFind).Enabled = true;
                oCustImpForm.Items.Item(btnImp).Enabled = true;
            }
            catch (Exception) { }
        }
    }
    private void ExpImportThread()
    {
        try
        {
            oCustImpForm.Items.Item(btnExpFImp).Enabled = false;
            oCustImpForm.Items.Item(btnFImp).Enabled = false;
            oCustImpForm.Items.Item(btnFind).Enabled = false;
            oCustImpForm.Items.Item(btnImp).Enabled = false;

            List<ImportData> dataList = new List<ImportData>();
            string fn = getFormItemVal(oCustImpForm, editExpFileName);

            var fileLines = File.ReadAllLines(fn).Skip(1).ToList();
            foreach (string line in fileLines)
            {
                string[] i = line.Split('\t');
                ImportData data = new ImportData();
                try
                {
                    data.CustomerID = i[0];
                    data.CustNum = i[1];
                    /*
                    if(data.CustomerID != "")
                    {
                        data.CustomerID = "A1";
                        data.CustNum = "8545319";
                    }
                    */
                }
                catch (Exception)
                { }
                if (data.CustomerID != null && data.CustomerID != "")
                    dataList.Add(data);
            }
            foreach (ImportData imp in dataList)
            {
                SAPCust sapcust = new SAPCust();
                try
                {
                    sapcust.cccust = new CCCUST();
                    sapcust.cccust.CustomerID = imp.CustomerID;
                    sapcust.cccust.GroupName = getGroupName(imp.CustomerID);
                    string code = "";
                    sapcust.cccust.CardName = getCardName(sapcust.cccust.GroupName, sapcust.cccust.cardType, "", ref code);
                    sapcust.cccust.CCAccountID = code;
                    sapcust.cccust.CustNum = imp.CustNum;
                    SecurityToken token = getToken(sapcust.cccust.CCAccountID);
                    sapcust.custObj = ebiz.GetCustomer(token, imp.CustomerID, "");
                    sapcust.cccust.city = sapcust.custObj.BillingAddress.City;
                    sapcust.cccust.email = sapcust.custObj.Email;
                    sapcust.cccust.street = sapcust.custObj.BillingAddress.Address1;
                    sapcust.cccust.state = sapcust.custObj.BillingAddress.State;
                    sapcust.cccust.zip = sapcust.custObj.BillingAddress.ZipCode;
                    sapcust.cccust.firstName = sapcust.custObj.BillingAddress.FirstName;
                    sapcust.cccust.lastName = sapcust.custObj.BillingAddress.LastName;
                    sapcust.cccust.MethodID = sapcust.custObj.PaymentMethodProfiles[0].MethodID;
                    sapcust.cccust.methodName = sapcust.custObj.PaymentMethodProfiles[0].MethodName;
                    try
                    {
                        if (imp.defaultCard != null && imp.defaultCard != "")
                            sapcust.cccust.@default = imp.defaultCard[0];
                    }
                    catch (Exception) { }


                    CCCUST ccust = sapcust.cccust;
                    int id = getNextTableID("@CCCUST");
                    string desc = sapcust.custObj.BillingAddress.FirstName + " " + sapcust.custObj.BillingAddress.LastName;
                    if (sapcust.custObj.PaymentMethodProfiles[0].MethodName != "" && sapcust.custObj.PaymentMethodProfiles[0].MethodName != null)
                        desc = sapcust.custObj.PaymentMethodProfiles[0].MethodName;
                    if (sapcust.custObj.PaymentMethodProfiles[0].MethodType == "check")
                    {
                        sapcust.cccust.methodDescription = id.ToString() + "_" + sapcust.custObj.PaymentMethodProfiles[0].Routing + " " + sapcust.custObj.PaymentMethodProfiles[0].Account + "(" + desc + ")";
                        sapcust.cccust.checkingAccount = sapcust.custObj.PaymentMethodProfiles[0].Account;
                        sapcust.cccust.routingNumber = sapcust.custObj.PaymentMethodProfiles[0].Routing;
                    }
                    else
                    {
                        sapcust.cccust.expDate = DateTime.Parse(sapcust.custObj.PaymentMethodProfiles[0].CardExpiration).ToString("MMyy");
                        sapcust.cccust.cardNum = sapcust.custObj.PaymentMethodProfiles[0].CardNumber;
                        sapcust.cccust.methodDescription = id.ToString() + "_" + sapcust.custObj.PaymentMethodProfiles[0].CardNumber + " " + sapcust.custObj.PaymentMethodProfiles[0].CardExpiration + "(" + desc + ")";
                        sapcust.cccust.cardType = sapcust.custObj.PaymentMethodProfiles[0].CardType;
                        sapcust.cccust.CardName = getCardName(sapcust.cccust.GroupName, sapcust.cccust.cardType, "", ref code);
                        sapcust.cccust.CCAccountID = code;
                    }
                    insert(ccust);
                    SetText("\r\n" + string.Format("Customer {0} card {1} added", sapcust.cccust.CustomerID, sapcust.cccust.cardNum));
                }
                catch (Exception ex)
                {
                    SetText("\r\n" + string.Format("Failed to add customer {0} card {1}. {2}", sapcust.cccust.CustomerID, sapcust.cccust.cardNum, ex.Message));

                }
            }
            SetText("\r\nImport Completed.");
        }
        catch (Exception ex)
        {
            showMessage(ex.Message);
        }
        finally
        {
            try
            {
                oCustImpForm.Items.Item(btnExpFImp).Enabled = true;
                oCustImpForm.Items.Item(btnFind).Enabled = true;
                oCustImpForm.Items.Item(btnFImp).Enabled = true;
                oCustImpForm.Items.Item(btnFind).Enabled = true;
                oCustImpForm.Items.Item(btnImp).Enabled = true;
            }
            catch (Exception) { }
        }
    }
    private void SetText(string text)
    {
        string s = getFormItemVal(oCustImpForm, editImportLog);
        s = s + text;
        setFormEditVal(oCustImpForm, editImportLog, s);
    }

}


