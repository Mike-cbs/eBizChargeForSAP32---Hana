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
using eBizChargeForSAP.ServiceReference1;

partial class SAP
{
    const string menueBizConnect = "eBizeBizConnectMenu";
    const string formeBizConnect = "eBizeBizConnectForm";
    const string matrixeBizConnect = "mBConn";
    const string dteBizConnect = "dtConn";
    SAPbouiCOM.Form oeBizConnectForm;
    public MerchantTransactionData merchantData = new MerchantTransactionData();
    private void AddeBizConnectMenu()
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
            oCreationPackage.UniqueID = menueBizConnect;
            oCreationPackage.String = "eBizCharge Connect";
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
    private void CreateeBizConnectForm()
    {
        try
        {
            ManualScan = "0";
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Button oButton = null;
            //SAPbouiCOM.StaticText oStaticText = null;
            // SAPbouiCOM.EditText oEditText = null;
            //SAPbouiCOM.ComboBox oComboBox = null;

            // add a new form
            SAPbouiCOM.FormCreationParams oCreationParams = null;

            oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.FormType = formeBizConnect;

            oCreationParams.UniqueID = formeBizConnect;
            try
            {
                oeBizConnectForm = SBO_Application.Forms.AddEx(oCreationParams);
            }
            catch (Exception)
            {
                oeBizConnectForm = SBO_Application.Forms.Item(formeBizConnect);

            }

            // set the form properties
            oeBizConnectForm.Title = "eBizCharge Connect";
            oeBizConnectForm.Left = 400;
            oeBizConnectForm.Top = 100;
            oeBizConnectForm.ClientHeight = 460;
            oeBizConnectForm.ClientWidth = 750;





            //************************
            // Adding a Rectangle
            //***********************
            int margin = 5;
            oItem = oeBizConnectForm.Items.Add("Rect1", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = margin;
            oItem.Width = oeBizConnectForm.ClientWidth - 2 * margin;
            oItem.Top = margin;
            oItem.Height = oeBizConnectForm.ClientHeight - 40;

            int nTopGap = 25;
            int left = 6;
            int wBtn = 70;
            int hBtn = 19;
            int span = 80;


            oItem = oeBizConnectForm.Items.Add(btnRefresh, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oeBizConnectForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Refresh";
            left += span;

            oItem = oeBizConnectForm.Items.Add(btnAdd, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oeBizConnectForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Upload All";
            left += span;
            /*
            oItem = oeBizConnectForm.Items.Add(btnRun, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oeBizConnectForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Run";

            left += span;
            */
            oItem = oeBizConnectForm.Items.Add(btnClose, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oeBizConnectForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";

            margin = 8;
            int top = 15;
            int edL = 150; //oItm.Left + oItm.Width;
            int edW = 100;
            int edH = 15;
            int nGap = 26;

            oItem = addPaneItem(oeBizConnectForm, editCustomerID, edL, top, edW, edH, "Customer ID:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1000);
            oItem = addPaneItem(oeBizConnectForm, cbCustomerID, edL + 110, top, edW, edH, "", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 1000);

            oItem = oeBizConnectForm.Items.Add(btnFind, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = edL + 220;
            oItem.Width = wBtn;
            oItem.Top = top - 2;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Find";

            oItem = addPaneItem(oeBizConnectForm, cbGroup, edL + 400, top, edW, edH, "Customer Group", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 1030);

            SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)oItem.Specific;
            ComboAddItem(oCB, "");
            ComboAddItem(oCB, "All");
            List<string> list = GetGroupNames();
            foreach (string c in list)
            {
                ComboAddItem(oCB, c);
            }

            top = oItem.Top + nGap;


            oItem = oeBizConnectForm.Items.Add(matrixeBizConnect, SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oItem.FromPane = 0;
            oItem.ToPane = 0;
            oItem.Left = 5 + margin;
            oItem.Width = oeBizConnectForm.ClientWidth - 2 * margin - 10;
            oItem.Top = top;
            oItem.Height = oeBizConnectForm.ClientHeight - 100;
            top = oItem.Height + oItem.Top + 2;
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oeBizConnectForm.Items.Item(matrixeBizConnect).Specific;
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Add("DocNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Invoice No.";
            oColumn.Width = 80;

            oColumn = oMatrix.Columns.Add("Status", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Status";
            oColumn.Width = 80;

            oColumn = oMatrix.Columns.Add("CustID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Customer";
            oColumn.Width = 80;

            SAPbouiCOM.LinkedButton btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_BusinessPartner).ToString();

            oColumn = oMatrix.Columns.Add("InvID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Origin No.";
            btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_Invoice).ToString();
            oColumn.Width = 80;



            oColumn = oMatrix.Columns.Add("UploadDT", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Upload Date";
            oColumn.Width = 80;

            oColumn = oMatrix.Columns.Add("PaidDT", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Payment Date";
            oColumn.Width = 80;

            oColumn = oMatrix.Columns.Add("Balance", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Balance";
            oColumn.Width = 80;
            oColumn = oMatrix.Columns.Add("UpBal", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Balance Uploaded";
            oColumn.Width = 80;
            oColumn = oMatrix.Columns.Add("AmtPaid", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Last ePayment";
            oColumn.Width = 80;

        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
        oeBizConnectForm.Visible = true;

    }
    const string CONNECT_STATUS_UPLOADED = "Uploaded";
    const string CONNECT_STATUS_PAID = "Paid";
    const string CONNECT_STATUS_PARTIAL = "Partial";

    private bool HandleInvSelect()
    {
        bool bSelected = false;
        try
        {

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oeBizConnectForm.Items.Item(matrixeBizConnect).Specific;
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                if (oMatrix.IsRowSelected(i))
                {
                    bSelected = true;
                    string invnum = getMatrixItem(oMatrix, "DocNum", i);
                    string status = getMatrixItem(oMatrix, "Status", i);
                    string CustID = getMatrixItem(oMatrix, "CustID", i);
                    string InvID = getMatrixItem(oMatrix, "InvID", i);
                    string balance = getMatrixItem(oMatrix, "Balance", i);

                    if (status == "")
                    {
                        string str = string.Format("Upload invoice {0} to eBizConnect?", invnum);
                        if (SBO_Application.MessageBox(str + "?", 1, "Yes", "No") == 1)
                        {
                            if (IsInvoicePaid(CustID, invnum))
                            {
                                showMessage("Invoice paid.");
                                updateStatus();
                            }
                            else
                            {
                                addConnectInvoice(CustID, InvID, balance);
                                populateeBizConnectMatrix();
                            }
                        }
                    }
                    else if (status == CONNECT_STATUS_UPLOADED)
                    {
                        string str = string.Format("Reload invoice {0} to eBizConnect?", invnum);
                        if (SBO_Application.MessageBox(str + "?", 1, "Yes", "No") == 1)
                        {
                            if (IsInvoicePaid(CustID, invnum))
                            {
                                showMessage("Invoice paid.");
                                updateStatus();
                            }
                            else
                            {
                                reloadConnectInvoice(CustID, InvID, balance);
                                populateeBizConnectMatrix();
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return bSelected;
    }
    private void CheckStatusSelect()
    {
        try
        {

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oeBizConnectForm.Items.Item(matrixeBizConnect).Specific;
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                string status = getMatrixItem(oMatrix, "Status", i);
                if (status == "")
                    oMatrix.SelectRow(i, true, true);



            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void UploadAll()
    {
        try
        {
            updateStatus();
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oeBizConnectForm.Items.Item(matrixeBizConnect).Specific;
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                string invnum = getMatrixItem(oMatrix, "DocNum", i);
                string status = getMatrixItem(oMatrix, "Status", i);
                string CustID = getMatrixItem(oMatrix, "CustID", i);
                string InvID = getMatrixItem(oMatrix, "InvID", i);
                string balance = getMatrixItem(oMatrix, "Balance", i);

                if (status == "")
                {
                    addConnectInvoice(CustID, InvID, balance);
                }
            }
            populateeBizConnectMatrix();
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void eBizConnectFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            oeBizConnectForm = form;
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
                                case cbCustomerID:
                                    string id = getFormItemVal(form, cbCustomerID);
                                    if (id != "")
                                    {
                                        setFormEditVal(form, editCustomerID, id);
                                        getConnectCustomer(id);
                                        updateStatus();
                                    }
                                    break;
                                case cbGroup:
                                    string grp = getFormItemVal(form, cbGroup);
                                    if (grp != "")
                                    {
                                        setFormEditVal(form, editCustomerID, "");
                                        setComboValue(form, cbCustomerID, "");
                                        updateStatus();

                                    }
                                    break;
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == matrixeBizConnect)
                        {
                            if (!HandleInvSelect())
                            {
                                switch (pVal.ColUID)
                                {
                                    case "PaidDT":
                                        populateeBizConnectMatrix("\"U_PaidDate\" DESC");
                                        break;
                                    case "UploadDT":
                                        populateeBizConnectMatrix("\"U_UploadDate\" DESC");
                                        break;
                                    case "Balance":
                                        populateeBizConnectMatrix("(a.\"DocTotal\" - a.\"PaidToDate\") ");
                                        break;
                                    case "UpBal":
                                        populateeBizConnectMatrix("\"U_UploadedBalance\"");
                                        break;
                                    case "Status":
                                        populateeBizConnectMatrix("\"U_Status\" DESC");
                                        break;
                                    case "CustID":
                                        populateeBizConnectMatrix("a.\"CardCode\"");
                                        break;
                                    case "InvID":
                                        populateeBizConnectMatrix("a.\"DocEntry\"");
                                        break;
                                }

                            }


                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        {
                            switch (pVal.ItemUID)
                            {
                                case btnClose:
                                    oeBizConnectForm.Close();
                                    oeBizConnectForm = null;
                                    break;
                                case btnAdd:
                                    string str = string.Format("Upload all invoices to eBizConnect?");
                                    if (SBO_Application.MessageBox(str + "?", 1, "Yes", "No") == 1)
                                    {
                                        tProcess = new Thread(UploadAll);
                                        tProcess.Start();
                                    }
                                    break;
                                case btnRefresh:
                                    try
                                    {
                                        updateStatus();

                                    }
                                    catch (Exception ex)
                                    {
                                        SBO_Application.MessageBox(ex.Message);
                                    }
                                    break;
                                case btnRun:
                                    try
                                    {
                                        //ProcessJob();
                                        //showMessage("Job completed");
                                        populateeBizConnectMatrix();

                                    }
                                    catch (Exception ex)
                                    {
                                        SBO_Application.MessageBox(ex.Message);
                                    }
                                    break;
                                case btnFind:
                                    string id = getFormItemVal(oeBizConnectForm, editCustomerID);
                                    showStatus("Finding customer please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                    List<String> list = FindCustomer(id);
                                    if (list.Count() == 0)
                                    {
                                        showStatus("No customer found.", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                    }
                                    else
                                    {
                                        SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbCustomerID).Specific;
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
                                        showStatus("Found " + list.Count().ToString() + " customer(s).", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

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
    private SAPbouiCOM.DataTable geteBizConnectMatrixDataTable()
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = oeBizConnectForm.DataSources.DataTables.Item(dteBizConnect);
        }
        catch (Exception)
        {
            oDT = oeBizConnectForm.DataSources.DataTables.Add(dteBizConnect);

        }

        return oDT;
    }
    private void populateeBizConnectMatrix(string orderBy = "")
    {
        string sql = "";
        try
        {
            // formatConnectNum();
            if (oeBizConnectForm == null)
                return;
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oeBizConnectForm.Items.Item(matrixeBizConnect).Specific;
            oMatrix.Clear();
            string where = "";
            string id = getFormItemVal(oeBizConnectForm, editCustomerID);
            if (id != "")
            {
                where = string.Format(" and \"CardCode\"='{0}' ", id);
            }
            string grp = getFormItemVal(oeBizConnectForm, cbGroup);
            if (grp != "")
            {
                string s = getCardCodeByGroup(grp);
                if (s != "")
                    where = string.Format(" and \"CardCode\" in ({0}) ", s);

            }
            SAPbouiCOM.DataTable oDts = geteBizConnectMatrixDataTable();
            sql = "select (a.\"DocTotal\" - a.\"PaidToDate\") as Balance \r\n" +
" , a.\"DocEntry\", a.\"DocNum\", a.\"CardCode\",  \r\n" +
 " b.\"U_UploadedBalance\", b.\"U_PaidAmount\", b.\"U_Status\", b.\"U_UploadDate\", b.\"U_PaidDate\"  \r\n" +
 " from OINV a left outer join \"@CCCONNECTINVOICE\" b  \r\n" +
 " on a.\"DocNum\" = \"U_InvoiceID\" and \"CardCode\" = \"U_CustomerID\"  \r\n" +
 " where ((\"DocTotal\" - \"PaidToDate\") > 0 or b.\"U_InvoiceID\" <> '')  \r\n" +
 " and CANCELED = 'N'  \r\n" + where;

            if (orderBy == "")
                orderBy = " a.\"DocEntry\" desc";

            sql = sql + " order by " + orderBy;
            oDts.ExecuteQuery(sql);
            //trace(sql);
            BindMatrix(oMatrix, "DocNum", "DocNum", dteBizConnect);
            BindMatrix(oMatrix, "InvID", "DocEntry", dteBizConnect);
            BindMatrix(oMatrix, "CustID", "CardCode", dteBizConnect);
            BindMatrix(oMatrix, "Status", "U_Status", dteBizConnect);
            BindMatrix(oMatrix, "PaidDT", "U_PaidDate", dteBizConnect);
            BindMatrix(oMatrix, "UploadDT", "U_UploadDate", dteBizConnect);
            BindMatrix(oMatrix, "Balance", "Balance", dteBizConnect);
            BindMatrix(oMatrix, "AmtPaid", "U_PaidAmount", dteBizConnect);
            BindMatrix(oMatrix, "UpBal", "U_UploadedBalance", dteBizConnect);
            oMatrix.LoadFromDataSource();


        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private SecurityToken getToken(string CardCode = "")
    {
        SecurityToken securityToken = new SecurityToken();
        //    securityToken.UserId = "qbouser";
        //  securityToken.SecurityId = "99359f03-b254-4adf-b446-24957fcb46cb";//old
        securityToken.SecurityId = "22092678-5a07-48d0-9b1b-67cf4234fc83";//sap
        //securityToken.SecurityId = "e5c10fab-6bde-407a-ad3a-dd28e9a3f365";// Android
        //      securityToken.Password = "qbouser";
        return securityToken;

        SecurityToken t = new SecurityToken();
        try
        {
            string pin = "";
            string sourceKey = "";
            string merchantID = cfgMerchantID;
            if (CardCode != "")
            {
                getPinSKFromCardAcct(CardCode, ref pin, ref sourceKey);
                if (sourceKey != "" && sourceKey != null)
                {
                    merchantID = sourceKey;
                }
            }
            else
            {
                string branchid = getBranchID();
                if (branchid != "0")
                {
                    var q1 = from x in cardAcctList
                             where x.U_BranchID == branchid
                             select x;
                    foreach (CCCardAcct acct in q1)
                    {
                        trace(string.Format("Found Card using brank ID: branchid = {0}, grp={0}, sourceKey={1}", acct.U_BranchID, acct.U_SourceKey));
                        if (acct.U_SourceKey != "" && acct.U_SourceKey != null)
                            merchantID = acct.U_SourceKey;
                    }
                }
            }



            t.UserId = "";
            t.SecurityId = merchantID;
            t.Password = "";
            return t;
        }
        catch (Exception)
        {

        }
        return t;
    }
    public Customer getConnectCustomer(string custID)
    {
        Customer cust = null;
        try
        {
            string cguid = getCustomerGUID(custID);
            if (cguid == "")
                return addConnectCustomer(custID);

            SecurityToken securityToken = getToken();
            Customer customer = new Customer();

            cust = ebiz.GetCustomer(securityToken, custID, null);

            return cust;

        }
        catch (Exception ex)
        {
            //errorLog(ex);
        }
        return cust;
    }
    public Customer addConnectCustomer(string custid)
    {
        try
        {

            SecurityToken securityToken = getToken();
            Customer customer = new Customer();

            customer = new Customer();
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sql = string.Format("Select a.\"CardName\", b.\"Street\",b.\"City\",b.\"State\",b.\"ZipCode\",a.\"E_Mail\" \r\n" +
" from OCRD a left outer join CRD1 b on a.\"CardCode\" = b.\"CardCode\"  \r\n" +
" where a.\"BillToDef\" = b.\"Address\" and a.\"CardCode\"='{0}'", custid);

            oRS.DoQuery(sql);

            if (!oRS.EoF)
            {

                string f = GetFieldVal(oRS, 0);
                string[] l = f.Split(' ');

                customer.FirstName = l[0];
                if (l.Count() > 1)
                    customer.LastName = l[1];
                else
                    customer.LastName = l[0];
                customer.CompanyName = GetFieldVal(oRS, 0);
                customer.CustomerId = custid;
                customer.BillingAddress = new eBizChargeForSAP.ServiceReference1.Address();
                customer.BillingAddress.Address1 = GetFieldVal(oRS, 1);
                customer.BillingAddress.Address2 = "";
                customer.BillingAddress.City = GetFieldVal(oRS, 2);
                customer.BillingAddress.ZipCode = GetFieldVal(oRS, 3);
                customer.BillingAddress.State = GetFieldVal(oRS, 4);
                CustomerResponse CustomerResponse = ebiz.AddCustomer(securityToken, customer);
                if (CustomerResponse.ErrorCode == 0)
                {
                    addCustomerGUID(custid, CustomerResponse.CustomerInternalId);
                    customer.CustomerInternalId = CustomerResponse.CustomerInternalId;
                    return customer;
                }
                else
                {
                    if (CustomerResponse.ErrorCode == 2)
                    {
                        Customer c = ebiz.GetCustomer(securityToken, custid, null);
                        addCustomerGUID(custid, c.CustomerInternalId);
                        return customer;
                    }
                    else
                    {
                        showMessage(CustomerResponse.Error);
                    }
                    return null;
                }
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return null;
    }
    public void reloadConnectInvoice(string custid, string invid, string balance)
    {
        try
        {

            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            if (!oDoc.GetByKey(int.Parse(invid)))
            {
                errorLog("Inv not found: " + invid);
                return;
            }
            string invnum = oDoc.DocNum.ToString();
            SecurityToken securityToken = getToken();
            Customer customer = getConnectCustomer(custid);


            Invoice invoice = new Invoice();
            SAPbobsCOM.Document_Lines olines = oDoc.Lines;
            Item[] Lineitems = new Item[olines.Count];
            for (int i = 0; i < olines.Count; i++)
            {
                olines.SetCurrentLine(i);
                Item item = new Item();
                item.Name = olines.ItemCode; ;
                item.Qty = (decimal)olines.Quantity;
                item.UnitPrice = (decimal)olines.Price;
                item.ItemId = olines.ItemCode;
                item.Description = olines.ItemDescription;
                item.TotalLineAmount = (decimal)olines.LineTotal;
                item.TotalLineTax = (decimal)olines.TaxTotal;
                Lineitems[i] = item;

            }

            invoice.InvoiceNumber = invnum;
            invoice.InvoiceDate = oDoc.DocDate.ToString("MM/dd/yyyy");
            invoice.CustomerId = customer.CustomerId;
            invoice.InvoiceDueDate = oDoc.DocDueDate.ToString("MM/dd/yyyy");
            invoice.InvoiceAmount = (decimal)oDoc.DocTotal;
            invoice.AmountDue = decimal.Parse(balance);
            invoice.DivisionId = "X";
            invoice.Items = Lineitems;
            InvoiceResponse resp = ebiz.UpdateInvoice(securityToken, invoice, customer.CustomerId, "", invnum, "");
            if (resp.Error != "" && resp.ErrorCode != 2)
            {
                showMessage(resp.Error);
                return;
            }
            Invoice inv = ebiz.GetInvoice(securityToken, customer.CustomerId, "", invnum, "");
            List<CCConnectInvoice> q = findCCConnectInvoice(string.Format(" where \"U_InvoiceGUID\" = '{0}'", invid));

            foreach (CCConnectInvoice i in q)
            {
                i.UploadedBalance = Math.Round(decimal.Parse(balance), 2).ToString();
                i.UploadDate = DateTime.Today;
                update(i);
            }


            populateeBizConnectMatrix();
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    public void addConnectInvoice(string custid, string invid, string balance)
    {
        try
        {
            showStatus(string.Format("Upload invoice: {0} for customer: {1}. Balance:{2}. Please wait...", invid, custid, balance), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            if (!oDoc.GetByKey(int.Parse(invid)))
            {
                errorLog("Inv not found: " + invid);
                return;
            }
            string invnum = oDoc.DocNum.ToString();

            SecurityToken securityToken = getToken();
            //Customer customer = getConnectCustomer(custid);

            Invoice invoice = new Invoice();
            invoice.UploadedBy = oCompany.UserName;
            SAPbobsCOM.Document_Lines olines = oDoc.Lines;
            Item[] Lineitems = new Item[olines.Count];
            for (int i = 0; i < olines.Count; i++)
            {
                olines.SetCurrentLine(i);
                Item item = new Item();
                item.Name = olines.ItemCode; ;
                item.Qty = (decimal)olines.Quantity;
                item.UnitPrice = (decimal)olines.Price;
                item.ItemId = olines.ItemCode;
                item.Description = olines.ItemDescription;
                item.TotalLineAmount = (decimal)olines.LineTotal;
                item.TotalLineTax = (decimal)olines.TaxTotal;
                Lineitems[i] = item;

            }

            invoice.InvoiceNumber = invnum;
            invoice.InvoiceDate = oDoc.DocDate.ToString("MM/dd/yyyy");
            invoice.CustomerId = oDoc.CardCode;
            invoice.InvoiceDueDate = oDoc.DocDueDate.ToString("MM/dd/yyyy");
            invoice.InvoiceAmount = (decimal)oDoc.DocTotal;
            invoice.AmountDue = decimal.Parse(balance);
            invoice.DivisionId = "X";
            invoice.Items = Lineitems;
            InvoiceResponse resp = ebiz.AddInvoice(securityToken, invoice);
            if (resp.Error != "" && resp.ErrorCode != 2)
            {
                showMessage(resp.Error);
                return;
            }
            CCConnectInvoice ccinv = new CCConnectInvoice();
            ccinv.CustomerID = oDoc.CardCode;
            ccinv.InvoiceID = invnum;
            ccinv.InvoiceGUID = invid;
            ccinv.UploadedBalance = Math.Round(decimal.Parse(balance), 2).ToString();
            ccinv.Status = "Uploaded";
            ccinv.UploadDate = DateTime.Today;
            insert(ccinv);
            showStatus(string.Format("Invoice: {0} for customer: {1} added. Balance:{2}.", invid, custid, balance), SAPbouiCOM.BoMessageTime.bmt_Medium, false);


        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void updateStatus()
    {
        try
        {
            string cid = getFormItemVal(oeBizConnectForm, editCustomerID);
            if (cid != "")
            {
                checkPaidStatus(cid);
            }
            else
            {
                string grp = getFormItemVal(oeBizConnectForm, cbGroup);
                if (grp != "")
                {

                    string s = getCardCodeByGroup(grp);
                    string[] c = s.Split(',');
                    foreach (string i in c)
                    {
                        checkPaidStatus(i.Replace("'", "").Trim());
                    }
                    populateeBizConnectMatrix();
                }
            }
            populateeBizConnectMatrix();


        }
        catch (Exception)
        {

        }
    }
    public bool IsInvoicePaid(string custid, string invid)
    {
        try
        {
            if (!IsConnectCustomerExist(custid))
                return false;
            showStatus("Checking payment status for customer: " + custid + ". Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);


            SecurityToken securityToken = getToken();

            Payment[] pmts = ebiz.GetPayments(securityToken, custid, "", DateTime.Today.AddDays(-30), DateTime.Today.AddDays(1), 0, 200, "");
            foreach (Payment p in pmts)
            {
                if (p.InvoiceNumber == invid)
                    return true;
            }
        }
        catch (Exception)
        {

        }
        return false;
    }

    public string getCustomerGUID(string customerID)
    {


        string sql = string.Format("select \"U_CustomerGUID\" from \"@CCCONNECTCUSTOMER\" where \"U_CustomerID\"='{0}'", customerID);
        try
        {
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                return (string)oRS.Fields.Item(0).Value;
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return "";
    }
    public void addCustomerGUID(string customerID, string cguid)
    {
        int id = getNextTableID("@CCCONNECTCUSTOMER");
        string sql = string.Format("insert into \"@CCCONNECTCUSTOMER\"(\"DocEntry\", \"U_CustomerID\", \"U_CustomerGUID\") values({2},'{0}','{1}')", customerID, cguid, id);
        try
        {
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }

    public void checkPaidStatus(string custid)
    {
        try
        {
            if (!IsConnectCustomerExist(custid))
                return;
            showStatus("Checking payment stauts for customer: " + custid + ". Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);


            SecurityToken securityToken = getToken();

            Payment[] pmts = ebiz.GetPayments(securityToken, custid, "", DateTime.Today.AddDays(-30), DateTime.Today.AddDays(1), 0, 200, "");
            foreach (Payment p in pmts)
            {

                var q = findCCConnectInvoice(string.Format(" where \"U_InvoiceID\" = '{0}'", p.InvoiceNumber));
                foreach (CCConnectInvoice i in q)
                {
                    try
                    {
                        int id = getDocEntry("13", p.InvoiceNumber, i.CustomerID);
                        SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                        if (!oDoc.GetByKey(id))
                        {
                            errorLog("Inv not found: " + id);
                            i.Status = "Invoice not found: " + i.InvoiceID;
                            update(i);
                            continue;
                        }
                        if (oDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Close)
                        {
                            i.PaidAmount = Math.Round(decimal.Parse(p.PaidAmount), 2).ToString();
                            i.PaidDate = DateTime.Parse(p.DatePaid);
                            i.Status = "Paid";
                            update(i);
                        }
                        else
                        {
                            string where = string.Format(" where \"U_CustomerID\"='{0}' and \"U_InvoiceID\"='{1}' and \"U_PaymentID\"='{2}'", p.CustomerId, p.InvoiceNumber, p.AuthCode);

                            List<CCCPayment> qcc = findCCPayment(where);
                            if (qcc.Count() == 0)
                            {
                                try
                                {
                                    Customer custObj = ebiz.GetCustomer(securityToken, i.CustomerID, "");
                                    foreach (PaymentMethodProfile pm in custObj.PaymentMethodProfiles)
                                    {
                                        if (!isPaymentMethodExists(custObj.CustomerId, pm.MethodID))
                                        {
                                            CCCUST cccust = new CCCUST();
                                            int idTB = getNextTableID("@CCCUST");
                                            cccust.CustomerID = custObj.CustomerId;
                                            cccust.CustNum = custObj.CustomerToken;
                                            cccust.MethodID = pm.MethodID;
                                            cccust.cardCode = "";
                                            cccust.expDate = pm.CardExpiration;
                                            cccust.cardNum = pm.CardNumber;
                                            cccust.cardType = pm.CardType;
                                            cccust.routingNumber = pm.Routing;
                                            cccust.checkingAccount = pm.Account;
                                            cccust.GroupName = getGroupName(cccust.CustomerID);
                                            string desc = custObj.BillingAddress.FirstName + " " + custObj.BillingAddress.LastName;
                                            if (pm.MethodName != "" && pm.MethodName != null)
                                                desc = pm.MethodName;
                                            if (pm.MethodType == "check")
                                            {
                                                cccust.methodDescription = idTB.ToString() + "_" + pm.Routing + " " + pm.Account + "(" + desc + ")";
                                            }
                                            else
                                            {
                                                cccust.methodDescription = idTB.ToString() + "_" + pm.CardNumber + " " + pm.CardExpiration + "(" + desc + ")";
                                            }
                                            if (!IsWalkIn(cccust.CustomerID))
                                            {
                                                insert(cccust);
                                            }
                                        }

                                    }
                                }
                                catch (Exception ex)
                                {
                                    errorLog(ex);
                                }
                                SAPbobsCOM.Payments oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                                int bplid = getBranchIdFromINV(id.ToString());
                                if (bplid != 0)
                                    oPmt.BPLID = bplid;
                                oPmt.CardCode = i.CustomerID;
                                oPmt.DocDate = DateTime.Now;
                                oPmt.Remarks = "eBizCharge Connect Payment";

                                oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
                                oPmt.JournalRemarks = "eBizCharge Connect Payment";
                                oPmt.Invoices.DocEntry = int.Parse(i.InvoiceGUID);
                                oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice;
                                oPmt.Invoices.SumApplied = getDoubleValue(p.PaidAmount);
                                SAPCust sapcust = getCustomerByID(i.CustomerID, "");
                                AddCreditCardPayment(oPmt, sapcust, getDoubleValue(p.PaidAmount), true);
                                int r = oPmt.Add();
                                if (r != 0)
                                {
                                    errorLog("eBizConnect Failed Payment. ID:" + i.CustomerID + "\r\n" +
                                        "InvoiceID:" + i.InvoiceID + "\r\n" +
                                        "CCAcountID:" + sapcust.cccust.CCAccountID + "\r\n" +
                                        getErrorString());

                                }
                                else
                                {

                                    CCCPayment pmt = new CCCPayment();
                                    pmt.InvoiceID = p.InvoiceNumber;
                                    pmt.CustomerID = p.CustomerId;
                                    pmt.PaymentID = p.AuthCode;
                                    pmt.Amount = p.PaidAmount;
                                    pmt.DateImported = DateTime.Now;
                                    insert(pmt);

                                }
                            }
                            i.PaidAmount = Math.Round(decimal.Parse(p.PaidAmount), 2).ToString();
                            i.PaidDate = DateTime.Parse(p.DatePaid);
                            if (getDoubleValue(p.PaidAmount) < getDoubleValue(p.InvoiceAmount))
                                i.Status = "Partial";
                            else
                                i.Status = "Paid";
                            update(i);
                        }
                    }
                    catch (Exception ex)
                    {
                        errorLog(ex);
                    }

                    showStatus("Status check for customer: " + custid + " completed.", SAPbouiCOM.BoMessageTime.bmt_Medium, false);


                }
            }

            var qUploaded = findCCConnectInvoice(string.Format("  where \"U_Status\" = 'Uploaded' and \"U_CustomerID\" = '{0}' ", custid));
            foreach (CCConnectInvoice i in qUploaded)
            {
                try
                {
                    Invoice inv = ebiz.GetInvoice(securityToken, i.CustomerID, "", i.InvoiceID, "");
                    if (inv == null)
                    {
                        delete(i);
                    }
                }
                catch (Exception)
                {
                    delete(i);
                }
            }


        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    public void setDefaultTaxRate()
    {
        try
        {
            if (cfgMerchantID == "" || cfgMerchantID == null)
                return;

            SecurityToken securityToken = getToken();

            merchantData = ebiz.GetMerchantTransactionData(securityToken);
            trace(string.Format("merchantData for {2}\r\ntax rate: '{0}'\r\nCommodity Code: '{1}'", merchantData.ItemTaxRate, merchantData.ItemCommodityCode, cfgMerchantID));
            cfgDefaultTaxRate = (merchantData.ItemTaxRate).ToString();
            validDefaultTax();
        }
        catch (Exception ex)
        {
            errorLog("GetMerchantTransactionData error: " + ex.Message);
        }

    }
    public bool SourceKeyCheck()
    {
        string s = "";
        try
        {
            s = ebiz.AuthenticateSK(cfgSourceKey);
            if (s == "0")
            {
                showMessage("Source key: '" + cfgSourceKey + "' not found in the integration server. ");
                return false;
            }
        }
        catch (Exception ex)
        {
            merchantData = new MerchantTransactionData();
            errorLog(ex);
        }
        return true;
    }

    public void eConnectSyncInvoice()
    {
        int interval = 600000;
        try
        {
            interval = int.Parse(cfgAutoSync.Split(':')[0]) * 60000;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        trace("eConnectSyncProc started. interval: " + interval + "; cfgAutoSync=" + cfgAutoSync);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            if (cfgMerchantID == "" || cfgMerchantID == null)
            {
                trace("cfgMerchantID is null eConnectSyncProc thread exited.");
                return;
            }
            //FROM B1 to Connect
            //   string sql = "select CustomerID, InvoiceGUID, UploadedBalance from CCConnectInvoice where Status != 'Paid'";
            String sql = "select [U_CustomerID[, [U_DocEntry[, cast([U_Balance[ as nvarchar(50)),[U_RefNum[ as [RefNum[ from [@CCINVLOG[ where [U_PaymentStatus[ != 'Paid' ".Replace("[", "\"");
            while (true)
            {
                Thread.Sleep(1000);
                bConnectSync = true;

                #region SAP B1 to eConnect
                try
                {
                    oRS.DoQuery(sql);
                    while (!oRS.EoF)
                    {
                        string CustomerID = (string)oRS.Fields.Item(0).Value;
                        string InvoiceID = (string)oRS.Fields.Item(1).Value;
                        string RefNum = (string)oRS.Fields.Item("RefNum").Value;
                        string bal = (string)oRS.Fields.Item(2).Value;
                        double b = getBalance(InvoiceID);

                        if (b != getDoubleValue(bal))
                        {
                            HandleeConnectInvoice(InvoiceID, "", CustomerID, bal, RefNum);
                        }
                        oRS.MoveNext();
                    }
                }
                catch (Exception ex)
                {
                    errorLog(ex);
                }
                #endregion

                #region Connect to SAP B1
                List<String> SecurityIDList = new List<string>();
                SecurityIDList.Add(getToken().SecurityId);
                foreach (string sid in SecurityIDList)
                {
                    try
                    {
                        #region Connect to B1
                        SecurityToken securityToken = new SecurityToken();
                        securityToken.SecurityId = sid;
                        List<Customer> custList = new List<Customer>();
                        #region Collect Bp with Payments
                        int count = getSQLInt("select Count(1) from OCRD where \"CardType\"='C'");
                        trace("Total Customer Count:" + count);
                        var filters = new SearchFilter[1];
                        filters[0] = new SearchFilter();
                        filters[0].FieldName = "SoftwareID";
                        filters[0].ComparisonOperator = "eq";
                        filters[0].FieldValue = Software;

                        CustomerListSearchResult result = ebiz.SearchCustomerList(securityToken, filters, 0, count, "", false, false, false);
                        foreach (Customer c in result.CustomerList)
                        {

                            if (c.CustomerId != "" && getSQLString(string.Format("select \"CardCode\" from OCRD where \"CardCode\"='{0}' and \"validFor\"='Y'", c.CustomerId)) != "")
                            {
                                Payment[] pmts = ebiz.GetPayments(securityToken, c.CustomerId, "", DateTime.Today.AddDays(-90), DateTime.Today.AddDays(1), 0, 100, "");
                                var count1 = pmts.Count();
                                if (pmts.Count() > 0)
                                    custList.Add(c);
                                else
                                {
                                    pmts = ebiz.SearchEbizWebFormReceivedPayments(securityToken, c.CustomerId, DateTime.Today.AddDays(-90), DateTime.Today.AddDays(1), null, 0, 100, "");
                                    if (pmts.Count() > 0)
                                        custList.Add(c);
                                }
                            }

                        }
                        if (custList.Count() == 0)
                            break;
                        trace("Sync from Connect to B1.  Security ID: " + sid + " total paid customer: " + custList.Count());
                        #endregion
                        //custList = custList.Where(a => a.CustomerId == "C000011").ToList();
                        foreach (Customer c2 in custList)
                        {
                            #region Pay Paid
                            try
                            {
                                Customer c1 = B1BPToConnectCustomer(c2.CustomerId);
                                CustomerResponse CustResp = ebiz.UpdateCustomer(securityToken, c1, c2.CustomerId, "");
                            }
                            catch (Exception ex)
                            {
                                errorLog("eConnectSyncProc update customer exception: " + ex.Message);
                            }
                            PostInvoicePayment(c2.CustomerId, sid);
                            // checkePayPaidStatus(c2.CustomerId, sid);
                            #endregion
                        }

                        foreach (Customer c2 in custList)
                        {
                            #region PaymentMethodProfile
                            /*
                            try
                            {
                                Customer custObj = ebiz.GetCustomer(securityToken, c2.CustomerId, "");
                                foreach (PaymentMethodProfile pm in custObj.PaymentMethodProfiles)
                                {
                                    if (!isPaymentMethodExists(custObj.CustomerId, pm.MethodID))
                                    {
                                        #region CCCUST
                                        CCCUST cccust = new CCCUST();
                                        int id = getNextCCCUSTID();
                                        cccust.CustomerID = custObj.CustomerId;
                                        cccust.CustNum = custObj.CustomerToken;
                                        cccust.MethodID = pm.MethodID;
                                        cccust.cardCode = "";
                                        cccust.expDate = pm.CardExpiration;
                                        cccust.cardNum = pm.CardNumber;
                                        cccust.cardType = pm.CardType;
                                        cccust.routingNumber = pm.Routing;
                                        cccust.checkingAccount = pm.Account;
                                        cccust.GroupName = getGroupName(cccust.CustomerID);
                                        string currency = getSQLString(string.Format("Select Currency from OCRD where CardCode = '{0}'", cccust.CustomerID));
                                        string branchID = getBranchIdByCustomer(cccust.CustomerID).ToString();
                                        string cardCode = "";
                                        string desc = custObj.BillingAddress.FirstName + " " + custObj.BillingAddress.LastName;
                                        string acctname = getAccountName(cccust.CustomerID);
                                        string cardName = getCardName(cccust.GroupName, cccust.cardType, currency, ref cardCode, acctname, branchID);
                                        cccust.CCAccountID = cardCode;
                                        cccust.CardName = cardName;
                                        if (pm.MethodName != "" && pm.MethodName != null)
                                            desc = pm.MethodName;
                                        if (pm.MethodType == "check")
                                        {
                                            cccust.methodDescription = id.ToString() + "_" + pm.Routing + " " + pm.Account + "(" + desc + ")";
                                        }
                                        else
                                        {
                                            cccust.methodDescription = id.ToString() + "_" + pm.CardNumber + " " + pm.CardExpiration + "(" + desc + ")";
                                        }
                                        if (!IsWalkIn(cccust.CustomerID))
                                        {
                                            insert(ref cccust);
                                            trace("Payment method added: " + cccust.methodDescription);
                                        }
                                        #endregion
                                    }

                                }
                            }
                            catch (Exception ex)
                            {
                                errorLog(ex);
                            }
                            */
                            #endregion
                        }
                        #endregion 
                    }
                    catch (Exception ex)
                    {
                        errorLog(ex);
                    }
                }
                //execute("Update CCCUST set active = 'Y' where active = ''");
                //Paid in B1. Sync up

                //UpdateConnectInvThreadProc();
                //setUDF();
                bConnectSync = false;
                #endregion

            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;

        }

    }
    public void eConnectSyncSalesOrder()
    {
        int interval = 600000;
        bool ValidFlag = false;
        String ValidMsg = "";
        int DocEntry = 0;
        string ArDownPaymentDocNum = "";
        try
        {
            interval = int.Parse(cfgAutoSync.Split(':')[0]) * 60000;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        trace("eConnectSync SO started. interval: " + interval + "; cfgAutoSync=" + cfgAutoSync);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            if (cfgMerchantID == "" || cfgMerchantID == null)
            {
                trace("cfgMerchantID is null eConnectSyncProc thread exited.");
                return;
            }
            //FROM B1 to Connect
            String sql = "select [U_CustomerID[, [U_DocEntry[, cast([U_Balance[ as nvarchar(50)),[U_RefNum[ as [RefNum[ from [@CCSOLOG[ where [U_PaymentStatus[ != 'Paid' and [U_DocState[ ='Add' and \"U_Balance\">0 ".Replace("[", "\"");
            while (true)
            {
                //Thread.Sleep(interval);
                Thread.Sleep(1000);
                bConnectSync = true;
                #region SAP B1 to eConnect
                try
                {
                    oRS.DoQuery(sql);
                    while (!oRS.EoF)
                    {
                        string CustomerID = (string)oRS.Fields.Item(0).Value;
                        string OrderID = (string)oRS.Fields.Item(1).Value;
                        string RefNum = (string)oRS.Fields.Item("RefNum").Value;
                        string RemainingBal = (string)oRS.Fields.Item(2).Value;
                        double PaidAmt = getBalance(OrderID, "ORDR");
                        double RemainingbalDouble = getDoubleValue(RemainingBal);
                        if (PaidAmt > 0)
                        {
                           // if (RemainingbalDouble != PaidAmt)
                            {
                                    HandleeConnectSales(OrderID, "", CustomerID, RemainingBal, RefNum);
                            }
                        }
                        oRS.MoveNext();
                    }
                    
                }
                catch (Exception ex)
                {
                    errorLog(ex);
                }
                #endregion
                
                //FROM Connect to B1
                List<String> SecurityIDList = new List<string>();
                SecurityIDList.Add(getToken().SecurityId);
                foreach (string sid in SecurityIDList)
                {
                    try
                    {
                        #region Connect to B1
                        SecurityToken securityToken = new SecurityToken();
                        securityToken.SecurityId = sid;
                        List<Customer> custList = new List<Customer>();
                        #region Collect Bp with Payments
                        int count = getSQLInt("select Count(1) from OCRD where \"CardType\"='C'");
                        trace("Total Customer Count:" + count);
                        var filters = new SearchFilter[1];
                        filters[0] = new SearchFilter();
                        filters[0].FieldName = "SoftwareID";
                        filters[0].ComparisonOperator = "eq";
                        filters[0].FieldValue = Software;

                        CustomerListSearchResult result = ebiz.SearchCustomerList(securityToken, filters, 0, count, "", false, false, false);
                        foreach (Customer c in result.CustomerList)
                        {

                            if (c.CustomerId != "" && getSQLString(string.Format("select \"CardCode\" from OCRD where \"CardCode\"='{0}' and \"validFor\"='Y'", c.CustomerId)) != "")
                            {
                                filters = null;
                                filters = new SearchFilter[4];
                                filters[0] = new SearchFilter();
                                filters[0].FieldName = "DateTime";
                                filters[0].ComparisonOperator = "gt";
                                filters[0].FieldValue = "2016-12-01 00:00:00";
                                filters[1] = new SearchFilter();
                                filters[1].FieldName = "CustomerID";
                                filters[1].ComparisonOperator = "eq";
                                filters[1].FieldValue = c.CustomerId;
                                filters[2] = new SearchFilter();
                                filters[2].FieldName = "IsTransactionApplied";
                                filters[2].ComparisonOperator = "eq";
                                filters[2].FieldValue ="false";
                                filters[3] = new SearchFilter();
                                filters[3].FieldName = "SoftwareId";
                                filters[3].ComparisonOperator = "eq";
                                filters[3].FieldValue = "SBO";
                                ApplicationTransactionSearchResult TranList=ebiz.SearchApplicationTransactions(securityToken, filters, false, 0, 100, "1");
                                var count1 = TranList.ApplicationTransactions.Count();
                                for (int i = 0; i < count1; i++)
                                {
                                    ApplicationTransactionDetails trans=TranList.ApplicationTransactions[i];
                                    String SalesOrderDocNum = trans.LinkedToTypeId;
                                    String AmountPaid = trans.TransactionAmount;
                                    #region Sales order payment
                                    PostSOPayment(trans);
                                    #endregion


                                }


                                //Payment[] pmts = ebiz.GetPayments(securityToken, c.CustomerId, "", DateTime.Today.AddDays(-90), DateTime.Today.AddDays(1), 0, 100, "");
                                //var count1 = pmts.Count();
                                //if (pmts.Count() > 0)
                                //    custList.Add(c);
                                //else
                                //{
                                //    pmts = ebiz.SearchEbizWebFormReceivedPayments(securityToken, c.CustomerId, DateTime.Today.AddDays(-90), DateTime.Today.AddDays(1), null, 0, 100, "");
                                //    if (pmts.Count() > 0)
                                //        custList.Add(c);
                                //}
                            }

                        }
                        if (custList.Count() == 0)
                            break;
                        trace("Sync from Connect to B1.  Security ID: " + sid + " total paid customer: " + custList.Count());
                        #endregion
                        //custList = custList.Where(a => a.CustomerId == "C0000148").ToList();
                       
                        #endregion 
                    }
                    catch (Exception ex)
                    {
                        errorLog(ex);
                    }
                }
                bConnectSync = false;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
    }

    public void setUDF()
    {
        if (cfgReceiptEmailUDF == null || cfgReceiptEmailUDF == "")
            return;
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            oRS.DoQuery(string.Format("select CardCode, {0} from OCRD where ISNULL({0}, '') <> ''", cfgReceiptEmailUDF));
            while (!oRS.EoF)
            {
                try
                {
                    string em = (string)oRS.Fields.Item(1).Value;
                    string custid = (string)oRS.Fields.Item(0).Value;
                    SecurityToken securityToken = getTokenEx(custid, "", "", "");
                    Customer custObj = ebiz.GetCustomer(securityToken, custid, "");
                    if (custObj.Email != em)
                    {
                        custObj.Email = em;
                        ebiz.UpdateCustomer(securityToken, custObj, custObj.CustomerId, custObj.CustomerInternalId);
                        trace("update connect customer: " + custid + " email to " + em);
                    }


                }
                catch (Exception)
                {

                }
                oRS.MoveNext();
            }

        }
        catch (Exception)
        { }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }

    }
    public void UpdateConnectInvThreadProc()
    {
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {

            string sql = "select distinct i.DocEntry,i.DocNum  as InvoiceNum, i.PaidToDate , c.Status from OINV i (NOLOCK), CCConnectInvoice c (NOLOCK) where  i.DocNum = c.InvoiceID and c.Status <> 'Paid' and i.PaidToDate > 0";
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                trace(string.Format("UpdateConnectInvThreadProc {0}, {1}, {2}", oRS.Fields.Item(0).Value, oRS.Fields.Item(1).Value, oRS.Fields.Item(2).Value));
                int DocEntry = (int)oRS.Fields.Item(0).Value;
                int invNum = (int)oRS.Fields.Item(1).Value;
                decimal amt = (decimal)oRS.Fields.Item(2).Value;
                HandleeConnectInvoice(DocEntry.ToString(), amt.ToString());
                oRS.MoveNext();
            }

        }
        catch (Exception ex)
        {
            errorLog("UpdateConnectInvThreadProc exception: " + ex.Message);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
    }

    bool bConnectSync = false;
    public Customer B1BPToConnectCustomer(string cardcode, SAPbouiCOM.Form form = null)
    {
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        string sql = "";
        Customer customer = new Customer();
        try
        {
            if (form != null)
            {
                customer.BillingAddress = new eBizChargeForSAP.ServiceReference1.Address();
                customer.CompanyName = getFormItemVal(form, "7");
                customer.Email = getFormItemVal(form, editHolderEmail);
                customer.FirstName = getFirstName(getFormItemVal(form, editHolderName));
                customer.LastName = getLastName(getFormItemVal(form, editHolderName));

                customer.BillingAddress.Address1 = getFormItemVal(form, editHolderAddr);
                customer.BillingAddress.City = getFormItemVal(form, editHolderCity);
                customer.BillingAddress.State = getFormItemVal(form, editHolderState);
                customer.BillingAddress.ZipCode = getFormItemVal(form, editHolderZip);
                return customer;
            }
            //Select a.CardName, b.Street,b.City,b.State,b.ZipCode,a.E_Mail, (select top 1 E_MailL from OCPR (nolock) where CardCode=a.CardCode and E_MailL <> '') as CEMail from ocrd (NOLOCK) a, CRD1 (NOLOCK) b where a.CardCode = b.CardCode and a.BillToDef = b.Address and a.cardcode='{0}'
            //sql = string.Format("Select a.CardName, b.Street,b.City,b.State,b.ZipCode,a.E_Mail from ocrd (NOLOCK) a, CRD1 (NOLOCK) b where a.CardCode = b.CardCode and a.BillToDef = b.Address and a.cardcode='{0}'", cardcode);
            sql = "SELECT T0.[CardName[, T1.[Street[,T1.[City[,T1.[State[,T1.[ZipCode[,T0.[E_Mail[ FROM OCRD T0  INNER JOIN CRD1 T1 ON T0.[CardCode[ = T1.[CardCode[ WHERE T1.[CardCode[ ='1'and T0.[BillToDef[=T1.[Address[".Replace("[", "\"");
            oRS.DoQuery(sql);

            if (!oRS.EoF)
            {
                trace(string.Format("Get customer info from db. ID={6}, {0}, {1}, {2}, {3}, {4}, {5}", GetFieldVal(oRS, 0), GetFieldVal(oRS, 1), GetFieldVal(oRS, 2), GetFieldVal(oRS, 3), GetFieldVal(oRS, 4), GetFieldVal(oRS, 5), cardcode));
                string f = getSQLString(string.Format("select \"Name\" from OCPR where \"CardCode\"='{0}'", cardcode)); ;
                string[] l = f.Split(' ');

                customer.FirstName = l[0];
                if (l.Count() > 1)
                    customer.LastName = l[1];
                else
                    customer.LastName = l[0];

                customer.CompanyName = GetFieldVal(oRS, 0);
                customer.CustomerId = cardcode;
                customer.BillingAddress = new eBizChargeForSAP.ServiceReference1.Address();
                customer.BillingAddress.Address1 = GetFieldVal(oRS, 1);
                customer.BillingAddress.Address2 = "";
                customer.BillingAddress.City = GetFieldVal(oRS, 2);
                customer.BillingAddress.ZipCode = GetFieldVal(oRS, 4);
                customer.BillingAddress.State = GetFieldVal(oRS, 3);
                string email = getSQLString(string.Format("select b.\"E_MailL\" from OCRD a, OCPR b where a.\"CntctPrsn\" = b.\"Name\" and a.\"CardCode\" = b.\"CardCode\" and b.\"CardCode\" = '{0}'", cardcode));
                if (email == "")
                    email = GetFieldVal(oRS, 5);
                customer.Email = email;
                try
                {
                    if (cfgDefaultContactID != "")
                    {
                        email = getSQLString(string.Format(" select \"E_MailL\" from OCPR  where \"CardCode\" ='{0}' and \"Name\" = '{1}'", cardcode, cfgDefaultContactID));
                        if (email != "")
                            customer.Email = email;
                    }
                }
                catch (Exception) { }
                try
                {
                    if (cfgReceiptEmailUDF != null && cfgReceiptEmailUDF != "")
                    {
                        SAPbobsCOM.BusinessPartners oBP = (SAPbobsCOM.BusinessPartners)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                        if (oBP.GetByKey(cardcode))
                        {
                            if (oBP.UserFields.Fields.Item(cfgReceiptEmailUDF) != null)
                            {
                                string em = oBP.UserFields.Fields.Item(cfgReceiptEmailUDF).Value.ToString();
                                if (em != "")
                                {
                                    customer.Email = em;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {

                }
            }
        }
        catch (Exception ex)
        {
            errorLog(sql);
            errorLog(ex);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
        return customer;
    }
    public void PostInvoicePayment(string custid, string securityID = "")
    {
        String InvoiceNum = "";
        SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
        try
        {
            #region OINV
            if (custid == "" || custid == null)
                return;

            string invType = "Invoice";
            if (oeBizConnectForm != null)
            {
                invType = getFormItemVal(oeBizConnectForm, cbInvType);
                showStatus("Checking " + invType + " payment status for customer: " + custid + ". Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
            else if (cfgConnectUseDP == "Y")
                invType = "Down Payment";
            if (invType == "Down Payment")
            {
                oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments);

            }
            string invSql = "OINV";
            if (invType == "Down Payment")
                invSql = " ODPI";
            #endregion

            SecurityToken securityToken = getTokenEx(custid, "", "", "");
            if (securityID != "")
                securityToken.SecurityId = securityID;
            //CheckConnectUploadProc(securityToken, custid);
            Payment[] pmts = ebiz.GetPayments(securityToken, custid, "", DateTime.Today.AddDays(-90), DateTime.Today.AddDays(1), 0, 1000, "");
            

            trace(string.Format("ebiz.GetPayments found {0} for custid {1}", pmts.Count(), custid));
            var refNumQ = pmts.Select(p => p.RefNum).Distinct();
            foreach (string refnum in refNumQ)
            {
                var q = from x in pmts
                        where x.RefNum == refnum
                        select x;
                SAPbobsCOM.Payments oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                double total = 0;
                int id = 0;
                string invDesc = "";
                string cardtype = "";
                string AuthCode = "";
                foreach (Payment p in q)
                {
                    AuthCode = p.AuthCode;
                    string sql = string.Format("select \"DocEntry\" from {2} where \"DocNum\"='{0}' and \"CardCode\" = '{1}'", p.InvoiceNumber, custid, invSql);
                    id = getSQLInt(sql);
                    int DocEntryInvoice = getSQLInt(string.Format("select \"DocEntry\" from {0} where \"CardCode\"='{1}' and \"DocNum\" = {2}", invSql, p.CustomerId, p.InvoiceNumber));
                    InvoiceNum = p.InvoiceNumber;
                    #region validation
                    if (id == 0)
                    {
                        trace("OINV DocEntry not found.\r\n" + sql);
                        continue;
                    }
                    sql = string.Format("select 1 from RCT3 a, ORCT b, RCT2 c where a.\"DocNum\" = b.\"DocEntry\" and c.\"DocNum\" = b.\"DocEntry\" and a.\"VoucherNum\" = '{0}' and b.\"Canceled\" = 'N' and c.\"DocEntry\" = {1}", refnum, id);
                    if (!IsRSEOF(sql))
                    {
                        double PaidAmt = ConvertDouble(p.PaidAmount);
                        double PaidToDate = new SAP().getBalance(id.ToString(),"OINV");//getSQLDouble(String.Format("SELECT \"PaidToDate\"  FROM OINV T0 WHERE T0.\"DocEntry\"='{0}'", id));
                        double DocTotal = getSQLDouble(String.Format("SELECT \"DocTotal\"  FROM OINV T0 WHERE T0.\"DocEntry\"='{0}'", id));
                        double AmtDue = DocTotal - PaidToDate;
                        double AmountDue = DocTotal - (PaidAmt+ PaidToDate);
                        if (false)
                        {
                        }
                        else
                        {
                            trace("Payment exists.  refNum: " + refnum + ";  query:" + sql);
                            continue;
                        }
                    }
                    cardtype = p.PaymentMethod;
                    trace(string.Format("pmt = Inv:{0}, CustID:{1}, RefNum:{2}, PAmount:{3}, Count:{4}", p.InvoiceNumber, p.CustomerId, p.RefNum, p.PaidAmount, q.Count()));

                    if (!oDoc.GetByKey(id))
                    {
                        trace("oDoc GetKey is false. id=" + id);
                        continue;
                    }
                    #endregion
                    invDesc = invDesc + string.Format("inv:{0},{2} amt:{1}  ", p.InvoiceNumber, p.PaidAmount, id);
                    cardtype = p.PaymentMethod;
                    var CCInvoiceLogSql = String.Format("select count(*) from [@CCINVLOG[ where [U_DocNum[='{0}'".Replace("[", "\""), p.InvoiceNumber);
                    int countCCInvoiceLog = getSQLInt(CCInvoiceLogSql);
                    if (countCCInvoiceLog == 0)
                    {
                        //Add Invoice in CCInvLog 
                        #region Add Invoice Log
                        LogTable oExportSOLogTable = new LogTable();
                        oExportSOLogTable.DocNum = p.InvoiceNumber;
                        oExportSOLogTable.DocType = "OINV";
                        oExportSOLogTable.PaymentStatus = "";
                        oExportSOLogTable.RefNum = "";
                        oExportSOLogTable.Balance = Convert.ToDecimal(p.InvoiceAmount.ToString());
                        oExportSOLogTable.DocEntry = DocEntryInvoice.ToString();
                        oExportSOLogTable.CardCode = p.CustomerId;
                        oExportSOLogTable.CreateDt = DateTime.Now;
                        oExportSOLogTable.Status = "Uploaded";
                        AddInvoiceLog(oExportSOLogTable);
                        #endregion
                    }
                    #region surcharge
                    /*
                if (p.TypeId.ToLower() == "surcharge")
                {

                    try
                    {
                        trace("Connect receive surcharge customer ID. " + p.CustomerId);
                        var qcc = from x in db.CCCPayments
                                  where x.CustomerID == p.CustomerId && x.PaymentID == p.AuthCode && p.TypeId == x.TypeId
                                  select x;
                        if (qcc.Count() == 0)
                        {
                            if (AddPaymentOnAccount(p.AuthCode, getDoubleValue(p.PaidAmount), custid))
                            {
                                showStatus("Receive surcharge of " + p.PaidAmount + ". Create payment on account: " + custid, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                trace("Connect payment on account customer ID: " + p.CustomerId + "; amount:" + p.PaidAmount);
                            }

                            CCCPayment pmt = new CCCPayment();
                            pmt.InvoiceID = p.InvoiceNumber;
                            pmt.CustomerID = p.CustomerId;
                            pmt.PaymentID = p.AuthCode;
                            pmt.Amount = p.PaidAmount;
                            pmt.TypeId = p.TypeId;
                            pmt.DateImported = DateTime.Now;
                            db.CCCPayments.InsertOnSubmit(pmt);
                            db.SubmitChanges();

                        }
                    }
                    catch (Exception ex)
                    {

                    }
                    continue;
                   
                }
                 */
                    #endregion

                    if (oDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Close)
                    {
                        #region Close Invoice Log
                        String sqlUpdateBalance1 = String.Format("update [@CCINVLOG[ set [U_Balance[='0',[U_PaidAmt[='{1}',[U_UpdateDt[=current_timestamp , [U_PaymentStatus[='Paid'where [U_DocEntry[='{0}'".Replace("[", "\""), DocEntryInvoice, Convert.ToDouble(p.InvoiceAmount));
                        execute(sqlUpdateBalance1);
                        trace("Connect Inv : " + DocEntryInvoice + " is closed" + "; Customer: " + p.CustomerId);
                        continue;
                        #endregion
                    }
                    trace("eConnect checkPaidStatus: " + custid + ", create payment for " + invType + ": " + p.InvoiceNumber);
                    #region OINV
                    int bplid = getBranchIdFromINV(p.InvoiceNumber);
                    if (invType == "Down Payment")
                        bplid = getBranchIdFromODPI(p.InvoiceNumber);
                    if (invType == "Credit Memo")
                        bplid = getBranchIdFromCM(p.InvoiceNumber);
                    if (bplid != 0)
                        oPmt.BPLID = bplid;
                    oPmt.CardCode = custid;
                    oPmt.DocDate = DateTime.Now;
                    //  oPmt.Remarks = "eBizCharge Connect Payment";
                    oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
                    //oPmt.JournalRemarks = "eBizCharge Connect Payment";
                    oPmt.Invoices.DocEntry = DocEntryInvoice;
                    oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice;
                    if (invType == "Down Payment")
                        oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_DownPayment;
                    if (invType == "Credit Memo")
                        oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote;
                    #endregion
                    double sumapplied = getDoubleValue(p.PaidAmount);
                    double balance = getBalance(oDoc.DocEntry.ToString());
                    if (sumapplied > balance)
                        sumapplied = balance;

                    oPmt.Invoices.SumApplied = sumapplied;

                    trace("Invoice sumapplied: " + oPmt.Invoices.SumApplied + " invoice : " + p.InvoiceNumber + "; balance=" + balance + "; p.PaidAmount = " + getDoubleValue(p.PaidAmount));

                    oPmt.Invoices.Add();
                    String Code = getSQLInt("select IFNULL(max(cast(\"DocNum\" as int) ),0) + 1 from \"@CCCPAYMENT\"").ToString();
                    String CCPaymentsql = String.Format("insert into [@CCCPAYMENT[ ([U_InvoiceID[, [U_CustomerID[, [U_PaymentID[,[U_Amount[, [U_DateImported[,[U_Type[,[DocEntry[,[DocNum[) values ('{0}','{1}','{2}','{3}',{4},'{5}','{6}','{7}')".Replace("[", "\""), p.InvoiceNumber, p.CustomerId, p.AuthCode, Convert.ToDecimal(p.PaidAmount), "current_timestamp", p.TypeId, Code, Code);
                    execute(CCPaymentsql);
                    //Update invoice
                    String Status = "";
                    if (getDoubleValue(p.PaidAmount) < getDoubleValue(p.InvoiceAmount))
                    {
                        Status = "Partial";
                    }
                    else
                    {
                        Status = "Paid";
                    }
                    var sqlUpdateBalance = String.Format("update [@CCINVLOG[ set [U_Balance[='{3}',[U_PaidAmt[='{1}',[U_UpdateDt[=current_timestamp , [U_PaymentStatus[='{2}'where [U_DocEntry[='{0}'".Replace("[", "\""), DocEntryInvoice, Convert.ToDouble(p.PaidAmount), Status, Convert.ToDouble(p.AmountDue));
                    execute(sqlUpdateBalance);

                    total = total + sumapplied;

                }
                if (total > 0)
                {
                    SAPCust sapcust = getCustomerByID(custid, "");
                    sapcust.cccust.CCAccountID = geteConnectCardCode(cardtype);
                    AddCreditCardPayment(oPmt, sapcust, total, true, refnum);

                    int r = oPmt.Add();
                    if (r != 0)
                    {
                        var errorMsg = "eBizConnect Failed Payment. ID:" + custid + "\r\n" +
                            invDesc + "\r\n" +
                            "CCAcountID:" + sapcust.cccust.CCAccountID + "\r\n" +
                            oCompany.GetLastErrorDescription();
                        errorLog(errorMsg);
                        execute(string.Format("delete from \"@CCCPAYMENT\" where \"U_PaymentID\"='{0}'", AuthCode));
                    }
                    else
                    {
                        if (invDesc != "")
                        {
                            trace(invDesc + " payment added");
                            
                        }
                        String msg = String.Format("Invoice:{0} Payment Added From Connect to SAPB1", InvoiceNum);
                        showStatus(msg, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    }
                }

            }
            #region Comment
            /*
            var qUploaded = from x in db.CCConnectInvoices
                            where x.Status == "Uploaded" && x.CustomerID == custid
                            select x;
            foreach (CCConnectInvoice i in qUploaded)
            {
                try
                {
                    Invoice inv = ebiz.GetInvoice(securityToken, i.CustomerID, "", i.InvoiceID, "");
                    if (inv == null)
                    {
                        db.CCConnectInvoices.DeleteOnSubmit(i);
                    }
                    else
                    {
                        HandleeConnectInvoice(i.InvoiceID, "", i.CustomerID, i.UploadedBalance);
                    }

                }
                catch (Exception)
                {
                    db.CCConnectInvoices.DeleteOnSubmit(i);
                }
            }
            db.SubmitChanges();
            */
            #endregion
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
            oDoc = null;
        }
        //showStatus("Status check for customer: " + custid + " completed.", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

    }
}

