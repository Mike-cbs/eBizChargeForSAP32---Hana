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
    const string menuePaymentForm = "eBizePaymentFormMenu";
    const string formePaymentForm = "eBizePaymentFormForm";
    const string matrixePaymentForm = "mEPFrm";
    const string dtePaymentForm = "dtEPFrm";
    const string cbEMTemplate = "cbFRMTMP";
    SAPbouiCOM.Form oePaymentFormForm;
    eBizChargeForSAP.ServiceReference1.EmailTemplate[] eTempList = null;

    private void AddePaymentFormMenu()
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
            oCreationPackage.UniqueID = menuePaymentForm;
            oCreationPackage.String = "eBizCharge ePayment Form";
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
    private void CreateePaymentFormForm()
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
            oCreationParams.FormType = formePaymentForm;

            oCreationParams.UniqueID = formePaymentForm;
            try
            {
                oePaymentFormForm = SBO_Application.Forms.AddEx(oCreationParams);
            }
            catch (Exception)
            {
                oePaymentFormForm = SBO_Application.Forms.Item(formePaymentForm);

            }

            // set the form properties
            oePaymentFormForm.Title = "eBizCharge ePayment Form";
            oePaymentFormForm.Left = 400;
            oePaymentFormForm.Top = 100;
            oePaymentFormForm.ClientHeight = 460;
            oePaymentFormForm.ClientWidth = 900;





            //************************
            // Adding a Rectangle
            //***********************
            int margin = 5;
            oItem = oePaymentFormForm.Items.Add("Rect1", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = margin;
            oItem.Width = oePaymentFormForm.ClientWidth - 2 * margin;
            oItem.Top = margin;
            oItem.Height = oePaymentFormForm.ClientHeight - 40;

            int nTopGap = 25;
            int left = 6;
            int wBtn = 70;
            int hBtn = 19;
            int span = 80;


            oItem = oePaymentFormForm.Items.Add(btnRefresh, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oePaymentFormForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Refresh";
            left += span;

            oItem = oePaymentFormForm.Items.Add(btnAdd, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oePaymentFormForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Send All";
            left += span;
            /*
            oItem = oePaymentFormForm.Items.Add(btnRun, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oePaymentFormForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Run";

            left += span;
            */
            oItem = oePaymentFormForm.Items.Add(btnClose, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oePaymentFormForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";

            margin = 8;
            int top = 15;
            int edL = 100; //oItm.Left + oItm.Width;
            int edW = 100;
            int edH = 15;
            int nGap = 26;

            oItem = addPaneItem(oePaymentFormForm, editCustomerID, edL, top, edW, edH, "Customer ID:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1000,80);
            oItem = addPaneItem(oePaymentFormForm, cbCustomerID, edL + 110, top, edW, edH, "", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 1000);

            oItem = oePaymentFormForm.Items.Add(btnFind, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = edL + 220;
            oItem.Width = wBtn;
            oItem.Top = top - 2;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Find";

            oItem = addPaneItem(oePaymentFormForm, cbGroup, edL + 400, top, edW, edH, "Customer Group", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 1030);

            SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)oItem.Specific;
            ComboAddItem(oCB, "");
            ComboAddItem(oCB, "All");
            List<string> list = GetGroupNames();
            foreach (string c in list)
            {
                ComboAddItem(oCB, c);
            }
            oItem = addPaneItem(oePaymentFormForm, cbEMTemplate, edL + 640, top, edW, edH, "Email Template", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 1060);
            oCB = (SAPbouiCOM.ComboBox)oItem.Specific;
            foreach (string s in ListEmailTemplate)
            {
                ComboAddItem(oCB, s);
            }
            string t = getDefaultEmailTemplate();
            if (t != "")
            {
                try
                {
                    oCB.Select(t);
                }catch(Exception)
                {
                    try
                    {
                        oCB.Select(0);
                    }
                    catch (Exception) { };
                }
            }
            else
            {
                try
                {
                    oCB.Select(0);
                }
                catch (Exception) { };
            }
            top = oItem.Top + nGap;


            oItem = oePaymentFormForm.Items.Add(matrixePaymentForm, SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oItem.FromPane = 0;
            oItem.ToPane = 0;
            oItem.Left = 5 + margin;
            oItem.Width = oePaymentFormForm.ClientWidth - 2 * margin - 10;
            oItem.Top = top;
            oItem.Height = oePaymentFormForm.ClientHeight - 100;
            top = oItem.Height + oItem.Top + 2;
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oePaymentFormForm.Items.Item(matrixePaymentForm).Specific;
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

            oColumn = oMatrix.Columns.Add("CName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Customer Name";
            oColumn.Width = 80;
            oColumn = oMatrix.Columns.Add("E_Mail", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "EMail";
            oColumn.Width = 80;
         
            oColumn = oMatrix.Columns.Add("InvID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Origin No.";
            btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_Invoice).ToString();
            oColumn.Width = 80;



            oColumn = oMatrix.Columns.Add("UploadDT", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Sent Date";
            oColumn.Width = 80;

            oColumn = oMatrix.Columns.Add("PaidDT", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Payment Date";
            oColumn.Width = 80;

            oColumn = oMatrix.Columns.Add("Balance", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Balance";
            oColumn.Width = 80;
            oColumn = oMatrix.Columns.Add("UpBal", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Balance Sented";
            oColumn.Width = 80;
            oColumn = oMatrix.Columns.Add("AmtPaid", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Last ePayment";
            oColumn.Width = 80;
            // populateePaymentFormMatrix();
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;


        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
        oePaymentFormForm.Visible = true;

    }
    const string EPAY_STATUS_SENT = "Sent";
    const string EPAY_STATUS_PAID = "Paid";
    const string EPAY_STATUS_PARTIAL = "Partial";

    private bool ePayHandleInvSelect()
    {
        bool bSelected = false;
        try
        {

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oePaymentFormForm.Items.Item(matrixePaymentForm).Specific;
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                if (oMatrix.IsRowSelected(i))
                {
                    bSelected = true;
                    string status = getMatrixItem(oMatrix, "Status", i);
                    string invnum = getMatrixItem(oMatrix, "DocNum", i);
          string name = getMatrixItem(oMatrix, "CName", i);
                    if (status == "")
                    {
                        string str = string.Format("Send invoice {0} to {1}?", invnum, name);
                        if (SBO_Application.MessageBox(str + "?", 1, "Yes", "No") == 1)
                        {
                            addePayInvoice(oMatrix, i);
                        }
                    }
                    else if (status == EPAY_STATUS_SENT)
                    {
                        string str = string.Format("Re-send invoice {0} to {1}?", invnum, name);
                        if (SBO_Application.MessageBox(str + "?", 1, "Yes", "No") == 1)
                        {
                            addePayInvoice(oMatrix, i);
                        }
                    }
                    populateePaymentFormMatrix();
                }
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return bSelected;
    }
    
    private void ePaymentFormFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
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
                                case cbEMTemplate:
                                    string t = getFormItemVal(form, cbEMTemplate);
                                    updateDefaultEmailTemplate(t);
                                    break;
                                case cbCustomerID:
                                    string id = getFormItemVal(form, cbCustomerID);
                                    if (id != "")
                                    {
                                        setFormEditVal(form, editCustomerID, id);
                                        checkePayPaidStatus(id);
                                        setComboValue(form, cbGroup, "");
                                        populateePaymentFormMatrix();
                                    }
                                    break;
                                case cbGroup:
                                    string grp = getFormItemVal(form, cbGroup);
                                    if (grp != "")
                                    {
                                        setFormEditVal(form, editCustomerID, "");
                                        setComboValue(form, cbCustomerID, "");
                                        string s = getCardCodeByGroup(grp);
                                        string[] c = s.Split(',');
                                        foreach (string i in c)
                                        {
                                            checkePayPaidStatus(i.Replace("'", "").Trim());
                                        }
                                        populateePaymentFormMatrix();
                                    }
                                    break;
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == matrixePaymentForm)
                        {
                            if (!ePayHandleInvSelect())
                            {
                                switch (pVal.ColUID)
                                {
                                    case "PaidDT":
                                        populateePaymentFormMatrix(" \"U_PaidDate\" DESC,");
                                        break;
                                    case "UploadDT":
                                        populateePaymentFormMatrix(" \"U_UploadDate\" DESC,");
                                        break;
                                    case "Balance":
                                        populateePaymentFormMatrix("balance,");
                                        break;
                                    case "UpBal":
                                        populateePaymentFormMatrix(" \"U_UploadedBalance\",");
                                        break;
                                    case "Status":
                                        populateePaymentFormMatrix(" \"U_Status\" DESC,");
                                        break;
                                    case "CustID":
                                    case "CName":
                                        populateePaymentFormMatrix(" a.\"CardCode\" DESC,");
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
                                    oePaymentFormForm.Close();
                                    oePaymentFormForm = null;
                                    break;
                                case btnAdd:
                                    string str = string.Format("Send all invoices?");
                                    if (SBO_Application.MessageBox(str + "?", 1, "Yes", "No") == 1)
                                    {
                                        tProcess = new Thread(SendAll);
                                        tProcess.Start();
                                    }
                                    break;
                                case btnRefresh:
                                    try
                                    {
                                        string cid = getFormItemVal(oePaymentFormForm, editCustomerID);
                                        if (cid != "")
                                        {
                                            checkePayPaidStatus(cid);
                                        }else
                                        {
                                            string grp = getFormItemVal(form, cbGroup);
                                            if (grp != "")
                                            {
                                                string s = getCardCodeByGroup(grp);
                                                string[] c = s.Split(',');
                                                foreach (string i in c)
                                                {
                                                    checkePayPaidStatus(i.Replace("'", "").Trim());
                                                }
                                            }
                                        }
                                        populateePaymentFormMatrix();

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
                                        populateePaymentFormMatrix();

                                    }
                                    catch (Exception ex)
                                    {
                                        SBO_Application.MessageBox(ex.Message);
                                    }
                                    break;
                                case btnFind:
                                    string id = getFormItemVal(oePaymentFormForm, editCustomerID);
                                    SBO_Application.SetStatusBarMessage("Finding customer please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                    List<String> list = FindCustomer(id);
                                    if (list.Count() == 0)
                                    {
                                        SBO_Application.SetStatusBarMessage("No customer found.", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
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
                                        ComboAddItem(oCB, "");
                                        foreach (string c in list)
                                        {
                                            ComboAddItem(oCB, c);
                                        }
                                        SBO_Application.SetStatusBarMessage("Found " + list.Count().ToString() + " customer(s).", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

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
    private SAPbouiCOM.DataTable getePaymentFormMatrixDataTable()
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = oePaymentFormForm.DataSources.DataTables.Item(dtePaymentForm);
        }
        catch (Exception)
        {
            oDT = oePaymentFormForm.DataSources.DataTables.Add(dtePaymentForm);

        }

        return oDT;
    }
    private void populateePaymentFormMatrix(string orderBy = "")
    {
        string sql = "";
        try
        {
             if (oePaymentFormForm == null)
                return;
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oePaymentFormForm.Items.Item(matrixePaymentForm).Specific;
            oMatrix.Clear();
            //oMatrix.AutoResizeColumns();
            string where = "";
            string id = getFormItemVal(oePaymentFormForm, editCustomerID);
            if (id != "")
            {
                where = string.Format(" and c.\"CardCode\"='{0}' ", id);
            }
            string grp = getFormItemVal(oePaymentFormForm, cbGroup);
            if (grp != "")
            {
                string s = getCardCodeByGroup(grp);
                if (s != "")
                    where = string.Format(" and c.\"CardCode\" in ({0}) ", s);
                else
                    where = " and 1=0";

            }
            SAPbouiCOM.DataTable oDts = getePaymentFormMatrixDataTable();
            sql = string.Format("select (a.\"DocTotal\" - a.\"PaidToDate\") as Balance " +
" , a.\"DocEntry\", a.\"DocNum\", a.\"CardCode\",  b.\"U_UploadedBalance\" " +
" , b.\"U_PaidAmount\", b.\"U_Status\", b.\"U_UploadDate\", b.\"U_PaidDate\" " +
" , c.\"CardName\", c.\"E_Mail\"  " +
" from OINV a INNER JOIN OCRD c on a.\"CardCode\" = c.\"CardCode\" " + 
" left outer join \"@CCPAYFORMINVOICE\" b  on a.\"DocNum\" = b.\"U_InvoiceID\" and c.\"CardCode\" = b.\"U_CustomerID\" " +
" where (((a.\"DocTotal\" - a.\"PaidToDate\") > 0 and a.CANCELED = 'N' and a.\"DocStatus\" <> 'C') or b.\"U_InvoiceID\" is not null)   " +
" {0} ", where);


            sql = sql + " order by " + orderBy + " a.\"DocEntry\" desc";
          
            oDts.ExecuteQuery(sql);
            BindMatrix(oMatrix, "DocNum", "DocNum", dtePaymentForm);
            BindMatrix(oMatrix, "InvID", "DocEntry", dtePaymentForm);
            BindMatrix(oMatrix, "CustID", "CardCode", dtePaymentForm);
            BindMatrix(oMatrix, "Status", "U_Status", dtePaymentForm);
            BindMatrix(oMatrix, "PaidDT", "U_PaidDate", dtePaymentForm);
            BindMatrix(oMatrix, "UploadDT", "U_UploadDate", dtePaymentForm);
            BindMatrix(oMatrix, "Balance", "Balance", dtePaymentForm);
            BindMatrix(oMatrix, "AmtPaid", "U_PaidAmount", dtePaymentForm);
            BindMatrix(oMatrix, "UpBal", "U_UploadedBalance", dtePaymentForm);
            BindMatrix(oMatrix, "CName", "CardName", dtePaymentForm);
            BindMatrix(oMatrix, "E_Mail", "E_Mail", dtePaymentForm, true);
            oMatrix.LoadFromDataSource();


        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    
  
    public void addePayInvoice(SAPbouiCOM.Matrix oMatrix, int i)
    {
        try
        {
            string invnum = getMatrixItem(oMatrix, "DocNum", i);

            string custid = getMatrixItem(oMatrix, "CustID", i);
            string invid = getMatrixItem(oMatrix, "InvID", i);
            string balance = getMatrixItem(oMatrix, "Balance", i);
            string name = getMatrixItem(oMatrix, "CName", i);
            string email = getMatrixItem(oMatrix, "E_Mail", i);
            SBO_Application.SetStatusBarMessage(string.Format("Sending invoice: {0} to customer: {1}. Balance:{2}. Please wait...", invid, name, balance), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            if (!oDoc.GetByKey(int.Parse(invid)))
            {
                errorLog("Inv not found: " + invid);
                return;
            }


            SecurityToken securityToken = getToken();
          
            string detail = "<table>";
            SAPbobsCOM.Document_Lines olines = oDoc.Lines;
            for (int n = 0; n < olines.Count; n++)
            {
                olines.SetCurrentLine(n);
                detail += string.Format("<tr><td nowrap>{0}<td><td>&nbsp;&nbsp;{1}&nbsp;X&nbsp;{2}</td></tr>", olines.ItemDescription, olines.Quantity, olines.Price);
            }
            detail += "</table>";
            EbizWebForm pf = new EbizWebForm();
            pf.Clerk = oCompany.UserName;
            pf.AmountDue= Decimal.Parse(balance);
            pf.TotalAmount = Decimal.Parse(balance);
            pf.CustFullName = name;
            pf.EmailAddress = email;
           
            pf.EmailNotes = "SAP email payment form";
            pf.EmailSubject = cfgEmailSubject.Replace("[CUSTOMER]", name).Replace("[INVOICE]", invnum);
            pf.FromEmail = cfgmerchantEmail;
            pf.CustomerId = custid;
            pf.InvoiceNumber = invnum;
            pf.Description = detail;
            pf.EmailTemplateName = getFormItemVal(oePaymentFormForm, cbEMTemplate);
            pf.SendEmailToCustomer = true;
            string url = ebiz.GetEbizWebFormURL(securityToken, pf);
            List<CCPayFormInvoice> q = findCCPayFormInvoice(string.Format(" where \"U_InvoiceGUID\" = '{0}' and \"U_CustomerID\" = '{1}' ", invid,custid));
            if (q.Count() == 0)
            {
                CCPayFormInvoice ccinv = new CCPayFormInvoice();
                ccinv.CustomerID = custid;
                ccinv.InvoiceID = invnum;
                ccinv.InvoiceGUID = invid;
                ccinv.UploadedBalance = balance;
                ccinv.Status = EPAY_STATUS_SENT;
                ccinv.UploadDate = DateTime.Today;
                insert(ccinv);
            }
            SBO_Application.SetStatusBarMessage(string.Format("Invoice: {0} sent to customer: {1}. Balance:{2}.", invid, custid, balance), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
   

    public void checkePayPaidStatus(string custid)
    {
        try
        {
            if (custid == "" || custid == null)
                return;
            if (!IsPayFormCustomerExist(custid))
                return;
            SBO_Application.SetStatusBarMessage("Checking payment status for customer: " + custid + ". Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);


            SecurityToken securityToken = getToken();
           
            Payment[] pmts = ebiz.SearchEbizWebFormReceivedPayments(securityToken, custid, DateTime.Today.AddDays(-90), DateTime.Today.AddDays(1),null, 0, 1000, "");
            foreach (Payment p in pmts)
            {
                List<CCPayFormInvoice> q = findCCPayFormInvoice(string.Format(" where \"U_InvoiceID\" = '{0}' and \"U_CustomerID\"='{1}'", p.InvoiceNumber, p.CustomerId));
                foreach (CCPayFormInvoice i in q)
                {

                    try
                    {
                        SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                        if (!oDoc.GetByKey(int.Parse(i.InvoiceGUID)))
                        {
                            errorLog("Inv not found: " + i.InvoiceGUID);
                            i.Status = "Invoice not found: " + i.InvoiceGUID;

                            update(i);
                            continue;
                        }
                        if (oDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Close)
                        {
                            i.PaidAmount = p.PaidAmount;
                            i.PaidDate = DateTime.Parse(p.DatePaid);
                            i.Status = "Paid";
                            update(i);
                        }
                        else
                        {
                            string where = string.Format(" where \"U_CustomerID\"='{0}' and \"U_InvoiceID\"='{1}' and \"U_PaymentID\"='{2}'", p.CustomerId, p.InvoiceNumber, p.AuthCode);

                            var qcc = findCCPayFormPayment(where);
                            if (qcc.Count() == 0)
                            {
                                SAPbobsCOM.Payments oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                                int bplid = getBranchIdFromINV(i.InvoiceGUID); 
                                if (bplid != 0)
                                    oPmt.BPLID = bplid;
                                oPmt.CardCode = i.CustomerID;
                                oPmt.DocDate = DateTime.Now;
                                oPmt.Remarks = "eBizCharge ePayment Form Payment";

                                oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
                                oPmt.JournalRemarks = "eBizCharge ePayment Form payment";
                                oPmt.Invoices.DocEntry = int.Parse(i.InvoiceGUID);
                                oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice;
                                oPmt.Invoices.SumApplied = double.Parse(p.PaidAmount);
                                SAPCust sapcust = getCustomerByID(i.CustomerID, "");
                                AddCreditCardPayment(oPmt, sapcust, double.Parse(p.PaidAmount), true);
                                int r = oPmt.Add();
                                if (r != 0)
                                {
                                    errorLog("ePaymentForm Failed Payment. ID:" + i.CustomerID + "\r\n" +
                                        "InvoiceID:" + i.InvoiceGUID + "\r\n" +
                                          oCompany.GetLastErrorDescription());

                                }
                                else
                                {

                                    CCPayFormPayment pmt = new CCPayFormPayment();
                                    pmt.InvoiceID = p.InvoiceNumber;
                                    pmt.CustomerID = p.CustomerId;
                                    pmt.PaymentID = p.AuthCode;
                                    pmt.Amount = p.PaidAmount;
                                    pmt.DateImported = DateTime.Now;
                                    insert(pmt);
                                
                                }
                            }
                            i.PaidAmount = p.PaidAmount;
                            i.PaidDate = DateTime.Parse(p.DatePaid);
                            if (double.Parse(p.PaidAmount) < double.Parse(p.InvoiceAmount))
                                i.Status = EPAY_STATUS_PARTIAL;
                            else
                                i.Status = EPAY_STATUS_PAID;
                            update(i);
                        }
                    }
                    catch (Exception ex)
                    {
                        errorLog(ex);
                    }
                }
            }
         }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        SBO_Application.SetStatusBarMessage("Status check for customer: " + custid + " completed.", SAPbouiCOM.BoMessageTime.bmt_Medium, false);     
    }
    List<String> ListEmailTemplate = new List<string>();
    public void getEmailTemplate()
    {
        try
        {
            ListEmailTemplate = new List<string>();
            if (cfgMerchantID == "" || cfgMerchantID == null)
                return ;
            SecurityToken securityToken = getToken();

            //var eTempList = ebiz.GetEmailTemplates(securityToken, null, null);
            //foreach (EmailTemplate t in eTempList)
            //{
            //    ListEmailTemplate.Add(t.TemplateName);
            //}
            if (ListEmailTemplate.Count==0)
                throw new Exception("No email temaplate found.");
            
        }
        catch (Exception ex)
        {
            //showMessage(ex.Message);
             errorLog(ex);
        }
    }
    public bool IsPayFormCustomerExist(string customerID)
    {


        string sql = string.Format("select 1 from \"@CCPAYFORMINVOICE\" where \"U_CustomerID\" = '{0}' and \"U_Status\" <> 'Paid'", customerID);
        try
        {
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                return true;
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        return false;
    }
    private void SendAll()
    {
        try
        {

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oePaymentFormForm.Items.Item(matrixePaymentForm).Specific;
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                string status = getMatrixItem(oMatrix, "Status", i);
               
                if (status == "")
                {
                    addePayInvoice(oMatrix, i);
                }
            }
            populateePaymentFormMatrix();
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    public void checkePayPaidStatus(string custid, string securityID = "")
    {
        SAPDBDataContext db = new SAPDBDataContext(connStr);
        try
        {
            if (custid == "" || custid == null)
                return;
            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            string invType = "Invoice";
            if (oePaymentFormForm != null)
            {
                invType = getFormItemVal(oePaymentFormForm, cbInvType);
                showStatus("Checking " + invType + " payment status for customer: " + custid + ". Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium,false);
            }
            else if (cfgConnectUseDP == "Y")
                invType = "Down Payment";
            showStatus("Checking payment status for customer: " + custid + ". Please wait...",  SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            if (invType == "Down Payment")
            {
                oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments);

            }
            string invSql = " OINV (NOLOCK) ";
            if (invType == "Down Payment")
                invSql = " ODPI (NOLOCK) ";
            SecurityToken securityToken = getTokenEx(custid, "", "", "");
            if (securityID != "")
                securityToken.SecurityId = securityID;

            Payment[] pmts = ebiz.SearchEbizWebFormReceivedPayments(securityToken, custid, DateTime.Today.AddDays(-90), DateTime.Today.AddDays(1), null, 0, 1000, "");
            trace(string.Format("ebiz.SearchEbizWebFormReceivedPayments found {0} for custid {1}", pmts.Count(), custid));

            foreach (Payment p in pmts)
            {
                trace(string.Format("eMail Form pmt = {0}, {1}, {2}, {3}, {4}", p.InvoiceNumber, p.CustomerId, p.RefNum, p.AuthCode, p.PaidAmount));
                if (p.InvoiceNumber == "Method Request")
                    continue;
                //Invoice inv = ebiz.GetInvoice(securityToken, p.DatePaid, "", p.InvoiceNumber, "");
                var q = from x in db.CCPayFormInvoices
                        where x.InvoiceID == p.InvoiceNumber
                        select x;
                if (q.Count() == 0)
                {
                    trace("EMAIL Form invoice with payment not found in CCPayFormInvoice invoice no. " + p.InvoiceNumber);

                    CCPayFormInvoice ccinv = new CCPayFormInvoice();
                    ccinv.CustomerID = p.CustomerId;
                    ccinv.InvoiceID = p.InvoiceNumber;
                    ccinv.InvoiceGUID = getSQLInt(string.Format("select DocEntry from {0} where cardcode='{1}' and docnum = {2}", invSql, p.CustomerId, p.InvoiceNumber)).ToString();
                    ccinv.UploadedBalance = p.InvoiceAmount.ToString();
                    ccinv.PaidAmount = "0";
                    ccinv.Status = "Uploaded";
                    ccinv.UploadDate = DateTime.Today;
                    db.CCPayFormInvoices.InsertOnSubmit(ccinv);
                    db.SubmitChanges();
                    q = from x in db.CCPayFormInvoices
                        where x.InvoiceID == p.InvoiceNumber
                        select x;
                }
                foreach (CCPayFormInvoice i in q)
                {
                    try
                    {
                        if (!oDoc.GetByKey(int.Parse(i.InvoiceGUID)))
                        {
                            errorLog("Inv not found: " + i.InvoiceGUID);
                            i.Status = "Invoice not found: " + i.InvoiceGUID;

                            db.SubmitChanges();
                            continue;
                        }
                        if (oDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Close)
                        {
                            i.PaidAmount = p.PaidAmount;
                            i.PaidDate = DateTime.Parse(p.DatePaid);
                            i.Status = "Paid";
                            db.SubmitChanges();
                        }
                        else
                        {
                            var qcc = from x in db.CCPayFormPayments
                                      where x.CustomerID == p.CustomerId && x.InvoiceID == p.InvoiceNumber && x.PaymentID == p.AuthCode
                                      select x;
                            if (qcc.Count() == 0)
                            {
                                SAPbobsCOM.Payments oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                                int bplid = getBranchIdFromINV(p.InvoiceNumber);
                                if (bplid != 0)
                                    oPmt.BPLID = bplid;
                                oPmt.CardCode = i.CustomerID;
                                oPmt.DocDate = DateTime.Now;
                                //  oPmt.Remarks = "eBizCharge ePayment Form Payment";

                                oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
                                // oPmt.JournalRemarks = "eBizCharge ePayment Form payment";
                                oPmt.Invoices.DocEntry = int.Parse(i.InvoiceGUID);
                                oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice;
                                if (oDoc.DocObjectCode == SAPbobsCOM.BoObjectTypes.oDownPayments)
                                    oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_DownPayment;

                                oPmt.Invoices.SumApplied = getDoubleValue(p.PaidAmount);

                                SAPCust sapcust = getCustomerByID(i.CustomerID, "");
                                sapcust.cccust.CCAccountID = geteConnectCardCode(p.PaymentMethod);
                                AddCreditCardPayment(oPmt, sapcust, getDoubleValue(p.PaidAmount), true, p.RefNum);
                                int r = oPmt.Add();
                                if (r != 0)
                                {
                                    errorLog("ePaymentForm Failed Payment. ID:" + i.CustomerID + "\r\n" +
                                        "InvoiceID:" + i.InvoiceGUID + "\r\n" +
                                          oCompany.GetLastErrorDescription());

                                }
                                else
                                {

                                    CCPayFormPayment pmt = new CCPayFormPayment();
                                    pmt.InvoiceID = p.InvoiceNumber;
                                    pmt.CustomerID = p.CustomerId;
                                    pmt.PaymentID = p.AuthCode;
                                    pmt.Amount = p.PaidAmount;
                                    pmt.DateImported = DateTime.Now;
                                    db.CCPayFormPayments.InsertOnSubmit(pmt);
                                    db.SubmitChanges();
                                }
                            }
                            i.PaidAmount = p.PaidAmount;
                            i.PaidDate = DateTime.Parse(p.DatePaid);
                            if (getDoubleValue(p.PaidAmount) < getDoubleValue(p.InvoiceAmount))
                                i.Status = EPAY_STATUS_PARTIAL;
                            else
                                i.Status = EPAY_STATUS_PAID;
                            db.SubmitChanges();
                        }
                    }
                    catch (Exception ex)
                    {
                        errorLog(ex);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        showStatus("Status check for customer: " + custid + " completed.",SAPbouiCOM.BoMessageTime.bmt_Medium ,false);
    }
}

