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

partial class SAP
{
    const string menuRBilling = "eBizRBillingMenu";
    const string formRBilling = "eBizRBillingForm";
    const string matrixRBilling = "mBInv";
    const string dtRBilling = "dtRBill";
    SAPbouiCOM.Form oRBillingForm;

   
    private void AddRBillingMenu()
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
            oCreationPackage.UniqueID = menuRBilling;
            oCreationPackage.String = "eBizCharge Recurring Billing";
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
    private void CreateRBillingForm()
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
            oCreationParams.FormType = formRBilling;

            oCreationParams.UniqueID = formRBilling;
            try
            {
                oRBillingForm = SBO_Application.Forms.AddEx(oCreationParams);
            }
            catch (Exception)
            {
                oRBillingForm = SBO_Application.Forms.Item(formRBilling);

            }

            // set the form properties
            oRBillingForm.Title = "eBizCharge Recurring Billing";
            oRBillingForm.Left = 400;
            oRBillingForm.Top = 100;
            oRBillingForm.ClientHeight = 460;
            oRBillingForm.ClientWidth = 975;





            //************************
            // Adding a Rectangle
            //***********************
            int margin = 5;
            oItem = oRBillingForm.Items.Add("Rect1", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = margin;
            oItem.Width = oRBillingForm.ClientWidth - 2 * margin;
            oItem.Top = margin;
            oItem.Height = oRBillingForm.ClientHeight - 40;

            int nTopGap = 25;
            int left = 6;
            int wBtn = 70;
            int hBtn = 19;
            int span = 80;
           

                oItem = oRBillingForm.Items.Add(btnRefresh, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = left;
                oItem.Width = wBtn;
                oItem.Top = oRBillingForm.ClientHeight - nTopGap;
                oItem.Height = hBtn;

                oButton = ((SAPbouiCOM.Button)(oItem.Specific));

                oButton.Caption = "Refresh";
                left += span;
            
            oItem = oRBillingForm.Items.Add(btnAdd, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oRBillingForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Add";
            left += span;

            oItem = oRBillingForm.Items.Add(btnRun, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oRBillingForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Run";

            left += span;

            oItem = oRBillingForm.Items.Add(btnClose, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oRBillingForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";

            margin = 8;
            int top = 15;
            int edL = 150; //oItm.Left + oItm.Width;
            int edW = 100;
            int edH = 15;
            int nGap = 26;

            oItem = addPaneItem(oRBillingForm, editCustomerID, edL, top, edW, edH, "Customer ID:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1000);
            oItem = addPaneItem(oRBillingForm, cbCustomerID, edL + 110, top, edW, edH, "", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 1000);

            oItem = oRBillingForm.Items.Add(btnFind, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = edL + 220;
            oItem.Width = wBtn;
            oItem.Top = top - 2;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Find";

            top = oItem.Top + nGap;


            oItem = oRBillingForm.Items.Add(matrixRBilling, SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oItem.FromPane = 0;
            oItem.ToPane = 0;
            oItem.Left = 5 + margin;
            oItem.Width = oRBillingForm.ClientWidth - 2 * margin - 10;
            oItem.Top = top;
            oItem.Height = oRBillingForm.ClientHeight - 100;
            top = oItem.Height + oItem.Top + 2;
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oRBillingForm.Items.Item(matrixRBilling).Specific;
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Add("jobID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "ID";
            oColumn.Width = 20;
            oColumn = oMatrix.Columns.Add("CustName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Name";
            oColumn.Width = 50;
            oColumn = oMatrix.Columns.Add("CustomerID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Customer";
            oColumn.Width = 80;
            SAPbouiCOM.LinkedButton btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_BusinessPartner).ToString();

            oColumn = oMatrix.Columns.Add("InvID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Invoice";
            btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_Invoice).ToString();
            oColumn.Width = 40;

            oColumn = oMatrix.Columns.Add("OrderID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Order";
            btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_Order).ToString();
            oColumn.Width = 40;

            oColumn = oMatrix.Columns.Add("Frequency", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Frequency";
            oColumn.Width = 50;

            oColumn = oMatrix.Columns.Add("StartDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Start Date";
            oColumn.Width = 50;

            oColumn = oMatrix.Columns.Add("EndDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "End Date";
            oColumn.Width = 50;
        /*
            oColumn = oMatrix.Columns.Add("CancelDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Cancel Date";
            oColumn.Width =50;
        */
            oColumn = oMatrix.Columns.Add("NextRun", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Next Run";
            oColumn.Width = 50;

            oColumn = oMatrix.Columns.Add("Amount", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Amount";
            oColumn.Width = 50;


            oColumn = oMatrix.Columns.Add("Desc", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Description";
            oColumn.Width = 200;
            oColumn = oMatrix.Columns.Add("LastRun", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Last Run";
            oColumn.Width = 50;
            oColumn = oMatrix.Columns.Add("Result", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Result";
            oColumn.Width = 200;
            populateRBillingMatrix();
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
        oRBillingForm.Visible = true;

    }
    private void RBillingFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            oRBillingForm = form;
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

                                        populateRBillingMatrix("", string.Format(" AND \"U_CustomerID\"='{0}'", getFormItemVal(oRBillingForm, editCustomerID)));
                                    }
                                    break;

                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == matrixRBilling)
                        {
                            if (pVal.ColUID != "jobID")
                            {
                                switch(pVal.ColUID)
                                {
                                    case "LastRun":
                                        populateRBillingMatrix("\"U_LastRunDate\" DESC,");
                                        break;
                                    case "Desc":
                                        populateRBillingMatrix("\"U_Description\",");
                                        break;
                                    case "NextRun":
                                        populateRBillingMatrix();
                                        break;
                                    case "InvID":
                                        populateRBillingMatrix("\"U_InvoiceID\", ");
                                        break;
                                    case "CustName":
                                        populateRBillingMatrix("\"U_CustomerName\", ");
                                        break;
                                    case "CustomerID":
                                        populateRBillingMatrix("\"U_CustomerID\", ");
                                        break;
                                   
                                    default:
                                        populateRBillingMatrix();
                                        break;
                                }

                            }
                            else
                                HandleJobSelect();
                            
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        {
                            switch (pVal.ItemUID)
                            {
                                case btnClose:
                                    oRBillingForm.Close();
                                    oRBillingForm = null;
                                    break;
                                case btnAdd:
                                    JobCustID = getFormItemVal(form, cbCustomerID);
                                    if (JobCustID == "")
                                    {
                                        showMessage("Please select a customer");
                                        return;
                                    }
                                        
                                    resetGlobal();
                                    JobCustID = getFormItemVal(form, cbCustomerID);

                                    CreateJobForm();
                                    bAuto = false;
                                    break;
                                case btnRefresh:
                                    try
                                    {
                                        populateRBillingMatrix();

                                    }
                                    catch (Exception ex)
                                    {
                                        showMessage(ex.Message);
                                    }
                                    break;
                                case btnRun:
                                    try
                                    {
                                        ProcessJob();
                                        
                                        populateRBillingMatrix();

                                    }
                                    catch (Exception ex)
                                    {
                                        showMessage(ex.Message);
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
                                        ComboAddItem(oCB, "");
                                        foreach(string c in list)
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
    private SAPbouiCOM.DataTable getRBillingMatrixDataTable()
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = oRBillingForm.DataSources.DataTables.Item(dtRBilling);
        }
        catch (Exception)
        {
            oDT = oRBillingForm.DataSources.DataTables.Add(dtRBilling);

        }

        return oDT;
    }
    private void populateRBillingMatrix(string orderBy = "", string where = "")
    {
        string sql = "";
        try
        {
            if (oRBillingForm == null)
                return;
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oRBillingForm.Items.Item(matrixRBilling).Specific;
            oMatrix.Clear();
            //oMatrix.AutoResizeColumns();
            SAPbouiCOM.DataTable oDts = getRBillingMatrixDataTable();
            sql = string.Format("select \"DocEntry\" as jobID, \r\n" +
"(select \"DocEntry\" from ORDR where CAST(\"DocEntry\" AS VARCHAR) = \"U_OrderID\") as OrderID,  \r\n" +
"\"U_Description\",  \r\n" +
"\"U_Result\",  \r\n" +
"\"U_CustomerID\",  \r\n" +
"\"U_Frequency\", \r\n" +
"\"U_StartDate\", \r\n" +
"\"U_EndDate\", \r\n" +
"\"U_NextRunDate\",  \r\n" +
"\"U_LastRunDate\", \r\n" +
"\"U_CancelledDate\",  \r\n" +
"\"U_Amount\",  \r\n" +
"(select \"CardName\" from OCRD where \"CardCode\" = \"U_CustomerID\") as CustomerName,  \r\n" +
"(select \"DocEntry\" from OINV where \"CardCode\" = \"U_CustomerID\" and CAST(\"DocNum\" as VARCHAR) = \"U_InvoiceID\") as InvDocNum  \r\n" +
"from \"@CCJOB\"  \r\n" +
"where (\"U_Cancelled\"<>'Y' OR \"U_Cancelled\" is NULL)  {0}  order By  {1} \"U_NextRunDate\" ", where, orderBy);

            //trace(sql);            
            oDts.ExecuteQuery(sql);
            BindMatrix(oMatrix, "jobID", "jobID", dtRBilling);
            BindMatrix(oMatrix, "CustName", "CustomerName", dtRBilling);
            BindMatrix(oMatrix, "InvID", "InvDocNum", dtRBilling);
            BindMatrix(oMatrix, "CustomerID", "U_CustomerID", dtRBilling);
            BindMatrix(oMatrix, "Desc", "U_Description", dtRBilling);
            BindMatrix(oMatrix, "Frequency", "U_Frequency", dtRBilling);
            BindMatrix(oMatrix, "StartDate", "U_StartDate", dtRBilling);
            BindMatrix(oMatrix, "EndDate", "U_EndDate", dtRBilling);
            BindMatrix(oMatrix, "NextRun", "U_NextRunDate", dtRBilling);
            BindMatrix(oMatrix, "LastRun", "U_LastRunDate", dtRBilling);
            BindMatrix(oMatrix, "Result", "U_Result", dtRBilling);
            //BindMatrix(oMatrix, "CancelDate", "CancelledDate", dtRBilling);
            BindMatrix(oMatrix, "OrderID", "OrderID", dtRBilling);
            BindMatrix(oMatrix, "Amount", "U_Amount", dtRBilling);
            oMatrix.LoadFromDataSource();


        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void HandleJobSelect()
    {
        try
        {
            
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oRBillingForm.Items.Item(matrixRBilling).Specific;
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                if (oMatrix.IsRowSelected(i))
                {
                    
                    resetGlobal();
                    JobOrderID= getMatrixItem(oMatrix, "OrderID", i);
                    JobID = getMatrixItem(oMatrix, "jobID", i);
                    CreateJobForm();
                    oMatrix.ClearSelections();
                    return;
                }
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }

}

