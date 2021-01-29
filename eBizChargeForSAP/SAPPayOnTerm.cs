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
    const string menuPayOnTerm = "eBizPayOnTermMenu";
    const string formPayOnTerm = "eBizPayOnTermForm";
   
    const string matrixPayOnTerm = "mPOTrm";

    SAPbouiCOM.Form oPayOnTermForm;
    const string dtPayOnTerm = "dtPOTrm";
    private void AddPayOnTermMenu()
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
            oCreationPackage.UniqueID = menuPayOnTerm;
            oCreationPackage.String = "eBizCharge Pay On Term";
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
    private void CreatePayOnTermForm()
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
            oCreationParams.FormType = formPayOnTerm;

            oCreationParams.UniqueID = formPayOnTerm;
            try
            {
                oPayOnTermForm = SBO_Application.Forms.AddEx(oCreationParams);
            }
            catch (Exception)
            {
                oPayOnTermForm = SBO_Application.Forms.Item(formPayOnTerm);

            }

            // set the form properties
            oPayOnTermForm.Title = "eBizCharge Batch Invoicing for Preauth";
            oPayOnTermForm.Left = 400;
            oPayOnTermForm.Top = 100;
            oPayOnTermForm.ClientHeight = 460;
            oPayOnTermForm.ClientWidth = 750;





            //************************
            // Adding a Rectangle
            //***********************
            int margin = 5;
            oItem = oPayOnTermForm.Items.Add("Rect1", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = margin;
            oItem.Width = oPayOnTermForm.ClientWidth - 2 * margin;
            oItem.Top = margin;
            oItem.Height = oPayOnTermForm.ClientHeight - 40;

            int nTopGap = 25;
            int left = 6;
            int wBtn = 70;
            int hBtn = 19;
            int span = 80;
            if (cfgBatchAutoMode == "Y")
            {
                oItem = oPayOnTermForm.Items.Add(btnRefresh, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = left;
                oItem.Width = wBtn;
                oItem.Top = oPayOnTermForm.ClientHeight - nTopGap;
                oItem.Height = hBtn;

                oButton = ((SAPbouiCOM.Button)(oItem.Specific));

                oButton.Caption = "Orders";
                left += span;
            }
            oItem = oPayOnTermForm.Items.Add(btnInvoice, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oPayOnTermForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Process";
            left += span;

            oItem = oPayOnTermForm.Items.Add(btnClose, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oPayOnTermForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";

            margin = 8;
            int top = 15;
            int edL = 150; //oItm.Left + oItm.Width;
            int edW = 100;
            int edH = 15;
            int nGap = 26;

            if (cfgBatchAutoMode == "Y")
                oItem = addPaneItem(oPayOnTermForm, editStartDate, edL, top, edW, edH, "Start Date:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1);
            else
                oItem = addPaneItem(oPayOnTermForm, editStartDate, edL, top, edW, edH, "Delivery Note:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1);
            /*
            oItem = oPayOnTermForm.Items.Add(btnFind, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = edL + 120;
            oItem.Width = wBtn;
            oItem.Top = top - 2;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            if (cfgBatchAutoMode == "Y")
                oButton.Caption = "Report";
            else
                oButton.Caption = "Scan";
            */
            // oItem = addPaneItem(oPayOnTermForm, editEndDate, edL + 400, top, edW, edH, "Invoice Date:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 99);


            top = oItem.Top + nGap;

            oItem = oPayOnTermForm.Items.Add(matrixPayOnTerm, SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oItem.FromPane = 0;
            oItem.ToPane = 0;
            oItem.Left = 5 + margin;
            oItem.Width = oPayOnTermForm.ClientWidth - 2 * margin - 10;
            oItem.Top = top;
            oItem.Height = oPayOnTermForm.ClientHeight - 100;
            top = oItem.Height + oItem.Top + 2;
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oPayOnTermForm.Items.Item(matrixPayOnTerm).Specific;
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Add("RefNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Auth Code";
            oColumn.Width = 50;
            oColumn = oMatrix.Columns.Add("InvoiceID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Invoice";
            SAPbouiCOM.LinkedButton btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_Invoice).ToString();
            oColumn.Width = 40;

            oColumn = oMatrix.Columns.Add("DeliveryID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Delivery";
            btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_DeliveryNotes).ToString();
            oColumn.Width = 40;

            oColumn = oMatrix.Columns.Add("DeliveryNo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Delivery No.";
            oColumn.Width = 80;

            oColumn = oMatrix.Columns.Add("OrderID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Order";
            oColumn.Width = 40;
            btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_Order).ToString();


            oColumn = oMatrix.Columns.Add("customerID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Customer";
            oColumn.Width = 40;
            btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_BusinessPartner).ToString();

            oColumn = oMatrix.Columns.Add("CardHolder", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Name";
            oColumn.Width = 100;

            oColumn = oMatrix.Columns.Add("Amount", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Amount";
            oColumn.Width = 50;


            oColumn = oMatrix.Columns.Add("avsResult", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "AVS";
            oColumn.Width = 200;
            oColumn = oMatrix.Columns.Add("Result", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Result";
            oColumn.Width = 200;
            /*
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
            DateTime dt = getLastBatchDate();
            
            if (dt < DateTime.Parse("01/01/2000"))
                dt = DateTime.Today.AddDays(-3);
                */
            // if(cfgBatchAutoMode=="Y")
            setFormEditVal(oPayOnTermForm, editStartDate, DateTime.Today.AddDays(-15).ToString("MM/dd/yyyy"));

            populatePayOnTermMatrix(false);

        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
        oPayOnTermForm.Visible = true;

    }
    private void PayOnTermFormHandler(SAPbouiCOM.ItemEvent pVal)
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
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:

                        if ((pVal.CharPressed == 13 || pVal.CharPressed == 10 || pVal.CharPressed == 9) && pVal.ItemUID == editStartDate)
                        {
                            try
                            {
                                string deliveryID = getFormItemVal(oPayOnTermForm, editStartDate);

                                if (deliveryID != "")
                                {
                                    if (cfgBatchAutoMode != "Y")
                                    {
                                        if (getFormItemVal(oPayOnTermForm, editStartDate) != "")
                                        {
                                            if (ManualScan == "")
                                                ManualScan = getFormItemVal(oPayOnTermForm, editStartDate);
                                            else
                                                ManualScan = ManualScan + "," + getFormItemVal(oPayOnTermForm, editStartDate);
                                            setFormEditVal(oPayOnTermForm, editStartDate, "");
                                        }
                                    }
                                    populatePayOnTermMatrix(true, getFormItemVal(oPayOnTermForm, editStartDate));
                                }
                            }
                            catch (Exception e)
                            {
                                errorLog(e);
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        {
                            switch (pVal.ItemUID)
                            {
                                case btnClose:
                                    oPayOnTermForm.Close();
                                    break;
                                case btnInvoice:
                                    //DateTime dt = DateTime.Parse(getFormItemVal(oPayOnTermForm, editStartDate));
                                    bAuto = false;
                                    populatePayOnTermMatrix(false);
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oPayOnTermForm.Items.Item(matrixPayOnTerm).Specific;
                                    if (oMatrix.RowCount == 0)
                                    {
                                        SBO_Application.MessageBox("No order to process.");
                                        return;
                                    }
                                    if (SBO_Application.MessageBox("Capture and create invoices for delivered order?", 1, "Yes", "No") == 1)
                                    {
                                        tProcess = new Thread(BatchCreateInvoice);
                                        tProcess.Start();
                                        /*
                                        BatchCreateInvoice();
                                        populatePayOnTermMatrix(true);
                                        SetPaymentID();
                                        */
                                    }
                                    break;
                                case btnRefresh:
                                    try
                                    {

                                        populatePayOnTermMatrix(false);

                                    }
                                    catch (Exception ex)
                                    {
                                        SBO_Application.MessageBox(ex.Message);
                                    }
                                    break;
                                case btnFind:
                                    {

                                        if (cfgBatchAutoMode != "Y")
                                        {
                                            if (getFormItemVal(oPayOnTermForm, editStartDate) != "")
                                            {
                                                if (ManualScan == "")
                                                    ManualScan = getFormItemVal(oPayOnTermForm, editStartDate);
                                                else
                                                    ManualScan = ManualScan + "," + getFormItemVal(oPayOnTermForm, editStartDate);
                                                setFormEditVal(oPayOnTermForm, editStartDate, "");
                                            }
                                        }
                                        populatePayOnTermMatrix(true, getFormItemVal(oPayOnTermForm, editStartDate));
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
    private SAPbouiCOM.DataTable getPayOnTermMatrixDataTable()
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = oPayOnTermForm.DataSources.DataTables.Item(dtPayOnTerm);
        }
        catch (Exception)
        {
            oDT = oPayOnTermForm.DataSources.DataTables.Add(dtPayOnTerm);

        }

        return oDT;
    }
    private void populatePayOnTermMatrix(bool showResult, string batchDate = "")
    {
        string sql = "";
        try
        {
            setDeliveryID();
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oPayOnTermForm.Items.Item(matrixPayOnTerm).Specific;
            oMatrix.Clear();
            //oMatrix.AutoResizeColumns();
            SAPbouiCOM.DataTable oDts = getPayOnTermMatrixDataTable();
            string manualSQL = string.Format("and exists( select 1 from DLN1 l,ODLN d  where l.\"DocEntry\" = d.\"DocEntry\" and CAST(l.\"BaseEntry\" AS VARCHAR) = a.\"U_OrderID\" and l.\"BaseType\" = 17 and d.\"DocNum\" in ({0})", ManualScan);

            sql = "select \"U_amount\", \"U_OrderID\", \"U_InvoiceID\" \r\n" +
", \"U_DeliveryID\" as deliveryID  \r\n" +
", ( select d.\"DocNum\" from ODLN d  where CAST(d.\"DocEntry\" as VARCHAR) = \"U_DeliveryID\") as deliveryNo \r\n" +
", \"U_CardHolder\", \"U_avsResult\", \"U_refNum\", \"U_BatchResult\", \"U_customerID\"  \r\n" +
" from \"@CCTRANS\" a , ORDR b  \r\n" +
" where a.\"U_OrderID\" = CAST(b.\"DocEntry\" AS VARCHAR) and \"U_command\" in ( 'cc:authonly', 'cc:authonly-remain') and \"U_OrderID\" <> '' \r\n" +
"and b.CANCELED = 'N' and \"U_DeliveryID\" is not null \r\n" +
string.Format(" and \"U_recDate\" > TO_DATE('{0}', 'YYYYMMDD')  \r\n", getDateDiff());
            string orderBy = "\r\n order by \"U_OrderID\" desc";
            /*
            if (!showResult)
            {
                sql = sql + " and not exists(select 1 from ORDR where CAST(\"DocEntry\" as VARCHAR) = a.\"U_OrderID\" and CANCELED = 'Y')  \r\n" +
string.Format(" and \"U_recDate\" > ADD_DAYS(CURRENT_DATE, - {0})  \r\n", getDateDiff()) +
" and not exists(select 1 from OINV where \"DocStatus\" = 'C' and CAST(\"DocEntry\" as VARCHAR) = a.\"U_InvoiceID\")  \r\n";

                if(cfgBatchAutoMode != "Y")
                {
                    sql = sql + manualSQL;
                }
            }
            else
            {
                if (cfgBatchAutoMode == "Y")
                {

                    sql = sql + " and \"U_recDate\" > ADD_DAYS(CURRENT_DATE, - 30)  \r\n";
                }else
                    sql = sql + manualSQL;
            }
             */
            sql = sql + orderBy;

            oDts.ExecuteQuery(sql);
            //trace(sql);
            BindMatrix(oMatrix, "RefNum", "U_refNum", dtPayOnTerm);
            BindMatrix(oMatrix, "CardHolder", "U_CardHolder", dtPayOnTerm);
            BindMatrix(oMatrix, "customerID", "U_customerID", dtPayOnTerm);
            BindMatrix(oMatrix, "InvoiceID", "U_InvoiceID", dtPayOnTerm);
            BindMatrix(oMatrix, "DeliveryID", "deliveryID", dtPayOnTerm);
            BindMatrix(oMatrix, "DeliveryNo", "deliveryNo", dtPayOnTerm);
            BindMatrix(oMatrix, "OrderID", "U_OrderID", dtPayOnTerm);
            BindMatrix(oMatrix, "Amount", "U_amount", dtPayOnTerm);
            BindMatrix(oMatrix, "avsResult", "U_avsResult", dtPayOnTerm);
            BindMatrix(oMatrix, "Result", "U_BatchResult", dtPayOnTerm);
            oMatrix.LoadFromDataSource();


        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
  
}

