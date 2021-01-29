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
    const string menuBatchInv = "eBizBatchInvMenu";
    const string formBatchInv= "eBizBatchInvForm";
    const string btnRefresh = "btnRef";
    const string btnRun = "btnRun";
    const string matrixBatchInv = "mBInv";
    const string btnSelAll = "btnSelA";

    SAPbouiCOM.Form oBatchInvForm;
    const string dtBatchInv = "dtBInv";
    string ManualScan = "0";
    string ProcessedOrder = "";
    public class ColSort
    {
        public string col { get; set; }
        public string sort { get; set; }
    }
    private void AddBatchInvMenu()
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
            oCreationPackage.UniqueID = menuBatchInv;
            oCreationPackage.String = "eBizCharge Batch Invoicing";
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
    List<ColSort> listColSort = new List<ColSort>();
    private void CreateBatchInvForm()
    {
        try
        {
            listColSort = new List<ColSort>();
            ColSort cs = new ColSort();
            cs.col = "RefNum";
            cs.sort = "refNum";
            listColSort.Add(cs);
            cs = new ColSort();
            cs.col = "CardHolder";
            cs.sort = "CardHolder";
            listColSort.Add(cs);
            cs = new ColSort();
            cs.col = "customerID";
            cs.sort = "customerID";
            listColSort.Add(cs);
            cs = new ColSort();
            cs.col = "InvoiceID";
            cs.sort = "InvoiceID";
            listColSort.Add(cs);
            cs = new ColSort();
            cs.col = "InvoiceNo";
            cs.sort = "InvoiceID";
            listColSort.Add(cs);
            cs = new ColSort();
            cs.col = "DeliveryID";
            cs.sort = "DeliveryID";
            listColSort.Add(cs);
            cs = new ColSort();
            cs.col = "DeliveryNo";
            cs.sort = "DeliveryID";
            listColSort.Add(cs);
            cs = new ColSort();
            cs.col = "OrderID";
            cs.sort = "OrderID";
            listColSort.Add(cs);
            cs = new ColSort();
            cs.col = "OrderNo";
            cs.sort = "OrderID";
            listColSort.Add(cs);
            cs = new ColSort();
            cs.col = "Amount";
            cs.sort = "dAmount";
            listColSort.Add(cs);
            cs = new ColSort();
            cs.col = "avsResult";
            cs.sort = "avsResult";
            listColSort.Add(cs);
            cs = new ColSort();
            cs.col = "Result";
            cs.sort = "BatchResult";
            listColSort.Add(cs);
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
            oCreationParams.FormType = formBatchInv;

            oCreationParams.UniqueID = formBatchInv;
            try
            {
                oBatchInvForm = SBO_Application.Forms.AddEx(oCreationParams);
            }
            catch (Exception)
            {
                oBatchInvForm = SBO_Application.Forms.Item(formBatchInv);

            }

            // set the form properties
            oBatchInvForm.Title = "eBizCharge Batch Invoicing for Preauth";
            oBatchInvForm.Left = 400;
            oBatchInvForm.Top = 100;
            oBatchInvForm.ClientHeight = 460;
            oBatchInvForm.ClientWidth = 950;





            //************************
            // Adding a Rectangle
            //***********************
            int margin = 5;
            oItem = oBatchInvForm.Items.Add("Rect1", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = margin;
            oItem.Width = oBatchInvForm.ClientWidth - 2 * margin;
            oItem.Top = margin;
            oItem.Height = oBatchInvForm.ClientHeight - 40;

            int nTopGap = 25;
            int left = 6;
            int wBtn = 70;
            int hBtn = 19;
            int span = 80;
            if (cfgBatchAutoMode == "Y")
            {
                oItem = oBatchInvForm.Items.Add(btnRefresh, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = left;
                oItem.Width = wBtn;
                oItem.Top = oBatchInvForm.ClientHeight - nTopGap;
                oItem.Height = hBtn;

                oButton = ((SAPbouiCOM.Button)(oItem.Specific));

                oButton.Caption = "Orders";
                left += span;
            }
            oItem = oBatchInvForm.Items.Add(btnInvoice, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oBatchInvForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Process";
            left += span;
            oItem = oBatchInvForm.Items.Add(btnSelAll, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oBatchInvForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Select All";
            left += span;
            oItem = oBatchInvForm.Items.Add(btnClose, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oBatchInvForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";
            
            margin = 8;
            int top = 15;
            int edL = 150; //oItm.Left + oItm.Width;
            int edW = 100;
            int edH = 15;
            int nGap = 26;
            
            if(cfgBatchAutoMode=="Y")
                oItem = addPaneItem(oBatchInvForm, editStartDate, edL, top, edW, edH, "Start Date:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1);
            else
                oItem = addPaneItem(oBatchInvForm, editStartDate, edL, top, edW, edH, "Delivery Note:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1);
            /*
            oItem = oBatchInvForm.Items.Add(btnFind, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
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
            oItem = addPaneItem(oBatchInvForm, editEndDate, edL + 400, top, edW, edH, "Invoice Date:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 99);


            top = oItem.Top + nGap;
            
            oItem = oBatchInvForm.Items.Add(matrixBatchInv, SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oItem.FromPane = 0;
            oItem.ToPane = 0;
            oItem.Left = 5 + margin;
            oItem.Width = oBatchInvForm.ClientWidth - 2 * margin - 10;
            oItem.Top = top;
            oItem.Height = oBatchInvForm.ClientHeight - 100;
            top = oItem.Height + oItem.Top + 2;
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oBatchInvForm.Items.Item(matrixBatchInv).Specific;
             SAPbouiCOM.Column oColumn = oMatrix.Columns.Add("sel", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "";
            oColumn.Width = 30;
            oColumn.ValOff = "N";
            oColumn.ValOn = "Y";
            oColumn.Editable = true;
            oColumn = oMatrix.Columns.Add("RefNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Auth Code";
            oColumn.Width = 50;
            oColumn = oMatrix.Columns.Add("InvoiceID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "";
            SAPbouiCOM.LinkedButton btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_Invoice).ToString();
            oColumn.Width = 18;

            oColumn = oMatrix.Columns.Add("InvoiceNo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Invoice No.";
            oColumn.Width = 70;

            oColumn = oMatrix.Columns.Add("DeliveryID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "";
            btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_DeliveryNotes).ToString();
            oColumn.Width = 18;

            oColumn = oMatrix.Columns.Add("DeliveryNo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Delivery No.";
            oColumn.Width = 70;

            oColumn = oMatrix.Columns.Add("OrderID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "";
            oColumn.Width = 18;
            btn = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            btn.LinkedObjectType = ((int)SAPbouiCOM.BoLinkedObject.lf_Order).ToString();


            oColumn = oMatrix.Columns.Add("OrderNo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Order No.";
            oColumn.Width = 70;

            oColumn = oMatrix.Columns.Add("customerID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "Customer";
            oColumn.Width = 60;
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
                setFormEditVal(oBatchInvForm, editStartDate, DateTime.Today.AddDays(-15).ToString("MM/dd/yyyy"));
            
            populateBatchInvMatrix(false);

        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
        oBatchInvForm.Visible = true;

    }
    private void BatchInvFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            oBatchInvForm = form;
            if (pVal.BeforeAction)
            {

            }
            else
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == matrixBatchInv)
                        {
                            if (pVal.Row == 0)
                            {
                                var q = from x in listColSort
                                        where x.col == pVal.ColUID
                                        select x;
                                foreach (ColSort cs in q)
                                {
                                    //populateBatchInvMatrix(cs.sort);
                                    string desc = " DESC ";
                                    if (cs.sort.IndexOf(desc) >= 0)
                                        cs.sort = cs.sort.Replace(desc, " ");
                                    else
                                        cs.sort = cs.sort + desc;

                                }
                            }
                            else
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oBatchInvForm.Items.Item(matrixBatchInv).Specific;
                                for (int i = 1; i <= oMatrix.RowCount; i++)
                                {
                                    SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("sel").Cells.Item(i).Specific;
                                    if (oCheckBox.Checked)
                                    {
                                        string DeliveryID = getMatrixItem(oMatrix, "DeliveryID", i);
                                        if(DeliveryID != "")
                                            selectSameDelivery(DeliveryID);
                                    }
                                }
                            }
                        }
                        break;
                        
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        
                        if ((pVal.CharPressed == 13 || pVal.CharPressed == 10 || pVal.CharPressed == 9) && pVal.ItemUID == editStartDate)
                        {
                            try
                            {
                                string deliveryID = getFormItemVal(oBatchInvForm, editStartDate);
                              
                                if (deliveryID != "")
                                {
                                    if (cfgBatchAutoMode != "Y")
                                    {
                                        if (getFormItemVal(oBatchInvForm, editStartDate) != "")
                                        {
                                            if (ManualScan == "")
                                                ManualScan = getFormItemVal(oBatchInvForm, editStartDate);
                                            else
                                                ManualScan = ManualScan + "," + getFormItemVal(oBatchInvForm, editStartDate);
                                            setFormEditVal(oBatchInvForm, editStartDate, "");
                                        }
                                    }
                                    populateBatchInvMatrix(true, getFormItemVal(oBatchInvForm, editStartDate));
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
                                    oBatchInvForm.Close();
                                    break;
                                case btnInvoice:
                                    //DateTime dt = DateTime.Parse(getFormItemVal(oBatchInvForm, editStartDate));
                                    bAuto = false;
                                    
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oBatchInvForm.Items.Item(matrixBatchInv).Specific;
                                    if (oMatrix.RowCount == 0)
                                    {
                                        SBO_Application.MessageBox("No order to process.");
                                        return;
                                    }
                                    string msg = "Create invoices and capture funds?";

                                    if (cfgCaptureOnly == "Y")
                                        msg = "Capture and create payment on invoices?";

                                    if (SBO_Application.MessageBox(msg, 1, "Yes", "No") == 1)
                                    {
                                        oBatchInvForm = form;
                                        if (cfgBINoDelivery == "Y")
                                        {
                                            tProcess = new Thread(BatchCaptureInvoice);
                                        }
                                        else
                                        {
                                            tProcess = new Thread(BatchCreateInvoice);
                                        }
                                        tProcess.Start();
                                        /*
                                        BatchCreateInvoice();
                                        populateBatchInvMatrix(true);
                                        SetPaymentID();
                                        */
                                    }
                                
                                    break;
                                case btnRefresh:
                                    try
                                    {

                                        populateBatchInvMatrix(false);

                                    }
                                    catch (Exception ex)
                                    {
                                        SBO_Application.MessageBox(ex.Message);
                                    }
                                    break;
                                case btnFind:
                                    {
                                        
                                        if (cfgBatchAutoMode!="Y")
                                        {
                                            if (getFormItemVal(oBatchInvForm, editStartDate) != "")
                                            {
                                                if (ManualScan == "")
                                                    ManualScan = getFormItemVal(oBatchInvForm, editStartDate);
                                                else
                                                    ManualScan = ManualScan + "," + getFormItemVal(oBatchInvForm, editStartDate);
                                                setFormEditVal(oBatchInvForm, editStartDate, "");
                                            }
                                        }
                                        populateBatchInvMatrix(true, getFormItemVal(oBatchInvForm, editStartDate));
                                    }
                                    break;
                                case btnSelAll:
                                    try
                                    {
                                        string sel = getFormItemVal(oBatchInvForm, btnSelAll);
                                        if (sel == "Select All")
                                        {
                                            populateBatchInvMatrix(false, "", "", " 'Y' as sel, ");
                                            setButtonCaption(oBatchInvForm, btnSelAll, "Unselect All");
                                        }
                                        else
                                        {
                                            populateBatchInvMatrix(false, "", "", " 'N' as sel, ");
                                            setButtonCaption(oBatchInvForm, btnSelAll, "Select All");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        showMessage(ex.Message);
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
    private void DocWizardFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
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
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        {
                            switch (pVal.ItemUID)
                            {
                                case "4":
                                    if (cfgCaptureOnPrint == "Y")
                                    {
                                        if (getFormItemVal(form, "4") == "Finish")
                                        {
                                           tProcess = new Thread(captureOnAdd);
                                           tProcess.Start();
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
    private SAPbouiCOM.DataTable getBatchInvMatrixDataTable()
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = oBatchInvForm.DataSources.DataTables.Item(dtBatchInv);
        }
        catch (Exception)
        {
            oDT = oBatchInvForm.DataSources.DataTables.Add(dtBatchInv);

        }

        return oDT;
    }
    private void populateBatchInvMatrix(bool showResult,string batchDate = "", string orderBy = "", string selAll = "")
    {
        string sql = "";
        try
        {
            setDeliveryID();
            setInvoiceID();
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oBatchInvForm.Items.Item(matrixBatchInv).Specific;
            oMatrix.Clear();
            //oMatrix.AutoResizeColumns();
            SAPbouiCOM.DataTable oDts = getBatchInvMatrixDataTable();
            string manualSQL = string.Format("and exists( select 1 from DLN1 l,ODLN d  where l.\"DocEntry\" = d.\"DocEntry\" and CAST(l.\"BaseEntry\" AS VARCHAR) = a.\"U_OrderID\" and l.\"BaseType\" = 17 and d.\"DocNum\" in ({0})", ManualScan);
            string cv = "";
            if (selAll == "")
                cv = " 'N' as sel, ";
            else
                cv = selAll;
            sql = "select \"U_amount\", \"U_OrderID\", \"U_InvoiceID\", " + cv + " \r\n" +
" \"U_DeliveryID\" as deliveryID, b.\"DocNum\" as orderNo  \r\n" +
", ( select i.\"DocNum\" from OINV i  where CAST(i.\"DocEntry\" as VARCHAR) = \"U_InvoiceID\") as invoiceNo \r\n" +
", ( select d.\"DocNum\" from ODLN d  where CAST(d.\"DocEntry\" as VARCHAR) = \"U_DeliveryID\") as deliveryNo \r\n" +
", \"U_CardHolder\", \"U_avsResult\", \"U_refNum\", \"U_BatchResult\", \"U_customerID\"  \r\n" +
" from \"@CCTRANS\" a , ORDR b  \r\n" +
" where a.\"U_OrderID\" = CAST(b.\"DocEntry\" AS VARCHAR) and b.CANCELED = 'N' and \"U_OrderID\" <> '' " +
string.Format(" and \"U_recDate\" > TO_DATE('{0}', 'YYYYMMDD')  \r\n", getDateDiff());

            if (showResult)
            {
                sql = sql + " and \"U_command\" in ( 'cc:authonly-captured','cc:authonly-error', 'cc:authonly', 'cc:authonly-remain')  \r\n";
                if(ProcessedOrder != "")
                    sql = sql + string.Format(" and a.\"U_OrderID\" in ({0}) \r\n", ProcessedOrder);
            } else
            {
                sql = sql + " and \"U_command\" in ( 'cc:authonly', 'cc:authonly-remain') \r\n" +
                " and not exists(select 1 from OINV where CAST(\"DocEntry\" AS VARCHAR) = a.\"U_InvoiceID\" and \"DocStatus\" = 'C') \r\n";
             }
            if (cfgBINoDelivery != "Y")
                sql = sql + " and \"U_DeliveryID\" <> '' \r\n";
            if(cfgCaptureOnly == "Y")
                sql = sql + " and \"U_InvoiceID\" <> '' \r\n";

            orderBy = "\r\n order by \"U_OrderID\" desc";
           
          
            sql = sql + orderBy;
            
            oDts.ExecuteQuery(sql);
            trace(sql);
            BindMatrix(oMatrix, "RefNum", "U_refNum", dtBatchInv);
            BindMatrix(oMatrix, "CardHolder", "U_CardHolder", dtBatchInv);
            BindMatrix(oMatrix, "customerID", "U_customerID", dtBatchInv);
            BindMatrix(oMatrix, "InvoiceID", "U_InvoiceID", dtBatchInv);
            BindMatrix(oMatrix, "InvoiceNo", "invoiceNo", dtBatchInv);
            BindMatrix(oMatrix, "DeliveryID", "deliveryID", dtBatchInv);
            BindMatrix(oMatrix, "DeliveryNo", "deliveryNo", dtBatchInv);
            BindMatrix(oMatrix, "OrderID", "U_OrderID", dtBatchInv);
            BindMatrix(oMatrix, "OrderNo", "orderNo", dtBatchInv);
            BindMatrix(oMatrix, "Amount", "U_amount", dtBatchInv);
            BindMatrix(oMatrix, "avsResult", "U_avsResult", dtBatchInv);
            BindMatrix(oMatrix, "Result", "U_BatchResult", dtBatchInv);
            BindMatrix(oMatrix, "sel", "sel", dtBatchInv, true);
            oMatrix.LoadFromDataSource();


        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private string getDateDiff()
    {
       
        try
        {
              string s = getFormItemVal(oBatchInvForm, editStartDate);
             return DateTime.Parse(s).ToString("yyyyMMdd");
        }catch(Exception)
        {
              setFormEditVal(oBatchInvForm, editStartDate, DateTime.Today.AddDays(-15).ToString("MM/dd/yyyy"));
        }
        return "20100101";
    }
    private void SetPaymentID()
    {
        string refNum = "";
        try
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oBatchInvForm.Items.Item(matrixBatchInv).Specific;
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                refNum = getMatrixItem(oMatrix, "RefNum", i);
                CCTRANUpdateIncomingPaymentID(refNum);
            } 
        }
        catch (Exception ex)
        {
            errorLog(ex.Message);
        }
    }
    private void BatchCaptureInvoice()
    {
        string refNum = "";
        ProcessedOrder = "";
        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oBatchInvForm.Items.Item(matrixBatchInv).Specific;
        int i = 0;
        try
        {
            FormFreeze(oBatchInvForm, true);
            DateTime invDate = DateTime.Today;
            string dt = getFormItemVal(oBatchInvForm, editEndDate);
            if (dt != "")
            {
                try
                {
                    invDate = DateTime.Parse(dt);
                }
                catch (Exception e)
                {
                    errorLog(e);
                }
            }

            for (i = 1; i <= oMatrix.RowCount; i++)
            {
                if (getMatrixItem(oMatrix, "Result", i) == "Success")
                    continue;
                SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("sel").Cells.Item(i).Specific;
                if (!oCheckBox.Checked)
                    continue;
              
                // showMessage(string.Format("Batch invoicing process i=" + i + "row count=" + oMatrix.RowCount));
                string customerID = getMatrixItem(oMatrix, "customerID", i);
                refNum = getMatrixItem(oMatrix, "RefNum", i);
                SAPCust sapcust = getCustomerByID(customerID, refNum);

                string OrderID = getMatrixItem(oMatrix, "OrderID", i);
                if (ProcessedOrder == "")
                    ProcessedOrder = "'" + OrderID + "'";
                else
                    ProcessedOrder = ProcessedOrder + ",'" + OrderID + "'";

                showStatus("eBizCharge processing order: " + OrderID + ". Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                string DeliveryID = getMatrixItem(oMatrix, "DeliveryID", i);
                string InvoiceID = getMatrixItem(oMatrix, "InvoiceID", i);
                string Amount = getMatrixItem(oMatrix, "Amount", i);
                SAPbobsCOM.Documents oSODoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                if (!oSODoc.GetByKey(int.Parse(OrderID)))
                {
                    setBatchResult(refNum, string.Format("Order Not found."), oMatrix, "Result", i);
                }
                else
                {
                    InvoiceID = batch_createInvoiceFromSO(oSODoc, invDate).ToString();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSODoc);
                    oSODoc = null;
                    if (InvoiceID == "0")
                    {
                        setBatchResult(refNum, string.Format("Cannot create invoice."), oMatrix, "Result", i);
                    }
                    else
                    {
                        SAPbobsCOM.Documents oInvDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                        if (!oInvDoc.GetByKey(int.Parse(InvoiceID)))
                        {
                            setBatchResult(refNum, string.Format("Invoice not found."), oMatrix, "Result", i);
                        } else
                        {
                            if (oInvDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Close)
                            {
                                setBatchResult(refNum, string.Format("Invoice is closed."), oMatrix, "Result", i);
                            }
                            else if (oInvDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Paid)
                            {
                                setBatchResult(refNum, string.Format("Invoice is paid."), oMatrix, "Result", i);
                            }
                            else
                            {
                                double bal = getInvoiceBalance(InvoiceID);
                                string camount = getDoubleValue(Amount) > bal ? bal.ToString() : Amount;

                                if (refNum == "" || refNum == null)
                                {
                                    CustomerTransactionRequest req = Batch_createCustomRequest(oInvDoc);
                                    req.Details.Amount = bal;
                                    req.Details.Subtotal = 0; // Math.Round(req.Details.Amount - (req.Details.Shipping + req.Details.Tax - req.Details.Discount), 2);

                                    req.Command = "cc:sale";

                                    if (Batch_runCustomRequest(req, sapcust, OrderID, InvoiceID, ref refNum))
                                    {
                                        confirmNum = refNum;

                                        AddPayment(bal, sapcust, InvoiceID);
                                        BatchSetTransInvoiceIDByDeliveryID(InvoiceID, refNum, DeliveryID);
                                        double remain = getDoubleValue(Amount) - bal;
                                        if (remain > 0)
                                            DuplicateTransLogForPartial(remain.ToString(), DeliveryID, refNum);
                                    }
                                }
                                else
                                {
                                    if (CaptureTrans(refNum, sapcust, camount, oInvDoc))
                                    {

                                        if (!AddPayment(getDoubleValue(camount), sapcust, InvoiceID))
                                        {
                                            setBatchResult(refNum, string.Format("Add payment on capture of {0} failed.", camount), oMatrix, "Result", i);
                                            continue;
                                        }
                                        setBatchResult(refNum, "Success", oMatrix, "Result", i);

                                        BatchSetTransInvoiceID(refNum, InvoiceID, OrderID, sapcust.cccust.CCAccountID, customerID);
                                        setCaptureResult(refNum, "cc:authonly-captured");
                                        bool hasNextT = hasNextPreauth(DeliveryID, refNum, i);
                                        if (hasNextT)
                                        {
                                            showStatus("Batch Invoicing Inv has multiple transaction on sales order deliveryID=" + DeliveryID, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                            continue;
                                        }
                                        double remain = Math.Round(getDoubleValue(Amount) - getDoubleValue(camount), 2);
                                        if (remain > 0)
                                        {
                                            trace("Batch Invoicing Inv partial shipment create dummy record. remain=" + remain);
                                            DuplicateTransLogForPartial(remain.ToString(), DeliveryID, refNum);
                                        }
                                        bal = getInvoiceBalance(InvoiceID);
                                        if (bal > 0)
                                        {
                                            showStatus("Batch Invoicing Inv charge remain after capture.  Balance=" + bal, SAPbouiCOM.BoMessageTime.bmt_Medium, false);


                                            CustomerTransactionRequest req = Batch_createCustomRequest(oInvDoc);
                                            req.Details.Amount = bal;
                                            req.Details.Subtotal = 0; // Math.Round(req.Details.Amount - (req.Details.Shipping + req.Details.Tax - req.Details.Discount), 2);

                                            req.Command = "cc:sale";

                                            if (Batch_runCustomRequest(req, sapcust, OrderID, InvoiceID, ref refNum))
                                            {
                                                confirmNum = refNum;

                                                if (!AddPayment(bal, sapcust, InvoiceID))
                                                {
                                                    setBatchResult(refNum, string.Format("Add remain payment of {0} failed.", bal), oMatrix, "Result", i);
                                                    continue;
                                                }
                                                BatchSetTransInvoiceIDByDeliveryID(InvoiceID, refNum, DeliveryID);

                                            }
                                        }
                                    }
                                    else
                                    {

                                        setBatchResult(refNum, "Failed to capture transaction.", oMatrix, "Result", i);
                                    }
                                }
                               
                            }
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvDoc);
                        oInvDoc = null;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            errorLog(ex.Message);
            setBatchResult(refNum, "Error: " + ex.Message, oMatrix, "Result", i);
        }
        finally
        {
            //populateBatchInvMatrix(true);
            FormFreeze(oBatchInvForm, false);
        }
    }
    
    private void BatchCreateInvoice()
    {
        string refNum = "";
        ProcessedOrder = "";
        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oBatchInvForm.Items.Item(matrixBatchInv).Specific;
        int i = 0;
        try
        {
            FormFreeze(oBatchInvForm,true);
            DateTime invDate = DateTime.Today;
            string dt = getFormItemVal(oBatchInvForm, editEndDate);
            if (dt != "")
            {
                try
                {
                    invDate = DateTime.Parse(dt);
                }
                catch (Exception e)
                {
                    errorLog(e);
                }
            }

              for (i = 1; i <= oMatrix.RowCount; i++)
            {
                if (getMatrixItem(oMatrix, "Result", i) == "Success")
                    continue;
                SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("sel").Cells.Item(i).Specific;
                if (!oCheckBox.Checked)
                    continue;
               // showMessage(string.Format("Batch invoicing process i=" + i + "row count=" + oMatrix.RowCount));
                string customerID = getMatrixItem(oMatrix, "customerID", i);
                refNum = getMatrixItem(oMatrix, "RefNum", i);
                SAPCust sapcust = getCustomerByID(customerID, refNum);

                string OrderID = getMatrixItem(oMatrix, "OrderID", i);
                if (ProcessedOrder == "")
                    ProcessedOrder = "'" + OrderID + "'";
                else
                    ProcessedOrder = ProcessedOrder + ",'" + OrderID + "'";
                showStatus("eBizCharge processing order: " + OrderID + ". Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                string DeliveryID = getMatrixItem(oMatrix, "DeliveryID", i); 
                string InvoiceID = getMatrixItem(oMatrix, "InvoiceID", i);
                string Amount = getMatrixItem(oMatrix, "Amount", i);
                SAPbobsCOM.Documents oSODoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                if (!oSODoc.GetByKey(int.Parse(OrderID)))
                {
                    setBatchResult(refNum, string.Format("order not found."), oMatrix, "Result", i);
                }
                else
                {
                    if (InvoiceID == "" && cfgCaptureOnly != "Y")
                    {
                        SAPbobsCOM.Documents oDeliveryDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
                        if (!oDeliveryDoc.GetByKey(int.Parse(DeliveryID)))
                        {
                            InvoiceID = "0";
                            setBatchResult(refNum, "No delivery note", oMatrix, "Result", i);
                        }
                        else
                        {
                            InvoiceID = batch_createInvoice(oDeliveryDoc, invDate).ToString();
                            if (InvoiceID == "0")
                                setBatchResult(refNum, "Can not create invoice.  " + getErrorString(), oMatrix, "Result", i);
                            else
                            {
                                try
                                {
                                    oDeliveryDoc.Close();
                                }
                                catch (Exception) { }
                                CCTRANAdjustSOWithInvId(int.Parse(InvoiceID), int.Parse(OrderID));
                            }
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oDeliveryDoc);
                        oDeliveryDoc = null;
                    }
                    if (InvoiceID == "0")
                    {
                        setBatchResult(refNum, string.Format("Cannot create invoice."), oMatrix, "Result", i);
                    }
                    else
                    {
                        SAPbobsCOM.Documents oInvDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                        if(!oInvDoc.GetByKey(int.Parse(InvoiceID)))
                        {
                            setBatchResult(refNum, "Invoice not found", oMatrix, "Result", i);
                        }
                        else
                        {
                            if (oInvDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Close)
                            {
                                setBatchResult(refNum, string.Format("Invoice is closed."), oMatrix, "Result", i);
                            }
                            else if (oInvDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Paid)
                            {
                                setBatchResult(refNum, string.Format("Invoice is paid."), oMatrix, "Result", i);
                            }
                            else
                            {
                                double bal = getInvoiceBalance(InvoiceID);
                                string camount = getDoubleValue(Amount) > bal ? bal.ToString() : Amount;

                                if (refNum == "" || refNum == null)
                                {
                                    CustomerTransactionRequest req = Batch_createCustomRequest(oInvDoc);
                                    req.Details.Amount = bal;
                                    req.Details.Subtotal = 0; // Math.Round(req.Details.Amount - (req.Details.Shipping + req.Details.Tax - req.Details.Discount), 2);

                                    req.Command = "cc:sale";

                                    if (!Batch_runCustomRequest(req, sapcust, OrderID, InvoiceID, ref refNum))
                                    {
                                        setBatchResult(refNum, "Failed to charge " + bal, oMatrix, "Result", i);
        }
                                    else
                                    {
                                        confirmNum = refNum;
                                        
                                        AddPayment(bal, sapcust, InvoiceID);
                                        BatchSetTransInvoiceIDByDeliveryID(InvoiceID, refNum, DeliveryID);
                                        double remain = getDoubleValue(Amount) - bal;
                                        if (remain > 0)
                                            DuplicateTransLogForPartial(remain.ToString(), DeliveryID, refNum);
                                    }
                                }
                                else
                                {
                                    if (!CaptureTrans(refNum, sapcust, camount, oInvDoc))
                                    {
                                        setBatchResult(refNum, "Failed to capture transaction.", oMatrix, "Result", i);
                                    }
                                    else
                                    {

                                        if (!AddPayment(getDoubleValue(camount), sapcust, InvoiceID))
                                        {
                                            setBatchResult(refNum, string.Format("Add payment on capture of {0} failed.", camount), oMatrix, "Result", i);
                                        }
                                        else
                                        {
                                            setBatchResult(refNum, "Success", oMatrix, "Result", i);

                                            BatchSetTransInvoiceID(refNum, InvoiceID, OrderID, sapcust.cccust.CCAccountID, customerID);
                                            setCaptureResult(refNum, "cc:authonly-captured");
                                            bool hasNextT = hasNextPreauth(DeliveryID, refNum, i);
                                            if (hasNextT)
                                            {
                                                trace("Batch Invoicing Inv has multiple transaction on sales order deliveryID=" + DeliveryID);
                                                continue;
                                            }
                                            double remain = Math.Round(getDoubleValue(Amount) - getDoubleValue(camount), 2);
                                            if (remain > 0)
                                            {
                                                trace("Batch Invoicing Inv partial shipment create dummy record. remain=" + remain);
                                                DuplicateTransLogForPartial(remain.ToString(), DeliveryID, refNum);
                                            }
                                            bal = getInvoiceBalance(InvoiceID);
                                            if (bal > 0)
                                            {
                                                trace("Batch Invoicing Inv charge remain after capture.  Balance=" + bal);


                                                CustomerTransactionRequest req = Batch_createCustomRequest(oInvDoc);
                                                req.Details.Amount = bal;
                                                req.Details.Subtotal = 0; // Math.Round(req.Details.Amount - (req.Details.Shipping + req.Details.Tax - req.Details.Discount), 2);

                                                req.Command = "cc:sale";

                                                if (Batch_runCustomRequest(req, sapcust, OrderID, InvoiceID, ref refNum))
                                                {
                                                    confirmNum = refNum;

                                                    if (!AddPayment(bal, sapcust, InvoiceID))
                                                    {
                                                        setBatchResult(refNum, string.Format("Add remain payment of {0} failed.", bal), oMatrix, "Result", i);
                                                        continue;
                                                    }
                                                    BatchSetTransInvoiceIDByDeliveryID(InvoiceID, refNum, DeliveryID);

                                                }
                                            }
                                        }
                                    }                                        
                                }
                            }
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvDoc);
                        oInvDoc = null;
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oSODoc);
                oSODoc = null;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex.Message);
            setBatchResult(refNum, "Error: " + ex.Message, oMatrix, "Result", i);
        }
        finally
        {
          //  populateBatchInvMatrix(true);
            FormFreeze(oBatchInvForm,false);
        }
    }
    public bool CaptureTrans(string refNum, SAPCust sapcust, string amount, SAPbobsCOM.Documents doc)
    {
        try
        {
            TransactionRequestObject req = Batch_createRequest(doc);
            req.CustReceipt = cfgsendCustomerEmail == "Y" ? true : false;
            req.CustReceiptName = "vterm_customer";
            try
            {
                req.BillingAddress.Email = sapcust.custObj.Email;
            }
            catch (Exception) { }
            req.RefNum = refNum;
            req.Details.Amount = getDoubleValue(amount);
            req.Command = "cc:capture";

            SecurityToken token = getToken(sapcust.cccust.CCAccountID);
            TransactionResponse resp = new TransactionResponse();

            return runCustomRequest(req, sapcust, ref refNum);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
      
        return true;
    }
    public SAPCust getCustomerByID(string ID, string refNum )
    {
       // trace("getCustomerByID ID=" + ID + ",refNum=" + refNum);
        SAPCust sapcust = new SAPCust();
        sapcust.cccust = new CCCUST();
        sapcust.custObj = new Customer();
        try { 
          
            if (ID != "")
            {
                List<CCCUST> q = findCCCust(string.Format(" where \"U_CustomerID\"='{0}' AND \"U_CustNum\" <> '' and \"U_checkingAccount\" = '' ", ID));
                foreach (CCCUST cust in q)
                {
                   // trace("Found CCust by ID = " + ID);
                    sapcust.cccust = cust;
                    sapcust.key = cust.methodDescription;
                    sapcust.custObj = new Customer();
                    sapcust.custObj.CustomerToken = cust.CustNum;
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
                    sapcust.custObj.PaymentMethodProfiles[0].MethodID = cust.MethodID;
                    sapcust.custObj.PaymentMethodProfiles[0].Routing = cust.routingNumber;
                    sapcust.custObj.PaymentMethodProfiles[0].Account = cust.checkingAccount;
                    if (cust.checkingAccount != null)
                        sapcust.custObj.PaymentMethodProfiles[0].MethodType = "check";
                    else
                        sapcust.custObj.PaymentMethodProfiles[0].MethodType = "cc";
                    if (cust.methodName != null)
                    {
                        if (cust.methodName.ToLower() == "primary")
                            return sapcust;
                    }
                    //trace("Found CCust by ID = " + ID + " OK");
                }
                if (q.Count() == 0)
                {
                    //Needed for eConnect customer without card.
                 //   trace("CCust not found ID = " + ID);
                    CCCUST cust = new CCCUST();
                    cust.recID = 0;
                    cust.CustomerID = ID;
                    cust.CustNum = "";
                    cust.MethodID = "";
                    cust.methodDescription = "";
                    cust.cardNum = "";
                    cust.expDate = "";
                    cust.routingNumber = "";
                    cust.checkingAccount = "";
                    cust.firstName = "";
                    cust.lastName = "";
                    cust.email = "";
                    cust.street = "";
                    cust.city = "";
                    cust.state = "";
                    cust.zip = "";
                    cust.cardType = "";
                    cust.GroupName = getGroupName(ID);
                    cust.CardName = "";
                    cust.CCAccountID = getCardCode();
                    sapcust.cccust = cust;
                    sapcust.key = cust.methodDescription;
                    sapcust.custObj = new Customer();
                    sapcust.custObj.CustomerToken = cust.CustNum;
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
                    sapcust.custObj.PaymentMethodProfiles[0].MethodID = cust.MethodID;
                    sapcust.custObj.PaymentMethodProfiles[0].Routing = cust.routingNumber;
                    sapcust.custObj.PaymentMethodProfiles[0].Account = cust.checkingAccount;
                }
            }
            if(refNum != "")
            {
                List<CCTRAN> list = findTrans(string.Format(" where \"U_customerID\" = '{0}' and \"U_refNum\" = '{1}' and \"U_command\"<>'cc:void'", ID, refNum));
                foreach(CCTRAN tran in list)
                {
                  //  trace("Found CCTRANS refNum = " + refNum);
                    if (sapcust.cccust == null)
                    {
                        sapcust.cccust = new CCCUST();
                        sapcust.cccust.CustNum = tran.custNum;
                    }
                    if(tran.CCAccountID != null)
                        sapcust.cccust.CCAccountID = tran.CCAccountID.ToString();
                    sapcust.custObj = new  Customer();
                    sapcust.custObj.CustomerToken = sapcust.cccust.CustNum;
                    sapcust.custObj.BillingAddress = new Address();
                    sapcust.custObj.Email = "";
                    sapcust.custObj.BillingAddress.FirstName = "";
                    sapcust.custObj.BillingAddress.LastName = "";
                    sapcust.custObj.BillingAddress.State = "";
                    sapcust.custObj.BillingAddress.City = "";
                    sapcust.custObj.BillingAddress.ZipCode = "";
                    sapcust.custObj.BillingAddress.Address1 = "";
                    sapcust.custObj.PaymentMethodProfiles = new PaymentMethodProfile[1]; 
                    sapcust.custObj.PaymentMethodProfiles[0] = new PaymentMethodProfile(); 
                    sapcust.custObj.PaymentMethodProfiles[0].CardNumber = tran.crCardNum;
                    sapcust.custObj.PaymentMethodProfiles[0].CardType = "";
                    sapcust.custObj.PaymentMethodProfiles[0].CardExpiration = "";
                    sapcust.custObj.PaymentMethodProfiles[0].CardCode = "";
                    sapcust.custObj.PaymentMethodProfiles[0].MethodID = tran.MethodID;
                    sapcust.custObj.PaymentMethodProfiles[0].Routing = "";
                    sapcust.custObj.PaymentMethodProfiles[0].Account = "";
                    sapcust.custObj.PaymentMethodProfiles[0].MethodType = "cc";
                   // trace("Found CCTRANS refNum = " + refNum + " return OK" );
                    return sapcust;
                } 
            }
           
        }
        catch(Exception ex)
        {
            errorLog(ex);
        }
        return sapcust;
    }
    public List<SAPCust> getCustomerListByID(string ID)
    {
        // trace("getCustomerByID ID=" + ID + ",refNum=" + refNum);
        List<SAPCust> list = new List<SAPCust>();
      
        try
        {
                List<CCCUST> q = findCCCust(string.Format(" where \"U_CustomerID\"='{0}' AND \"U_CustNum\" <> '' and \"U_checkingAccount\" = '' Order By \"U_default\" desc , \"DocEntry\" desc ", ID));
                foreach (CCCUST cust in q)
                {
                    SAPCust sapcust = new SAPCust();
                    sapcust.cccust = new CCCUST();
                    sapcust.custObj = new Customer();
                    // trace("Found CCust by ID = " + ID);
                    sapcust.cccust = cust;
                    sapcust.key = cust.methodDescription;
                    sapcust.custObj = new Customer();
                    sapcust.custObj.CustomerToken = cust.CustNum;
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
                    sapcust.custObj.PaymentMethodProfiles[0].MethodID = cust.MethodID;
                    sapcust.custObj.PaymentMethodProfiles[0].Routing = cust.routingNumber;
                    sapcust.custObj.PaymentMethodProfiles[0].Account = cust.checkingAccount;
                    if (cust.checkingAccount != null)
                        sapcust.custObj.PaymentMethodProfiles[0].MethodType = "check";
                    else
                        sapcust.custObj.PaymentMethodProfiles[0].MethodType = "cc";
                    list.Add(sapcust);
                }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return list;
    }
    public int batch_createInvoiceFromSO(SAPbobsCOM.Documents oRefDoc, DateTime invDate)
    {
        try
        {
            int nInv = getInvDocEntryByOrder(oRefDoc.CardCode, oRefDoc.DocEntry);
            if (nInv != 0)
            {
                trace(string.Format("Batch invoicing Retrived Invoice {0} from delivery {1}", nInv, oRefDoc.DocEntry));
                return nInv;
            }
            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            oDoc.DocNum = getNextTableNum("OINV");
            oDoc.BPL_IDAssignedToInvoice = oRefDoc.BPL_IDAssignedToInvoice;
            showStatus(string.Format("Batch invoicing Create Invoice Brach ID = {0} from order entry {1}", oRefDoc.BPL_IDAssignedToInvoice, oRefDoc.DocEntry), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

            oDoc.CardCode = oRefDoc.CardCode;
            oDoc.CardName = oRefDoc.CardName;
            oDoc.DocDueDate = invDate.AddDays(5);
            oDoc.DocDate = invDate;

            oDoc.Comments = "Created by eBizCharge Batch Invoicing base on sales Order Number: " + oRefDoc.DocNum;
            //oDoc.DeferredTax = oSODoc.DeferredTax;
            oDoc.DiscountPercent = oRefDoc.DiscountPercent;
            oDoc.SalesPersonCode = oRefDoc.SalesPersonCode;
            oDoc.Rounding = oRefDoc.Rounding;
            oDoc.RoundingDiffAmount = oRefDoc.RoundingDiffAmount;
            SAPbobsCOM.DocumentsAdditionalExpenses expLines = oRefDoc.Expenses;

            for (int n = 0; n < expLines.Count; n++)
            {
                expLines.SetCurrentLine(n);
                if (expLines.ExpenseCode != 0)
                {
                    oDoc.Expenses.ExpenseCode = expLines.ExpenseCode;
                    oDoc.Expenses.TaxCode = expLines.TaxCode;
                    oDoc.Expenses.LineTotal = expLines.LineTotal;
                    if (expLines.LineTotalFC != 0)
                        oDoc.Expenses.LineTotal = expLines.LineTotalFC;
                    oDoc.Expenses.Add();
                }
            }
            SAPbobsCOM.Document_Lines olines = oRefDoc.Lines;
            double total = 0;
            oDoc.DocType = oRefDoc.DocType;
            for (int i = 0; i < olines.Count; i++)
            {
                olines.SetCurrentLine(i);
                if (oRefDoc.DocType == SAPbobsCOM.BoDocumentTypes.dDocument_Items)
                {

                    oDoc.Lines.ItemCode = olines.ItemCode;
                    oDoc.Lines.SerialNum = olines.SerialNum;
                    oDoc.Lines.Quantity = olines.Quantity;
                    oDoc.Lines.TaxCode = olines.TaxCode;
                    oDoc.Lines.Price = olines.Price;
                    oDoc.Lines.DiscountPercent = olines.DiscountPercent;

                }
                else
                {
                    oDoc.Lines.AccountCode = olines.AccountCode;

                }
                SAPbobsCOM.Document_LinesAdditionalExpenses exp = olines.Expenses;

                for (int n = 0; n < exp.Count; n++)
                {
                    exp.SetCurrentLine(n);
                    if (exp.ExpenseCode != 0)
                    {
                        oDoc.Lines.Expenses.ExpenseCode = exp.ExpenseCode;
                        oDoc.Lines.Expenses.TaxCode = exp.TaxCode;
                        oDoc.Lines.Expenses.LineTotal = exp.LineTotal;
                        if (exp.LineTotalFC != 0)
                            oDoc.Lines.Expenses.LineTotal = exp.LineTotalFC;
                        oDoc.Lines.Expenses.Add();
                    }
                }

                oDoc.Lines.Currency = olines.Currency;
                oDoc.Lines.BaseEntry = oRefDoc.DocEntry;
                oDoc.Lines.BaseType = 17;
                oDoc.Lines.BaseLine = olines.LineNum;
                oDoc.Lines.ItemDescription = olines.ItemDescription;
                oDoc.Lines.TaxCode = olines.TaxCode;
                oDoc.Lines.LineTotal = olines.LineTotal;
                oDoc.Lines.Add();
                total += olines.LineTotal;
            }
            int r = oDoc.Add();
            if (r != 0)
            {
                string err = "Batch Invoicing create invoice failed. " + getErrorString();
                errorLog(err);
                return 0;
            }
            nInv = getInvDocEntryByOrder(oRefDoc.CardCode, oRefDoc.DocEntry);
            trace("Batch Invoicing create Invoice: " + nInv + " On Order: " + oRefDoc.DocEntry);
            /*
            if (SBO_Application.MessageBox("Batch Invoicing create Invoice: " + nInv + ", continue?", 1, "Yes", "No") != 1)
            {
                return 0;
            }
             * */
            if (!oDoc.GetByKey(nInv))
            {
                string err = "Batch Invoicing create invoice failed. n = " + nInv.ToString() + ", OrderID = " + oRefDoc.DocEntry.ToString();
                showMessage(err);
                return 0;
            }
            if (oDoc.DocTotal != oRefDoc.DocTotal)
            {
                string err = "Batch Invoicing create invoice failed. Invoice DocTotal = " + oDoc.DocTotal.ToString() + ", Delivery Total = " + oRefDoc.DocTotal.ToString();
                showMessage(err);
                return 0;
            }
            trace("Batch Invoicing create Invoice: " + nInv + ", total:" + oDoc.DocTotal.ToString());
            /*
            if (SBO_Application.MessageBox("Batch Invoicing create Invoice: " + nInv + ", total:" + oDoc.DocTotal.ToString() + ", continue?", 1, "Yes", "No") != 1)
            {
                return 0;
            }
             */
            return nInv;
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }
        return 0;
    }
    public int batch_createInvoice(SAPbobsCOM.Documents oRefDoc, DateTime invDate)
    {
        try
        {
            int nInv = getDocEntryByDelivery(oRefDoc.CardCode, oRefDoc.DocEntry);
            if (nInv != 0)
            {
                trace(string.Format("Batch invoicing Retrived Invoice {0} from delivery {1}", nInv, oRefDoc.DocEntry));
                return nInv;
            }
            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            oDoc.DocNum = getNextTableNum("OINV");
            oDoc.BPL_IDAssignedToInvoice = oRefDoc.BPL_IDAssignedToInvoice;
            trace(string.Format("Batch invoicing Create Invoice Brach ID = {0} from delivery {1}", oRefDoc.BPL_IDAssignedToInvoice, oRefDoc.DocEntry));

            oDoc.CardCode = oRefDoc.CardCode;
            oDoc.CardName = oRefDoc.CardName;
            oDoc.DocDueDate = invDate.AddDays(5);
            oDoc.DocDate = invDate;

            oDoc.Comments = "Created by eBizCharge Batch Invoicing base on Delivery Note Number: " + oRefDoc.DocNum;
            //oDoc.DeferredTax = oSODoc.DeferredTax;
            oDoc.DiscountPercent = oRefDoc.DiscountPercent;
            oDoc.SalesPersonCode = oRefDoc.SalesPersonCode;
            oDoc.Rounding = oRefDoc.Rounding;
            oDoc.RoundingDiffAmount = oRefDoc.RoundingDiffAmount;
            SAPbobsCOM.DocumentsAdditionalExpenses expLines = oRefDoc.Expenses;

            for (int n = 0; n < expLines.Count; n++)
            {
                expLines.SetCurrentLine(n);
                if (expLines.ExpenseCode != 0)
                {
                    oDoc.Expenses.ExpenseCode = expLines.ExpenseCode;
                    oDoc.Expenses.TaxCode = expLines.TaxCode;
                    oDoc.Expenses.LineTotal = expLines.LineTotal;
                    if (expLines.LineTotalFC != 0)
                        oDoc.Expenses.LineTotal = expLines.LineTotalFC;
                    oDoc.Expenses.Add();
                }
            }
            SAPbobsCOM.Document_Lines olines = oRefDoc.Lines;
            double total = 0;
            oDoc.DocType = oRefDoc.DocType;
            for (int i = 0; i < olines.Count; i++)
            {
                olines.SetCurrentLine(i);
                if (oRefDoc.DocType == SAPbobsCOM.BoDocumentTypes.dDocument_Items)
                {

                    oDoc.Lines.ItemCode = olines.ItemCode;
                    oDoc.Lines.SerialNum = olines.SerialNum;
                    oDoc.Lines.Quantity = olines.Quantity;
                    oDoc.Lines.TaxCode = olines.TaxCode;
                    oDoc.Lines.Price = olines.Price;
                    oDoc.Lines.DiscountPercent = olines.DiscountPercent;

                }
                else
                {
                    oDoc.Lines.AccountCode = olines.AccountCode;

                }
                SAPbobsCOM.Document_LinesAdditionalExpenses exp = olines.Expenses;

                for (int n = 0; n < exp.Count; n++)
                {
                    exp.SetCurrentLine(n);
                    if (exp.ExpenseCode != 0)
                    {
                        oDoc.Lines.Expenses.ExpenseCode = exp.ExpenseCode;
                        oDoc.Lines.Expenses.TaxCode = exp.TaxCode;
                        oDoc.Lines.Expenses.LineTotal = exp.LineTotal;
                        if (exp.LineTotalFC != 0)
                            oDoc.Lines.Expenses.LineTotal = exp.LineTotalFC;
                        oDoc.Lines.Expenses.Add();
                    }
                }

                oDoc.Lines.Currency = olines.Currency;
                oDoc.Lines.BaseEntry = oRefDoc.DocEntry;
                oDoc.Lines.BaseType = 15;
                oDoc.Lines.BaseLine = olines.LineNum;
                oDoc.Lines.ItemDescription = olines.ItemDescription;
                oDoc.Lines.TaxCode = olines.TaxCode;
                oDoc.Lines.LineTotal = olines.LineTotal;
                oDoc.Lines.Add();
                total += olines.LineTotal;
            }
            int r = oDoc.Add();
            if (r != 0)
            {
                string err = "Batch Invoicing create invoice failed. " + getErrorString();
                errorLog(err);
                return 0;
            }
            nInv = getDocEntryByDelivery(oRefDoc.CardCode, oRefDoc.DocEntry);
            trace("Batch Invoicing create Invoice: " + nInv);
            /*
            if (SBO_Application.MessageBox("Batch Invoicing create Invoice: " + nInv + ", continue?", 1, "Yes", "No") != 1)
            {
                return 0;
            }
             * */
            if (!oDoc.GetByKey(nInv))
            {
                string err = "Batch Invoicing create invoice failed. n = " + nInv.ToString() + ", DeliveryID = " + oRefDoc.DocEntry.ToString();
                showMessage(err);
                return 0;
            }
            if (oDoc.DocTotal != oRefDoc.DocTotal)
            {
                string err = "Batch Invoicing create invoice failed. Invoice DocTotal = " + oDoc.DocTotal.ToString() + ", Delivery Total = " + oRefDoc.DocTotal.ToString();
                showMessage(err);
                return 0;
            }
            trace("Batch Invoicing create Invoice: " + nInv + ", total:" + oDoc.DocTotal.ToString());
            /*
            if (SBO_Application.MessageBox("Batch Invoicing create Invoice: " + nInv + ", total:" + oDoc.DocTotal.ToString() + ", continue?", 1, "Yes", "No") != 1)
            {
                return 0;
            }
             */
            return nInv;
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }
        return 0;

    }
    private void selectSameDelivery(string dID)
    {
        try
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oBatchInvForm.Items.Item(matrixBatchInv).Specific;
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                string DeliveryID = getMatrixItem(oMatrix, "DeliveryID", i);

                SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("sel").Cells.Item(i).Specific;
                if (!oCheckBox.Checked && DeliveryID == dID)
                {
                    oCheckBox.Checked = true;
                }
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private bool hasNextPreauth(string dID, string refNum, int nStart)
    {
        try
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oBatchInvForm.Items.Item(matrixBatchInv).Specific;
            for (int i = nStart; i <= oMatrix.RowCount; i++)
            {
                string DeliveryID = getMatrixItem(oMatrix, "DeliveryID", i);
                string nrefNum = getMatrixItem(oMatrix, "RefNum", i);

                if (refNum != nrefNum && DeliveryID == dID)
                {
                    trace("Has next preauth refNum: " + refNum + ", deliveryID = " + dID);
                    return true;
                }
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    public void captureOnAdd()
    {
        try
        {
            Thread.Sleep(3000);
            captureService("nodelivery");
            captureService();

        }
        catch(Exception ex)
        {
            errorLog(ex);
        }
       
    }
    public void captureService(string opt = "")
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();
        SAPbobsCOM.Documents oInvDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);

        try
        {
            if (!oCompany.Connected)
            {
                trace("Restart connection.");
                if (!connectToDB())
                    return;
            }
            // trace("Capture Service Started.  opt=" + opt);
            clearInvoiceIDForPreAuth();
            setDeliveryID();
            setInvoiceID();
            List<TransForCapture> list = getTransForCapture(opt);
            foreach (TransForCapture t in list)
            {
                if (!oInvDoc.GetByKey(int.Parse(t.InvoiceID)))
                {
                    trace("Invoice: " + t.InvoiceID + " not found");
                    setCaptureResult(t.RefNum, "cc:authonly-noinv");
                    continue;
                }
                if (oInvDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Close)
                {
                    trace("Invoice: " + oInvDoc.DocNum + " is closed");
                    setCaptureResult(t.RefNum, "cc:authonly-closed");
                    continue;
                }
                if (t == null)
                {
                    trace("Transaction not found");
                    continue;
                }
                double bal = getInvoiceBalance(t.InvoiceID);
                if (bal == 0)
                {
                    trace("Invoice: " + t.InvoiceID + " balance is 0");
                    setCaptureResult(t.RefNum, "cc:authonly-0bal");
                    continue;
                }
                string camount = getDoubleValue(t.Amount) > bal ? bal.ToString() : t.Amount;
                trace(string.Format("Capture service processing Invoice {0}, Amount {1}, refNum {2}, bal {3}", t.InvoiceID, t.Amount, t.RefNum, bal));
                SAPCust sapcust = getCustomerByID(t.CustomerID, t.RefNum);
                sapcust.cccust.CCAccountID = t.U_CCAccountID;
                if (t.RefNum == "" || t.RefNum == null)
                {

                    trace(string.Format("Capture service refNum not found.  Invoice {0}, Amount {1}, refNum {2}, bal {3}", t.InvoiceID, t.Amount, t.RefNum, bal));


                }
                else
                {
                    oInvDoc.GetByKey(int.Parse(t.InvoiceID));
                    if (oInvDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Close)
                    {
                        trace("Invoice: " + oInvDoc.DocNum + " is closed");
                        setCaptureResult(t.RefNum, "cc:authonly-closed");
                        continue;
                    }

                    if (CaptureTrans(t.RefNum, sapcust, camount, oInvDoc))
                    {

                        AddPayment(getDoubleValue(camount), sapcust, t.InvoiceID.ToString());
                        BatchSetTransInvoiceID(t.RefNum, t.InvoiceID, t.OrderID, sapcust.cccust.CCAccountID, camount);
                        setCaptureResult(t.RefNum, "cc:authonly-captured");
                    }
                    else
                    {
                        setCaptureResult(t.RefNum, "cc:authonly-error");
                        string msg = string.Format("Failed to capture invoice {0} of {1}", t.InvoiceID, t.Amount);
                        Console.Write(msg);
                        errorLog(msg);

                    }
                }

            }
            if(oInvDoc == null)
                oInvDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            list = getTransForErrorCheck();
           // trace(string.Format("TransForErrorCheck list count = {0}", list.Count));

            foreach (TransForCapture t in list)
            {
                //trace(string.Format("TransForErrorCheck t = {0}/{1}/{2}", t.CustomerID, t.InvoiceID, t.RefNum));
                if (!oInvDoc.GetByKey(int.Parse(t.InvoiceID)))
                {
                    trace("TransForErrorCheck Invoice: " + t.InvoiceID + " not found");
                    setCaptureResult(t.RefNum, "cc:authonly-noinv");
                    continue;
                }
                //trace(string.Format("TransForErrorCheck oDoc Staus = {0}", oInvDoc.DocumentStatus));

                if (oInvDoc.DocumentStatus == SAPbobsCOM.BoStatus.bost_Close)
                {
                    trace("TransForErrorCheck Invoice: " + oInvDoc.DocNum + " is closed");
                    setCaptureResult(t.RefNum, "cc:authonly-closed");
                    continue;
                }
               
                double bal = getInvoiceBalance(t.InvoiceID);
                if (bal == 0)
                {
                    trace("TransForErrorCheck Invoice: " + t.InvoiceID + " balance is 0");
                    setCaptureResult(t.RefNum, "cc:authonly-0bal");
                    continue;
                }
                //trace(string.Format("TransForErrorCheck bal={0}", bal));
                try
                {
                    SAPCust sapcust = getCustomerByID(t.CustomerID, t.RefNum);
                    if (sapcust != null)
                    {
                        if (sapcust.cccust != null)
                        {
                          //  trace(string.Format("TransForErrorCheck accID={0}", sapcust.cccust.CCAccountID));
                            SecurityToken token = getToken(sapcust.cccust.CCAccountID);
                            TransactionObject tObj = ebiz.GetTransactionDetails(token, t.RefNum);
                            if (tObj != null)
                            {
                                trace(string.Format("TransForErrorCheck Transaction Details {0},{1},{2},{3},{4}",
                                    tObj.Response.Error, tObj.Response.Result, tObj.Status, tObj.TransactionType, tObj.Response.AuthAmount));
                                if (tObj.TransactionType.ToLower() == "sale" && tObj.Response.Result.ToLower() == "approved")
                                {
                                    AddPayment(tObj.Response.AuthAmount, sapcust, t.InvoiceID.ToString());
                                    setCaptureResult(t.RefNum, "cc:authonly-captured");
                                }
                                else if (tObj.TransactionType.ToLower() == "auth only" && tObj.Response.Result.ToLower() == "approved")
                                {
                                    trace(string.Format("TransErrorCheck t = {0}/{1}/{2}  restored to authonly", t.CustomerID, t.InvoiceID, t.RefNum));
                                    setErrorCheckResult(t.RefNum, "cc:authonly");
                                }

                            }
                        }
                    }
                }catch(Exception ex)
                {
                    //errorLog(ex);
                    trace(string.Format("TransErrorCheck t = {0}/{1}/{2}  transnotfound {3}", t.CustomerID, t.InvoiceID, t.RefNum, ex.Message));
                    setErrorCheckResult(t.RefNum, "transnotfound");
                }
            }
        }
        catch (Exception ex)
        {
            trace("Capture service exception: " + ex.Message);
            if (bRunJob)
                System.Environment.Exit(0);
        }
        finally
        {
            try
            {

                transReset();
                if (oInvDoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvDoc);
                    oInvDoc = null;
                }
                watch.Stop();
                // trace("Capture Service End. Time elapsed: " + watch.ElapsedMilliseconds);

            }
            catch (Exception) { }
        }
    }
    public List<TransForCapture> getTransForErrorCheck()
    {
        List<TransForCapture> list = new List<TransForCapture>();
        string sql = "select \"U_refNum\", \"U_InvoiceID\", \"U_DeliveryID\", \"U_OrderID\", \"U_amount\", \"U_customerID\",\"U_CCAccountID\" from \"@CCTRANS\" where \"U_command\" in ('cc:authonly-error') and \"U_InvoiceID\" <> '' and not exists (select 1 from OINV where Cast(\"DocEntry\" as varchar) = \"U_InvoiceID\" and \"DocStatus\" = 'C')";
        //trace(sql);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            //trace(sql);
             oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                TransForCapture t = new TransForCapture();
                t.RefNum = (string)oRS.Fields.Item(0).Value;
                t.InvoiceID = (string)oRS.Fields.Item(1).Value;
                t.DeliveryID = (string)oRS.Fields.Item(2).Value;
                t.OrderID = (string)oRS.Fields.Item(3).Value;
                t.Amount = (string)oRS.Fields.Item(4).Value;
                t.CustomerID = (string)oRS.Fields.Item(5).Value;
                t.U_CCAccountID = (string)oRS.Fields.Item(6).Value;
                list.Add(t);
                oRS.MoveNext();
            }
           
            
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }

        return list;
    }
    public List<TransForCapture> getTransForCapture(string opt = "")
    {
        List<TransForCapture> list = new List<TransForCapture>();
        string sql = string.Format("select \"U_refNum\", \"U_InvoiceID\", \"U_DeliveryID\", \"U_OrderID\", \"U_amount\", \"U_customerID\",\"U_CCAccountID\" from \"@CCTRANS\" where \"U_command\" in ('cc:authonly','cc:authonly-remain') and \"U_InvoiceID\" <> '' and \"U_DeliveryID\" is not null and not exists (select 1 from OINV where Cast(\"DocEntry\" as varchar) = \"U_InvoiceID\" and \"DocStatus\" = 'C')");
        if (opt == "nodelivery")
            sql = "select \"U_refNum\", \"U_InvoiceID\", \"U_DeliveryID\", \"U_OrderID\", \"U_amount\", \"U_customerID\",\"U_CCAccountID\" from \"@CCTRANS\" where \"U_command\" in ('cc:authonly','cc:authonly-remain') and \"U_InvoiceID\" <> '' and not exists (select 1 from OINV where Cast(\"DocEntry\" as varchar) = \"U_InvoiceID\" and \"DocStatus\" = 'C')";
        string invList = "('NotUsed'";
        try
        {
            //trace(sql);
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);
            while (!oRS.EoF)
            {
                TransForCapture t = new TransForCapture();
                t.RefNum = (string)oRS.Fields.Item(0).Value;
                t.InvoiceID = (string)oRS.Fields.Item(1).Value;
                t.DeliveryID = (string)oRS.Fields.Item(2).Value;
                t.OrderID = (string)oRS.Fields.Item(3).Value;
                t.Amount = (string)oRS.Fields.Item(4).Value;
                t.CustomerID = (string)oRS.Fields.Item(5).Value;
                t.U_CCAccountID = (string)oRS.Fields.Item(6).Value;
                invList = invList + ",'" + t.InvoiceID + "'";
                list.Add(t);
                oRS.MoveNext();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
            invList = invList + ")";
            invList = invList.Replace("'NotUsed',", "");
            if (list.Count() > 0)
            {
                sql = string.Format("UPDATE \"@CCTRANS\" set  \"U_command\"='cc:processing' where \"U_command\"='cc:authonly' and  \"U_InvoiceID\" in {0}", invList);
               
                oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery(sql);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                oRS = null;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }

        return list;
    }
    public void transReset()
    {
         string sql = string.Format("UPDATE \"@CCTRANS\" set  \"U_command\"='cc:authonly' where \"U_command\"='cc:processing'");
        try
        {
            execute(sql);
          
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            if (bRunJob)
            {
                trace("Encounter db error. Program Existing...");
                System.Environment.Exit(0);
            }
        }

       
    }
}
public class TransForCapture
{
    public string InvoiceID { set; get; }
    public string DeliveryID { set; get; }
    public string OrderID { set; get; }
    public string CustomerID { set; get; }
    public string RefNum { set; get; }
    public string Amount { set; get; }
    public string U_CCAccountID { set; get; }

}



