using eBizChargeForSAP;
using eBizChargeForSAP.Screen;
using eBizChargeForSAP.ServiceReference1;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

partial class SAP
{


    #region DataMember
    String TesingFilter = "";
    int DownloadRowCount = 1000;
    bool flgItemDfolder = true, flgBPDfolder = true, flgSODfolder = true, flgSOPfolder = true;
    SAPbouiCOM.Form FormSyncDownload;
    string MtxSyncSOFrmCon = "mtMain";
    String SyncDownloadTabBP = "TabBp", MtxSyncDBpListStr = "TbpL", MtxSyncDBpLogStr = "TbpLog";
    String MtxSyncDownloadBPListStr = "mtxBPLst", MtxSyncDownloadBPLogStr = "mtxBPLog";
    CustomerListSearchResult ListBp = null;
    ItemDetails[] OutputListItem = null;
    SalesOrder[] OutputListSO = null;
    ApplicationTransactionSearchResult TransactionList = null;

    String SyncDownloadTabItem = "TabItm", MtxSyncDItemListStr = "TltmList", MtxSyncDItemLogStr = "TItmLog";
    String MtxSyncDownloadItemListStr = "mtxItmLst", MtxSyncDownloadItemLogStr = "mtxItmLog";
    String MtxSyncDownloadSOListStr = "MtxSOList", MtxSyncDownloadSOLogStr = "mtxItmLog";

    string dtSyncSOFrmCon = "dtemp";
    SAPbouiCOM.DataTable odtSyncBP, odtSyncBPLog, odtSyncItem, odtSyncItemLog, odtSyncSOFrmCon, odtSyncSOFrmConLog, odtSyncSOPFrmCon, odtSyncSOPFrmConLog;
    bool flgDocInitSyncSOFrmCon = false;
    SAPbouiCOM.Matrix OMtxSyncSOFrmCon;
    SAPbouiCOM.Matrix OMtxSyncBPList, OMtxSyncBPLog, OMtxSyncItemList, OMtxSyncItemLog, OMtxSyncSOList, OMtxSyncSOLog, OMtxSyncSOPList;

    bool FlgdtSyncSOFrmCon = true;
    SAPbouiCOM.EditText ed_fromDateSyncSO, ed_ToDateSyncSO;
    SAPbouiCOM.ComboBox Cmb_MenuSyncSO;
    string txtFromdtSyncSO = "txtFromdt", txtTodtSyncSO = "txtTodt", cmbMenuStrSyncSO = "cmbMenu";
    SalesOrder[] ListSalesOrder = null;


    SAPbobsCOM.Documents DocSalesOrder, SAPDoc;
    SAPbobsCOM.Document_Lines DocSalesOrderLines;
    #endregion

    #region FormEvent
    private void CreateSyncSO()
    {
        try
        {
            flgDocInitSyncSOFrmCon = false;
            AddXMLForm("Presentation_Layer.SyncDownload.xml");
            FormSyncDownload.Freeze(true);
            SyncSOFrmConInit();
            FormSyncDownload.Freeze(false);
            FlgdtSyncSOFrmCon = true;
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }
    private void SyncDownloadFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            FormSyncDownload = form;
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:

                        break;
                }
            }
            else
            {
                switch (pVal.EventType)
                {

                    case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
                        currentForm = form;
                        //tProcess = new Thread(BPMoveThreadProc);
                        //tProcess.Start();
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:

                        try
                        {
                            formSelTarget = form;

                            //form.Items.Item(tbeBizCharge).Visible = false;
                            //   bcontrolCreated = false;

                        }
                        catch (Exception ex)
                        {
                            errorLog(ex);
                        }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        try
                        {
                            //Reload DATA
                            //if (!bcontrolCreated)
                            //    return;
                            //if (pVal.ItemChanged == true && pVal.ItemUID == editCardNum)
                            //{
                            //    setCardName(form, selSAPCust);
                            //}
                        }
                        catch (Exception)
                        { }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        {

                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == "btnFind")
                        {
                            if (VAlidDateSyncSO())
                            {
                                // ShowDataaSyncSO();
                            }

                        }
                        if (pVal.ItemUID == "btnPost")
                        {
                            HandleItemSyncSO(form);
                        }
                        if (pVal.ItemUID == MtxSyncSOFrmCon && pVal.ColUID.Equals("clSel") && pVal.Row == 0 && !pVal.BeforeAction)
                        {
                            HandleSyncSOSelectALL(form, MtxSyncSOFrmCon, odtSyncSOFrmCon);
                        }
                        if (pVal.ItemUID == MtxSyncSOFrmCon)
                        {
                            if (form.Items.Item(MtxSyncSOFrmCon) == null)
                                return;

                            OMtxSyncSOFrmCon.FlushToDataSource();
                        }
                        #region Download Button

                        //Business Partner
                        if (pVal.ItemUID.Equals("btnDBp"))
                        {
                            new Thread(SyncDownloadBPProc).Start();
                        }
                        //Download Item Button 
                        if (pVal.ItemUID.Equals("btnDitm"))
                        {
                            new Thread(SyncDownloadItemProc).Start();
                        }//
                        if (pVal.ItemUID.Equals("btnDSo"))
                        {
                            new Thread(SyncDownloadSOProc).Start();
                        }
                        if (pVal.ItemUID.Equals("btnDSop"))
                        {
                            new Thread(SyncDownloadSOPProc).Start();
                        }
                        #endregion
                        #region Tab Functionality

                        //BP Tab
                        if (pVal.ItemUID.Equals("TbpLog"))
                        {
                            ShowSyncDownloadBPLOG(form);
                        }
                        if (pVal.ItemUID.Equals("TabBp") || pVal.ItemUID.Equals("TbpL"))//Item List Tab
                        {
                            if (flgBPDfolder)
                            {
                                flgBPDfolder = false;
                                SetFolder(FormSyncDownload, "TabBp");
                                SetFolder(FormSyncDownload, "TbpL");
                            }
                        }
                        //Item TaB
                        if (pVal.ItemUID.Equals("TItmLog"))
                        {
                            ShowSyncDownloadItemLOG(form);
                        }
                        if (pVal.ItemUID.Equals("TabItm") || pVal.ItemUID.Equals("TltmList"))
                        {
                            if (flgItemDfolder)
                            {
                                flgItemDfolder = false;
                                SetFolder(FormSyncDownload, "TabItm");
                                SetFolder(FormSyncDownload, "TltmList");
                            }
                        }
                        //SO Tab
                        if (pVal.ItemUID.Equals("TbSoLg"))
                        {
                            ShowSyncDownloadSOLOG(form);
                        }
                        if (pVal.ItemUID.Equals("TabSO") || pVal.ItemUID.Equals("TbSoL"))
                        {
                            if (flgSODfolder)
                            {
                                flgSODfolder = false;
                                SetFolder(FormSyncDownload, "TabSO");
                                SetFolder(FormSyncDownload, "TbSoL");
                            }
                        }
                        //Sales Order Payment Tab
                        if (pVal.ItemUID.Equals("TbSoPLog"))
                        {
                            ShowSyncDownloadSOPLOG(form);
                        }
                        if (pVal.ItemUID.Equals("TabSOP") || pVal.ItemUID.Equals("TabSopLst"))
                        {
                            if (flgSOPfolder)
                            {
                                flgSOPfolder = false;
                                SetFolder(FormSyncDownload, "TabSOP");
                                SetFolder(FormSyncDownload, "TabSopLst");
                            }
                        }
                        #endregion
                        #region Select All
                        if (pVal.ItemUID.Equals("chkBpL1"))
                        {
                            SelectAllCheckBox(form, "mtxBPLst", "chkBpL1");
                        }
                        if (pVal.ItemUID.Equals("chkItmL1"))
                        {
                            SelectAllCheckBox(form, "mtxItmLst", "chkItmL1");
                        }
                        if (pVal.ItemUID.Equals("chkSoL1"))
                        {
                            SelectAllCheckBox(form, "MtxSOList", "chkSoL1");
                        }
                        if (pVal.ItemUID.Equals("chkSopL1"))
                        {
                            SelectAllCheckBox(form, "mtxSOP", "chkSopL1");
                        }
                        #endregion
                        #region Import Functionality
                        //Business Partner
                        if (pVal.ItemUID.Equals("btnBpImp"))
                        {
                            if (SBO_Application.MessageBox("Are you sure you want to import into SAP B1?", 1, "Yes", "No") == 1)
                            {
                                DownloadDataBP("mtxBPLst", form);
                            }
                        }
                        if (pVal.ItemUID.Equals("btnImpSap"))
                        {
                            if (SBO_Application.MessageBox("Are you sure you want to import into SAP B1?", 1, "Yes", "No") == 1)
                            {
                                DownloadDataItem("mtxItmLst", form);
                            }
                        }
                        if (pVal.ItemUID.Equals("btnSOSap"))
                        {
                            if (SBO_Application.MessageBox("Are you sure you want to import into SAP B1?", 1, "Yes", "No") == 1)
                            {
                                DownloadDataSO("MtxSOList", form);
                            }
                        }
                        if (pVal.ItemUID.Equals("btnSOPSap"))
                        {
                            if (SBO_Application.MessageBox("Are you sure you want to import into SAP B1?", 1, "Yes", "No") == 1)
                            {
                                DownloadDataSOP("mtxSOP", form);
                            }
                        }
                        #endregion
                        #region Export Log
                        if (pVal.ItemUID.Equals("btnExp11") || pVal.ItemUID.Equals("btnExp12") || pVal.ItemUID.Equals("btnExp21") || pVal.ItemUID.Equals("btnExp22"))
                        {
                            FileBrowserSyncDown(pVal.ItemUID, form);
                        }
                        if (pVal.ItemUID.Equals("btnExp31") || pVal.ItemUID.Equals("btnExp32") || pVal.ItemUID.Equals("btnExp41") || pVal.ItemUID.Equals("btnExp42"))
                        {
                            FileBrowserSyncDown(pVal.ItemUID, form);
                        }
                        #endregion
                        #region Clear Log
                        //btnclr1,btnclr2,btnclr3,btnclr4
                        if (pVal.ItemUID.Equals("btnclr1"))
                        {
                            if (SBO_Application.MessageBox("Are you sure you want to Clear Log?", 1, "Yes", "No") == 1)
                            {
                                ClearLog("CCBPIMP");
                            }
                        }
                        if (pVal.ItemUID.Equals("btnclr2"))
                        {
                            if (SBO_Application.MessageBox("Are you sure you want to Clear Log?", 1, "Yes", "No") == 1)
                            {
                                ClearLog("CCITMIMP");
                            }
                        }
                        if (pVal.ItemUID.Equals("btnclr3"))
                        {
                            if (SBO_Application.MessageBox("Are you sure you want to Clear Log?", 1, "Yes", "No") == 1)
                            {
                                ClearLog("CCSOIMP");
                            }
                        }
                        if (pVal.ItemUID.Equals("btnclr4"))
                        {
                            if (SBO_Application.MessageBox("Are you sure you want to Clear Log?", 1, "Yes", "No") == 1)
                            {
                                ClearLog("CCSOPIMP");
                            }
                        }
                        #endregion
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        // handleItemPress(form, pVal);
                        break;
                }

            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }

    private void HandleSyncSOSelectALL(SAPbouiCOM.Form form, String MatxStr, SAPbouiCOM.DataTable oDataTable)
    {
        bool flgChange = false;
        try
        {
            if (form.Items.Item(MatxStr) == null)
                return;

            form.Items.Item(MatxStr).AffectsFormMode = false;
            SAPbouiCOM.Matrix OMatrix = (SAPbouiCOM.Matrix)form.Items.Item(MatxStr).Specific;
            int RowCount = oDataTable.Rows.Count;
            form.Freeze(true);
            if (RowCount > 0)
            {  //for (int i = oDataTable.Rows.Count - 1; i >= 0; i--)
                ShowMsgCustom("Selecting ...!");
                for (int i = 0; i < oDataTable.Rows.Count; i++)
                {
                    var chk = (SAPbouiCOM.CheckBox)OMatrix.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    if (FlgdtSyncSOFrmCon)
                    {
                        if (chk.Checked)
                        {
                        }
                        else
                        {
                            flgChange = true;
                            OMatrix.Columns.Item("clSel").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                    }
                    else
                    {
                        if (chk.Checked)
                        {
                            flgChange = true;
                            OMatrix.Columns.Item("clSel").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else
                        {
                        }
                    }
                }
            }
            form.Freeze(false);
            form.Update();
            if (flgChange)
            {
                FlgdtSyncSOFrmCon = !FlgdtSyncSOFrmCon;
            }

        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }

    private void HandleItemSyncSO(SAPbouiCOM.Form form)
    {
        bool FlgChange = false;
        try
        {
            if (form.Items.Item(MtxSyncSOFrmCon) == null)
                return;

            form.Items.Item(MtxSyncSOFrmCon).AffectsFormMode = false;
            OMtxSyncSOFrmCon = (SAPbouiCOM.Matrix)form.Items.Item(MtxSyncSOFrmCon).Specific;
            int RowCount = odtSyncSOFrmCon.Rows.Count;
            if (RowCount > 0)
            {
                for (int i = odtSyncSOFrmCon.Rows.Count - 1; i >= 0; i--)
                {
                    var EdEbixCharge = GetMatrixValue(OMtxSyncSOFrmCon, i, "clEbizkey").ToString();

                    var chk = (SAPbouiCOM.CheckBox)OMtxSyncSOFrmCon.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    if (chk.Checked)
                    {
                        FlgChange = true;
                        PostSyncSO(EdEbixCharge, ListSalesOrder[i], i, OMtxSyncSOFrmCon, form);
                    }
                    else
                    {

                    }

                }
                if (FlgChange)
                {
                    //   oMatrixBpExport.LoadFromDataSource();
                    OMtxSyncSOFrmCon.FlushToDataSource();
                }
            }
        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }

    #endregion

    #region Helping Function
    private bool VAlidDateSyncSO()
    {
        bool flag = false;
        try
        {
            var FromValue = GetItemValue(txtFromdtSyncSO, FormSyncDownload);
            var ToDateValue = GetItemValue(txtTodtSyncSO, FormSyncDownload);

            if (string.IsNullOrEmpty(FromValue))
            {
                MsgWarning("From date can't be left empty.");
                return false;
            }
            else if (string.IsNullOrEmpty(ToDateValue))
            {
                MsgWarning("To Date can't be left empty.");
                return false;
            }

            else
            {
                return true;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return flag;
    }

    private void SyncSOFrmConInit()
    {
        try
        {
            flgDocInitSyncSOFrmCon = true;
            //  ShowDataaSyncSO();
            //FillDataSyncSOFrmCon();
            ShowDataaSyncSO();

        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
    }
    private void FillDataSyncSOFrmCon()
    {
        try
        {

            ed_fromDateSyncSO = ((SAPbouiCOM.EditText)(FormSyncDownload.Items.Item(txtFromdtSyncSO).Specific));
            FormSyncDownload.DataSources.UserDataSources.Add(txtFromdtSyncSO, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_fromDateSyncSO.DataBind.SetBound(true, "", txtFromdtSyncSO);
            ed_fromDateSyncSO.Value = DatetimeTosapFormatDate(DateTime.Now);
            ed_fromDateSyncSO.Item.Visible = false;

            ed_ToDateSyncSO = ((SAPbouiCOM.EditText)(FormSyncDownload.Items.Item(txtTodtSyncSO).Specific));
            FormSyncDownload.DataSources.UserDataSources.Add(txtTodtSyncSO, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_ToDateSyncSO.DataBind.SetBound(true, "", txtTodtSyncSO);
            ed_ToDateSyncSO.Value = DatetimeTosapFormatDate(DateTime.Now);
            ed_ToDateSyncSO.Item.Visible = false;

            Cmb_MenuSyncSO = ((SAPbouiCOM.ComboBox)(FormSyncDownload.Items.Item(cmbMenuStrSyncSO).Specific));
            Cmb_MenuSyncSO.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    private void ShowDataaSyncSO()
    {
        try
        {
            tProcess = new Thread(SyncSOThreadProc);
            tProcess.Start();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            throw;
        }
    }

    public void SyncSOThreadProc()
    {
        try
        {
            if (formSelTarget == null)
            {
                if (FormSyncDownload != null)
                    formSelTarget = FormSyncDownload;
                else
                    return;
            }
            if (formSelTarget.TypeEx != formSyncDownload)
                return;

            if (!flgDocInitSyncSOFrmCon) return;

            //  ShowSyncSO(FormSyncDownload);
            SetFolder(FormSyncDownload, "TabBp");
            SetFolder(FormSyncDownload, "TbpL");
            InitializeCheckBoxSyncD(FormSyncDownload);

        }
        catch (Exception ex)
        {
        }

    }
    private void ShowSyncSO(SAPbouiCOM.Form form)
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxSyncSOFrmCon) == null)
                return;


            var filters = new SearchFilter[1];
            filters[0] = new SearchFilter();

            filters[0].FieldName = "SoftwareId";
            filters[0].ComparisonOperator = "ne";
            filters[0].FieldValue = Software + "^";

            //filters
            ListSalesOrder = ebiz.SearchSalesOrders(getToken(), "", "", "", "", filters, 0, 1000, "1", true);

            form.Items.Item(MtxSyncSOFrmCon).AffectsFormMode = false;
            OMtxSyncSOFrmCon = (SAPbouiCOM.Matrix)form.Items.Item(MtxSyncSOFrmCon).Specific;
            OMtxSyncSOFrmCon.Clear();
            odtSyncSOFrmCon = GetDataTableSyncDownload(form);

            sql = String.Format("select  top 1'N' as [clSel],ROW_NUMBER() OVER (ORDER BY [DocEntry]) AS [ID],[DocNum],[DocDate],[DocTotal],[CardCode],[CardName],IFNULL([U_EbizChargeID],'') as [EBiz],CASE WHEN IFNULL([U_EbizChargeID], '') = '' THEN 'False' ELSE 'True'END as [Sync] from ORDR where  [DocStatus]='100' ").Replace("]", "\"").Replace("[", "\"");
            odtSyncSOFrmCon.ExecuteQuery(sql);
            BindMtxSyncSOFrmCon(OMtxSyncSOFrmCon, "clSel", "clSel", dtItem, true);
            BindMtxSyncSOFrmCon(OMtxSyncSOFrmCon, "clNo", "ID");
            BindMtxSyncSOFrmCon(OMtxSyncSOFrmCon, "DocNum", "DocNum");
            BindMtxSyncSOFrmCon(OMtxSyncSOFrmCon, "DocDate", "DocDate");
            BindMtxSyncSOFrmCon(OMtxSyncSOFrmCon, "CardCode", "CardCode");
            BindMtxSyncSOFrmCon(OMtxSyncSOFrmCon, "Software", "CardName");
            BindMtxSyncSOFrmCon(OMtxSyncSOFrmCon, "DocTotal", "DocTotal");
            BindMtxSyncSOFrmCon(OMtxSyncSOFrmCon, "clEbizkey", "EBiz");
            BindMtxSyncSOFrmCon(OMtxSyncSOFrmCon, "clSync", "Sync");

            OMtxSyncSOFrmCon.LoadFromDataSource();
            OMtxSyncSOFrmCon.Clear();
            foreach (var itm in ListSalesOrder)
            {
                OMtxSyncSOFrmCon.AddRow(1);
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)OMtxSyncSOFrmCon.Columns.Item("CardCode").Cells.Item(OMtxSyncSOFrmCon.RowCount).Specific;
                oEdit.Value = itm.CustomerId.Trim();
                oEdit = (SAPbouiCOM.EditText)OMtxSyncSOFrmCon.Columns.Item("DocTotal").Cells.Item(OMtxSyncSOFrmCon.RowCount).Specific;
                oEdit.Value = itm.Amount.ToString();
                oEdit = (SAPbouiCOM.EditText)OMtxSyncSOFrmCon.Columns.Item("DocNum").Cells.Item(OMtxSyncSOFrmCon.RowCount).Specific;
                oEdit.Value = itm.SalesOrderNumber;
                oEdit = (SAPbouiCOM.EditText)OMtxSyncSOFrmCon.Columns.Item("clEbizkey").Cells.Item(OMtxSyncSOFrmCon.RowCount).Specific;
                oEdit.Value = itm.SalesOrderInternalId;
                oEdit = (SAPbouiCOM.EditText)OMtxSyncSOFrmCon.Columns.Item("Software").Cells.Item(OMtxSyncSOFrmCon.RowCount).Specific;
                oEdit.Value = itm.Software.Trim();
                if (itm.Software.Equals(Software))
                {
                    var t = "";
                }
                oEdit = (SAPbouiCOM.EditText)OMtxSyncSOFrmCon.Columns.Item("DocDate").Cells.Item(OMtxSyncSOFrmCon.RowCount).Specific;
                var date = DatetimeTosapFormatDate(Convert.ToDateTime(itm.DueDate));
                oEdit.Value = date;
            }
            OMtxSyncSOFrmCon.AutoResizeColumns();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }

    private SAPbouiCOM.DataTable GetDataTableSyncDownload(SAPbouiCOM.Form form)
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = form.DataSources.DataTables.Item(MtxSyncSOFrmCon);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
            oDT = form.DataSources.DataTables.Add(MtxSyncSOFrmCon);

        }

        return oDT;
    }
    private SAPbouiCOM.DataTable GetDataTableSyncDownload(SAPbouiCOM.Form form, String Mtx)
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = form.DataSources.DataTables.Item(Mtx);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
            oDT = form.DataSources.DataTables.Add(Mtx);

        }

        return oDT;
    }
    private void BindMtxSyncSOFrmCon(SAPbouiCOM.Matrix oMatrix1, string mtCol, string dtCol, string dt = null, bool editable = false)
    {
        try
        {
            SAPbouiCOM.Column oColumn = oMatrix1.Columns.Item(mtCol);
            oColumn.Editable = editable;
            if (mtCol.Equals("clEbizkey"))
            {
                oColumn.Visible = false;
            }

            if (dt == null)
            {
                oColumn.DataBind.Bind(MtxSyncSOFrmCon, dtCol);
            }
            else
                oColumn.DataBind.Bind(dt, dtCol);
        }
        catch (Exception ex)
        {
            errorLog("Can not bind " + mtCol + " to " + dtCol + ".  error: " + ex.Message);
        }

    }
    private void BindMatrix(string MtxStr, SAPbouiCOM.Matrix oMatrix1, string mtCol, string dtCol, string dt = null, bool editable = false)
    {
        try
        {
            SAPbouiCOM.Column oColumn = oMatrix1.Columns.Item(mtCol);
            oColumn.Editable = editable;
            if (dt == null)
            {
                oColumn.DataBind.Bind(MtxStr, dtCol);
            }
            else
                oColumn.DataBind.Bind(dt, dtCol);
        }
        catch (Exception ex)
        {
            errorLog("Can not bind " + mtCol + " to " + dtCol + ".  error: " + ex.Message);
        }

    }
    private void SelectAllCheckBox(SAPbouiCOM.Form form, String MatxStr, string CheckboxID)
    {
        try
        {
            if (form.Items.Item(MatxStr) == null)
                return;

            form.Items.Item(MatxStr).AffectsFormMode = false;
            SAPbouiCOM.Matrix OMatrix = (SAPbouiCOM.Matrix)form.Items.Item(MatxStr).Specific;
            var SelectAllCheckBox = (SAPbouiCOM.CheckBox)form.Items.Item(CheckboxID).Specific;
            int RowCount = OMatrix.RowCount;
            form.Freeze(true);
            bool flgUpdateCheckbox = false;
            if (RowCount > 0)
            {
                ShowMsgCustom("Selecting ...!");
                for (int i = RowCount - 1; i >= 0; i--)
                {
                    var CheckBoxMatrixState = (SAPbouiCOM.CheckBox)OMatrix.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    if (!SelectAllCheckBox.Checked)
                    {
                        if (!CheckBoxMatrixState.Checked)
                        {
                            OMatrix.Columns.Item("clSel").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            flgUpdateCheckbox = true;
                        }
                    }
                    else
                    {
                        if (CheckBoxMatrixState.Checked)
                        {
                            OMatrix.Columns.Item("clSel").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                    }
                }
            }

            form.Freeze(false);
            form.Update();
            if (flgUpdateCheckbox)
            {
                form.DataSources.UserDataSources.Item(CheckboxID).Value = "Y";
            }
            else
            {
                form.DataSources.UserDataSources.Item(CheckboxID).Value = "N";
            }
        }
        catch (Exception ex)
        {
            errorLog(ex.Message);
        }
    }
    private void InitializeCheckBoxSyncD(SAPbouiCOM.Form oform)
    {
        try
        {
            String checkID = "chkBpL1";
            oCheckboxSOSelectAll = ((SAPbouiCOM.CheckBox)(oform.Items.Item(checkID).Specific));
            oform.DataSources.UserDataSources.Add(checkID, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oCheckboxSOSelectAll.DataBind.SetBound(true, "", checkID);
            oform.DataSources.UserDataSources.Item(checkID).Value = "N";

            checkID = "chkItmL1";
            oCheckboxItemSelectAll = ((SAPbouiCOM.CheckBox)(oform.Items.Item(checkID).Specific));
            oform.DataSources.UserDataSources.Add(checkID, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oCheckboxItemSelectAll.DataBind.SetBound(true, "", checkID);
            oform.DataSources.UserDataSources.Item(checkID).Value = "N";

            checkID = "chkSoL1";
            oCheckboxBPSelectAll = ((SAPbouiCOM.CheckBox)(oform.Items.Item(checkID).Specific));
            oform.DataSources.UserDataSources.Add(checkID, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oCheckboxBPSelectAll.DataBind.SetBound(true, "", checkID);
            oform.DataSources.UserDataSources.Item(checkID).Value = "N";

            checkID = "chkSopL1";
            oCheckboxBPSelectAll = ((SAPbouiCOM.CheckBox)(oform.Items.Item(checkID).Specific));
            oform.DataSources.UserDataSources.Add(checkID, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oCheckboxBPSelectAll.DataBind.SetBound(true, "", checkID);
            oform.DataSources.UserDataSources.Item(checkID).Value = "N";


        }
        catch (Exception ex)
        {

            var error = ex.Message;
        }
    }
    #endregion

    #region Post Logic
    private void PostSyncSO(String EdEbixCharge, SalesOrder oSalesOrder, int i, SAPbouiCOM.Matrix oMatx, SAPbouiCOM.Form Oform)
    {
        String EbizKey = "";
        bool PostStatus = false;
        bool ValidFlag = false;
        String ValidMsg = "";
        int DocEntry = 0;
        try
        {
            String Msg = "";
            if (CheckBPFromExternal(oSalesOrder.CustomerId.Trim(), ref Msg))
            {
                int count = 1;
                foreach (var item in oSalesOrder.Items)
                {
                    if (!CheckItemForExternal(item.ItemId.Trim(), ref Msg))
                    {
                        if (count == oSalesOrder.Items.Length)
                        {
                            //
                            #region Posting
                            var Obj = oSalesOrder;
                            ValidSalesOrder(Obj, ref ValidMsg, ref ValidFlag);

                            if (ValidFlag)
                            {
                                if (!IsSODocNumExist(Obj.SalesOrderNumber))
                                {
                                    // Add Order
                                    //CreateSOObject(Obj, ref ValidMsg, ref ValidFlag, ref DocEntry);
                                    if (ValidFlag && DocEntry > 0)
                                    {
                                        #region Update
                                        try
                                        {
                                            var SearchOrder = ebiz.GetSalesOrder(getToken(), "", "", "", Obj.SalesOrderInternalId);
                                            var oldSoNum = SearchOrder.SalesOrderNumber;

                                            if (SearchOrder != null)
                                            {
                                                SalesOrder updateSO = new SalesOrder();
                                                updateSO.CustomerId = SearchOrder.CustomerId.Trim();
                                                updateSO.SalesOrderInternalId = null;
                                                updateSO.Software = Software;
                                                updateSO.SalesOrderNumber = getSODocNum(DocEntry);
                                                updateSO.Date = Obj.Date;
                                                updateSO.DueDate = Obj.DueDate;
                                                updateSO.Amount = Obj.Amount;
                                                updateSO.AmountDue = Obj.AmountDue;
                                                updateSO.NotifyCustomer = false;
                                                updateSO.TotalTaxAmount = 0;
                                                updateSO.Memo = "";
                                                updateSO.ShipDate = SearchOrder.DueDate;
                                                updateSO.ShipVia = "";
                                                updateSO.Items = Obj.Items;

                                                var updateSoResp = ebiz.UpdateSalesOrder(getToken(), updateSO, Obj.CustomerId.Trim(), null, oldSoNum, null);
                                                if (updateSoResp.Status.Equals(""))
                                                {
                                                    ValidMsg = "Update To Connect Status:" + updateSoResp.Status;
                                                }
                                                else if (updateSoResp.Status.Equals("Success"))
                                                {

                                                    PostStatus = true;
                                                    ShowSystemLog(String.Format("Sync Doc:{0} Downloded Successfully", oSalesOrder.SalesOrderNumber));
                                                }
                                            }


                                        }
                                        catch { }
                                        #endregion
                                    }
                                }
                                else
                                {
                                    //Update order
                                    //   CreateSOObject(Obj, ref ValidMsg, ref ValidFlag, ref DocEntry, "Update");
                                    if (ValidFlag && DocEntry > 0)
                                    {
                                        PostStatus = true;
                                        ShowSystemLog(String.Format("Sync Doc:{0} Uploaded Successfully", oSalesOrder.SalesOrderNumber));
                                    }

                                }
                            }

                            #region status
                            if (PostStatus)
                            {
                                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
                                oEdit.Value = "True";
                            }
                            else
                            {
                                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
                                oEdit.Value = "False";
                            }
                            #endregion
                            #endregion
                            //
                        }
                    }

                    else
                    {
                        ShowSystemLog(String.Format("Sync Doc:{0} Failed as {1}", oSalesOrder.SalesOrderNumber, Msg));
                    }

                    count++;
                }
            }
            else
            {

                ShowSystemLog(String.Format("Sync Doc:{0} Failed as {1}", oSalesOrder.SalesOrderNumber, Msg));
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }
    }
    private void CreateSOObject(SalesOrder oSalesOrder, ref String Msg, ref bool Flag, ref int DocEntry, String Mode = "Add")
    {
        try
        {
            Flag = true;
            Msg = "";
            DocEntry = 0;
            DocSalesOrder = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
            if (!Mode.Equals("Add"))
            {
                int DocEntry1 = getDocEntry("17", oSalesOrder.SalesOrderNumber.Trim().ToString(), oSalesOrder.CustomerId.Trim().ToString());
                DocSalesOrder.GetByKey(DocEntry1);
            }
            DocSalesOrder.CardCode = oSalesOrder.CustomerId.Trim();
            DateTime dt;
            try
            {
                dt = Convert.ToDateTime(oSalesOrder.Date);
                DocSalesOrder.DocDate = dt;
            }
            catch (Exception ex)
            {
                DocSalesOrder.DocDate = DateTime.Now;
            }
            try
            {
                dt = Convert.ToDateTime(oSalesOrder.DueDate);
                DocSalesOrder.DocDueDate = dt;
            }
            catch (Exception ex)
            {
                DocSalesOrder.DocDueDate = DateTime.Now;
            }
            DocSalesOrder.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
            DocSalesOrder.Comments = oSalesOrder.Description;
            //DocSalesOrder.DiscountPercent = 0;
            try
            {
                DocSalesOrder.UserFields.Fields.Item("U_" + U_EbizChargeKey).Value = oSalesOrder.SalesOrderInternalId;
                DocSalesOrder.UserFields.Fields.Item("U_ConnectDocNum").Value = oSalesOrder.SalesOrderNumber;//ConnectDocNum
            }
            catch (Exception)
            {
            }
            if (!Mode.Equals("Add"))
            {
                for (int i = DocSalesOrder.Lines.Count - 1; i >= 0; i--)
                {
                    DocSalesOrder.Lines.SetCurrentLine(i);
                    DocSalesOrder.Lines.Delete();
                }
            }
            foreach (var itm in oSalesOrder.Items)
            {
                DocSalesOrderLines = DocSalesOrder.Lines;
                DocSalesOrderLines.ItemCode = itm.ItemId.Trim();
                DocSalesOrderLines.Quantity = ConvertDouble(itm.Qty);
                DocSalesOrderLines.UnitPrice = ConvertDouble(itm.UnitPrice);
                //DocSalesOrderLines.VatGroup = "S2";
                DocSalesOrder.Lines.Add();
            }
            int sta = 0;
            if (!Mode.Equals("Add"))
            {
                sta = DocSalesOrder.Update();
            }
            else
            {
                sta = DocSalesOrder.Add();
            }
            var sapDocEntry = oCompany.GetNewObjectKey();
            if (sta != 0)
            {
                sapDocEntry = "";
                int erroCode = 0;
                string errDescr = "";
                oCompany.GetLastError(out erroCode, out errDescr);
                Flag = false;
                Msg = errDescr;
            }
            else
            {// Posted
                Flag = true;
                Msg = "";
                DocEntry = Convert.ToInt32(sapDocEntry);
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            Flag = false;
            Msg = "Issue:" + ex.Message;
        }

    }
    public void ValidSalesOrder(SalesOrder oSalesOrder, ref String Msg, ref bool Flag)
    {
        try
        {
            Msg = String.Empty; Flag = true;
            ExportBpModel oBP = GetSAPBP(oSalesOrder.CustomerId.Trim());
            if (oBP != null)
            {
                ItemMasterSAP oItem = null;
                foreach (var item in oSalesOrder.Items)
                {
                    oItem = GetSAPItem(item.ItemId.Trim());
                    if (oItem == null)
                    {
                        Flag = false;
                        Msg = String.Format("Item {0} don't exist on SAP", item.ItemId);
                        return;
                    }
                }
            }
            else
            {
                Flag = false;
                Msg = "Business Partner Dont Exist on SAP";
                return;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }

    #endregion

    #region Download Logic of All Tabs
    public void SyncDownloadBPProc()
    {
        string sql = "";
        try
        {
            if (formSelTarget == null)
            {
                if (FormSyncDownload != null)
                    formSelTarget = FormSyncDownload;
                else
                    return;
            }
            if (formSelTarget.TypeEx != formSyncDownload)
                return;

            if (!flgDocInitSyncSOFrmCon) return;

            SAPbouiCOM.Form form = FormSyncDownload;

            if (form.Items.Item(SyncDownloadTabBP) == null)
                return;


            var filters = new SearchFilter[1];
            filters[0] = new SearchFilter();
            filters[0].FieldName = "SoftwareID";
            filters[0].ComparisonOperator = "ne";
            filters[0].FieldValue = Software+ TesingFilter;
            ListBp = ebiz.SearchCustomerList(getToken(), filters, 0, DownloadRowCount, "1", false, false, false);

            form.Items.Item(MtxSyncDownloadBPListStr).AffectsFormMode = false;
            OMtxSyncBPList = (SAPbouiCOM.Matrix)form.Items.Item(MtxSyncDownloadBPListStr).Specific;
            OMtxSyncBPList.Clear();

            odtSyncBP = GetDataTableSyncDownload(form);

            sql = String.Format("select  top 1'N' as [clSel],ROW_NUMBER() OVER (ORDER BY [DocEntry]) AS [ID],[DocDate],[DocTotal],cast([CardName] as nvarchar(200)) as [CardName] from ORDR ").Replace("]", "\"").Replace("[", "\"");
            odtSyncBP.ExecuteQuery(sql);
            BindMatrix(MtxSyncDownloadBPListStr, OMtxSyncBPList, "clSel", "clSel", dtItem, true);
            BindMatrix(MtxSyncDownloadBPListStr, OMtxSyncBPList, "CardCode", "CardName", dtItem, false);
            BindMatrix(MtxSyncDownloadBPListStr, OMtxSyncBPList, "CardName", "CardName", dtItem, false);
            BindMatrix(MtxSyncDownloadBPListStr, OMtxSyncBPList, "Email", "CardName", dtItem, false);
            BindMatrix(MtxSyncDownloadBPListStr, OMtxSyncBPList, "Balance", "DocTotal", dtItem, false);
            BindMatrix(MtxSyncDownloadBPListStr, OMtxSyncBPList, "Source", "CardName", dtItem, false);
            BindMatrix(MtxSyncDownloadBPListStr, OMtxSyncBPList, "IntID", "CardName", dtItem, false);
            OMtxSyncBPList.LoadFromDataSource();
            OMtxSyncBPList.Clear();

            for (int i = 0; i < ListBp.Count; i++)
            {
                Customer oBP = ListBp.CustomerList[i];
                OMtxSyncBPList.AddRow(1);
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)OMtxSyncBPList.Columns.Item("CardCode").Cells.Item(OMtxSyncBPList.RowCount).Specific;
                oEdit.Value = oBP.CustomerId.Trim();
                oEdit = (SAPbouiCOM.EditText)OMtxSyncBPList.Columns.Item("CardName").Cells.Item(OMtxSyncBPList.RowCount).Specific;
                oEdit.Value = String.Format("{0} {1}", oBP.FirstName.Trim(), oBP.LastName.Trim());
                oEdit = (SAPbouiCOM.EditText)OMtxSyncBPList.Columns.Item("Email").Cells.Item(OMtxSyncBPList.RowCount).Specific;
                oEdit.Value = String.Format("{0}", oBP.Email.Trim());
                oEdit = (SAPbouiCOM.EditText)OMtxSyncBPList.Columns.Item("Source").Cells.Item(OMtxSyncBPList.RowCount).Specific;
                oEdit.Value = String.Format("{0}", oBP.SoftwareId.Trim());
                oEdit = (SAPbouiCOM.EditText)OMtxSyncBPList.Columns.Item("IntID").Cells.Item(OMtxSyncBPList.RowCount).Specific;
                oEdit.Value = String.Format("{0}", oBP.CustomerInternalId.Trim());
            }
            OMtxSyncBPList.Columns.Item("IntID").Visible = false;
            OMtxSyncBPList.AutoResizeColumns();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }

    }

    public void SyncDownloadItemProc()
    {
        string sql = "";
        String MtrxStrID = MtxSyncDownloadItemListStr;
        SAPbouiCOM.Matrix oMtrx = null;
        SAPbouiCOM.DataTable oDt = null;
        try
        {
            if (formSelTarget == null)
            {
                if (FormSyncDownload != null)
                    formSelTarget = FormSyncDownload;
                else
                    return;
            }
            if (formSelTarget.TypeEx != formSyncDownload)
                return;

            if (!flgDocInitSyncSOFrmCon) return;

            SAPbouiCOM.Form form = FormSyncDownload;

            if (form.Items.Item(MtrxStrID) == null)
                return;


            var filters = new SearchFilter[1];
            filters[0] = new SearchFilter();
            filters[0].FieldName = "SoftwareID";
            filters[0].ComparisonOperator = "notequal";
            filters[0].FieldValue = Software+ TesingFilter;
            OutputListItem = ebiz.SearchItems(getToken(), "", "", filters, 0, 1000, "1");

            form.Items.Item(MtrxStrID).AffectsFormMode = false;
            oMtrx = (SAPbouiCOM.Matrix)form.Items.Item(MtrxStrID).Specific;
            oMtrx.Clear();

            oDt = GetDataTableSyncDownload(form, MtrxStrID);

            sql = String.Format("select  top 1'N' as [clSel],ROW_NUMBER() OVER (ORDER BY [DocEntry]) AS [ID],[DocDate],[DocTotal] ,cast([CardName] as nvarchar(200)) as [CardName] from ORDR ").Replace("]", "\"").Replace("[", "\"");
            oDt.ExecuteQuery(sql);
            BindMatrix(MtrxStrID, oMtrx, "clSel", "clSel", dtItem, true);
            BindMatrix(MtrxStrID, oMtrx, "ItemCode", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "ItemName", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "Quantity", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "UnitPrice", "DocTotal", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "Source", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "IntID", "CardName", dtItem, false);
            oMtrx.LoadFromDataSource();
            oMtrx.Clear();

            for (int i = 0; i < OutputListItem.Length; i++)
            {
                ItemDetails Oitem = OutputListItem[i];
                oMtrx.AddRow(1);
                try
                {
                    SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("ItemCode").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = Oitem.ItemId.Trim();
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("ItemName").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", Oitem.Name.Trim());
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("Quantity").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", Oitem.QtyOnHand.ToString());
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("UnitPrice").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", Oitem.UnitPrice.ToString());
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("Source").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", Oitem.SoftwareId.ToString());
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("IntID").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", Oitem.ItemInternalId.Trim());
                }
                catch (Exception ex)
                {
                    errorLog(ex);
                }
            }
            oMtrx.Columns.Item("IntID").Visible = false;
            oMtrx.AutoResizeColumns();
            OMtxSyncItemList = oMtrx;
            odtSyncItem = oDt;
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }

    }

    public void SyncDownloadSOProc()
    {
        string sql = "";
        String MtrxStrID = MtxSyncDownloadSOListStr;
        SAPbouiCOM.Matrix oMtrx = null;
        SAPbouiCOM.DataTable oDt = null;
        try
        {
            if (formSelTarget == null)
            {
                if (FormSyncDownload != null)
                    formSelTarget = FormSyncDownload;
                else
                    return;
            }
            if (formSelTarget.TypeEx != formSyncDownload)
                return;

            if (!flgDocInitSyncSOFrmCon) return;

            SAPbouiCOM.Form form = FormSyncDownload;

            if (form.Items.Item(MtrxStrID) == null)
                return;


            var filters = new SearchFilter[1];
            filters[0] = new SearchFilter();
            filters[0].FieldName = "SoftwareID";
            filters[0].ComparisonOperator = "ne";
            filters[0].FieldValue = Software+ TesingFilter;
            OutputListSO = ebiz.SearchSalesOrders(getToken(), "", "", "", "", filters, 0, 1000, "1", true);

            form.Items.Item(MtrxStrID).AffectsFormMode = false;
            oMtrx = (SAPbouiCOM.Matrix)form.Items.Item(MtrxStrID).Specific;
            oMtrx.Clear();

            oDt = GetDataTableSyncDownload(form, MtrxStrID);

            sql = String.Format("select  top 1'N' as [clSel],ROW_NUMBER() OVER (ORDER BY [DocEntry]) AS [ID],[DocDate],[DocTotal] ,cast([CardName] as nvarchar(200)) as [CardName] from ORDR ").Replace("]", "\"").Replace("[", "\"");
            oDt.ExecuteQuery(sql);
            BindMatrix(MtrxStrID, oMtrx, "clSel", "clSel", dtItem, true);
            BindMatrix(MtrxStrID, oMtrx, "DocEntry", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "DocNum", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "CardCode", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "CardName", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "Type", "Type", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "DocDate", "DocDate", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "AmtPaid", "DocTotal", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "Source", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "IntID", "CardName", dtItem, false);
            oMtrx.LoadFromDataSource();
            oMtrx.Clear();

            for (int i = 0; i < OutputListSO.Length; i++)
            {
                SalesOrder oSales = OutputListSO[i];
                oMtrx.AddRow(1);
                try
                {
                    SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("DocEntry").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = oSales.SalesOrderNumber.Trim();
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("DocNum").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", oSales.SalesOrderNumber.Trim());
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("CardCode").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", oSales.CustomerId.Trim().ToString());
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("CardName").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", oSales.CustomerId.Trim().ToString());
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("Type").Cells.Item(oMtrx.RowCount).Specific;
                    if (oSales.TypeId != null)
                    {
                        oEdit.Value = String.Format("{0}", oSales.TypeId.ToString());
                    }
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("DocDate").Cells.Item(oMtrx.RowCount).Specific;
                    var date = DatetimeTosapFormatDate(Convert.ToDateTime(oSales.DueDate));
                    oEdit.Value = date;
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("AmtPaid").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", oSales.AmountDue.ToString());
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("Source").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", oSales.Software.ToString());
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("IntID").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", oSales.SalesOrderInternalId.Trim());
                }
                catch (Exception ex)
                {
                    errorLog(ex);
                }
            }
            oMtrx.Columns.Item("IntID").Visible = false;
            oMtrx.AutoResizeColumns();
            OMtxSyncSOList = oMtrx;
            odtSyncSOFrmCon = oDt;
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }

    }

    public void SyncDownloadSOPProc()
    {
        string sql = "";
        String MtrxStrID = "mtxSOP";
        SAPbouiCOM.Matrix oMtrx = null;
        SAPbouiCOM.DataTable oDt = null;
        try
        {
            if (formSelTarget == null)
            {
                if (FormSyncDownload != null)
                    formSelTarget = FormSyncDownload;
                else
                    return;
            }
            if (formSelTarget.TypeEx != formSyncDownload)
                return;

            if (!flgDocInitSyncSOFrmCon) return;

            SAPbouiCOM.Form form = FormSyncDownload;

            if (form.Items.Item(MtrxStrID) == null)
                return;


            var filters = new SearchFilter[2];
            filters[0] = new SearchFilter();
            filters[0].FieldName = "SoftwareID";
            filters[0].ComparisonOperator = "ne";
            filters[0].FieldValue = Software+ TesingFilter;
            filters[1] = new SearchFilter();
            filters[1].FieldName = "IsTransactionApplied";
            filters[1].ComparisonOperator = "ne";
            filters[1].FieldValue = "1";
            TransactionList = ebiz.SearchApplicationTransactions(getToken(), filters, false, 0, 1000, "");


            form.Items.Item(MtrxStrID).AffectsFormMode = false;
            oMtrx = (SAPbouiCOM.Matrix)form.Items.Item(MtrxStrID).Specific;
            oMtrx.Clear();

            oDt = GetDataTableSyncDownload(form, MtrxStrID);

            sql = String.Format("select  top 1'N' as [clSel],ROW_NUMBER() OVER (ORDER BY [DocEntry]) AS [ID],[DocDate],[DocTotal] ,cast([CardName] as nvarchar(200)) as [CardName] from ORDR ").Replace("]", "\"").Replace("[", "\"");
            oDt.ExecuteQuery(sql);
            BindMatrix(MtrxStrID, oMtrx, "clSel", "clSel", dtItem, true);
            BindMatrix(MtrxStrID, oMtrx, "DocEntry", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "DocNum", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "CardCode", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "CardName", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "Type", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "DatePaid", "DocDate", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "AmtPaid", "DocTotal", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "PymtM", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "Source", "CardName", dtItem, false);
            BindMatrix(MtrxStrID, oMtrx, "IntID", "CardName", dtItem, false);
            oMtrx.LoadFromDataSource();
            oMtrx.Clear();

            for (int i = 0; i < TransactionList.Count; i++)
            {
                var transaction = TransactionList.ApplicationTransactions[i];
                var transactionDetail = ebiz.GetTransactionDetails(getToken(), transaction.TransactionId);

                oMtrx.AddRow(1);
                try
                {
                    SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("DocEntry").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = getDocEntry(transaction.LinkedToTypeId.Trim(), "ORDR").ToString();
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("DocNum").Cells.Item(oMtrx.RowCount).Specific;
                    var DocNum = GetSalesOrderDocNum(transaction.LinkedToTypeId.Trim());
                    oEdit.Value = String.Format("{0}", transaction.LinkedToTypeId.Trim());
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("CardCode").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", transactionDetail.CustomerID.Trim().ToString());
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("CardName").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", FindCustomerNameByID(transactionDetail.CustomerID.Trim().ToString()));
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("Type").Cells.Item(oMtrx.RowCount).Specific;
                    if (transaction.TransactionTypeId != null)
                    {
                        oEdit.Value = String.Format("{0}", transaction.TransactionTypeId.ToString());
                    }
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("PymtM").Cells.Item(oMtrx.RowCount).Specific;
                    if (transaction.PaymentMethodType != null)
                    {
                        var payMethod = transaction.PaymentMethodType + " ending in " + transaction.PaymentMethodLast4;
                        oEdit.Value = String.Format("{0}", payMethod);
                    }
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("DatePaid").Cells.Item(oMtrx.RowCount).Specific;
                    var date = DatetimeTosapFormatDate(Convert.ToDateTime(transaction.TransactionDate));
                    oEdit.Value = date;
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("AmtPaid").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", transactionDetail.Details.Amount);
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("Source").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", transaction.SoftwareId.ToString());
                    oEdit = (SAPbouiCOM.EditText)oMtrx.Columns.Item("IntID").Cells.Item(oMtrx.RowCount).Specific;
                    oEdit.Value = String.Format("{0}", transaction.ApplicationTransactionInternalId.Trim());
                }
                catch (Exception ex)
                {
                    errorLog(ex);
                }
            }
            //oMtrx.Columns.Item("IntID").Visible = false;
            oMtrx.AutoResizeColumns();
            OMtxSyncSOPList = oMtrx;
            odtSyncSOPFrmCon = oDt;
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }

    }
    #endregion

    #region Posting Data from Connect to SAP B1

    #region Business Partner
    private void DownloadDataBP(String MatxStr, SAPbouiCOM.Form form)
    {
        try
        {
            bool FlgStatus = false;
            bool FlgHasSelected = false;
            String Resp = "";
            String Primarykey = "";
            SAPbouiCOM.Matrix OMatrix = (SAPbouiCOM.Matrix)form.Items.Item(MatxStr).Specific;
            int RowCount = OMatrix.RowCount;
            if (RowCount > 0)
            {
                ShowMsgCustom("Downloading ...!", "S", "S");
                for (int i = RowCount - 1; i >= 0; i--)
                {
                    Customer oCust = ListBp.CustomerList[i];
                    var CheckBoxMatrixState = (SAPbouiCOM.CheckBox)OMatrix.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    var InternalBpId = (SAPbouiCOM.EditText)OMatrix.Columns.Item("IntID").Cells.Item(i + 1).Specific;
                    if (CheckBoxMatrixState.Checked)
                    {
                        FlgHasSelected = true;
                        //Create Object
                        if (oCust.CustomerInternalId.Equals(InternalBpId.Value))
                        {
                            var Output = FindCustomerByID(oCust.CustomerId);
                            if (Output.Equals(""))
                            {
                                //Add
                                PostBP(oCust, ref FlgStatus, ref Resp, ref Primarykey);
                            }
                            else
                            {
                                //Update
                                PostBP(oCust, ref FlgStatus, ref Resp, ref Primarykey, true);
                            }
                            //AddLog
                            #region Log
                            LogTable oExportBPLogTable = new LogTable();
                            oExportBPLogTable.CardCode = oCust.CustomerId;
                            oExportBPLogTable.CardName = oCust.FirstName + " " + oCust.LastName;
                            oExportBPLogTable.Email = oCust.Email;
                            oExportBPLogTable.Balance = 0;
                            oExportBPLogTable.DocEntry = Primarykey;
                            oExportBPLogTable.Software = oCust.SoftwareId;
                            oExportBPLogTable.Status = Resp;
                            oExportBPLogTable.CreateDt = DateTime.Now;
                            oExportBPLogTable.UpdateDt = DateTime.Now;
                            AddUploadDataLog(oExportBPLogTable, "CCBPIMP");

                            if (FlgStatus)
                            {
                                ShowMsgCustom(String.Format("Downloading of Customer;{0}", oCust.CustomerId), "S", "S");
                            }
                            else
                            {
                                ShowMsgCustom(String.Format("Downloading of Customer:{0} failed status:{1}", oCust.CustomerId, Resp), "S", "W");
                            }
                            //Update on Connect
                            if (FlgStatus)
                            {
                                Thread myNewThread = new Thread(() => UpdateBpToConnect(ListBp.CustomerList[i]));
                                myNewThread.Start();
                            }
                            #endregion
                        }
                    }
                }
                if (FlgHasSelected)
                {
                    ShowMsgCustom("Downloading done.", "S", "S");
                }
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void UpdateBpToConnect(Customer oCust)
    {
        try
        {
            oCust.SoftwareId = Software;
            var resp = ebiz.UpdateCustomer(getToken(), oCust, oCust.CustomerId, oCust.CustomerInternalId);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void PostBP(Customer oCust, ref bool flgStatus, ref String Response, ref String DocEntry, bool Update = false)
    {
        try
        {
            SAPbobsCOM.BusinessPartners BP_Doc;
            if (!oCompany.Connected)
            {
                oCompany.Connect();
            }

            {
                BP_Doc = (SAPbobsCOM.BusinessPartners)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                BP_Doc.CardName = oCust.CustomerId;

                if (Update)
                {
                    BP_Doc.GetByKey(oCust.CustomerId);
                }
                else
                {
                    BP_Doc.CardCode = oCust.CustomerId;
                }
                BP_Doc.CardForeignName = String.Format("{0} {1}", oCust.FirstName, oCust.LastName);

                var BillingAdd = oCust.BillingAddress;
                BP_Doc.Address = BillingAdd.Address1;
                BP_Doc.Addresses.AddressName = "Bill To";
                BP_Doc.Addresses.Block = "1";
                BP_Doc.ZipCode = BillingAdd.ZipCode;

                // BP_Doc.Addresses.Street = BillingAdd.;
                BP_Doc.Addresses.AddressName2 = BillingAdd.Address2;
                BP_Doc.Addresses.AddressName3 = BillingAdd.Address3;
                //BP_Doc.Addresses.Street = BP.Address1;
                BP_Doc.Addresses.City = BillingAdd.City;
                BP_Doc.Addresses.Country = BillingAdd.Country;
                BP_Doc.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                BP_Doc.Addresses.Add();
                try
                {
                    BP_Doc.ContactEmployees.Name = oCust.FirstName;
                    BP_Doc.ContactEmployees.MobilePhone = oCust.CellPhone;
                    BP_Doc.ContactEmployees.Add();
                    BP_Doc.Phone2 = oCust.Phone;
                }
                catch (Exception)
                {
                }
                BP_Doc.Valid = SAPbobsCOM.BoYesNoEnum.tYES;
                try
                {
                    BP_Doc.UserFields.Fields.Item("U_EbizChargeID").Value = oCust.CustomerInternalId;
                }
                catch { }
                int sta = 0;
                if (Update)
                {
                    sta = BP_Doc.Update();
                }
                else
                {
                    sta = BP_Doc.Add();
                }
                String pp = BP_Doc.CardCode;
                var sapDocEntry = oCompany.GetNewObjectKey();
                String a = "";
                oCompany.GetNewObjectCode(out a);
                if (sta != 0)
                {
                    sapDocEntry = "";
                    int erroCode = 0;
                    string errDescr = "";
                    oCompany.GetLastError(out erroCode, out errDescr);
                    flgStatus = false;
                    Response = String.Format("Error Posting:{0}", errDescr);
                }
                else
                {
                    flgStatus = true;
                    if (Update)
                    {
                        Response = String.Format("Updated", sapDocEntry);
                    }
                    else
                    {
                        Response = String.Format("Imported", sapDocEntry);
                    }
                }
            }

        }
        catch (Exception ex)
        {
            flgStatus = false;
            Response = String.Format("Error in Document:{0}", ex.Message);

        }

    }
    #endregion
    #region Item Master
    private void DownloadDataItem(String MatxStr, SAPbouiCOM.Form form)
    {
        try
        {
            bool FlgStatus = false;
            bool FlgHasSelected = false;
            String Resp = "";
            String Primarykey = "";
            SAPbouiCOM.Matrix OMatrix = (SAPbouiCOM.Matrix)form.Items.Item(MatxStr).Specific;
            int RowCount = OMatrix.RowCount;
            if (RowCount > 0)
            {
                ShowMsgCustom("Downloading ...!", "S", "S");
                for (int i = RowCount - 1; i >= 0; i--)
                {
                    ItemDetails oitem = OutputListItem[i];
                    var CheckBoxMatrixState = (SAPbouiCOM.CheckBox)OMatrix.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    var InternalId = (SAPbouiCOM.EditText)OMatrix.Columns.Item("IntID").Cells.Item(i + 1).Specific;
                    if (CheckBoxMatrixState.Checked)
                    {
                        FlgHasSelected = true;
                        //Create Object
                        if (oitem.ItemInternalId.Equals(InternalId.Value))
                        {
                            var Output = FindItemByID(oitem.ItemId);
                            if (Output.Equals(""))
                            {
                                //Add
                                PostItem(oitem, ref FlgStatus, ref Resp, ref Primarykey);
                            }
                            else
                            {
                                //Update
                                PostItem(oitem, ref FlgStatus, ref Resp, ref Primarykey, true);
                            }
                            //AddLog
                            #region Log


                            if (FlgStatus)
                            {
                                ShowMsgCustom(String.Format("Downloading of Item:{0}", oitem.ItemId), "S", "S");
                            }
                            else
                            {
                                ShowMsgCustom(String.Format("Downloading of Item:{0} failed status:{1}", oitem.ItemId, Resp), "S", "W");
                            }
                            //Update on Connect

                            Thread myNewThread = new Thread(() => UpdateItemToConnect(oitem, Resp, FlgStatus));
                            myNewThread.Start();

                            #endregion
                        }
                    }
                }
                if (FlgHasSelected)
                {
                    ShowMsgCustom("Downloading done.", "S", "S");
                }
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void UpdateItemToConnect(ItemDetails oitem, String Resp, bool flgposted)
    {
        try
        {
            oitem.SoftwareId = Software;
            ItemDetailsResponse resp = ebiz.UpdateItem(getToken(), oitem, oitem.ItemId, oitem.ItemInternalId);

            if (resp.Status.Equals("Error"))
            {
                Resp = "SAP Status:" + Resp + " Connect Status:" + resp.Error;
            }

            LogTable oLogTable = new LogTable();
            oLogTable.ItemCode = oitem.ItemId;
            oLogTable.ItemName = oitem.Description;
            oLogTable.Quantity = oitem.QtyOnHand.Value;
            oLogTable.UnitPrice = oitem.UnitPrice.Value;
            oLogTable.DocEntry = oitem.ItemId;
            oLogTable.Software = oitem.SoftwareId;
            oLogTable.Status = Resp;
            oLogTable.CreateDt = DateTime.Now;
            oLogTable.UpdateDt = DateTime.Now;
            AddUploadDataLog(oLogTable, "CCITMIMP");
            //if(resp.//)
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void PostItem(ItemDetails oitem, ref bool flgStatus, ref String Response, ref String DocEntry, bool Update = false)
    {
        try
        {
            SAPbobsCOM.Items oitem_Doc;
            if (!oCompany.Connected)
            {
                oCompany.Connect();
            }

            {
                oitem_Doc = (SAPbobsCOM.Items)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                if (Update)
                {
                    oitem_Doc.GetByKey(oitem.ItemId);
                }
                else
                {
                    oitem_Doc.ItemCode = oitem.ItemId;
                }
                oitem_Doc.ItemName = String.Format("{0}", oitem.Description);
                oitem_Doc.Valid = SAPbobsCOM.BoYesNoEnum.tYES;
                try
                {
                    oitem_Doc.UserFields.Fields.Item("U_EbizChargeID").Value = oitem.ItemInternalId;
                }
                catch { }
                int sta = 0;
                if (Update)
                {
                    sta = oitem_Doc.Update();
                }
                else
                {
                    sta = oitem_Doc.Add();
                }
                String pp = oitem_Doc.ItemCode;
                var sapDocEntry = oCompany.GetNewObjectKey();
                String a = "";
                oCompany.GetNewObjectCode(out a);
                if (sta != 0)
                {
                    sapDocEntry = "";
                    int erroCode = 0;
                    string errDescr = "";
                    oCompany.GetLastError(out erroCode, out errDescr);
                    flgStatus = false;
                    Response = String.Format("Error Posting:{0}", errDescr);
                }
                else
                {
                    flgStatus = true;
                    if (Update)
                    {
                        Response = String.Format("Updated", sapDocEntry);
                    }
                    else
                    {
                        Response = String.Format("Imported", sapDocEntry);
                    }
                }
            }

        }
        catch (Exception ex)
        {
            flgStatus = false;
            Response = String.Format("Error in Document:{0}", ex.Message);

        }

    }
    #endregion
    #region Sales Order
    private void DownloadDataSO(String MatxStr, SAPbouiCOM.Form form)
    {
        try
        {
            bool FlgStatus = false;
            bool FlgHasSelected = false;
            String Resp = "";
            String Primarykey = "";
            SAPbouiCOM.Matrix OMatrix = (SAPbouiCOM.Matrix)form.Items.Item(MatxStr).Specific;
            int RowCount = OMatrix.RowCount;
            if (RowCount > 0)
            {
                ShowMsgCustom("Downloading ...!", "S", "S");
                for (int i = RowCount - 1; i >= 0; i--)
                {
                    SalesOrder oSalesOrder = OutputListSO[i];
                    var CheckBoxMatrixState = (SAPbouiCOM.CheckBox)OMatrix.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    var InternalId = (SAPbouiCOM.EditText)OMatrix.Columns.Item("IntID").Cells.Item(i + 1).Specific;
                    if (CheckBoxMatrixState.Checked)
                    {
                        FlgHasSelected = true;
                        //Create Object
                        if (oSalesOrder.SalesOrderInternalId.Equals(InternalId.Value))
                        {
                            var Output = "";// GetSOOfConnect(oSalesOrder.SalesOrderNumber);
                            if (Output.Equals(""))
                            {
                                //Add
                                PostSO(oSalesOrder);
                            }
                            else
                            {
                                //Update
                                //    PostItem(oSalesOrder, ref FlgStatus, ref Resp, ref Primarykey, true);
                            }
                            //AddLog
                            #region Log


                            if (FlgStatus)
                            {
                                ShowMsgCustom(String.Format("Downloading of order:{0}", oSalesOrder.SalesOrderNumber), "S", "S");
                            }
                            else
                            {
                                ShowMsgCustom(String.Format("Downloading of order:{0} failed status:{1}", oSalesOrder.SalesOrderNumber, Resp), "S", "W");
                            }
                            //Update on Connect

                            #endregion
                        }
                    }
                }
                if (FlgHasSelected)
                {
                    ShowMsgCustom("Downloading done.", "S", "S");
                }
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private bool FindItemInSalesOrder(SalesOrder oSaleOrder, ref String Item)
    {
        bool flgItem = true;
        try
        {
            for (int i = 0; i < oSaleOrder.Items.Length; i++)
            {
                if (FindItemByID(oSaleOrder.Items[i].ItemId.Trim()).Length == 0)
                {
                    Item = oSaleOrder.Items[i].ItemId;
                    return false;
                }
            }
        }
        catch (Exception e)
        {
            errorLog(e);
        }
        return flgItem;
    }
    private bool PostSO(SalesOrder oSaleOrder)
    {
        bool flg = false;
        string item = "";
        bool ValidFlag = false;
        String ValidMsg = "";
        int DocEntry = 0;

        try
        {
            oSaleOrder.CustomerId = oSaleOrder.CustomerId.Trim();
            if (FindCustomerByID(oSaleOrder.CustomerId.Trim()).ToString().Length > 1)
            {
                if (FindItemInSalesOrder(oSaleOrder, ref item))
                {
                    if (GetSOOfConnect(oSaleOrder.SalesOrderNumber.Trim()).Equals(""))
                    {
                        //Add
                        CreateSOObject(oSaleOrder, ref ValidMsg, ref ValidFlag, ref DocEntry);
                        #region Update Sales order to connect
                        var response = "";
                        if (ValidFlag)
                        {
                            response = string.Format("Sales order#: {0} posted sucessfully...! ", DocEntry);
                            //Update connect SAPB1
                            oSaleOrder.Software = Software;
                            SalesOrderResponse oSOResp = ebiz.UpdateSalesOrder(getToken(), oSaleOrder, oSaleOrder.CustomerId, "", oSaleOrder.SalesOrderNumber, oSaleOrder.SalesOrderInternalId);

                        }
                        else
                        {
                            response = string.Format("Posting failed Error:{0} ", ValidMsg);
                        }
                        #endregion
                        #region Log
                        LogTable oLogTable = new LogTable();
                        oLogTable.DocEntry = oSaleOrder.SalesOrderNumber;
                        oLogTable.DocNum = DocEntry.ToString();
                        oLogTable.CardCode = oSaleOrder.CustomerId;
                        oLogTable.CardName = FindCustomerNameByID(oSaleOrder.CustomerId);
                        oLogTable.DatePaid = oSaleOrder.Date;
                        oLogTable.Software = oSaleOrder.Software;
                        oLogTable.Status = response;
                        oLogTable.CreateDt = DateTime.Now;
                        oLogTable.UpdateDt = DateTime.Now;
                        AddUploadDataLog(oLogTable, "CCSOIMP");
                        #endregion
                    }
                    else
                    {
                        //Add
                        CreateSOObject(oSaleOrder, ref ValidMsg, ref ValidFlag, ref DocEntry, "Update");
                        #region Update Sales order to connect
                        var response = "";
                        if (ValidFlag)
                        {
                            response = string.Format("Sales order#: {0} Updated sucessfully...! ", DocEntry);
                            //Update connect SAPB1
                            oSaleOrder.Software = Software;
                            SalesOrderResponse oSOResp = ebiz.UpdateSalesOrder(getToken(), oSaleOrder, oSaleOrder.CustomerId, "", oSaleOrder.SalesOrderNumber, oSaleOrder.SalesOrderInternalId);

                        }
                        else
                        {
                            response = string.Format("Updating failed Error:{0} ", ValidMsg);
                        }
                        #endregion
                        #region Log
                        LogTable oLogTable = new LogTable();
                        oLogTable.DocEntry = oSaleOrder.SalesOrderNumber;
                        oLogTable.DocNum = DocEntry.ToString();
                        oLogTable.CardCode = oSaleOrder.CustomerId;
                        oLogTable.CardName = FindCustomerNameByID(oSaleOrder.CustomerId);
                        oLogTable.DatePaid = oSaleOrder.Date;
                        oLogTable.Software = oSaleOrder.Software;
                        oLogTable.Status = response;
                        oLogTable.CreateDt = DateTime.Now;
                        oLogTable.UpdateDt = DateTime.Now;
                        AddUploadDataLog(oLogTable, "CCSOIMP");
                        #endregion


                    }
                }
                else
                {
                    #region Log
                    LogTable oLogTable = new LogTable();
                    oLogTable.Software = oSaleOrder.Software;
                    oLogTable.Status = String.Format("Item: {0} dont' exist in SAP B1", item);
                    oLogTable.CreateDt = DateTime.Now;
                    oLogTable.UpdateDt = DateTime.Now;
                    AddUploadDataLog(oLogTable, "CCSOIMP");
                    #endregion

                }
            }
            else
            {
                #region Log
                LogTable oLogTable = new LogTable();
                oLogTable.Software = oSaleOrder.Software;
                oLogTable.Status = String.Format("Customer: {0} dont' exist in SAP B1", oSaleOrder.CustomerId);
                oLogTable.CreateDt = DateTime.Now;
                oLogTable.UpdateDt = DateTime.Now;
                AddUploadDataLog(oLogTable, "CCSOIMP");
                #endregion
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return flg;
    }
    #endregion
    #region Sales OrderPayment
    private void DownloadDataSOP(String MatxStr, SAPbouiCOM.Form form)
    {
        try
        {
            bool FlgStatus = false;
            bool FlgHasSelected = false;
            String Resp = "";
            SAPbouiCOM.Matrix OMatrix = (SAPbouiCOM.Matrix)form.Items.Item(MatxStr).Specific;
            int RowCount = OMatrix.RowCount;
            if (RowCount > 0)
            {
                ShowMsgCustom("Downloading ...!", "S", "S");
                for (int i = RowCount - 1; i >= 0; i--)
                {
                    ApplicationTransactionDetails oTransaction = TransactionList.ApplicationTransactions[i];
                    var CheckBoxMatrixState = (SAPbouiCOM.CheckBox)OMatrix.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    var InternalId = (SAPbouiCOM.EditText)OMatrix.Columns.Item("IntID").Cells.Item(i + 1).Specific;
                    if (CheckBoxMatrixState.Checked)
                    {
                        FlgHasSelected = true;
                        //Create Object
                        if (oTransaction.ApplicationTransactionInternalId.Equals(InternalId.Value))
                        {
                            var Output = "";// GetSOOfConnect(oSalesOrder.SalesOrderNumber);
                            if (Output.Equals(""))
                            {
                                //Add
                                PostSOPayment(oTransaction);
                            }
                            else
                            {
                                //Update
                                //    PostItem(oSalesOrder, ref FlgStatus, ref Resp, ref Primarykey, true);
                            }
                            //AddLog
                            #region Log


                            if (FlgStatus)
                            {
                                ShowMsgCustom(String.Format("Downloading of payment:{0}", oTransaction.LinkedToTypeId), "S", "S");
                            }
                            else
                            {
                                ShowMsgCustom(String.Format("Downloading of payment:{0} failed status:{1}", oTransaction.LinkedToTypeId, Resp), "S", "W");
                            }
                            //Update on Connect

                            #endregion
                        }
                    }
                }
                if (FlgHasSelected)
                {
                    ShowMsgCustom("Downloading done.", "S", "S");
                }
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private bool PostSOPayment(ApplicationTransactionDetails oApptransaction)
    {
        bool flg = false;
        string item = "";
        bool ValidFlag = false;
        String ValidMsg = "";
        int DocEntry = 0;
        string ArDownPaymentDocNum = "";
        try
        {
            TransactionObject trasactionObj = ebiz.GetTransactionDetails(getToken(), oApptransaction.TransactionId);
            if (FindCustomerByID(trasactionObj.CustomerID.Trim()).ToString().Length > 1)
            {
                SalesOrder oSaleOrder = ebiz.GetSalesOrder(getToken(), trasactionObj.CustomerID, "", oApptransaction.LinkedToTypeId, oApptransaction.LinkedToInternalId);
                if (getSODocEntry(oSaleOrder.SalesOrderNumber.Trim())=="") {
                    //So dont exist
                    return false;
                }
                if (FindItemInSalesOrder(oSaleOrder, ref item))
                {
                    var GetSo = GetSOOfConnect(oSaleOrder.SalesOrderNumber.Trim());
                    if (GetSo.Equals("NoResult") || GetSo.Equals(""))
                    {
                        //Add Payment if sales order
                        if (oApptransaction.TransactionTypeId.Equals("Sale") || oApptransaction.TransactionTypeId.Equals("sales"))
                        {
                            PostArDownpaymeny(oSaleOrder, ref ValidMsg, ref ValidFlag, ref ArDownPaymentDocNum);
                            if (ValidFlag)
                            {
                                SalesOrder oOrder = ebiz.GetSalesOrder(getToken(), oSaleOrder.CustomerId.Trim(), "", oSaleOrder.SalesOrderNumber.Trim().ToString(), "");
                                oOrder.AmountDue = oOrder.AmountDue-Convert.ToDecimal(oApptransaction.TransactionAmount);
                                SalesOrderResponse resp = ebiz.UpdateSalesOrder(getToken(), oSaleOrder, oSaleOrder.CustomerId.Trim(), "", oSaleOrder.SalesOrderNumber.Trim().ToString(), "");
                                updateCCConnectSalesOrder(oSaleOrder.SalesOrderNumber.Trim().ToString(), oApptransaction.TransactionAmount, oOrder.AmountDue.ToString());
                                showMessage(String.Format("Payment downloaded from Connect to SAPB1 against sales order:{0}", oSaleOrder.SalesOrderNumber.Trim()));
                                trace(String.Format("Payment downloaded from Connect to SAPB1 against sales order:{0}", oSaleOrder.SalesOrderNumber.Trim()));
                            }
                        }//
                        else if (oApptransaction.TransactionTypeId.Equals("AuthOnly") || oApptransaction.TransactionTypeId.Equals("Auth Only"))
                        {
                            #region Pre Auth logic
                            String IsCCTransLogAlreadyPosted = getSOAuthonlyAlreadyPostedStatus(getDocEntry(oSaleOrder.SalesOrderNumber, "ORDR").ToString());
                            if (IsCCTransLogAlreadyPosted.Equals("") || IsCCTransLogAlreadyPosted.Equals("0"))
                            {
                                try
                                {
                                    #region CCtrans
                                    SalesOrder Osales = GetSO(oSaleOrder.SalesOrderNumber);

                                    TransactionResponse resp = trasactionObj.Response;

                                    CCTRAN cctran = new CCTRAN();
                                    cctran.OrderID = getDocEntry(oSaleOrder.SalesOrderNumber, "ORDR").ToString();// EdDocNum;
                                    cctran.InvoiceID = "";
                                    cctran.CreditMemoID = "";
                                    cctran.customerID = Osales.CustomerId;
                                    cctran.MethodID = "";
                                    cctran.custNum = "";
                                    string currency = getCurrency("ORDR", getSODocEntry(Osales.SalesOrderNumber));
                                    string group = getGroupName(Osales.CustomerId);
                                    string cardtype = trasactionObj.CreditCardData.CardType;
                                    string acctid = "";
                                    string cardName = getCardName(group, cardtype, currency, ref acctid);
                                    cctran.CCAccountID = int.Parse(acctid);

                                    cctran.crCardNum = trasactionObj.CreditCardData.CardNumber;
                                    cctran.Description = trasactionObj.Details.Description;
                                    cctran.recID = trasactionObj.Details.Invoice;

                                    cctran.acsUrl = resp.AcsUrl;
                                    cctran.authAmount = resp.AuthAmount.ToString();
                                    cctran.authCode = resp.AuthCode;
                                    cctran.avsResult = resp.AvsResult;
                                    cctran.avsResultCode = resp.AvsResultCode;
                                    cctran.batchNum = resp.BatchNum;
                                    cctran.batchRef = resp.BatchRefNum;
                                    cctran.cardCodeResult = resp.CardCodeResult;
                                    cctran.cardCodeResultCode = resp.CardCodeResultCode;
                                    cctran.cardLevelResult = resp.CardLevelResult;
                                    cctran.cardLevelResultCode = resp.CardLevelResultCode;
                                    cctran.conversionRate = resp.ConversionRate.ToString();
                                    cctran.convertedAmount = resp.ConvertedAmount.ToString();
                                    cctran.convertedAmountCurrency = resp.ConvertedAmountCurrency.ToString();
                                    cctran.custNum = resp.CustNum;
                                    cctran.error = resp.Error;
                                    cctran.errorCode = resp.ErrorCode;
                                    cctran.isDuplicate = resp.isDuplicate.ToString();
                                    cctran.payload = resp.Payload;
                                    cctran.profilerScore = resp.ProfilerScore;
                                    cctran.profilerResponse = resp.ProfilerResponse;
                                    cctran.profilerReason = resp.ProfilerReason;
                                    cctran.refNum = resp.RefNum;
                                    cctran.remainingBalance = resp.RemainingBalance.ToString();
                                    cctran.result = resp.Result;
                                    cctran.resultCode = resp.ResultCode;
                                    cctran.status = resp.Status;
                                    cctran.statusCode = resp.StatusCode;
                                    cctran.vpasResultCode = resp.VpasResultCode;
                                    cctran.recDate = DateTime.Now;//Use local time not server time
                                    cctran.command = trasactionObj.TransactionType;
                                    if (trasactionObj.TransactionType == "Auth Only" || trasactionObj.TransactionType.Equals("AuthOnly"))
                                        cctran.command = "cc:authonly";
                                    cctran.amount = trasactionObj.Details.Amount.ToString();

                                    insert(cctran);

                                    #endregion CCtrans
                                }
                                catch (Exception ex)
                                {
                                    ValidMsg = ex.Message;
                                    ValidFlag = false;
                                }
                            }
                            else
                            {
                                //  oEdit.Value = "Paid";
                                ValidMsg = "Already auth only transaction done";
                                ValidFlag = false;
                                var resp = ebiz.MarkApplicationTransactionAsApplied(getToken(), oApptransaction.ApplicationTransactionInternalId);
                            }
                            #endregion Pre Auth logic

                        }
                        #region Update Sales order to connect
                        var response = "";
                        if (ValidFlag)
                        {
                            response = string.Format("A/R Down Payment #: {0} posted sucessfully...! ", ArDownPaymentDocNum);
                            var resp = ebiz.MarkApplicationTransactionAsApplied(getToken(), oApptransaction.ApplicationTransactionInternalId);

                        }
                        else
                        {
                            response = string.Format("Posting failed Error:{0} ", ValidMsg);
                        }
                        #endregion
                        #region Log
                        LogTable oLogTable = new LogTable();
                        oLogTable.DocEntry = getDocEntry(oSaleOrder.SalesOrderNumber, "ORDR").ToString();
                        oLogTable.DocNum = oSaleOrder.SalesOrderNumber;
                        oLogTable.CardCode = oSaleOrder.CustomerId.Trim();
                        oLogTable.CardName = FindCustomerNameByID(oSaleOrder.CustomerId.Trim());
                        oLogTable.DatePaid = oSaleOrder.Date;
                        oLogTable.Software = oSaleOrder.Software.Trim();
                        oLogTable.Method = oApptransaction.TransactionTypeId.Trim();
                        oLogTable.AmountPaid = Convert.ToDecimal(trasactionObj.Details.Amount);
                        oLogTable.Status = response;
                        oLogTable.CreateDt = DateTime.Now;
                        oLogTable.UpdateDt = DateTime.Now;
                        AddUploadDataLog(oLogTable, "CCSOPIMP");
                        #endregion
                    }
                    else
                    {
                        #region Log
                        LogTable oLogTable = new LogTable();
                        oLogTable.DocEntry = oSaleOrder.SalesOrderNumber;
                        oLogTable.DocNum = DocEntry.ToString();
                        oLogTable.CardCode = oSaleOrder.CustomerId;
                        oLogTable.CardName = FindCustomerNameByID(oSaleOrder.CustomerId);
                        oLogTable.DatePaid = oSaleOrder.Date;
                        oLogTable.Software = oSaleOrder.Software;
                        oLogTable.Status = String.Format("Sales order: {0} dont' exist in SAP B1", oSaleOrder.SalesOrderNumber);
                        oLogTable.CreateDt = DateTime.Now;
                        oLogTable.UpdateDt = DateTime.Now;
                        AddUploadDataLog(oLogTable, "CCSOPIMP");
                        #endregion


                    }
                }
                else
                {
                    #region Log
                    LogTable oLogTable = new LogTable();
                    oLogTable.Software = oSaleOrder.Software;
                    oLogTable.Status = String.Format("Item: {0} dont' exist in SAP B1", item);
                    oLogTable.CreateDt = DateTime.Now;
                    oLogTable.UpdateDt = DateTime.Now;
                    AddUploadDataLog(oLogTable, "CCSOPIMP");
                    #endregion

                }
            }
            else
            {
                #region Log
                LogTable oLogTable = new LogTable();
                oLogTable.Software = oApptransaction.SoftwareId;
                oLogTable.Status = String.Format("Customer: {0} dont' exist in SAP B1", trasactionObj.CustomerID.Trim().ToString());
                oLogTable.CreateDt = DateTime.Now;
                oLogTable.UpdateDt = DateTime.Now;
                AddUploadDataLog(oLogTable, "CCSOPIMP");
                #endregion
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return flg;
    }

    private void ProcessPayment(TransactionObject trasactionObj, ApplicationTransactionDetails oAppTransaction)
    {
        try
        {
            LogTable oLogTable = new LogTable();
            if (trasactionObj != null)
            {
                var PaymentMethod = trasactionObj.TransactionType;
                if (PaymentMethod.Equals("Sale"))
                {
                    if (true)
                    {
                        SalesOrder Osales = GetSO(trasactionObj.Details.OrderID);
                        oLogTable.DocSent = "Sent";
                        oLogTable.DocNum = Osales.SalesOrderNumber;
                        oLogTable.DocNum = GetSalesOrderDocNum(Osales.SalesOrderNumber);
                        oLogTable.Email = Osales.EmailTemplateID;
                        oLogTable.CardCode = Osales.CustomerId;
                        oLogTable.CardName = FindCustomerNameByID(Osales.CustomerId);
                        oLogTable.Software = oAppTransaction.SoftwareId;
                        oLogTable.DatePaid = trasactionObj.DateTime;
                        oLogTable.Type = trasactionObj.TransactionType;
                        Osales.Amount = ConvertDecimal(trasactionObj.Details.Amount);
                        oLogTable.AmountPaid = ConvertDecimal(trasactionObj.Response.AuthAmount);
                        var Method4Digit = oAppTransaction.PaymentMethodType + " ending in " + oAppTransaction.PaymentMethodLast4;
                        oLogTable.Method = Method4Digit;

                        bool flgPostingStatus = false;
                        String TransationId = "";
                        String Msg = "";
                        PostArDownpaymeny(Osales, ref Msg, ref flgPostingStatus, ref TransationId);
                        if (flgPostingStatus)
                        {
                            oLogTable.Status = "AR down payment posted Successfully";
                            ShowSystemLog(String.Format("Payment Received Action: {1} against Sales Order: {0} Status:Done", Osales.SalesOrderNumber, "AR DownPayment"));

                            var resp = ebiz.MarkApplicationTransactionAsApplied(getToken(), oAppTransaction.TransactionId);
                        }
                        oLogTable.Status = String.Format("AR DownPayment posting failed:{0}", Msg);
                        AddUploadDataLog(oLogTable, "CCSOPIMP");

                    }
                    else
                    {


                        //                                        oEdit.Value = "Paid";
                    }
                }
                else
                {
                    String IsCCTransLogAlreadyPosted = getSOAuthonlyAlreadyPostedStatus(getDocEntry(trasactionObj.Details.OrderID, "ORDR").ToString());
                    if (IsCCTransLogAlreadyPosted.Equals("") || IsCCTransLogAlreadyPosted.Equals("0"))
                    {
                        #region CCtrans
                        SalesOrder Osales = GetSO(trasactionObj.Details.OrderID);
                        TransactionResponse resp = trasactionObj.Response;

                        CCTRAN cctran = new CCTRAN();
                        cctran.OrderID = getDocEntry(trasactionObj.Details.OrderID, "ORDR").ToString();// EdDocNum;
                        cctran.InvoiceID = "";
                        cctran.CreditMemoID = "";
                        cctran.customerID = Osales.CustomerId;
                        cctran.MethodID = "";
                        cctran.custNum = "";
                        string currency = getCurrency("ORDR", getSODocEntry(Osales.SalesOrderNumber));
                        string group = getGroupName(Osales.CustomerId);
                        string cardtype = trasactionObj.CreditCardData.CardType;
                        string acctid = "";
                        string cardName = getCardName(group, cardtype, currency, ref acctid);
                        cctran.CCAccountID = int.Parse(acctid);

                        cctran.crCardNum = trasactionObj.CreditCardData.CardNumber;
                        cctran.Description = trasactionObj.Details.Description;
                        cctran.recID = trasactionObj.Details.Invoice;

                        cctran.acsUrl = resp.AcsUrl;
                        cctran.authAmount = resp.AuthAmount.ToString();
                        cctran.authCode = resp.AuthCode;
                        cctran.avsResult = resp.AvsResult;
                        cctran.avsResultCode = resp.AvsResultCode;
                        cctran.batchNum = resp.BatchNum;
                        cctran.batchRef = resp.BatchRefNum;
                        cctran.cardCodeResult = resp.CardCodeResult;
                        cctran.cardCodeResultCode = resp.CardCodeResultCode;
                        cctran.cardLevelResult = resp.CardLevelResult;
                        cctran.cardLevelResultCode = resp.CardLevelResultCode;
                        cctran.conversionRate = resp.ConversionRate.ToString();
                        cctran.convertedAmount = resp.ConvertedAmount.ToString();
                        cctran.convertedAmountCurrency = resp.ConvertedAmountCurrency.ToString();
                        cctran.custNum = resp.CustNum;
                        cctran.error = resp.Error;
                        cctran.errorCode = resp.ErrorCode;
                        cctran.isDuplicate = resp.isDuplicate.ToString();
                        cctran.payload = resp.Payload;
                        cctran.profilerScore = resp.ProfilerScore;
                        cctran.profilerResponse = resp.ProfilerResponse;
                        cctran.profilerReason = resp.ProfilerReason;
                        cctran.refNum = resp.RefNum;
                        cctran.remainingBalance = resp.RemainingBalance.ToString();
                        cctran.result = resp.Result;
                        cctran.resultCode = resp.ResultCode;
                        cctran.status = resp.Status;
                        cctran.statusCode = resp.StatusCode;
                        cctran.vpasResultCode = resp.VpasResultCode;
                        cctran.recDate = DateTime.Now;//Use local time not server time
                        cctran.command = trasactionObj.TransactionType;
                        if (trasactionObj.TransactionType == "Auth Only")
                            cctran.command = "cc:authonly";
                        cctran.amount = trasactionObj.Details.Amount.ToString();

                        insert(cctran);
                        #endregion CCtrans

                        oLogTable.DocSent = "Sent";
                        oLogTable.DocEntry = Osales.SalesOrderNumber;
                        oLogTable.DocNum = GetSalesOrderDocNum(Osales.SalesOrderNumber);
                        oLogTable.CardCode = Osales.CustomerId;
                        oLogTable.CardName = FindCustomerNameByID(Osales.CustomerId);
                        oLogTable.Software = oAppTransaction.SoftwareId;
                        oLogTable.DatePaid = trasactionObj.DateTime;
                        oLogTable.Type = trasactionObj.TransactionType;
                        Osales.Amount = ConvertDecimal(trasactionObj.Details.Amount);
                        oLogTable.AmountPaid = ConvertDecimal(trasactionObj.Response.AuthAmount);
                        var Method4Digit = oAppTransaction.PaymentMethodType + " ending in " + oAppTransaction.PaymentMethodLast4;
                        oLogTable.Method = Method4Digit;
                        AddUploadDataLog(oLogTable, "CCSOPIMP");
                    }
                    else
                    {
                        //  oEdit.Value = "Paid";
                    }
                }

            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    #endregion

    #endregion

    #region Log
    //Business Partner
    private void ShowSyncDownloadBPLOG(SAPbouiCOM.Form form, String Where = "")
    {
        string sql = "";
        var MtxID = "mtxBPLog";
        try
        {
            if (form.Items.Item(MtxID) == null)
                return;

            form.Items.Item(MtxID).AffectsFormMode = false;
            OMtxSyncBPLog = (SAPbouiCOM.Matrix)form.Items.Item(MtxID).Specific;
            OMtxSyncBPLog.Clear();
            odtSyncBPLog = getDataTableSyncUpload(form, MtxID);
            sql = String.Format("select ROW_NUMBER() OVER (ORDER BY ]Code]) AS ]ID],[U_CardCode] as [CardCode],[U_CardName] as [CardName],[U_Email] as [Email],[U_Balance] as [Balance],[U_Software] as [Software],[U_Status] as [Status],[U_CreateDt]  as [CreateDt] from [@CCBPIMP] order by [Code]").Replace("]", "\"").Replace("[", "\"");

            odtSyncBPLog.ExecuteQuery(sql);
            BindMatrixSyncUpload(MtxID, OMtxSyncBPLog, "clNo", "ID");
            BindMatrixSyncUpload(MtxID, OMtxSyncBPLog, "CardCode", "CardCode");
            BindMatrixSyncUpload(MtxID, OMtxSyncBPLog, "CardName", "CardName");
            BindMatrixSyncUpload(MtxID, OMtxSyncBPLog, "Email", "Email");
            BindMatrixSyncUpload(MtxID, OMtxSyncBPLog, "Balance", "Balance");
            BindMatrixSyncUpload(MtxID, OMtxSyncBPLog, "Software", "Software");
            BindMatrixSyncUpload(MtxID, OMtxSyncBPLog, "Status", "Status");//
            BindMatrixSyncUpload(MtxID, OMtxSyncBPLog, "DownDt", "CreateDt");
            OMtxSyncBPLog.LoadFromDataSource();
            OMtxSyncBPLog.AutoResizeColumns();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void ShowSyncDownloadItemLOG(SAPbouiCOM.Form form, String Where = "")
    {
        string sql = "";
        var MtxID = "mtxItmLog";
        SAPbouiCOM.Matrix oMtx;

        try
        {
            if (form.Items.Item(MtxID) == null)
                return;

            form.Items.Item(MtxID).AffectsFormMode = false;
            oMtx = (SAPbouiCOM.Matrix)form.Items.Item(MtxID).Specific;
            oMtx.Clear();
            odtSyncItemLog = getDataTableSyncUpload(form, MtxID);
            sql = String.Format("select ROW_NUMBER() OVER (ORDER BY ]Code]) AS ]ID],[U_ItemCode] as [ItemCode],[U_ItemName] as [ItemName],[U_Quantity] as [Quantity],[U_UnitPrice] as [UnitPrice],[U_Software] as [Software],[U_Status] as [Status],[U_CreateDt] as [CreateDt] from [@CCITMIMP] order by [Code]").Replace("]", "\"").Replace("[", "\"");

            odtSyncItemLog.ExecuteQuery(sql);
            BindMatrixSyncUpload(MtxID, oMtx, "clNo", "ID");
            BindMatrixSyncUpload(MtxID, oMtx, "ItemCode", "ItemCode");
            BindMatrixSyncUpload(MtxID, oMtx, "ItemName", "ItemName");
            BindMatrixSyncUpload(MtxID, oMtx, "Quantity", "Quantity");
            BindMatrixSyncUpload(MtxID, oMtx, "UnitPrice", "UnitPrice");
            BindMatrixSyncUpload(MtxID, oMtx, "Software", "Software");
            BindMatrixSyncUpload(MtxID, oMtx, "Status", "Status");
            BindMatrixSyncUpload(MtxID, oMtx, "DownDt", "CreateDt");
            oMtx.LoadFromDataSource();
            oMtx.AutoResizeColumns();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void ShowSyncDownloadSOLOG(SAPbouiCOM.Form form, String Where = "")
    {
        string sql = "";
        var MtxID = "mtxSOLog";
        SAPbouiCOM.Matrix oMtx;

        try
        {
            if (form.Items.Item(MtxID) == null)
                return;

            form.Items.Item(MtxID).AffectsFormMode = false;
            oMtx = (SAPbouiCOM.Matrix)form.Items.Item(MtxID).Specific;
            oMtx.Clear();
            odtSyncSOFrmConLog = getDataTableSyncUpload(form, MtxID);//[U_DocNum],[U_DocEntry],[U_CardCode],[U_Customer],[U_DatePaid],[U_Software],[U_Status],[U_CreateDt]
            sql = String.Format("select ROW_NUMBER() OVER (ORDER BY ]Code]) AS ]ID],[U_DocNum] as [DocNum],[U_DocEntry] as [DocEntry],[U_CardCode] as [CardCode],[U_Customer] as [CardName],[U_DatePaid] as [DatePaid],[U_Software] as [Software],[U_Status] as [Status],[U_CreateDt] as [CreateDt] from [@CCSOIMP] order by [Code]").Replace("]", "\"").Replace("[", "\"");

            odtSyncSOFrmConLog.ExecuteQuery(sql);
            BindMatrixSyncUpload(MtxID, oMtx, "clNo", "ID");
            BindMatrixSyncUpload(MtxID, oMtx, "DocNum", "DocNum");
            BindMatrixSyncUpload(MtxID, oMtx, "DocEntry", "DocEntry");
            BindMatrixSyncUpload(MtxID, oMtx, "CardCode", "CardCode");
            BindMatrixSyncUpload(MtxID, oMtx, "CardName", "CardName");
            BindMatrixSyncUpload(MtxID, oMtx, "Software", "Software");
            BindMatrixSyncUpload(MtxID, oMtx, "Status", "Status");
            BindMatrixSyncUpload(MtxID, oMtx, "DatePaid", "DatePaid");
            oMtx.LoadFromDataSource();
            oMtx.AutoResizeColumns();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void ShowSyncDownloadSOPLOG(SAPbouiCOM.Form form, String Where = "")
    {
        string sql = "";
        var MtxID = "MtxSOPlg";
        SAPbouiCOM.Matrix oMtx;

        try
        {
            if (form.Items.Item(MtxID) == null)
                return;

            form.Items.Item(MtxID).AffectsFormMode = false;
            oMtx = (SAPbouiCOM.Matrix)form.Items.Item(MtxID).Specific;
            oMtx.Clear();
            odtSyncSOFrmConLog = getDataTableSyncUpload(form, MtxID);//[U_DocNum],[U_DocEntry],[U_CardCode],[U_Customer],[U_DatePaid],[U_Software],[U_Status],[U_CreateDt]
            sql = String.Format("select ROW_NUMBER() OVER (ORDER BY ]Code]) AS ]ID],[U_DocNum] as [DocNum],[U_DocEntry] as [DocEntry],[U_CardCode] as [CardCode],[U_Customer] as [CardName],[U_DatePaid] as [DatePaid],[U_Software] as [Software],[U_Status] as [Status],[U_CreateDt] as [CreateDt],[U_PaymentType] as [PaymentType],[U_AmtPaid] as [AmtPaid]  from [@CCSOPIMP] order by [Code]").Replace("]", "\"").Replace("[", "\"");

            odtSyncSOFrmConLog.ExecuteQuery(sql);
            BindMatrixSyncUpload(MtxID, oMtx, "clNo", "ID");
            BindMatrixSyncUpload(MtxID, oMtx, "DocNum", "DocNum");
            BindMatrixSyncUpload(MtxID, oMtx, "DocEntry", "DocEntry");
            BindMatrixSyncUpload(MtxID, oMtx, "CardCode", "CardCode");
            BindMatrixSyncUpload(MtxID, oMtx, "Customer", "CardName");
            BindMatrixSyncUpload(MtxID, oMtx, "PymtType", "PaymentType");
            BindMatrixSyncUpload(MtxID, oMtx, "DatePaid", "DatePaid");
            BindMatrixSyncUpload(MtxID, oMtx, "AmtPaid", "AmtPaid");

            BindMatrixSyncUpload(MtxID, oMtx, "ImpSource", "Software");
            BindMatrixSyncUpload(MtxID, oMtx, "ImpStatus", "Status");

            oMtx.LoadFromDataSource();
            oMtx.AutoResizeColumns();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    #endregion

    #region  Export Button Functionality
    private void FileBrowserSyncDown(String Button, SAPbouiCOM.Form form)
    {
        try
        {
            SAPbouiCOM.Matrix oMtx = null;
            String FileName = "";
            String MatxID = "";
            if (Button.Equals("btnExp11"))
            {
                FileName = "BusinessDownload";
                MatxID = "mtxBPLst";
                oMtx = (SAPbouiCOM.Matrix)form.Items.Item(MatxID).Specific;
            }
            else if (Button.Equals("btnExp12"))
            {
                FileName = "BusinessDownloadLog";
                MatxID = "mtxBPLog";
                oMtx = (SAPbouiCOM.Matrix)form.Items.Item(MatxID).Specific;
            }
            else if (Button.Equals("btnExp21"))
            {
                FileName = "ItemDownload";
                MatxID = "mtxItmLst";
                oMtx = (SAPbouiCOM.Matrix)form.Items.Item(MatxID).Specific;
            }
            else if (Button.Equals("btnExp22"))
            {
                FileName = "ItemDownloadLog";
                MatxID = "mtxItmLog";
                oMtx = (SAPbouiCOM.Matrix)form.Items.Item(MatxID).Specific;
            }
            else if (Button.Equals("btnExp31"))
            {
                FileName = "SalesOrderDownload";
                MatxID = "MtxSOList";
                oMtx = (SAPbouiCOM.Matrix)form.Items.Item(MatxID).Specific;
            }
            else if (Button.Equals("btnExp32"))
            {
                FileName = "SalesOrderDownloadLog";
                MatxID = "mtxSOLog";
                oMtx = (SAPbouiCOM.Matrix)form.Items.Item(MatxID).Specific;
            }
            else if (Button.Equals("btnExp41"))
            {
                FileName = "SalesOrderPaymentDownload";
                MatxID = "mtxSOP";
                oMtx = (SAPbouiCOM.Matrix)form.Items.Item(MatxID).Specific;
            }
            else if (Button.Equals("btnExp42"))
            {
                FileName = "SalesOrderPaymentDownloadLog";
                MatxID = "MtxSOPlg";
                oMtx = (SAPbouiCOM.Matrix)form.Items.Item(MatxID).Specific;
            }
            if (oMtx != null)
            {
                SelectFileDialog dialog = new SelectFileDialog("C:\\", FileName,
                    "|.csv*", DialogType.SAVE);
                dialog.Open();
                if (!string.IsNullOrEmpty(dialog.SelectedFile))
                {
                    WriteToCSVSyncDown(oMtx, dialog.SelectedFile + ".csv");
                }
            }
        }
        catch (Exception ex)
        {


        }
    }
    private void WriteToCSVSyncDown(SAPbouiCOM.Matrix oMatrix, string FileName)
    {
        try
        {
            System.Text.StringBuilder sb = new StringBuilder();
            var ColList = (SAPbouiCOM.Columns)oMatrix.Columns;
            var ColumeCount = oMatrix.Columns.Count;
            int RowCount = oMatrix.RowCount;
            //Set colume header
            for (int H = 0; H < ColumeCount; H++)
            {
                try
                {
                    var ColDescp = ColList.Item(H).Description;
                    var cell = CSVColumeRules(ColDescp);
                    sb.Append(cell + ",");
                }
                catch { }
            }
            sb.AppendLine();
            if (RowCount > 0)
            {
                ShowMsgCustom("Downloading ...!", "S", "S");
                for (int i = 0; i < RowCount; i++)
                {
                    for (int j = 0; j < ColumeCount; j++)
                    {
                        try
                        {
                            var ColUnique = ColList.Item(j).UniqueID;
                            var Oitem = (SAPbouiCOM.EditText)oMatrix.Columns.Item(ColUnique).Cells.Item(i + 1).Specific;
                            var itemvalue = Oitem.Value;
                            if (j == 0 && itemvalue.ToString().Equals(""))
                            {
                                itemvalue = (i + 1).ToString();
                            }
                            var cell = CSVColumeRules(itemvalue);
                            sb.Append(cell + ",");
                        }
                        catch { }
                    }
                    sb.AppendLine();
                }
            }
            System.IO.File.WriteAllText(FileName, sb.ToString());
        }
        catch (Exception ex)
        {
            showMessage("Export File Issue" + ex.Message);
            SBO_Application.MessageBox("Export File Issue" + ex.Message);

        }

    }

    #endregion Export Button Functionality

    #region Clear 
    private void ClearLog(String Table)
    {
        try
        {
            List<String> KeyList = GetAllKeyList(Table);
            if (KeyList != null & KeyList.Count > 0)
            {
                ShowSystemLog("Clearing Log...!");
                for (int i = 0; i < KeyList.Count; i++)
                {
                    SAPbobsCOM.UserTable oUsrTbl = oCompany.UserTables.Item(Table);
                    oUsrTbl.GetByKey(KeyList[i]);
                    int Res = oUsrTbl.Remove();
                    int erroCode = 0;
                    string errDescr = "";
                    if (Res != 0)
                    {
                        oCompany.GetLastError(out erroCode, out errDescr);
                        trace(errDescr);
                    }

                }
                ShowSystemLog("Clearing Log Done...!");
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    #endregion
}

