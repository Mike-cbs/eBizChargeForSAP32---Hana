using eBizChargeForSAP.Screen;
using eBizChargeForSAP.ServiceReference1;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

partial class SAP
{

    #region DataMember
    bool flgItemfolder = true, flgBPfolder = true, flgInvfolder = true, flgReceivedfolder = true;

    SAPbouiCOM.Form FrmSyncUpload;
    bool SyncUploadflgDocInitiallization = false;
    string MtxSyncUpItemListStr = "mtxSUL1", MtxSyncUpBPListStr = "mtxSUB1", MtxSyncUpSOListStr = "mtxSUS1", MtxSyncUpInvoiceListStr = "mtxILog";
    SAPbouiCOM.Matrix oMatrixSyncUploadItemList1, oMatrixSyncUploadBPList1, oMatrixSyncUploadSOList1, oMatrixSyncUploadInvoiceList1;
    SAPbouiCOM.DataTable oDTSyncUploadItemList1, oDTSyncUploadBPList1, oDTSyncUploadSOList1, oDTSyncUploadInvList1;
    string MtxSyncUpSOListStrLOG = "mtxSO2", MtxSyncUpItemListStrLOG = "mtxSUL2", MtxSyncUpBPListStrLOG = "mtxSUB2", MtxSyncUpInvoiceListStrLOG = "mtxInLog";
    SAPbouiCOM.Matrix oMatrixSyncUploadItemList1LOG, oMatrixSyncUploadBPList1LOG, oMatrixSyncUploadSOList1LOG, oMatrixSyncUploadInvoiceList1LOG;
    SAPbouiCOM.DataTable oDTSyncUploadItemList1LOG, oDTSyncUploadBPList1LOG, oDTSyncUploadSOList1LOG, oDTSyncUploadInvoiceList1LOG;

    DataTable oDataTable;
    SAPbouiCOM.EditText ed_SUFromDate, ed_SUToDate;
    SAPbouiCOM.EditText ed_InFromDate, ed_InToDate;
    String ed_SUFromDateString = "edfdt", ed_SUToDateString = "edtdt";
    String ed_InFromDateString = "edfdt2", ed_InToDateString = "edtdt2";
    SAPbouiCOM.ComboBox Cmb_CustomerGrpSU, Cmb_PaymentsTermSU, Cmb_CustomerGrpSU2, Cmb_PaymentsTermSU2;
    SAPbouiCOM.CheckBox oCheckboxSOSelectAll, oCheckboxItemSelectAll, oCheckboxBPSelectAll;
    String oCheckboxSOSelectAllstring = "ckSOL1", oCheckboxItemSelectAllstring = "ckItmL1", oCheckboxBPSelectAllstring = "chkBPL1", oCheckboxInvoiceSelectAllstring = "ckInL1";

    SAPbouiCOM.Folder oFolder;

    #endregion

    #region FormEvent
    private void CreateSyncUpload()
    {
        try
        {
            SyncUploadflgDocInitiallization = false;
            AddXMLForm("Presentation_Layer.SyncUpload.xml");
            FrmSyncUpload.Freeze(true);
            InitializeSyncupload();
            FrmSyncUpload.Freeze(false);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void SyncUploadFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            FrmSyncUpload = form;
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
                        if (pVal.ItemUID.Equals(oCheckboxSOSelectAllstring))
                        {
                            SelectAllSU(form, MtxSyncUpSOListStr, oDTSyncUploadSOList1, oCheckboxSOSelectAllstring);
                        }
                        if (pVal.ItemUID.Equals(oCheckboxInvoiceSelectAllstring))
                        {
                            SelectAllSU(form, MtxSyncUpInvoiceListStr, oDTSyncUploadInvList1, oCheckboxInvoiceSelectAllstring);
                        }
                        if (pVal.ItemUID.Equals(oCheckboxItemSelectAllstring))
                        {
                            SelectAllSU(form, MtxSyncUpItemListStr, oDTSyncUploadItemList1, oCheckboxItemSelectAllstring);
                        }
                        if (pVal.ItemUID.Equals(oCheckboxBPSelectAllstring))
                        {
                            SelectAllSU(form, MtxSyncUpBPListStr, oDTSyncUploadBPList1, oCheckboxBPSelectAllstring);
                        }
                        if (pVal.ItemUID.Equals("btnU1"))
                        {//Btn Sales order Upload
                            UploadSOSyncObect(form);
                        }
                        if (pVal.ItemUID.Equals("btnUpInv"))
                        {//Btn Invoice Upload
                            UploadInvoiceSyncObect(form);
                        }

                        if (pVal.ItemUID.Equals("btnIU1"))
                        {//Btn Item  Upload
                            UploadItemSyncObect(form);
                        }
                        if (pVal.ItemUID.Equals("BtnBpU1"))
                        {//Btn Business Partner Upload
                            UploadBPSyncObect(form);
                        }
                        if (pVal.ItemUID.Equals("btnCSO") || pVal.ItemUID.Equals("btnCI2") || pVal.ItemUID.Equals("btnBPClear") || pVal.ItemUID.Equals("btnCInv"))
                        {
                            ClearLogItem(pVal.ItemUID);
                        }//btnInvEL
                        if (pVal.ItemUID.Equals("btnInvEL") || pVal.ItemUID.Equals("btnEInv") || pVal.ItemUID.Equals("btnESO1") || pVal.ItemUID.Equals("btnESO2") || pVal.ItemUID.Equals("btnEI1") || pVal.ItemUID.Equals("btnEI2") || pVal.ItemUID.Equals("btnBPE1") || pVal.ItemUID.Equals("btnEBp2"))
                        {
                            FileBrowser(pVal.ItemUID);
                        }
                        if (pVal.ItemUID.Equals("btnfind1"))
                        {
                            FindFunctionality(form);
                        }
                        if (pVal.ItemUID.Equals("btnfind2"))
                        {
                            FindFunctionality_Invoice(form);
                        }
                        if (pVal.ItemUID == "btnPost")
                        {
                            HandleItemPressItemExport(form);
                        }
                        if (pVal.ItemUID.Equals("btnRngClr"))
                        {
                            BtnClearRangeSO();
                        }
                        if (pVal.ItemUID.Equals("btnClr2"))
                        {
                            BtnClearRangeInvoice();
                        }
                        if (pVal.ItemUID.Equals("TabSOLog"))//SalesOrder Log Tab
                        {
                            ShowSyncUploadSOLOG(FrmSyncUpload);
                        }
                        if (pVal.ItemUID.Equals("TabSO") || pVal.ItemUID.Equals("TabSOList"))//SalesOrder List Tab
                        {
                            ShowSyncUploadSO(FrmSyncUpload);
                        }

                        if (pVal.ItemUID.Equals("TabItemlog"))//Item Log Tab
                        {
                            ShowSyncUpLoadItemListLOG(FrmSyncUpload);
                        }
                        if (pVal.ItemUID.Equals("TabItem") || pVal.ItemUID.Equals("TabItmLst"))//Item List Tab
                        {
                            if (flgItemfolder)
                            {
                                flgItemfolder = false;
                                SetFolder(FrmSyncUpload, "TabItem");
                                SetFolder(FrmSyncUpload, "TabItmLst");
                            }
                            ShowSyncUpLoadItemList(FrmSyncUpload);
                        }
                        //BP  Tab
                        if (pVal.ItemUID.Equals("TabBPLog"))
                        {
                            ShowSyncUploadBPListLOG(FrmSyncUpload);
                        }
                        if (pVal.ItemUID.Equals("TabBP") || pVal.ItemUID.Equals("TabBPList"))//BP List Tab
                        {
                            if (flgBPfolder)
                            {
                                flgBPfolder = false;
                                SetFolder(FrmSyncUpload, "TabBP");
                                SetFolder(FrmSyncUpload, "TabBPList");
                            }
                            ShowSyncUploadBPList(FrmSyncUpload);
                        }
                        if (pVal.ItemUID.Equals("TabInv") || pVal.ItemUID.Equals("TBInLst"))//Invoice List Tab
                        {
                            if (flgInvfolder)
                            {
                                flgInvfolder = false;
                                SetFolder(FrmSyncUpload, "TabInv");
                                SetFolder(FrmSyncUpload, "TBInLst");
                            }
                            ShowSyncUploadInvoice(FrmSyncUpload);
                        }
                        if (pVal.ItemUID.Equals("TbIInLog"))//SalesOrder Log Tab
                        {
                            ShowSyncUploadInvoiceLOG(FrmSyncUpload);
                        }
                        if (pVal.ItemUID == MtxItem && pVal.ColUID.Equals("clSel") && pVal.Row == 0 && !pVal.BeforeAction)
                        {
                            HandleExportItemSelectAll(form, MtxItem, oDTItemExport);
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                        try
                        {
                            if (pVal.ItemUID == MtxSyncUpSOListStr & pVal.ColUID == "DocNum")
                            {//SO
                                SAPbouiCOM.EditText DocNum = (SAPbouiCOM.EditText)oMatrixSyncUploadSOList1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific;
                                DoOpenLinkedObjectForm("139", "2050", "8", DocNum.Value);
                            }
                            if (pVal.ItemUID == MtxSyncUpSOListStr & pVal.ColUID == "DocEntry")
                            {//SO
                                SAPbouiCOM.EditText DocEntry = (SAPbouiCOM.EditText)oMatrixSyncUploadSOList1.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific;
                                var DocNum = getSODocNum(Convert.ToInt32(DocEntry.Value));
                                DoOpenLinkedObjectForm("139", "2050", "8", DocNum);
                            }
                            if (pVal.ItemUID == MtxSyncUpSOListStr & pVal.ColUID == "CardCode")
                            {//SO
                                SAPbouiCOM.EditText CardCode = (SAPbouiCOM.EditText)oMatrixSyncUploadSOList1.Columns.Item("CardCode").Cells.Item(pVal.Row).Specific;
                                DoOpenLinkedObjectForm("134", "2561", "5", CardCode.Value);
                            }
                            //SO Log
                            if (pVal.ItemUID == "mtxSO2" & pVal.ColUID == "DocNum")
                            {//SO
                                SAPbouiCOM.EditText DocNum = (SAPbouiCOM.EditText)oMatrixSyncUploadSOList1LOG.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific;
                                DoOpenLinkedObjectForm("139", "2050", "8", DocNum.Value);
                            }
                            if (pVal.ItemUID == "mtxSO2" & pVal.ColUID == "DocEntry")
                            {//SO
                                SAPbouiCOM.EditText DocEntry = (SAPbouiCOM.EditText)oMatrixSyncUploadSOList1LOG.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific;
                                var DocNum = getSODocNum(Convert.ToInt32(DocEntry.Value));
                                DoOpenLinkedObjectForm("139", "2050", "8", DocNum);
                            }
                            //Invoice 
                            if (pVal.ItemUID == MtxSyncUpInvoiceListStr & pVal.ColUID == "DocNum")
                            {
                                SAPbouiCOM.EditText DocNum = (SAPbouiCOM.EditText)oMatrixSyncUploadInvoiceList1.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific;
                                DoOpenLinkedObjectForm("133", "2053", "8", DocNum.Value);
                            }
                            if (pVal.ItemUID == MtxSyncUpInvoiceListStr & pVal.ColUID == "DocEntry")
                            {
                                SAPbouiCOM.EditText DocEntry = (SAPbouiCOM.EditText)oMatrixSyncUploadInvoiceList1.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific;
                                var DocNum = getInvoiceDocNum(Convert.ToInt32(DocEntry.Value));
                                DoOpenLinkedObjectForm("133", "2053", "8", DocNum);
                            }
                            if (pVal.ItemUID == MtxSyncUpInvoiceListStr & pVal.ColUID == "CardCode")
                            {
                                SAPbouiCOM.EditText CardCode = (SAPbouiCOM.EditText)oMatrixSyncUploadInvoiceList1.Columns.Item("CardCode").Cells.Item(pVal.Row).Specific;
                                DoOpenLinkedObjectForm("134", "2561", "5", CardCode.Value);
                            }
                            //Invoice Log
                            if (pVal.ItemUID == "mtxInLog" & pVal.ColUID == "DocNum")
                            {//SO
                                SAPbouiCOM.EditText DocNum = (SAPbouiCOM.EditText)oMatrixSyncUploadInvoiceList1LOG.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific;
                                DoOpenLinkedObjectForm("133", "2053", "8", DocNum.Value);
                            }
                            if (pVal.ItemUID == "mtxInLog" & pVal.ColUID == "DocEntry")
                            {//SO
                                SAPbouiCOM.EditText DocEntry = (SAPbouiCOM.EditText)oMatrixSyncUploadInvoiceList1LOG.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific;
                                var DocNum = getInvoiceDocNum(Convert.ToInt32(DocEntry.Value));
                                DoOpenLinkedObjectForm("133", "2053", "8", DocNum);
                            }



                            if (pVal.ItemUID == MtxSyncUpBPListStr & pVal.ColUID == "colCard")
                            {//BP
                                SAPbouiCOM.EditText CardCode = (SAPbouiCOM.EditText)oMatrixSyncUploadBPList1.Columns.Item("colCard").Cells.Item(pVal.Row).Specific;
                                DoOpenLinkedObjectForm("134", "2561", "5", CardCode.Value);
                            }
                            if (pVal.ItemUID == MtxSyncUpItemListStr & pVal.ColUID == "ItemCode")
                            {//Item
                                SAPbouiCOM.EditText ItemCode = (SAPbouiCOM.EditText)oMatrixSyncUploadItemList1.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific;
                                DoOpenLinkedObjectForm("150", "3073", "5", ItemCode.Value);
                            }
                        }
                        catch (Exception ex)
                        {
                            var error = ex.Message;

                        }
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
    #endregion

    #region Upload Logic
    #region Invoice 
    private void UploadInvoiceSyncObect(SAPbouiCOM.Form form)
    {
        bool FlgChange = false;
        try
        {//MtxSyncUpSOListStr, oDTSyncUploadSOList1, oCheckboxSOSelectAllstring);
            if (form.Items.Item(MtxSyncUpInvoiceListStr) == null)
                return;

            form.Items.Item(MtxSyncUpInvoiceListStr).AffectsFormMode = false;
            oMatrixExportSO = (SAPbouiCOM.Matrix)form.Items.Item(MtxSyncUpInvoiceListStr).Specific;
            int RowCount = oDTSyncUploadInvList1.Rows.Count;
            if (RowCount > 0)
            {
                for (int i = oDTSyncUploadInvList1.Rows.Count - 1; i >= 0; i--)
                {
                    var EdDocNum = GetMatrixValue(oMatrixSyncUploadInvoiceList1, i, "DocNum").ToString();

                    var chk = (SAPbouiCOM.CheckBox)oMatrixSyncUploadInvoiceList1.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    if (chk.Checked)
                    {
                        FlgChange = true;
                        Invoice oInvoice = GetInvoice(EdDocNum);
                        PostInvoiceSyncObj(oInvoice, i, oMatrixSyncUploadInvoiceList1);
                    }
                    else
                    {

                    }

                }
                if (FlgChange)
                {
                    //   oMatrixBpExport.LoadFromDataSource();
                    oMatrixSyncUploadInvoiceList1.FlushToDataSource();
                }
            }
        }
        catch (Exception ex)
        {
            errorLog(ex.Message);
        }
    }
    private void PostInvoiceSyncObj(Invoice OInvoice, int i, SAPbouiCOM.Matrix oMatx)
    {
        String EbizKey = "";
        bool PostStatus = false;
        try
        {
            if (CheckBP(OInvoice.CustomerId))
            {
                foreach (var item in OInvoice.Items)
                {
                    if (!CheckItem(item.ItemId))
                    {
                        var Eror = "";
                    }
                }
            }

            bool flgAddNewRecord = true;
            InvoiceResponse resp = ebiz.AddInvoice(getToken(), OInvoice);
            String Refnum = "";

            if (resp.ErrorCode == 2 && resp.Error.Equals("Record already exists") && resp.Status.Equals("Error"))
            {
                //Update
                Invoice OldSalesOrder = ebiz.GetInvoice(getToken(), OInvoice.CustomerId, "", OInvoice.InvoiceNumber, "");
                UpdateEbizChargeKey("OINV", OldSalesOrder.InvoiceInternalId, OInvoice.InvoiceNumber);
                Invoice updateInv = new Invoice();
                updateInv.CustomerId = OInvoice.CustomerId.Trim();
                updateInv.InvoiceInternalId = null;
                updateInv.Software = Software;
                updateInv.InvoiceNumber = OInvoice.InvoiceNumber;
                updateInv.InvoiceDate = OInvoice.InvoiceDate;
                updateInv.InvoiceDueDate = OInvoice.InvoiceDueDate;
                updateInv.InvoiceAmount = OInvoice.InvoiceAmount;
                updateInv.AmountDue = OInvoice.AmountDue;
                updateInv.NotifyCustomer = false;
                updateInv.TotalTaxAmount = 0;
                updateInv.InvoiceMemo = "";
                updateInv.InvoiceShipDate = OInvoice.InvoiceShipDate;
                updateInv.InvoiceShipVia = OInvoice.InvoiceShipVia;
                updateInv.Items = OInvoice.Items;

                resp = ebiz.UpdateInvoice(getToken(), updateInv, OInvoice.CustomerId.Trim(), null, OInvoice.InvoiceNumber, null);
                if (resp.Status.Equals(""))
                {
                    //ValidMsg = "Update To Connect Status:" + updateSoResp.Status;
                }
                ///
                flgAddNewRecord = false;
                EbizKey = OldSalesOrder.InvoiceInternalId;
                PostStatus = true;
                ShowSystemLog(String.Format("Invoice :{0} Updated successfully", OInvoice.InvoiceNumber));
            }
            else if (resp.ErrorCode == 0 && resp.Error.Equals("") && resp.Status.Equals("Success"))
            {
                // OldSalesOrder.InvoiceInternalId, OInvoice.InvoiceNumber
                UpdateEbizChargeKey("OINV", resp.InvoiceInternalId, OInvoice.InvoiceNumber);
                EbizKey = resp.InvoiceInternalId;
                PostStatus = true;
                #region Run Transaction
                TransactionResponse oTranResp = RunTransactionInvoice(OInvoice);
                Refnum = oTranResp.RefNum;
                String sql = string.Format("update [OINV[  set [U_RefNum[='{1}' Where [DocNum[='{0}'", OInvoice.InvoiceNumber, Refnum).Replace("[","\"");
                execute(sql);
                #endregion
                ShowSystemLog(String.Format("Invoice :{0} Upload successfully", OInvoice.InvoiceNumber));
            }

            if (PostStatus)
            {

                LogTable oExportSOLogTable = new LogTable();
                oExportSOLogTable.DocNum = OInvoice.InvoiceNumber;
                oExportSOLogTable.DocType = "OINV";
                oExportSOLogTable.PaymentStatus = "";
                oExportSOLogTable.RefNum = Refnum;
                oExportSOLogTable.Balance = Convert.ToDecimal(OInvoice.InvoiceAmount);
                oExportSOLogTable.DocEntry = OInvoice.PoNum;
                oExportSOLogTable.CardCode = OInvoice.CustomerId;

                if (flgAddNewRecord)
                {
                    oExportSOLogTable.CreateDt = DateTime.Now;
                    if (resp.Status.Equals("Success"))
                    {
                        oExportSOLogTable.Status = "Uploaded";
                    }
                }
                else
                {
                    oExportSOLogTable.CreateDt = DateTime.Now;
                    if (resp.Status.Equals("Success"))
                    {
                        oExportSOLogTable.Status = "Updated";
                    }
                }
                oExportSOLogTable.UpdateDt = DateTime.Now;

                SAPbouiCOM.EditText oEdit;
                oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("UploadDt").Cells.Item(i + 1).Specific;
                oEdit.Value = DatetimeTosapFormatDate(oExportSOLogTable.UpdateDt);
                AddInvoiceLog(oExportSOLogTable);

            }
            else
            {
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
                oEdit.Value = "False";
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }
    #endregion

    #region Sales Order
    private void UploadSOSyncObect(SAPbouiCOM.Form form)
    {
        bool FlgChange = false;
        try
        {//MtxSyncUpSOListStr, oDTSyncUploadSOList1, oCheckboxSOSelectAllstring);
            if (form.Items.Item(MtxSyncUpSOListStr) == null)
                return;

            form.Items.Item(MtxSyncUpSOListStr).AffectsFormMode = false;
            oMatrixExportSO = (SAPbouiCOM.Matrix)form.Items.Item(MtxSyncUpSOListStr).Specific;
            int RowCount = oDTSyncUploadSOList1.Rows.Count;
            if (RowCount > 0)
            {
                for (int i = oDTSyncUploadSOList1.Rows.Count - 1; i >= 0; i--)
                {
                    var EdDocNum = GetMatrixValue(oMatrixSyncUploadSOList1, i, "DocNum").ToString();

                    var chk = (SAPbouiCOM.CheckBox)oMatrixSyncUploadSOList1.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    if (chk.Checked)
                    {
                        FlgChange = true;
                        SalesOrder oSalesOrder = GetSO(EdDocNum);
                        PostSOSyncObj(oSalesOrder, i, oMatrixSyncUploadSOList1);
                    }
                    else
                    {

                    }

                }
                if (FlgChange)
                {
                    //   oMatrixBpExport.LoadFromDataSource();
                    oMatrixSyncUploadSOList1.FlushToDataSource();
                }
            }
        }
        catch (Exception ex)
        {
            errorLog(ex.Message);
        }
    }
    private void PostSOSyncObj(SalesOrder oSalesOrder, int i, SAPbouiCOM.Matrix oMatx)
    {
        String EbizKey = "", Refnum = "";
        bool PostStatus = false;
        try
        {
            if (CheckBP(oSalesOrder.CustomerId))
            {
                foreach (var item in oSalesOrder.Items)
                {
                    if (!CheckItem(item.ItemId))
                    {
                        var Eror = "";
                    }
                }
            }

            bool flgAddNewRecord = true;
            SalesOrderResponse resp = ebiz.AddSalesOrder(getToken(), oSalesOrder);
            if (resp.ErrorCode == 2 && resp.Error.Equals("Record already exists") && resp.Status.Equals("Error"))
            {
                //Update
                SalesOrder OldSalesOrder = ebiz.GetSalesOrder(getToken(), oSalesOrder.CustomerId, "", oSalesOrder.SalesOrderNumber, "");
                UpdateEbizChargeKey("ORDR", OldSalesOrder.SalesOrderInternalId, oSalesOrder.SalesOrderNumber);
                SalesOrder updateSO = new SalesOrder();
                updateSO.CustomerId = oSalesOrder.CustomerId.Trim();
                updateSO.SalesOrderInternalId = null;
                updateSO.Software = Software;
                updateSO.SalesOrderNumber = oSalesOrder.SalesOrderNumber;
                updateSO.Date = oSalesOrder.Date;
                updateSO.DueDate = oSalesOrder.DueDate;
                updateSO.Amount = oSalesOrder.Amount;
                updateSO.AmountDue = oSalesOrder.AmountDue;
                updateSO.Currency = oSalesOrder.Currency;
                updateSO.NotifyCustomer = false;
                updateSO.TotalTaxAmount = 0;
                updateSO.Memo = "";
                updateSO.ShipDate = oSalesOrder.DueDate;
                updateSO.ShipVia = "";
                updateSO.Items = oSalesOrder.Items;

                resp = ebiz.UpdateSalesOrder(getToken(), updateSO, oSalesOrder.CustomerId.Trim(), null, oSalesOrder.SalesOrderNumber, null);
                if (resp.Status.Equals(""))
                {
                    //ValidMsg = "Update To Connect Status:" + updateSoResp.Status;
                }
                ///
                flgAddNewRecord = false;
                EbizKey = OldSalesOrder.SalesOrderInternalId;
                PostStatus = true;
                ShowSystemLog(String.Format("Sales Order:{0} Updated successfully", oSalesOrder.SalesOrderNumber));
            }
            else if (resp.ErrorCode == 0 && resp.Error.Equals("") && resp.Status.Equals("Success"))
            {
                UpdateEbizChargeKey("ORDR", resp.SalesOrderInternalId, oSalesOrder.SalesOrderNumber);
                EbizKey = resp.SalesOrderInternalId;
                PostStatus = true;
                ShowSystemLog(String.Format("Sales Order:{0} Upload successfully", oSalesOrder.SalesOrderNumber));
                TransactionResponse oTranResp = RunTransactionOrder(oSalesOrder);
                Refnum = oTranResp.RefNum;
                String sql = string.Format("update [ORDR[  set [U_RefNum[='{1}' Where [DocNum[='{0}'", oSalesOrder.SalesOrderNumber, Refnum).Replace("[", "\"");
                execute(sql);
            }

            if (PostStatus)
            {

                LogTable oExportSOLogTable = new LogTable();
                oExportSOLogTable.DocNum = oSalesOrder.SalesOrderNumber;
                oExportSOLogTable.DocType = "SO";
                oExportSOLogTable.PaymentStatus = "";
                oExportSOLogTable.RefNum = Refnum;
                oExportSOLogTable.CardCode = oSalesOrder.CustomerId;
                oExportSOLogTable.DocEntry = oSalesOrder.PoNum;
                oExportSOLogTable.AmountPaid = 0;
                oExportSOLogTable.Amount = ConvertDecimal(oSalesOrder.Amount.ToString());
                oExportSOLogTable.AmountDue = ConvertDecimal(oSalesOrder.AmountDue.ToString());
                String DocState = "Update";
                if (flgAddNewRecord)
                {
                    oExportSOLogTable.CreateDt = DateTime.Now;
                    if (resp.Status.Equals("Success"))
                    {
                        oExportSOLogTable.Status = "Uploaded";
                    }
                    DocState = "Add";
                }
                else
                {
                    oExportSOLogTable.CreateDt = DateTime.Now;
                    if (resp.Status.Equals("Success"))
                    {
                        oExportSOLogTable.Status = "Uploaded";
                    }
                }
                oExportSOLogTable.UpdateDt = DateTime.Now;

                SAPbouiCOM.EditText oEdit;
                oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("UploadDt").Cells.Item(i + 1).Specific;
                oEdit.Value = DatetimeTosapFormatDate(oExportSOLogTable.UpdateDt);
                oExportSOLogTable.DocState = DocState;
                AddSOLog(oExportSOLogTable);
            }
            else
            {
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
                oEdit.Value = "False";
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }
    #endregion

    #region Item Master Data
    private void UploadItemSyncObect(SAPbouiCOM.Form form)
    {
        bool FlgChange = false;
        try
        {
            if (form.Items.Item(MtxSyncUpItemListStr) == null)
                return;

            form.Items.Item(MtxSyncUpItemListStr).AffectsFormMode = false;
            oMatrixSyncUploadItemList1 = (SAPbouiCOM.Matrix)form.Items.Item(MtxSyncUpItemListStr).Specific;
            int RowCount = oDTSyncUploadItemList1.Rows.Count;
            if (RowCount > 0)
            {
                for (int i = oDTSyncUploadItemList1.Rows.Count - 1; i >= 0; i--)
                {
                    var EdItemCode = GetMatrixValue(oMatrixSyncUploadItemList1, i, "ItemCode").ToString();
                    var EdItemName = GetMatrixValue(oMatrixSyncUploadItemList1, i, "ItemName").ToString();

                    var chk = (SAPbouiCOM.CheckBox)oMatrixSyncUploadItemList1.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    if (chk.Checked)
                    {
                        FlgChange = true;
                        ItemMasterSAP oItem = GetSAPItem(EdItemCode);
                        ItemDetails NewItem = new ItemDetails();
                        NewItem.ItemId = EdItemCode;
                        NewItem.Description = EdItemName;
                        NewItem.UnitPrice = oItem.Price;
                        NewItem.QtyOnHand = oItem.onHand;
                        NewItem.SoftwareId = Software;

                        PostItemSU(NewItem, i, oMatrixSyncUploadItemList1);
                    }
                    else
                    {
                    }

                }
                if (FlgChange)
                {
                    //oMatrixItemExport.FlushToDataSource();
                    //   oMatrixItemExport.LoadFromDataSource();
                }

            }
        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }
    private void PostItemSU(ItemDetails Oitem, int i, SAPbouiCOM.Matrix oMatx)
    {
        String EbizKey = "";
        try
        {
            var Token = getToken();
            ItemDetailsResponse resp = ebiz.AddItem(Token, Oitem);
            if (resp.ErrorCode == 2 && resp.Error.Equals("Record already exists") && resp.Status.Equals("Error"))
            {
                ItemDetails[] list = ebiz.SearchItems(Token, "", Oitem.ItemId, null, 0, 2, "");
                if (list != null && list.Count() > 0)
                {
                    resp = ebiz.UpdateItem(Token, Oitem, resp.ItemInternalId, Oitem.ItemId);
                    UpdateEbizChargeKey("OITM", list[0].ItemInternalId, Oitem.ItemId);
                    EbizKey = list[0].ItemInternalId;
                }
                ShowSystemLog(String.Format("ItemCode:{0} Updated successfully", Oitem.ItemId));
            }
            else if (resp.ErrorCode == 0 && resp.Error.Equals("") && resp.Status.Equals("Success"))
            {
                UpdateEbizChargeKey("OITM", resp.ItemInternalId, Oitem.ItemId);
                EbizKey = resp.ItemInternalId;
                ShowSystemLog(String.Format("ItemCode:{0} Added successfully", Oitem.ItemId));
            }
            try
            {
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("UploadDt").Cells.Item(i + 1).Specific;
                oEdit.Value = DatetimeTosapFormatDate(DateTime.Now) + "";
            }
            catch (Exception ex)
            {
                errorLog(ex);
            }

            //
            LogTable oExportItemLogTable = new LogTable();
            oExportItemLogTable.ItemCode = Oitem.ItemId;
            oExportItemLogTable.CreateDt = DateTime.Now;
            oExportItemLogTable.Status = resp.Status;
            oExportItemLogTable.Ebiz = resp.ItemInternalId;
            oExportItemLogTable.UpdateDt = DateTime.Now;
            if (resp.Status.Equals("Success"))
            {
                oExportItemLogTable.Status = "Uploaded";
            }
            AddUploadDataLog(oExportItemLogTable, "CCITEMLOGTAB");

            //

        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }
    #endregion

    #region Business Partner
    private void UploadBPSyncObect(SAPbouiCOM.Form form)
    {
        bool FlgChange = false;
        try
        {
            if (form.Items.Item(MtxSyncUpBPListStr) == null)
                return;

            form.Items.Item(MtxSyncUpBPListStr).AffectsFormMode = false;
            oMatrixSyncUploadBPList1 = (SAPbouiCOM.Matrix)form.Items.Item(MtxSyncUpBPListStr).Specific;
            int RowCount = oDTSyncUploadBPList1.Rows.Count;
            if (RowCount > 0)
            {
                for (int i = oDTSyncUploadBPList1.Rows.Count - 1; i >= 0; i--)
                {
                    var EdCardCode = GetMatrixValue(oMatrixSyncUploadBPList1, i, "colCard").ToString();
                    var EdCardName = GetMatrixValue(oMatrixSyncUploadBPList1, i, "CardName").ToString();
                    var EdEmail = GetMatrixValue(oMatrixSyncUploadBPList1, i, "E_Mail").ToString();

                    var chk = (SAPbouiCOM.CheckBox)oMatrixSyncUploadBPList1.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    if (chk.Checked)
                    {
                        FlgChange = true;
                        Customer oCust = NewCustomer(EdCardCode, EdCardName, EdEmail);
                        PostBPSY(oCust, i, oMatrixSyncUploadBPList1);
                    }
                    else
                    {

                    }

                }
                if (FlgChange)
                {
                    //   oMatrixBpExport.LoadFromDataSource();
                    //  oMatrixBpExport.FlushToDataSource();
                }
            }
        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }
    private void PostBPSY(Customer oCustomer, int i, SAPbouiCOM.Matrix oMatx)
    {
        String EbizKey = "";
        try
        {
            CustomerResponse resp = ebiz.AddCustomer(getToken(), oCustomer);
            if (resp.ErrorCode == 2 && resp.Error.Equals("Record already exists") && resp.Status.Equals("Failed"))
            {
                Customer OldCustomer = ebiz.GetCustomer(getToken(), oCustomer.CustomerId, "");
                resp = ebiz.UpdateCustomer(getToken(), oCustomer, OldCustomer.CustomerId, OldCustomer.CustomerInternalId);
                UpdateEbizChargeKey("OCRD", OldCustomer.CustomerInternalId, oCustomer.CustomerId);
                EbizKey = OldCustomer.CustomerInternalId;
                ShowSystemLog(String.Format("Bussiness Partner:{0} Updated successfully", oCustomer.CustomerId));
            }
            else if (resp.ErrorCode == 0 && resp.Error.Equals("") && resp.Status.Equals("Success"))
            {
                UpdateEbizChargeKey("OCRD", resp.CustomerInternalId, oCustomer.CustomerId);
                EbizKey = resp.CustomerInternalId;
                ShowSystemLog(String.Format("Bussiness Partner:{0} Added successfully", oCustomer.CustomerId));
            }
            SAPbouiCOM.EditText oEdit;
            try
            {
                oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("UpdateDt").Cells.Item(i + 1).Specific;
                oEdit.Value = DateTime.Now.ToShortDateString();
            }
            catch (Exception ex)
            {
                errorLog(ex);
            }
            LogTable oExportBPLogTable = new LogTable();
            oExportBPLogTable.CardCode = oCustomer.CustomerId;
            oExportBPLogTable.CreateDt = DateTime.Now;
            oExportBPLogTable.Status = resp.Status;
            oExportBPLogTable.Ebiz = resp.CustomerInternalId;
            oExportBPLogTable.UpdateDt = DateTime.Now;
            if (resp.Status.Equals("Success"))
            {
                oExportBPLogTable.Status = "Uploaded";
            }
            AddUploadDataLog(oExportBPLogTable, "CCBPLOGTAB");

        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }
    #endregion
    #endregion

    #region Helping Function

    private void InitializeSyncupload()
    {
        try
        {
            SyncUploadflgDocInitiallization = true;
            ShowDataSyncUpload();
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void ShowDataSyncUpload()
    {
        try
        {
            tProcess = new Thread(SyncUploadCreateTabThreadProc);
            tProcess.Start();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            throw;
        }
    }
    public void SyncUploadCreateTabThreadProc()
    {
        try
        {
            if (formSelTarget == null)
            {
                if (FrmSyncUpload != null)
                    formSelTarget = FrmSyncUpload;
                else
                    return;
            }
            if (formSelTarget.TypeEx != formSyncUpload)
                return;

            if (!SyncUploadflgDocInitiallization) return;



            ShowSyncUploadSO(FrmSyncUpload);
            FillDate();
            InitializeCheckBox();
            FillCustomerTypeSyncUpload();
            FillPaymentTermsSyncUpload();

        }
        catch (Exception ex)
        {
        }
    }
    private void ShowSyncUpLoadItemList(SAPbouiCOM.Form form, String Where = "")
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxSyncUpItemListStr) == null)
                return;

            form.Items.Item(MtxSyncUpItemListStr).AffectsFormMode = false;
            oMatrixSyncUploadItemList1 = (SAPbouiCOM.Matrix)form.Items.Item(MtxSyncUpItemListStr).Specific;
            oMatrixSyncUploadItemList1.Clear();
            oDTSyncUploadItemList1 = getDataTableSyncUpload(form, MtxSyncUpItemListStr);
            sql = String.Format("select 'N' as [clSel],ROW_NUMBER() OVER (ORDER BY [DocEntry]) AS [ID],[ItemCode],[ItemName],cast([OnHand] as int) as [InStock],[LastPurPrc] as [Price],(SELECT MAX(T0.[U_CreateDt]) FROM [@CCITEMLOGTAB]  T0 WHERE T0.[U_ItemCode[ =OITM.[ItemCode[) as [UploadDt] from OITM ").Replace("]", "\"").Replace("[", "\"");

            oDTSyncUploadItemList1.ExecuteQuery(sql);

            BindMatrixSyncUpload(MtxSyncUpItemListStr, oMatrixSyncUploadItemList1, "clSel", "clSel", null, true);
            BindMatrixSyncUpload(MtxSyncUpItemListStr, oMatrixSyncUploadItemList1, "clNo", "ID");
            BindMatrixSyncUpload(MtxSyncUpItemListStr, oMatrixSyncUploadItemList1, "ItemCode", "ItemCode");
            BindMatrixSyncUpload(MtxSyncUpItemListStr, oMatrixSyncUploadItemList1, "ItemName", "ItemName");
            BindMatrixSyncUpload(MtxSyncUpItemListStr, oMatrixSyncUploadItemList1, "UnitPrice", "Price");
            BindMatrixSyncUpload(MtxSyncUpItemListStr, oMatrixSyncUploadItemList1, "InStock", "InStock");
            BindMatrixSyncUpload(MtxSyncUpItemListStr, oMatrixSyncUploadItemList1, "UploadDt", "UploadDt");

            oMatrixSyncUploadItemList1.LoadFromDataSource();
            oMatrixSyncUploadItemList1.AutoResizeColumns();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void ShowSyncUploadBPList(SAPbouiCOM.Form form, String DefaultCheck = "N")
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxSyncUpBPListStr) == null)
                return;

            form.Items.Item(MtxSyncUpBPListStr).AffectsFormMode = false;
            oMatrixSyncUploadBPList1 = (SAPbouiCOM.Matrix)form.Items.Item(MtxSyncUpBPListStr).Specific;
            oMatrixSyncUploadBPList1.Clear();
            oDTSyncUploadBPList1 = getDataTableSyncUpload(form, MtxSyncUpBPListStr);
            sql = String.Format("SELECT '{0}' as [clSel],ROW_NUMBER() OVER (ORDER BY T0.[CardCode]) AS [ID],T0.[CardCode] as [c], T0.[CardName],T0.[Balance],T0.[E_Mail], IFNULL(T0.[U_EbizChargeID],'') as [EBiz],CASE WHEN IFNULL(T0.[U_EbizChargeID], '') = '' THEN 'False' ELSE 'True'END as [Sync],(select Max([U_CreateDt]) from [@CCBPLOGTAB[ where [U_CardCode]=T0.[CardCode]) as [UpdateDt] FROM OCRD T0 WHERE T0.[CardType] = 'C'", DefaultCheck).Replace("]", "\"").Replace("[", "\"");

            oDTSyncUploadBPList1.ExecuteQuery(sql);
            BindMatrixSyncUpload(MtxSyncUpBPListStr, oMatrixSyncUploadBPList1, "clSel", "clSel", null, true);
            BindMatrixSyncUpload(MtxSyncUpBPListStr, oMatrixSyncUploadBPList1, "clNo", "ID");
            BindMatrixSyncUpload(MtxSyncUpBPListStr, oMatrixSyncUploadBPList1, "colCard", "c");
            BindMatrixSyncUpload(MtxSyncUpBPListStr, oMatrixSyncUploadBPList1, "CardName", "CardName");
            BindMatrixSyncUpload(MtxSyncUpBPListStr, oMatrixSyncUploadBPList1, "Balance", "Balance");
            BindMatrixSyncUpload(MtxSyncUpBPListStr, oMatrixSyncUploadBPList1, "E_Mail", "E_Mail");
            BindMatrixSyncUpload(MtxSyncUpBPListStr, oMatrixSyncUploadBPList1, "UpdateDt", "UpdateDt");

            // BindMatrixBP(oMatrixBpExport, "clSync", "Sync");

            oMatrixSyncUploadBPList1.LoadFromDataSource();
            oMatrixSyncUploadBPList1.AutoResizeColumns();

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void ShowSyncUploadSO(SAPbouiCOM.Form form, String Where = "")
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxSyncUpSOListStr) == null)
                return;

            form.Items.Item(MtxSyncUpSOListStr).AffectsFormMode = false;
            oMatrixSyncUploadSOList1 = (SAPbouiCOM.Matrix)form.Items.Item(MtxSyncUpSOListStr).Specific;
            oMatrixSyncUploadSOList1.Clear();
            oDTSyncUploadSOList1 = getDataTableSyncUpload(form, MtxSyncUpSOListStr);

            if (Where.Equals(""))
            {
                sql = String.Format(" select 'N' as [clSel[,ROW_NUMBER() OVER (ORDER BY R.[DocEntry[) AS [ID[,R.[DocEntry[,R.[DocNum[,R.[DocDate[,R.[DocTotal[,R.[CardCode[,R.[CardName[  , (select Max([U_Status[) from[@CCSOLOG[where[U_DocNum[= Cast(R.[DocNum[ as nvarchar(20)))as [Status[ , (select Max([U_CreateDt[) from[@CCSOLOG[where[U_DocNum[= Cast(R.[DocNum[ as nvarchar(20)) )as [UploadDt[ ,OCRD.[Balance[,CASE  WHEN IFNULL(R.[U_EbizChargeMarkPayment[, '') = 'sales' THEN 'Deposit'  WHEN IFNULL(R.[U_EbizChargeMarkPayment[, '') = 'AuthOnly' THEN 'Pre-auth' ELSE '' END as [Type[ from ORDR R  inner join OCRD on OCRD.[CardCode[=R.[CardCode[ where R.[DocStatus[= 'O' and R.\"DocDate\" <=current_date  ").Replace("]", "\"").Replace("[", "\"");
            }
            else
            {
                var subQ = " ,( select Max([U_Status[) from [@CCSOLOG[  where [U_DocNum[=Cast(RDR.[DocNum[ as nvarchar(20)))as [Status[ ,(select Max([U_CreateDt[) from [@CCSOLOG[where[U_DocNum[= Cast(RDR.[DocNum[ as nvarchar(20)) )as [UploadDt[ ,C.[Balance[   ";
                sql = String.Format("select 'N' as ]clSel],ROW_NUMBER() OVER (ORDER BY RDR.[DocEntry]) AS ]ID],RDR.[DocEntry[,RDR.[DocNum],RDR.[DocDate],RDR.[DocTotal],RDR.[CardCode],  RDR.[CardName] {1}  from ORDR RDR  inner  join OCRD C on C.[CardCode] = RDR.[CardCode]  inner  join OCRG CRG on C.[GroupCode] = CRG.[GroupCode]  INNER JOIN OCTG CTG ON C.[GroupNum] = CTG.[GroupNum]  where RDR.[DocStatus]='O'   and IFNULL(RDR.[U_EbizChargeMarkPayment],'')<>'Paid' {0}", Where, subQ).Replace("[", "\"").Replace("]", "\"");
            }

            oDTSyncUploadSOList1.ExecuteQuery(sql);

            BindMatrixSyncUpload(MtxSyncUpSOListStr, oMatrixSyncUploadSOList1, "clSel", "clSel", null, true);
            BindMatrixSyncUpload(MtxSyncUpSOListStr, oMatrixSyncUploadSOList1, "clNo", "ID");
            BindMatrixSyncUpload(MtxSyncUpSOListStr, oMatrixSyncUploadSOList1, "DocNum", "DocNum");
            BindMatrixSyncUpload(MtxSyncUpSOListStr, oMatrixSyncUploadSOList1, "DocEntry", "DocEntry");
            BindMatrixSyncUpload(MtxSyncUpSOListStr, oMatrixSyncUploadSOList1, "DocDate", "DocDate");
            BindMatrixSyncUpload(MtxSyncUpSOListStr, oMatrixSyncUploadSOList1, "CardCode", "CardCode");
            BindMatrixSyncUpload(MtxSyncUpSOListStr, oMatrixSyncUploadSOList1, "CardName", "CardName");
            BindMatrixSyncUpload(MtxSyncUpSOListStr, oMatrixSyncUploadSOList1, "Type", "Type");
            BindMatrixSyncUpload(MtxSyncUpSOListStr, oMatrixSyncUploadSOList1, "DocTotal", "DocTotal");
            BindMatrixSyncUpload(MtxSyncUpSOListStr, oMatrixSyncUploadSOList1, "Balance", "Balance");
            BindMatrixSyncUpload(MtxSyncUpSOListStr, oMatrixSyncUploadSOList1, "UploadDt", "UploadDt");

            oMatrixSyncUploadSOList1.LoadFromDataSource();
            oMatrixSyncUploadSOList1.AutoResizeColumns();

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void ShowSyncUploadInvoice(SAPbouiCOM.Form form, String Where = "")
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxSyncUpInvoiceListStr) == null)
                return;

            form.Items.Item(MtxSyncUpInvoiceListStr).AffectsFormMode = false;
            oMatrixSyncUploadInvoiceList1 = (SAPbouiCOM.Matrix)form.Items.Item(MtxSyncUpInvoiceListStr).Specific;
            oMatrixSyncUploadInvoiceList1.Clear();
            oDTSyncUploadInvList1 = getDataTableSyncUpload(form, MtxSyncUpInvoiceListStr);

            if (Where.Equals(""))
            {
                sql = String.Format(" select 'N' as [clSel[,ROW_NUMBER() OVER (ORDER BY R.[DocEntry[) AS [ID[,R.[DocEntry[,R.[DocNum[,R.[DocDate[,R.[DocTotal[,R.[CardCode[,R.[CardName[  , (select Max([U_Status[) from[@CCINVLOG[where[U_DocNum[= Cast(R.[DocNum[ as nvarchar(20)))as [Status[ , (select Max([U_CreateDt[) from[@CCINVLOG[where[U_DocNum[= Cast(R.[DocNum[ as nvarchar(20)) )as [UploadDt[ ,OCRD.[Balance[,CASE  WHEN IFNULL(R.[U_EbizChargeMarkPayment[, '') = 'sales' THEN 'Deposit'  WHEN IFNULL(R.[U_EbizChargeMarkPayment[, '') = 'AuthOnly' THEN 'Pre-auth' ELSE '' END as [Type[ from OINV R  inner join OCRD on OCRD.[CardCode[=R.[CardCode[ where R.[DocStatus[ in  ('O','C') and R.\"CANCELED\" <> 'Y' and R.\"DocDate\" <=current_date  ").Replace("]", "\"").Replace("[", "\"");
            }
            else
            {
                var subQ = " ,( select Max([U_Status[) from [@CCINVLOG[  where [U_DocNum[=Cast(RDR.[DocNum[ as nvarchar(20)))as [Status[ ,(select Max([U_CreateDt[) from [@CCINVLOG[where[U_DocNum[= Cast(RDR.[DocNum[ as nvarchar(20)) )as [UploadDt[ ,C.[Balance[   ";
                sql = String.Format("select 'N' as ]clSel],ROW_NUMBER() OVER (ORDER BY RDR.[DocEntry]) AS ]ID],RDR.[DocEntry[,RDR.[DocNum],RDR.[DocDate],RDR.[DocTotal],RDR.[CardCode],  RDR.[CardName] {1}  from OINV RDR  inner  join OCRD C on C.[CardCode] = RDR.[CardCode]  inner  join OCRG CRG on C.[GroupCode] = CRG.[GroupCode]  INNER JOIN OCTG CTG ON C.[GroupNum] = CTG.[GroupNum]  where RDR.[DocStatus] in ('O','C') and RDR.\"CANCELED\" <> 'Y'   and IFNULL(RDR.[U_EbizChargeMarkPayment],'')<>'Paid' {0}", Where, subQ).Replace("[", "\"").Replace("]", "\"");
            }

            oDTSyncUploadInvList1.ExecuteQuery(sql);

            BindMatrixSyncUpload(MtxSyncUpInvoiceListStr, oMatrixSyncUploadInvoiceList1, "clSel", "clSel", null, true);
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStr, oMatrixSyncUploadInvoiceList1, "clNo", "ID");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStr, oMatrixSyncUploadInvoiceList1, "DocNum", "DocNum");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStr, oMatrixSyncUploadInvoiceList1, "DocEntry", "DocEntry");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStr, oMatrixSyncUploadInvoiceList1, "DocDate", "DocDate");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStr, oMatrixSyncUploadInvoiceList1, "CardCode", "CardCode");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStr, oMatrixSyncUploadInvoiceList1, "CardName", "CardName");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStr, oMatrixSyncUploadInvoiceList1, "Type", "Type");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStr, oMatrixSyncUploadInvoiceList1, "DocTotal", "DocTotal");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStr, oMatrixSyncUploadInvoiceList1, "Balance", "Balance");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStr, oMatrixSyncUploadInvoiceList1, "UploadDt", "UploadDt");

            oMatrixSyncUploadInvoiceList1.LoadFromDataSource();
            oMatrixSyncUploadInvoiceList1.AutoResizeColumns();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void FillDate()
    {
        try
        {
            ed_SUFromDate = ((SAPbouiCOM.EditText)(FrmSyncUpload.Items.Item(ed_SUFromDateString).Specific));
            FrmSyncUpload.DataSources.UserDataSources.Add(ed_SUFromDateString, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_SUFromDate.DataBind.SetBound(true, "", ed_SUFromDateString);
            ed_SUFromDate.Value = DatetimeTosapFormatDate(DateTime.Now);

            ed_SUToDate = ((SAPbouiCOM.EditText)(FrmSyncUpload.Items.Item(ed_SUToDateString).Specific));
            FrmSyncUpload.DataSources.UserDataSources.Add(ed_SUToDateString, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_SUToDate.DataBind.SetBound(true, "", ed_SUToDateString);
            ed_SUToDate.Value = DatetimeTosapFormatDate(DateTime.Now);

            ed_InFromDate = ((SAPbouiCOM.EditText)(FrmSyncUpload.Items.Item(ed_InFromDateString).Specific));
            FrmSyncUpload.DataSources.UserDataSources.Add(ed_InFromDateString, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_InFromDate.DataBind.SetBound(true, "", ed_InFromDateString);
            ed_InFromDate.Value = DatetimeTosapFormatDate(DateTime.Now);

            ed_InToDate = ((SAPbouiCOM.EditText)(FrmSyncUpload.Items.Item(ed_InToDateString).Specific));
            FrmSyncUpload.DataSources.UserDataSources.Add(ed_InToDateString, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_InToDate.DataBind.SetBound(true, "", ed_InToDateString);
            ed_InToDate.Value = DatetimeTosapFormatDate(DateTime.Now);



            //SetFolder("TabItem");
            //SetFolder("TabItmLst");
            SetFolder(FrmSyncUpload, "TabSO");
            SetFolder(FrmSyncUpload, "TabSOList");
            flgBPfolder = true;
            flgItemfolder = true;
        }
        catch (Exception ex)
        {
            var error = ex.Message;
            errorLog("SyncUpload Fill Date :" + ex.Message);
        }
    }
    private void SetFolder(SAPbouiCOM.Form form, String oFolderId)
    {
        try
        {
            oFolder = ((SAPbouiCOM.Folder)form.Items.Item(oFolderId).Specific);
            oFolder.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
            errorLog("SyncUpload SetFolder :" + ex.Message);
        }
    }
    private void InitializeCheckBox()
    {
        try
        {
            String checkID = oCheckboxSOSelectAllstring;
            oCheckboxSOSelectAll = ((SAPbouiCOM.CheckBox)(FrmSyncUpload.Items.Item(checkID).Specific));
            FrmSyncUpload.DataSources.UserDataSources.Add(checkID, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oCheckboxSOSelectAll.DataBind.SetBound(true, "", checkID);
            FrmSyncUpload.DataSources.UserDataSources.Item(checkID).Value = "N";

            checkID = oCheckboxItemSelectAllstring;
            oCheckboxItemSelectAll = ((SAPbouiCOM.CheckBox)(FrmSyncUpload.Items.Item(checkID).Specific));
            FrmSyncUpload.DataSources.UserDataSources.Add(checkID, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oCheckboxItemSelectAll.DataBind.SetBound(true, "", checkID);
            FrmSyncUpload.DataSources.UserDataSources.Item(checkID).Value = "N";

            checkID = oCheckboxBPSelectAllstring;
            oCheckboxBPSelectAll = ((SAPbouiCOM.CheckBox)(FrmSyncUpload.Items.Item(checkID).Specific));
            FrmSyncUpload.DataSources.UserDataSources.Add(checkID, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oCheckboxBPSelectAll.DataBind.SetBound(true, "", checkID);
            FrmSyncUpload.DataSources.UserDataSources.Item(checkID).Value = "N";

            checkID = oCheckboxInvoiceSelectAllstring;
            oCheckboxBPSelectAll = ((SAPbouiCOM.CheckBox)(FrmSyncUpload.Items.Item(checkID).Specific));
            FrmSyncUpload.DataSources.UserDataSources.Add(checkID, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oCheckboxBPSelectAll.DataBind.SetBound(true, "", checkID);
            FrmSyncUpload.DataSources.UserDataSources.Item(checkID).Value = "N";
        }
        catch (Exception ex)
        {

            var error = ex.Message;
        }
    }
    private void FillCustomerTypeSyncUpload()
    {
        try
        {
            Cmb_CustomerGrpSU = ((SAPbouiCOM.ComboBox)(FrmSyncUpload.Items.Item("cmbCgrp").Specific));
            Cmb_CustomerGrpSU2 = ((SAPbouiCOM.ComboBox)(FrmSyncUpload.Items.Item("cmbCgp2").Specific));
            ComboAddItem(Cmb_CustomerGrpSU, "All");
            ComboAddItem(Cmb_CustomerGrpSU2, "All");
            List<string> list = GetGroupNames();
            foreach (string c in list)
            {
                ComboAddItem(Cmb_CustomerGrpSU, c);
                ComboAddItem(Cmb_CustomerGrpSU2, c);
            }
            Cmb_CustomerGrpSU.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            Cmb_CustomerGrpSU2.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            //cmbCgp2
        }
        catch (Exception ex)
        {
            string Error = ex.Message;
        }
    }
    private void FillPaymentTermsSyncUpload()
    {
        try
        {
            Cmb_PaymentsTermSU = ((SAPbouiCOM.ComboBox)(FrmSyncUpload.Items.Item("cmbPtm").Specific));
            ComboAddItem(Cmb_PaymentsTermSU, "All");
            Cmb_PaymentsTermSU2 = ((SAPbouiCOM.ComboBox)(FrmSyncUpload.Items.Item("cmbPtm2").Specific));
            ComboAddItem(Cmb_PaymentsTermSU2, "All");
            List<string> list = GetPaymentTerms();
            foreach (string c in list)
            {
                ComboAddItem(Cmb_PaymentsTermSU, c);
                ComboAddItem(Cmb_PaymentsTermSU2, c);
            }
            Cmb_PaymentsTermSU.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            Cmb_PaymentsTermSU2.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        }
        catch (Exception ex)
        {
            string Error = ex.Message;
        }
    }
    private void FindFunctionality(SAPbouiCOM.Form form)
    {
        try
        {
            if (form.Items.Item(MtxSyncUpSOListStr) == null)
                return;

            var CustomerGroup = GetComboValue(FrmSyncUpload, "cmbCgrp");
            var PaymentTerms = GetComboValue(FrmSyncUpload, "cmbPtm");
            var CardCode = GetItemValue("edBPId", FrmSyncUpload);
            DateTime FromDate = getdateFor(ed_SUFromDate);
            DateTime ToDate = getdateFor(ed_SUToDate);

            String CustomerGroupParameter = CustomerGroup;
            String CardCodeParameter = CardCode;
            String PaymentTermsParameter = PaymentTerms;
            String FromDateParameter = "", ToDateParameter;
            string Where = "";
            if (CardCode.Equals(""))
            {
                CardCodeParameter = "%";
            }
            if (PaymentTermsParameter.Equals("All"))
            {
                PaymentTermsParameter = "%";
            }
            if (CustomerGroupParameter.Equals("All"))
            {
                CustomerGroupParameter = "%";
            }
            if (FromDate.Equals(DateTime.MinValue))
            {
                FromDateParameter = " like '%'";
            }
            else
            {
                FromDateParameter = string.Format(" >= cast('{0}' as date)", FromDate.ToString("yyyy/MM/dd"));
            }
            if (ToDate.Equals(DateTime.MinValue))
            {
                ToDateParameter = " like '%'";
            }
            else
            {
                ToDateParameter = string.Format(" <= cast('{0}' as date)", ToDate.ToString("yyyy/MM/dd"));
            }
            Where = String.Format("and C.[CardCode] like '{0}' and CRG.[GroupName] like '{2}'  and CTG.[PymntGroup] like '{3}'  and RDR.[DocDate]  {1} and RDR.[DocDate]  {4}", CardCodeParameter, FromDateParameter, CustomerGroupParameter, PaymentTermsParameter, ToDateParameter);
            ShowSyncUploadSO(form, Where);

        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }

    private void FindFunctionality_Invoice(SAPbouiCOM.Form form)
    {
        try
        {
            if (form.Items.Item(MtxSyncUpInvoiceListStr) == null)
                return;

            var CustomerGroup = GetComboValue(FrmSyncUpload, "cmbCgp2");
            var PaymentTerms = GetComboValue(FrmSyncUpload, "cmbPtm2");
            var CardCode = GetItemValue("edBp2", FrmSyncUpload);
            DateTime FromDate = getdateFor(ed_InFromDate);
            DateTime ToDate = getdateFor(ed_InToDate);

            String CustomerGroupParameter = CustomerGroup;
            String CardCodeParameter = CardCode;
            String PaymentTermsParameter = PaymentTerms;
            String FromDateParameter = "", ToDateParameter;
            string Where = "";
            if (CardCode.Equals(""))
            {
                CardCodeParameter = "%";
            }
            if (PaymentTermsParameter.Equals("All"))
            {
                PaymentTermsParameter = "%";
            }
            if (CustomerGroupParameter.Equals("All"))
            {
                CustomerGroupParameter = "%";
            }
            if (FromDate.Equals(DateTime.MinValue))
            {
                FromDateParameter = " like '%'";
            }
            else
            {
                FromDateParameter = string.Format(" >= cast('{0}' as date)", FromDate.ToString("yyyy/MM/dd"));
            }
            if (ToDate.Equals(DateTime.MinValue))
            {
                ToDateParameter = " like '%'";
            }
            else
            {
                ToDateParameter = string.Format(" <= cast('{0}' as date)", ToDate.ToString("yyyy/MM/dd"));
            }
            Where = String.Format("and C.[CardCode] like '{0}' and CRG.[GroupName] like '{2}'  and CTG.[PymntGroup] like '{3}'  and RDR.[DocDate]  {1} and RDR.[DocDate]  {4}", CardCodeParameter, FromDateParameter, CustomerGroupParameter, PaymentTermsParameter, ToDateParameter);
            ShowSyncUploadInvoice(form, Where);

        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }
    private void BtnClearRangeSO()
    {
        try
        {
            ed_SUFromDate.Value = "";
            ed_SUToDate.Value = "";
        }
        catch (Exception ex)
        {
            var error = ex.Message;
        }
    }
    private void BtnClearRangeInvoice()
    {
        try
        {
            ed_InFromDate.Value = "";
            ed_InToDate.Value = "";
        }
        catch (Exception ex)
        {
            var error = ex.Message;
        }
    }
    private void SelectAllSU(SAPbouiCOM.Form form, String MatxStr, SAPbouiCOM.DataTable oDataTable, string CheckboxID)
    {
        try
        {
            if (form.Items.Item(MatxStr) == null)
                return;

            form.Items.Item(MatxStr).AffectsFormMode = false;
            SAPbouiCOM.Matrix OMatrix = (SAPbouiCOM.Matrix)form.Items.Item(MatxStr).Specific;
            var SelectAllCheckBox = (SAPbouiCOM.CheckBox)form.Items.Item(CheckboxID).Specific;
            int RowCount = oDataTable.Rows.Count;
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
            //if (!SelectAllCheckBox.Checked)
            //{
            //    SelectAllCheckBox.Checked = true;
            //}
            //else {
            //    SelectAllCheckBox.Checked = false;
            //}
            form.Freeze(false);
            form.Update();
            if (flgUpdateCheckbox)
            {
                FrmSyncUpload.DataSources.UserDataSources.Item(CheckboxID).Value = "Y";
            }
            else
            {
                FrmSyncUpload.DataSources.UserDataSources.Item(CheckboxID).Value = "N";
            }
        }
        catch (Exception ex)
        {
            errorLog(ex.Message);
        }
    }

    #endregion

    #region DataTable
    private SAPbouiCOM.DataTable getDataTableSyncUpload(SAPbouiCOM.Form form, string Mtxstr)
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = form.DataSources.DataTables.Item(Mtxstr);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
            oDT = form.DataSources.DataTables.Add(Mtxstr);

        }

        return oDT;
    }
    private void BindMatrixSyncUpload(string MtxStr, SAPbouiCOM.Matrix oMatrix1, string mtCol, string dtCol, string dt = null, bool editable = false)
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


    #endregion

    #region Export CSV
    private void FileBrowser(String Button)
    {
        try
        {
            SAPbouiCOM.DataTable oDtable = null;
            SAPbouiCOM.Matrix oMtx = null;
            String FileName = "";
            if (Button.Equals("btnESO1"))
            {
                FileName = "ExportSO";
                oDtable = oDTSyncUploadSOList1;
                oMtx = oMatrixSyncUploadSOList1;
            }
            else if (Button.Equals("btnEInv"))
            {
                FileName = "ExportInvoice";
                oDtable = oDTSyncUploadInvList1;
                oMtx = oMatrixSyncUploadInvoiceList1;
            }
            else if (Button.Equals("btnESO2"))
            {
                FileName = "ExportSOLog";
                oDtable = oDTSyncUploadSOList1LOG;
                oMtx = oMatrixSyncUploadSOList1LOG;
            }//
            else if (Button.Equals("btnInvEL"))
            {
                FileName = "ExportInvoiceLog";
                oDtable = oDTSyncUploadInvoiceList1LOG;
                oMtx = oMatrixSyncUploadInvoiceList1LOG;
            }
            else if (Button.Equals("btnEI1"))
            {
                FileName = "ExportItem";
                oDtable = oDTSyncUploadItemList1;
                oMtx = oMatrixSyncUploadItemList1;
            }
            else if (Button.Equals("btnEI2"))
            {
                FileName = "ExportItemLog";
                oDtable = oDTSyncUploadItemList1LOG;
                oMtx = oMatrixSyncUploadItemList1LOG;
            }
            else if (Button.Equals("btnBPE1"))
            {
                FileName = "ExportBusinessPartner";
                oDtable = oDTSyncUploadBPList1;
                oMtx = oMatrixSyncUploadBPList1;
            }
            else if (Button.Equals("btnEBp2"))
            {
                FileName = "ExportBusinessPartnerLog";
                oDtable = oDTSyncUploadBPList1LOG;
                oMtx = oMatrixSyncUploadBPList1LOG;
            }
            else if (Button.Equals("ReceivedPayment"))
            {
                FileName = "ReceivedPayment";
                oDtable = odtReceivedPayment;
                oMtx = oMatrixReceivedPayment;
            }//

            if (oDtable != null && oMtx != null)
            {
                SelectFileDialog dialog = new SelectFileDialog("C:\\", FileName,
                    "|.csv*", DialogType.SAVE);
                dialog.Open();
                if (!string.IsNullOrEmpty(dialog.SelectedFile))
                {
                    WriteToCSV(oDtable, oMtx, dialog.SelectedFile + ".csv");
                }
            }
        }
        catch (Exception ex)
        {


        }
    }
    private void WriteToCSV(SAPbouiCOM.DataTable oDataTable, SAPbouiCOM.Matrix oMatrix, string FileName)
    {
        try
        {
            StringBuilder sb = new StringBuilder();
            int RowCount = oDataTable.Rows.Count;
            if (RowCount > 0)
            {

                var columnNamesList = oDataTable.Columns.Cast<SAPbouiCOM.DataColumn>().Select(column => column.Name).ToList();
                sb.AppendLine(string.Join(",", columnNamesList.ToList()));
                for (int i = 0; i < oDataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < columnNamesList.Count(); j++)
                    {
                        var cell = CSVColumeRules(GetMatrixValue(oMatrix, i, columnNamesList.ElementAt(j).ToString()).ToString());
                        sb.Append(cell + ",");
                    }
                    sb.AppendLine();
                }
                System.IO.File.WriteAllText(FileName, sb.ToString());
            }
        }
        catch (Exception ex)
        {
            showMessage("Export File Issue" + ex.Message);
            SBO_Application.MessageBox("Export File Issue" + ex.Message);

        }

    }

    private void WriteToCSVSO(SAPbouiCOM.DataTable oDataTable, SAPbouiCOM.Matrix oMatrix, string FileName)
    {
        try
        {
            StringBuilder sb = new StringBuilder();
            int RowCount = oDataTable.Rows.Count;
            if (RowCount > 0)
            {
                // sb.Append(GetDataTableColum(oDataTable));
                sb.AppendLine("clNo,DocEntry,DocNum,CardCode,CardName,Type,DocDate,DocTotal,Balance,UploadDt");
                for (int i = 0; i < oDataTable.Rows.Count; i++)
                {
                    //clNo,DocEntry,DocNum,CardCode,CardName,Type,DocDate,DocTotal,Balance,UploadDt

                    var clNo = GetMatrixValue(oMatrix, i, "clNo").ToString();
                    var DocEntry = GetMatrixValue(oMatrix, i, "DocEntry").ToString();
                    var DocNum = GetMatrixValue(oMatrix, i, "DocNum").ToString();
                    var CardCode = GetMatrixValue(oMatrix, i, "CardCode").ToString();
                    var CardName = GetMatrixValue(oMatrix, i, "CardName").ToString();
                    var Type = GetMatrixValue(oMatrix, i, "Type").ToString();
                    var DocDate = GetMatrixValue(oMatrix, i, "DocDate").ToString();
                    var DocTotal = CSVColumeRules(GetMatrixValue(oMatrix, i, "DocTotal").ToString());
                    var Balance = CSVColumeRules(GetMatrixValue(oMatrix, i, "Balance").ToString());
                    var UploadDt = GetMatrixValue(oMatrix, i, "UploadDt").ToString();
                    sb.AppendLine(clNo + "," + DocEntry + "," + DocNum + "," + CardCode + "," + CardName + "," + Type + "," + DocDate + "," + DocTotal + "," + Balance + "," + UploadDt);
                }
                System.IO.File.WriteAllText(FileName, sb.ToString());
            }
        }
        catch (Exception ex)
        {
            showMessage("Export File Issue" + ex.Message);
            SBO_Application.MessageBox("Export File Issue" + ex.Message);

        }

    }
    private String CSVColumeRules(String input)
    {
        String output = "";
        try
        {
            output = input.Replace(",", "");
        }
        catch (Exception ex)
        {
            showMessage("Export File Issue" + ex.Message);
        }
        return output;

    }
    private StringBuilder GetDataTableColum(SAPbouiCOM.DataTable oDataTable)
    {
        StringBuilder ColList = new StringBuilder();
        try
        {
            var columnNamesList = oDataTable.Columns.Cast<SAPbouiCOM.DataColumn>().Select(column => column.Name);
            for (int i = 0; i < columnNamesList.Count(); i++)
            {
                ColList.Append(columnNamesList.ElementAt(i).ToString() + ",");
            }
            ColList.Remove(ColList.Length - 1, 1);
            ColList.AppendLine();
        }
        catch (Exception ex)
        {

        }
        return ColList;
    }
    #endregion

    #region Show Data from Log
    private void ShowSyncUploadSOLOG(SAPbouiCOM.Form form, String Where = "")
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxSyncUpSOListStrLOG) == null)
                return;

            form.Items.Item(MtxSyncUpSOListStrLOG).AffectsFormMode = false;
            oMatrixSyncUploadSOList1LOG = (SAPbouiCOM.Matrix)form.Items.Item(MtxSyncUpSOListStrLOG).Specific;
            oMatrixSyncUploadSOList1LOG.Clear();
            oDTSyncUploadSOList1LOG = getDataTableSyncUpload(form, MtxSyncUpSOListStrLOG);


            sql = String.Format("select 'N' as [clSel[,ROW_NUMBER() OVER (ORDER BY R.[DocEntry[) AS [ID[,R.[DocEntry[,R.[DocNum[,R.[DocDate[,R.[DocTotal[,R.[CardCode[,R.[CardName[   ,Logg.[U_Status[ as [Status[ ,[U_CreateDt[ as [UploadDt[ ,OCRD.[Balance[,CASE  WHEN IFNULL(R.[U_EbizChargeMarkPayment[, '') = 'sales' THEN 'Deposit'  WHEN IFNULL(R.[U_EbizChargeMarkPayment[, '') = 'AuthOnly' THEN 'Pre-auth' ELSE '' END as [Type[  from [@CCSOLOG[ Logg inner JOIN ORDR R  on Logg.[U_DocNum[=R.[DocNum[  inner join OCRD on OCRD.[CardCode[=R.[CardCode[").Replace("[", "\"");
            oDTSyncUploadSOList1LOG.ExecuteQuery(sql);

            //  BindMatrixSyncUpload(MtxSyncUpSOListStrLOG, oMatrixSyncUploadSOList1LOG, "clSel", "clSel", null, true);
            BindMatrixSyncUpload(MtxSyncUpSOListStrLOG, oMatrixSyncUploadSOList1LOG, "clNo", "ID");
            BindMatrixSyncUpload(MtxSyncUpSOListStrLOG, oMatrixSyncUploadSOList1LOG, "DocNum", "DocNum");
            BindMatrixSyncUpload(MtxSyncUpSOListStrLOG, oMatrixSyncUploadSOList1LOG, "DocEntry", "DocEntry");
            BindMatrixSyncUpload(MtxSyncUpSOListStrLOG, oMatrixSyncUploadSOList1LOG, "DocDate", "DocDate");
            BindMatrixSyncUpload(MtxSyncUpSOListStrLOG, oMatrixSyncUploadSOList1LOG, "Status", "Status");
            BindMatrixSyncUpload(MtxSyncUpSOListStrLOG, oMatrixSyncUploadSOList1LOG, "CardName", "CardName");
            BindMatrixSyncUpload(MtxSyncUpSOListStrLOG, oMatrixSyncUploadSOList1LOG, "Type", "Type");
            BindMatrixSyncUpload(MtxSyncUpSOListStrLOG, oMatrixSyncUploadSOList1LOG, "DocTotal", "DocTotal");
            BindMatrixSyncUpload(MtxSyncUpSOListStrLOG, oMatrixSyncUploadSOList1LOG, "Balance", "Balance");
            BindMatrixSyncUpload(MtxSyncUpSOListStrLOG, oMatrixSyncUploadSOList1LOG, "UploadDt", "UploadDt");

            oMatrixSyncUploadSOList1LOG.LoadFromDataSource();
            oMatrixSyncUploadSOList1LOG.AutoResizeColumns();

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void ShowSyncUploadInvoiceLOG(SAPbouiCOM.Form form, String Where = "")
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxSyncUpInvoiceListStrLOG) == null)
                return;

            form.Items.Item(MtxSyncUpInvoiceListStrLOG).AffectsFormMode = false;
            oMatrixSyncUploadInvoiceList1LOG = (SAPbouiCOM.Matrix)form.Items.Item(MtxSyncUpInvoiceListStrLOG).Specific;
            oMatrixSyncUploadInvoiceList1LOG.Clear();
            oDTSyncUploadInvoiceList1LOG = getDataTableSyncUpload(form, MtxSyncUpInvoiceListStrLOG);


            sql = String.Format("select 'N' as [clSel[,ROW_NUMBER() OVER (ORDER BY R.[DocEntry[) AS [ID[,R.[DocEntry[,R.[DocNum[,R.[DocDate[,R.[DocTotal[,R.[CardCode[,R.[CardName[   ,Logg.[U_Status[ as [Status[ ,[U_CreateDt[ as [UploadDt[ ,OCRD.[Balance[,CASE  WHEN IFNULL(R.[U_EbizChargeMarkPayment[, '') = 'sales' THEN 'Deposit'  WHEN IFNULL(R.[U_EbizChargeMarkPayment[, '') = 'AuthOnly' THEN 'Pre-auth' ELSE '' END as [Type[  from [@CCINVLOG[ Logg inner JOIN OINV R  on Logg.[U_DocNum[=R.[DocNum[  inner join OCRD on OCRD.[CardCode[=R.[CardCode[  order by Logg.[Code[ desc ").Replace("[", "\"");
            oDTSyncUploadInvoiceList1LOG.ExecuteQuery(sql);

            BindMatrixSyncUpload(MtxSyncUpInvoiceListStrLOG, oMatrixSyncUploadInvoiceList1LOG, "clNo", "ID");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStrLOG, oMatrixSyncUploadInvoiceList1LOG, "DocNum", "DocNum");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStrLOG, oMatrixSyncUploadInvoiceList1LOG, "DocEntry", "DocEntry");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStrLOG, oMatrixSyncUploadInvoiceList1LOG, "DocDate", "DocDate");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStrLOG, oMatrixSyncUploadInvoiceList1LOG, "Status", "Status");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStrLOG, oMatrixSyncUploadInvoiceList1LOG, "CardName", "CardName");
            //BindMatrixSyncUpload(MtxSyncUpInvoiceListStrLOG, oMatrixSyncUploadInvoiceList1LOG, "Type", "Type");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStrLOG, oMatrixSyncUploadInvoiceList1LOG, "DocTotal", "DocTotal");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStrLOG, oMatrixSyncUploadInvoiceList1LOG, "Balance", "Balance");
            BindMatrixSyncUpload(MtxSyncUpInvoiceListStrLOG, oMatrixSyncUploadInvoiceList1LOG, "UploadDt", "UploadDt");

            oMatrixSyncUploadInvoiceList1LOG.LoadFromDataSource();
            oMatrixSyncUploadInvoiceList1LOG.AutoResizeColumns();

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void ShowSyncUpLoadItemListLOG(SAPbouiCOM.Form form, String Where = "")
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxSyncUpItemListStrLOG) == null)
                return;

            form.Items.Item(MtxSyncUpItemListStrLOG).AffectsFormMode = false;
            oMatrixSyncUploadItemList1LOG = (SAPbouiCOM.Matrix)form.Items.Item(MtxSyncUpItemListStrLOG).Specific;
            oMatrixSyncUploadItemList1LOG.Clear();
            oDTSyncUploadItemList1LOG = getDataTableSyncUpload(form, MtxSyncUpItemListStrLOG);
            sql = String.Format("SELECT T0.]U_Status] as [Status[, T0.]U_EbizKey], T0.]U_CreateDt], T0.]U_UpdateDt] as [UpdateDt[,ROW_NUMBER() OVER (ORDER BY ]DocEntry]) AS ]ID],]ItemCode],]ItemName],cast(]OnHand] as int) as ]InStock],]LastPurPrc] as ]UnitPrice] FROM ]@CCITEMLOGTAB]  T0 inner join OITM on OITM.]ItemCode]=T0.]U_ItemCode] ").Replace("]", "\"").Replace("[", "\"");

            oDTSyncUploadItemList1LOG.ExecuteQuery(sql);

            //BindMatrixSyncUpload(MtxSyncUpItemListStrLOG, oMatrixSyncUploadItemList1LOG, "clSel", "clSel", null, true);
            BindMatrixSyncUpload(MtxSyncUpItemListStrLOG, oMatrixSyncUploadItemList1LOG, "clNo", "ID");
            BindMatrixSyncUpload(MtxSyncUpItemListStrLOG, oMatrixSyncUploadItemList1LOG, "ItemCode", "ItemCode");
            BindMatrixSyncUpload(MtxSyncUpItemListStrLOG, oMatrixSyncUploadItemList1LOG, "ItemName", "ItemName");
            BindMatrixSyncUpload(MtxSyncUpItemListStrLOG, oMatrixSyncUploadItemList1LOG, "UnitPrice", "UnitPrice");
            BindMatrixSyncUpload(MtxSyncUpItemListStrLOG, oMatrixSyncUploadItemList1LOG, "InStock", "InStock");
            BindMatrixSyncUpload(MtxSyncUpItemListStrLOG, oMatrixSyncUploadItemList1LOG, "UpdateDt", "UpdateDt");
            BindMatrixSyncUpload(MtxSyncUpItemListStrLOG, oMatrixSyncUploadItemList1LOG, "Status", "Status");

            oMatrixSyncUploadItemList1LOG.LoadFromDataSource();
            oMatrixSyncUploadItemList1LOG.AutoResizeColumns();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void ShowSyncUploadBPListLOG(SAPbouiCOM.Form form, String DefaultCheck = "N")
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxSyncUpBPListStrLOG) == null)
                return;

            form.Items.Item(MtxSyncUpBPListStrLOG).AffectsFormMode = false;
            oMatrixSyncUploadBPList1LOG = (SAPbouiCOM.Matrix)form.Items.Item(MtxSyncUpBPListStrLOG).Specific;
            oMatrixSyncUploadBPList1LOG.Clear();
            oDTSyncUploadBPList1LOG = getDataTableSyncUpload(form, MtxSyncUpBPListStrLOG);
            sql = String.Format("SELECT '{0}' as [clSel],ROW_NUMBER() OVER (ORDER BY T0.[CardCode[) AS [ID[,T0.[CardCode[ as [c[, T0.[CardName[,T0.[Balance[,T0.[E_Mail[, IFNULL(T0.[U_EbizChargeID[,'') as [EBiz[,CASE  WHEN IFNULL(T0.[U_EbizChargeID[, '') = '' THEN 'False' ELSE 'True'END as [Sync[,T10.[U_CreateDt[ as [UpdateDt[,T10.[U_Status[ as [Status[ FROM   [@CCBPLOGTAB[ T10 inner join  OCRD T0 on T0.[CardCode[= T10.[U_CardCode[", DefaultCheck).Replace("]", "\"").Replace("[", "\"");

            oDTSyncUploadBPList1LOG.ExecuteQuery(sql);
            BindMatrixSyncUpload(MtxSyncUpBPListStrLOG, oMatrixSyncUploadBPList1LOG, "clNo", "ID");
            BindMatrixSyncUpload(MtxSyncUpBPListStrLOG, oMatrixSyncUploadBPList1LOG, "colCard", "c");
            BindMatrixSyncUpload(MtxSyncUpBPListStrLOG, oMatrixSyncUploadBPList1LOG, "CardName", "CardName");
            BindMatrixSyncUpload(MtxSyncUpBPListStrLOG, oMatrixSyncUploadBPList1LOG, "Balance", "Balance");
            BindMatrixSyncUpload(MtxSyncUpBPListStrLOG, oMatrixSyncUploadBPList1LOG, "E_Mail", "E_Mail");
            BindMatrixSyncUpload(MtxSyncUpBPListStrLOG, oMatrixSyncUploadBPList1LOG, "Status", "Status");
            BindMatrixSyncUpload(MtxSyncUpBPListStrLOG, oMatrixSyncUploadBPList1LOG, "UpdateDt", "UpdateDt");
            oMatrixSyncUploadBPList1LOG.LoadFromDataSource();
            oMatrixSyncUploadBPList1LOG.AutoResizeColumns();

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    #endregion

    #region Clear Button
    private void ClearLogItem(String ID)
    {
        try
        {
            String Table = "";
            if (ID.Equals("btnIU1"))
            {
                Table = "CCITEMLOGTAB";
            }
            else if (ID.Equals("btnBPClear"))
            {
                Table = "CCBPLOGTAB";
            }
            else if (ID.Equals("btnCSO"))
            {
                Table = "CCSOLOG";
            }
            else if (ID.Equals("btnCInv"))
            {

                Table = "CCINVLOG";
            }

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

    #region Common Function 

    private TransactionResponse RunTransactionInvoice(Invoice OInvoice)
    {
        TransactionResponse oTranResp = null;
        try
        {
            TransactionRequestObject oTran = new TransactionRequestObject();
            oTran.CustomerID = OInvoice.CustomerId;

            oTran.Software = Software;
            oTran.Command = "cc:sale";
            oTran.CustReceipt = cfgsendCustomerEmail == "Y" ? true : false;
            oTran.CustReceiptName = "vterm_customer";
            //oTran.
            LineItem[] Lineitems = new LineItem[OInvoice.Items.Count()];
            for (int i = 0; i < OInvoice.Items.Count(); i++)
            {
                Item item = OInvoice.Items[i];
                Lineitems[i] = new LineItem();
                Lineitems[i].ProductName = item.ItemId;
                Lineitems[i].Qty = ConvertString(item.Qty);
                Lineitems[i].UnitPrice = ConvertString(item.UnitPrice);
                //item.TotalLineAmount = ConvertDecimal(oRS.Fields.Item("LineTotal").Value);
            }
            oTran.LineItems = Lineitems;
            oTran.Details = new TransactionDetail();
            oTran.Details.Invoice = OInvoice.InvoiceNumber;
            oTran.Details.Amount = Convert.ToDouble(OInvoice.InvoiceAmount);
            oTran.Details.Currency = ConvertString(OInvoice.Currency);
            oTran.Details.Subtotal = 0;

            oTran.CreditCardData = new CreditCardData();
            oTran.CreditCardData.CardNumber = "4000100511112229";
            oTran.CreditCardData.CardExpiration = "0922";
            //oTran.CreditCardData.AvsZip = "YYY";

            //   oTran.CreditCardData.AvsStreet = "999";
            //   oTran.CreditCardData.AvsZip = "999";

            //oTran.CreditCardData.
            oTranResp = ebiz.runTransaction(getToken(), oTran);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
        }
        return oTranResp;
    }

    private TransactionResponse RunTransactionOrder(SalesOrder Order)
    {
        TransactionResponse oTranResp = null;
        try
        {
            TransactionRequestObject oTran = new TransactionRequestObject();
            Customer cust = ebiz.GetCustomer(getToken(), Order.CustomerId, null);

            oTran.CustomerID = Order.CustomerId;
            oTran.Software = Software;
            oTran.Command = "cc:sale";
            oTran.CustReceipt = cfgsendCustomerEmail == "Y" ? true : false;
           
            //oTran.AuthCode=
            // oTran.CustReceiptName = "vterm_customer";

            //oTran.
            LineItem[] Lineitems = new LineItem[Order.Items.Count()];
            for (int i = 0; i < Order.Items.Count(); i++)
            {
                Item item = Order.Items[i];
                Lineitems[i] = new LineItem();
                Lineitems[i].ProductName = item.ItemId;
                Lineitems[i].Qty = ConvertString(item.Qty);
                Lineitems[i].UnitPrice = ConvertString(item.UnitPrice);

                //item.TotalLineAmount = ConvertDecimal(oRS.Fields.Item("LineTotal").Value);
            }
            oTran.LineItems = Lineitems;
            oTran.Details = new TransactionDetail();
            //oTran.Details.Amount = Convert.ToDouble(Order.InvoiceAmount);
            oTran.Details.OrderID = Order.SalesOrderNumber;
        //    oTran.Details.Currency = ConvertString(Order.Currency);
            //oTran.Details.Subtotal = Convert.ToDouble(Order.Amount);
            oTran.Details.Amount = Convert.ToDouble(Order.Amount);
            //   oTran.Details.



            oTran.CreditCardData = new CreditCardData();
            oTran.CreditCardData.CardNumber = "4444555566667779";
            oTran.CreditCardData.CardExpiration = "0922";
            

            //   oTran.CreditCardData.AvsStreet = "999";
            //   oTran.CreditCardData.AvsZip = "999";

            //oTran.CreditCardData.
            oTranResp = ebiz.runTransaction(getToken(), oTran);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
        }
        return oTranResp;
    }

    #endregion
}

