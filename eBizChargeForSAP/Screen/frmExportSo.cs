using eBizChargeForSAP.ServiceReference1;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

partial class SAP
{


    #region DataMember
    SAPbouiCOM.Form theExportSO;
    string MtxExportSO = "mtMain";
    //mtxSO1
    string dtExportSO = "dtemp";
    SAPbouiCOM.DataTable oDTExportSO;
    bool ExportSOflgDocInitiallization = false;
    SAPbouiCOM.Matrix oMatrixExportSO;
    bool flStateExportSO = true;
    SAPbouiCOM.ComboBox Cmb_CustomerGrp, Cmb_PaymentsTerm;
    SAPbouiCOM.EditText ed_FindDocDate;
    string txtFindDocDate = "txtdocDt";
    #endregion

    #region FormEvent
    private void CreateExportSO()
    {
        try
        {
            ExportSOflgDocInitiallization = false;
            AddXMLForm("Presentation_Layer.ExportSO.xml");
            theExportSO.Freeze(true);
            InitializeFormExportSO();
            theExportSO.Freeze(false);
            flStateExportSO = true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void ExportSOFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            theExportSO = form;
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


                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        {

                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                        try
                        {
                            if (pVal.ItemUID == MtxExportSO & pVal.ColUID == "DocNum")
                            {//SO
                                SAPbouiCOM.EditText DocNum = (SAPbouiCOM.EditText)oMatrixExportSO.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific;
                                DoOpenLinkedObjectForm("139", "2050", "8", DocNum.Value);
                            }
                            if (pVal.ItemUID == MtxExportSO & pVal.ColUID == "DocEntry")
                            {//SO
                                SAPbouiCOM.EditText DocEntry = (SAPbouiCOM.EditText)oMatrixExportSO.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific;
                                var DocNum = getSODocNum(Convert.ToInt32(DocEntry.Value));
                                DoOpenLinkedObjectForm("139", "2050", "8", DocNum);
                            }
                            if (pVal.ItemUID == MtxExportSO & pVal.ColUID == "CardCode")
                            {//SO
                                SAPbouiCOM.EditText CardCode = (SAPbouiCOM.EditText)oMatrixExportSO.Columns.Item("CardCode").Cells.Item(pVal.Row).Specific;
                                DoOpenLinkedObjectForm("134", "2561", "5", CardCode.Value);
                            }
                        }
                        catch (Exception ex)
                        {
                            var error = ex.Message;

                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == "btnPost")
                        {
                            HandleItemPressIExportSO(form);
                        }
                        if (pVal.ItemUID == "btnsrch")
                        {
                            SearchExportSO(form);
                        }
                        if (pVal.ItemUID == MtxExportSO && pVal.ColUID.Equals("clSel") && pVal.Row == 0 && !pVal.BeforeAction)
                        {
                            HandleExportSOSelectAll(form, MtxExportSO, oDTExportSO);
                        }
                        if (pVal.ItemUID == MtxExportSO)
                        {
                            if (form.Items.Item(MtxExportSO) == null)
                                return;

                            oMatrixExportSO.FlushToDataSource();
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

    private void HandleExportSOSelectAll(SAPbouiCOM.Form form, String MatxStr, SAPbouiCOM.DataTable oDataTable)
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
            {
                ShowMsgCustom("Selecting ...!");
                for (int i = 0; i < oDataTable.Rows.Count; i++)
                {
                    var chk = (SAPbouiCOM.CheckBox)OMatrix.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    if (flStateExportSO)
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
                flStateExportSO = !flStateExportSO;

                //    oMatrixBpExport.FlushToDataSource();
            }

        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }

    private void HandleItemPressIExportSO(SAPbouiCOM.Form form)
    {
        bool FlgChange = false;
        try
        {
            if (form.Items.Item(MtxExportSO) == null)
                return;

            form.Items.Item(MtxExportSO).AffectsFormMode = false;
            oMatrixExportSO = (SAPbouiCOM.Matrix)form.Items.Item(MtxExportSO).Specific;
            int RowCount = oDTExportSO.Rows.Count;
            if (RowCount > 0)
            {
                for (int i = oDTExportSO.Rows.Count - 1; i >= 0; i--)
                {
                    var EdDocNum = GetMatrixValue(oMatrixExportSO, i, "DocNum").ToString();

                    var chk = (SAPbouiCOM.CheckBox)oMatrixExportSO.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    if (chk.Checked)
                    {
                        FlgChange = true;
                        SalesOrder oSalesOrder = GetSO(EdDocNum);
                        PostSO(oSalesOrder, i, oMatrixExportSO);
                    }
                    else
                    {

                    }

                }
                if (FlgChange)
                {
                    //   oMatrixBpExport.LoadFromDataSource();
                    oMatrixExportSO.FlushToDataSource();
                }
            }
        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }
    private void SearchExportSO(SAPbouiCOM.Form form)
    {
        try
        {
            if (form.Items.Item(MtxExportSO) == null)
                return;

            var CustomerGroup = GetComboValue(theExportSO, "cmbCgrp");
            var PaymentTerms = GetComboValue(theExportSO, "cmbPtm");
            var CardCode = GetItemValue("txtBp", theExportSO);
            DateTime DocDate = getdateFor(ed_FindDocDate);

            String CustomerGroupParameter = CustomerGroup;
            String CardCodeParameter = CardCode;
            String PaymentTermsParameter = PaymentTerms;
            String DateParameter = "";
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
            if (DocDate.Equals(DateTime.MinValue))
            {
                DateParameter = " like '%'";
            }
            else
            {
                DateParameter = string.Format(" >= cast('{0}' as date)", DocDate.ToString("yyyy/MM/dd"));
            }
            Where = String.Format("and C.[CardCode] like '{0}' and RDR.[DocDate]  {1} and CRG.[GroupName] like '{2}'  and CTG.[PymntGroup] like '{3}'", CardCodeParameter, DateParameter, CustomerGroupParameter, PaymentTermsParameter);
            ShowExportSO(form, Where);





        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }
    #endregion

    #region Helping Function
    private void InitializeFormExportSO()
    {
        try
        {
            ExportSOflgDocInitiallization = true;
            ShowDataExportSO();
        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
    }

    private void ShowDataExportSO()
    {
        try
        {
            tProcess = new Thread(ExportSOCreateTabThreadProc);
            tProcess.Start();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            throw;
        }
    }
    public void ExportSOCreateTabThreadProc()
    {
        try
        {
            if (formSelTarget == null)
            {
                if (theExportSO != null)
                    formSelTarget = theExportSO;
                else
                    return;
            }
            if (formSelTarget.TypeEx != formExportSO)
                return;

            if (!ExportSOflgDocInitiallization) return;


            ShowExportSO(theExportSO);
            FillExportSO();
        }
        catch (Exception ex)
        {
        }

    }
    private void ShowExportSO(SAPbouiCOM.Form form, String Where = "")
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxExportSO) == null)
                return;

            form.Items.Item(MtxExportSO).AffectsFormMode = false;
            oMatrixExportSO = (SAPbouiCOM.Matrix)form.Items.Item(MtxExportSO).Specific;
            oMatrixExportSO.Clear();
            oDTExportSO = getDataTableExportSo(form);
            if (Where.Equals(""))
            {
                sql = String.Format(" select 'N' as [clSel[,ROW_NUMBER() OVER (ORDER BY R.[DocEntry[) AS [ID[,R.[DocEntry[,R.[DocNum[,R.[DocDate[,R.[DocTotal[,R.[CardCode[,R.[CardName[,  IFNULL(R.[U_EbizChargeID[, '') as [EBiz[, 'False' as [Sync[  , (select Max([U_Status[) from[@CCSOLOG[where[U_DocNum[= Cast(R.[DocNum[ as nvarchar(20)))as [Status[ , (select Min([U_CreateDt[) from[@CCSOLOG[where[U_DocNum[= Cast(R.[DocNum[ as nvarchar(20)) )as [UploadDt[ , (select Max([U_UpdateDt[) from[@CCSOLOG[where[U_DocNum[= Cast(R.[DocNum[ as nvarchar(20))  )as [TransDt[ from ORDR R where R.[DocStatus[= 'O'   and IFNULL(R.[U_EbizChargeMarkPayment[, '') <> 'Paid' and R.\"DocDate\" =current_date  ").Replace("]", "\"").Replace("[", "\"");
            }
            else
            {
                var subQ = " ,( select Max([U_Status[) from [@CCSOLOG[  where [U_DocNum[=Cast(RDR.[DocNum[ as nvarchar(20)))as [Status[ ,(select Min([U_CreateDt[) from [@CCSOLOG[where[U_DocNum[= Cast(RDR.[DocNum[ as nvarchar(20)) )as [UploadDt[ , (select Max([U_UpdateDt[) from [@CCSOLOG[ where [U_DocNum[= Cast(RDR.[DocNum[ as nvarchar(20))  )as [TransDt[";
                sql = String.Format("select 'N' as ]clSel],ROW_NUMBER() OVER (ORDER BY RDR.[DocEntry]) AS ]ID],RDR.[DocEntry[,RDR.[DocNum],RDR.[DocDate],RDR.[DocTotal],RDR.[CardCode],  RDR.[CardName], IFNULL(RDR.[U_EbizChargeID], '') as ]EBiz], 'False'  as ]Sync] {1}  from ORDR RDR  inner  join OCRD C on C.[CardCode] = RDR.[CardCode]  inner  join OCRG CRG on C.[GroupCode] = CRG.[GroupCode]  INNER JOIN OCTG CTG ON C.[GroupNum] = CTG.[GroupNum]  where RDR.[DocStatus]='O'   and IFNULL(RDR.[U_EbizChargeMarkPayment],'')<>'Paid' {0}", Where, subQ).Replace("[", "\"").Replace("]", "\"");
            }

            oDTExportSO.ExecuteQuery(sql);

            BindMatrixExportSO(oMatrixExportSO, "clSel", "clSel", dtItem, true);
            BindMatrixExportSO(oMatrixExportSO, "clNo", "ID");
            BindMatrixExportSO(oMatrixExportSO, "DocNum", "DocNum");
            BindMatrixExportSO(oMatrixExportSO, "DocEntry", "DocEntry");
            BindMatrixExportSO(oMatrixExportSO, "DocDate", "DocDate");
            BindMatrixExportSO(oMatrixExportSO, "CardCode", "CardCode");
            BindMatrixExportSO(oMatrixExportSO, "CardName", "CardName");
            BindMatrixExportSO(oMatrixExportSO, "DocTotal", "DocTotal");
            BindMatrixExportSO(oMatrixExportSO, "clSync", "Sync");
            BindMatrixExportSO(oMatrixExportSO, "Status", "Status");
            BindMatrixExportSO(oMatrixExportSO, "UploadDt", "UploadDt");
            BindMatrixExportSO(oMatrixExportSO, "TransDt", "TransDt");


            oMatrixExportSO.LoadFromDataSource();
            oMatrixExportSO.AutoResizeColumns();

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }

    private SAPbouiCOM.DataTable getDataTableExportSo(SAPbouiCOM.Form form)
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = form.DataSources.DataTables.Item(MtxExportSO);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
            oDT = form.DataSources.DataTables.Add(MtxExportSO);

        }

        return oDT;
    }
    private void BindMatrixExportSO(SAPbouiCOM.Matrix oMatrix1, string mtCol, string dtCol, string dt = null, bool editable = false)
    {
        try
        {
            SAPbouiCOM.Column oColumn = oMatrix1.Columns.Item(mtCol);
            oColumn.Editable = editable;
            if (dt == null)
            {
                oColumn.DataBind.Bind(MtxExportSO, dtCol);
            }
            else
                oColumn.DataBind.Bind(dt, dtCol);
        }
        catch (Exception ex)
        {
            errorLog("Can not bind " + mtCol + " to " + dtCol + ".  error: " + ex.Message);
        }

    }
    private void FillExportSO()
    {
        try
        {
            FillCustomerType();
            FillPaymentTerms();
            FillFindDate();
        }
        catch
        {
        }
    }
    private void FillCustomerType()
    {
        try
        {
            Cmb_CustomerGrp = ((SAPbouiCOM.ComboBox)(theExportSO.Items.Item("cmbCgrp").Specific));
            //ComboAddItem(Cmb_CustomerGrp, "");
            ComboAddItem(Cmb_CustomerGrp, "All");
            List<string> list = GetGroupNames();
            foreach (string c in list)
            {
                ComboAddItem(Cmb_CustomerGrp, c);
            }
            Cmb_CustomerGrp.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        }
        catch
        {

        }
    }
    private void FillPaymentTerms()
    {
        try
        {
            Cmb_PaymentsTerm = ((SAPbouiCOM.ComboBox)(theExportSO.Items.Item("cmbPtm").Specific));
            ComboAddItem(Cmb_PaymentsTerm, "All");
            List<string> list = GetPaymentTerms();
            foreach (string c in list)
            {
                ComboAddItem(Cmb_PaymentsTerm, c);
            }
            Cmb_PaymentsTerm.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        }
        catch
        {

        }
    }
    private void FillFindDate()
    {
        try
        {
            ed_FindDocDate = ((SAPbouiCOM.EditText)(theExportSO.Items.Item(txtFindDocDate).Specific));
            theExportSO.DataSources.UserDataSources.Add(txtFindDocDate, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_FindDocDate.DataBind.SetBound(true, "", txtFindDocDate);
            //ed_ToDate.Value = DatetimeTosapFormatDate(DateTime.Now);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
        }
    }

    #endregion

    private void PostSO(SalesOrder oSalesOrder, int i, SAPbouiCOM.Matrix oMatx)
    {
        String EbizKey = "";
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
            String Refnum = "";
            SalesOrderResponse resp = ebiz.AddSalesOrder(getToken(), oSalesOrder);
            if (resp.ErrorCode == 2 && resp.Error.Equals("Record already exists") && resp.Status.Equals("Error"))
            {
                
                //Update
                SalesOrder OldSalesOrder = ebiz.GetSalesOrder(getToken(), oSalesOrder.CustomerId, "", oSalesOrder.SalesOrderNumber, "");
                UpdateEbizChargeKey("ORDR", OldSalesOrder.SalesOrderInternalId, oSalesOrder.SalesOrderNumber);
                ///
                SalesOrder updateSO = new SalesOrder();
                updateSO.CustomerId = oSalesOrder.CustomerId.Trim();
                updateSO.SalesOrderInternalId = null;
                updateSO.Software = Software;
                updateSO.SalesOrderNumber = oSalesOrder.SalesOrderNumber;
                updateSO.Currency = oSalesOrder.Currency;
                updateSO.Date = oSalesOrder.Date;
                updateSO.DueDate = oSalesOrder.DueDate;
                updateSO.Amount = oSalesOrder.Amount;
                updateSO.AmountDue = oSalesOrder.AmountDue;
                updateSO.NotifyCustomer = false;
                updateSO.TotalTaxAmount = 0;
                updateSO.Memo = "";
                updateSO.ShipDate = oSalesOrder.DueDate;
                updateSO.ShipVia = "";
                updateSO.Items = oSalesOrder.Items;

                var updateSoResp = ebiz.UpdateSalesOrder(getToken(), updateSO, oSalesOrder.CustomerId.Trim(), null, oSalesOrder.SalesOrderNumber, null);
                if (updateSoResp.Status.Equals(""))
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
            }

            if (PostStatus)
            {
                //SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clEbizkey").Cells.Item(i + 1).Specific;
                //oEdit.Value = EbizKey;
                //U_PaidAmt,U_RefNum,U_CustomerID,U_DocEntry,U_Balance
                LogTable oExportSOLogTable = new LogTable();
                oExportSOLogTable.DocNum = oSalesOrder.SalesOrderNumber;
                oExportSOLogTable.DocType = "SO";
                oExportSOLogTable.Status = "Uploaded";
                oExportSOLogTable.PaymentStatus = "";
                oExportSOLogTable.RefNum = Refnum;
                oExportSOLogTable.CardCode = oSalesOrder.CustomerId;
                oExportSOLogTable.DocEntry = oSalesOrder.PoNum;
                oExportSOLogTable.AmountPaid = 0;
                oExportSOLogTable.Amount = ConvertDecimal(oSalesOrder.Amount.ToString());
                oExportSOLogTable.AmountDue = ConvertDecimal(oSalesOrder.AmountDue.ToString());

                if (flgAddNewRecord)
                {
                    oExportSOLogTable.CreateDt = DateTime.Now;
                    oExportSOLogTable.DocState = "Add";
                }
                else
                {
                    oExportSOLogTable.CreateDt = DateTime.Now;
                    oExportSOLogTable.DocState = "Update";
                }
                oExportSOLogTable.UpdateDt = DateTime.Now;

                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
                oEdit.Value = "True";
                oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("Status").Cells.Item(i + 1).Specific;
                oEdit.Value = oExportSOLogTable.Status;
                oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("UploadDt").Cells.Item(i + 1).Specific;
                if (oExportSOLogTable.CreateDt != DateTime.MinValue)
                    oEdit.Value = DatetimeTosapFormatDate(oExportSOLogTable.CreateDt);
                oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("TransDt").Cells.Item(i + 1).Specific;
                oEdit.Value = DatetimeTosapFormatDate(oExportSOLogTable.UpdateDt);
                AddSOLog(oExportSOLogTable);

            }
            else
            {
                //SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clEbizkey").Cells.Item(i + 1).Specific;
                //oEdit.Value = resp.Error;
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
                oEdit.Value = "False";

            }
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }

}

public class LogTable
{
    public String DocEntry,DocNum = "",ItemName="",ItemCode="",CardCode="",CardName="",RefNum="";
    public String DocType = "",Type="",DocState="";
    public String Ebiz = "";
    public String Status = "";
    public String Email = "", Method = "", DocSent = "", Software = "";
    public decimal AmountPaid = 0,Balance=0,Quantity=0,UnitPrice=0,Amount,AmountDue;
    public String DatePaid;

    public String PaymentStatus = "";
    public DateTime CreateDt;
    public DateTime UpdateDt;

}

