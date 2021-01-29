using eBizChargeForSAP.ServiceReference1;
using System;
using System.Threading;

partial class SAP
{


    #region DataMember
    SAPbouiCOM.Form FormRecPay;
    string MtxRecPay = "mtMain";
    string dtRecPay = "dtemp";
    SAPbouiCOM.DataTable odtRecPay;
    bool flgDocInitRecPay = false;
    SAPbouiCOM.Matrix OMtxRecPay;
    bool FlgdtRecPay = true;
    SAPbouiCOM.EditText ed_fromDatePyR, ed_ToDatePyR;
    SAPbouiCOM.ComboBox Cmb_MenuPyR;
    string txtFromdtPyR = "txtFromdt", txtTodtPyR = "txtTodt", cmbMenuStrPyR = "cmbMenu";

    #endregion


    #region FormEvent
    private void CreateRecPay()
    {
        try
        {
            flgDocInitRecPay = false;
            AddXMLForm("Presentation_Layer.SendReceivedEmail.xml");
            FormRecPay.Freeze(true);
            RecPayInit();
            FormRecPay.Freeze(false);
            FlgdtRecPay = true;
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }
    private void RecPayFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            FormRecPay = form;
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
                            if (ValidDateRecPay())
                            {
                                ShowDataRecPay();
                            }

                        }
                        if (pVal.ItemUID == "btnPost")
                        {
                            HandleItemRecPay(form);
                        }
                        if (pVal.ItemUID == MtxRecPay && pVal.ColUID.Equals("clSel") && pVal.Row == 0 && !pVal.BeforeAction)
                        {
                            HandleRecPaySelectALL(form, MtxRecPay, odtRecPay);
                        }
                        if (pVal.ItemUID == MtxRecPay)
                        {
                            if (form.Items.Item(MtxRecPay) == null)
                                return;

                            OMtxRecPay.FlushToDataSource();
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

    private void HandleRecPaySelectALL(SAPbouiCOM.Form form, String MatxStr, SAPbouiCOM.DataTable oDataTable)
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
                    if (FlgdtRecPay)
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
                FlgdtRecPay = !FlgdtRecPay;
            }

        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }

    private void HandleItemRecPay(SAPbouiCOM.Form form)
    {
        bool FlgChange = false;
        try
        {
            if (form.Items.Item(MtxRecPay) == null)
                return;

            form.Items.Item(MtxRecPay).AffectsFormMode = false;
            OMtxRecPay = (SAPbouiCOM.Matrix)form.Items.Item(MtxRecPay).Specific;
            int RowCount = odtRecPay.Rows.Count;
            if (RowCount > 0)
            {
                for (int i = odtRecPay.Rows.Count - 1; i >= 0; i--)
                {
                    var EdTemplate = GetMatrixValue(OMtxRecPay, i, "clEbizkey").ToString();//DocNum
                    var EdDocNum = GetMatrixValue(OMtxRecPay, i, "DocNum").ToString();//

                    var chk = (SAPbouiCOM.CheckBox)OMtxRecPay.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    if (chk.Checked)
                    {
                        FlgChange = true;
                        PostRecPay(EdTemplate, i, OMtxRecPay, form, EdDocNum);
                    }
                    else
                    {

                    }

                }
                if (FlgChange)
                {
                    //   oMatrixBpExport.LoadFromDataSource();
                    OMtxRecPay.FlushToDataSource();
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
    private bool ValidDateRecPay()
    {
        bool flag = false;
        try
        {
            var FromValue = GetItemValue(txtFromdtPyR, FormRecPay);
            var ToDateValue = GetItemValue(txtTodtPyR, FormRecPay);

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
  
    private void RecPayInit()
    {
        try
        {
            flgDocInitRecPay = true;
            //  ShowDataRecPay();
            fillDateFieldsRecPay();
        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
    }
    private void fillDateFieldsRecPay()
    {
        try
        {

            ed_fromDatePyR = ((SAPbouiCOM.EditText)(FormRecPay.Items.Item(txtFromdtPyR).Specific));
            FormRecPay.DataSources.UserDataSources.Add(txtFromdtPyR, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_fromDatePyR.DataBind.SetBound(true, "", txtFromdtPyR);
            ed_fromDatePyR.Value = DatetimeTosapFormatDate(DateTime.Now);

            ed_ToDatePyR = ((SAPbouiCOM.EditText)(FormRecPay.Items.Item(txtTodtPyR).Specific));
            FormRecPay.DataSources.UserDataSources.Add(txtTodtPyR, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_ToDatePyR.DataBind.SetBound(true, "", txtTodtPyR);
            ed_ToDatePyR.Value = DatetimeTosapFormatDate(DateTime.Now);

            Cmb_MenuPyR = ((SAPbouiCOM.ComboBox)(FormRecPay.Items.Item(cmbMenuStrPyR).Specific));
            Cmb_MenuPyR.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    private void ShowDataRecPay()
    {
        try
        {
            tProcess = new Thread(RecPayThreadProc);
            tProcess.Start();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            throw;
        }
    }

    public void RecPayThreadProc()
    {
        try
        {
            if (formSelTarget == null)
            {
                if (FormRecPay != null)
                    formSelTarget = FormRecPay;
                else
                    return;
            }
            if (formSelTarget.TypeEx != formEmailPaymentReceived)
                return;

            if (!flgDocInitRecPay) return;

            ShowRecPay(FormRecPay);

        }
        catch (Exception ex)
        {
        }

    }

    private void ShowRecPay(SAPbouiCOM.Form form)
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxRecPay) == null)
                return;

            DateTime FromdtValue = getdate(ed_fromDatePyR);
            DateTime TodtValue = getdate(ed_ToDatePyR).Add(DateTime.MaxValue.TimeOfDay);
            var filters = new SearchFilter[1];
            filters[0] = new SearchFilter();

            filters[0].FieldName = "PaymentInternalId";
            filters[0].ComparisonOperator = "eq";
            filters[0].FieldValue = "7331bcf0-4ea4-4259-a14f-3343233d295a";

            var PendingPaymentsList = ebiz.SearchEbizWebFormReceivedPayments(getToken(), "", FromdtValue, TodtValue, null, 0, 100, "");
            


            form.Items.Item(MtxRecPay).AffectsFormMode = false;
            OMtxRecPay = (SAPbouiCOM.Matrix)form.Items.Item(MtxRecPay).Specific;
            OMtxRecPay.Clear();
            odtRecPay = GetDatatableRecPay(form);

            sql = String.Format("select  top 1'N' as [clSel],ROW_NUMBER() OVER (ORDER BY [DocEntry]) AS [ID],[DocNum],[DocDate],[DocTotal],[CardCode],[CardName],IFNULL([U_EbizChargeID],'') as [EBiz],CASE WHEN IFNULL([U_EbizChargeID], '') = '' THEN 'False' ELSE 'True'END as [Sync] from ORDR where  [DocStatus]='100' ").Replace("]", "\"").Replace("[", "\"");
            odtRecPay.ExecuteQuery(sql);
            BindMtxRecPay(OMtxRecPay, "clSel", "clSel", dtItem, true);
            BindMtxRecPay(OMtxRecPay, "clNo", "ID");
            BindMtxRecPay(OMtxRecPay, "DocNum", "DocNum");
            BindMtxRecPay(OMtxRecPay, "DocDate", "DocDate");
            BindMtxRecPay(OMtxRecPay, "CardCode", "CardCode");
            BindMtxRecPay(OMtxRecPay, "Email", "CardName");
            BindMtxRecPay(OMtxRecPay, "DocTotal", "DocTotal");
            BindMtxRecPay(OMtxRecPay, "clEbizkey", "EBiz");
            BindMtxRecPay(OMtxRecPay, "clSync", "Sync");

            OMtxRecPay.LoadFromDataSource();
            OMtxRecPay.Clear();
            foreach (var itm in PendingPaymentsList)
            {
                OMtxRecPay.AddRow(1);
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)OMtxRecPay.Columns.Item("CardCode").Cells.Item(OMtxRecPay.RowCount).Specific;
                oEdit.Value = itm.CustomerId;
                oEdit = (SAPbouiCOM.EditText)OMtxRecPay.Columns.Item("DocTotal").Cells.Item(OMtxRecPay.RowCount).Specific;
                oEdit.Value = itm.AmountDue;
                oEdit = (SAPbouiCOM.EditText)OMtxRecPay.Columns.Item("DocNum").Cells.Item(OMtxRecPay.RowCount).Specific;
                oEdit.Value = itm.InvoiceNumber;
                oEdit = (SAPbouiCOM.EditText)OMtxRecPay.Columns.Item("clEbizkey").Cells.Item(OMtxRecPay.RowCount).Specific;
                oEdit.Value = itm.PaymentInternalId;
                oEdit = (SAPbouiCOM.EditText)OMtxRecPay.Columns.Item("Email").Cells.Item(OMtxRecPay.RowCount).Specific;
                oEdit.Value = itm.CustomerEmailAddress;
                oEdit = (SAPbouiCOM.EditText)OMtxRecPay.Columns.Item("DocDate").Cells.Item(OMtxRecPay.RowCount).Specific;
                var date = DatetimeTosapFormatDate(Convert.ToDateTime(itm.PaymentRequestDateTime));
                oEdit.Value = date;
            }
            OMtxRecPay.AutoResizeColumns();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }

    private SAPbouiCOM.DataTable GetDatatableRecPay(SAPbouiCOM.Form form)
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = form.DataSources.DataTables.Item(MtxRecPay);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
            oDT = form.DataSources.DataTables.Add(MtxRecPay);

        }

        return oDT;
    }
    private void BindMtxRecPay(SAPbouiCOM.Matrix oMatrix1, string mtCol, string dtCol, string dt = null, bool editable = false)
    {
        try
        {
            SAPbouiCOM.Column oColumn = oMatrix1.Columns.Item(mtCol);
            oColumn.Editable = editable;
            if (dt == null)
            {
                oColumn.DataBind.Bind(MtxRecPay, dtCol);
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

    private void PostRecPay(String TemplateID, int i, SAPbouiCOM.Matrix oMatx, SAPbouiCOM.Form Oform,String DocNum)
    {
        String EbizKey = "";
        bool PostStatus = false;
        try
        {
            var Menu = GetComboValue(Oform, cmbMenuStrPyR);
            PaymentResponse oPaymentResponse;

            if (Menu.Equals("Mark Payment"))
            {
                oPaymentResponse = ebiz.MarkEbizWebFormPaymentAsApplied(getToken(), TemplateID);
                if (oPaymentResponse.Status.Equals("Success"))
                {
                    PostStatus = true;
                }
                else
                {
                    PostStatus = false;
                }
                
            } else if (Menu.Equals("Delete"))
            {
                var ReminderResp = ebiz.DeleteEbizWebFormPayment(getToken(), TemplateID);
                if (ReminderResp.Equals("1"))
                {
                    PostStatus = true;
                }
                else
                {
                    PostStatus = false;
                }

            }
            else if(Menu.Equals("Apply"))
            {//Delete
                oPaymentResponse = ebiz.MarkEbizWebFormPaymentAsApplied(getToken(), TemplateID);
                //Incoming Payment 
                if (oPaymentResponse.Status.Equals("Success"))
                {
                    PostStatus = true;
                    updateSOPayment(DocNum);
                }
                else
                {
                    PostStatus = false;
                }

            }

            if (PostStatus)
            {
                //SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clEbizkey").Cells.Item(i + 1).Specific;
                //oEdit.Value = EbizKey;
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
                oEdit.Value = "True";
            }
            else
            {
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
                oEdit.Value = "Fasle";

            }
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }

}
