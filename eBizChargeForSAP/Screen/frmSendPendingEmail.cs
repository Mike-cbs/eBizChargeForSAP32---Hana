using eBizChargeForSAP.ServiceReference1;
using System;
using System.Threading;

partial class SAP
{


    #region DataMember
    SAPbouiCOM.Form FormSendPendingEmail;
    string MtxSendPendingEmail = "mtMain";
    string dtSendPendingEmail = "dtemp";
    SAPbouiCOM.DataTable odtSendPendingEmail;
    bool SendPendingEmailflgDocInitiallization = false;
    SAPbouiCOM.Matrix oMatrixPendEmail;
    bool flStatePendEmail = true;
    SAPbouiCOM.EditText ed_fromDate, ed_ToDate, ed_SendNewfromDate, ed_SendNewToDate;
    SAPbouiCOM.ComboBox Cmb_Menu;
    string txtFromdt = "txtFromdt", txtTodt = "txtTodt", cmbMenuString = "cmbMenu";
    string txtSendNewFromdt = "txtFrmDt", txtSendNewTodt = "txtTDt", txtSendNewSO = "txtSO";

    #endregion


    #region FormEvent
    private void CreatePendEmail()
    {
        try
        {
            SendPendingEmailflgDocInitiallization = false;
            AddXMLForm("Presentation_Layer.SendPendingEmail.xml");
            FormSendPendingEmail.Freeze(true);
            SendPendingEmailInit();
            FormSendPendingEmail.Freeze(false);
            flStatePendEmail = true;
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }
    private void PendEmailFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            FormSendPendingEmail = form;
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
                            if (ValidPendingEmailSearch())
                            {
                                ShowDataSendPendingEmail();
                            }

                        }
                        if (pVal.ItemUID == "btnPost")
                        {
                            HandleItemPressPendEmail(form);
                        }
                        if (pVal.ItemUID == MtxSendPendingEmail && pVal.ColUID.Equals("clSel") && pVal.Row == 0 && !pVal.BeforeAction)
                        {
                            HandlePendEmailSelectALL(form, MtxSendPendingEmail, odtSendPendingEmail);
                        }
                        if (pVal.ItemUID == MtxSendPendingEmail)
                        {
                            if (form.Items.Item(MtxSendPendingEmail) == null)
                                return;

                            oMatrixPendEmail.FlushToDataSource();
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
    private void HandlePendEmailSelectALL(SAPbouiCOM.Form form, String MatxStr, SAPbouiCOM.DataTable oDataTable)
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
                    if (flStatePendEmail)
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
                flStatePendEmail = !flStatePendEmail;
            }

        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }
    private void HandleItemPressPendEmail(SAPbouiCOM.Form form)
    {
        bool FlgChange = false;
        try
        {
            if (form.Items.Item(MtxSendPendingEmail) == null)
                return;

            form.Items.Item(MtxSendPendingEmail).AffectsFormMode = false;
            oMatrixPendEmail = (SAPbouiCOM.Matrix)form.Items.Item(MtxSendPendingEmail).Specific;
            int RowCount = odtSendPendingEmail.Rows.Count;
            if (RowCount > 0)
            {
                for (int i = odtSendPendingEmail.Rows.Count - 1; i >= 0; i--)
                {
                    var EdTemplate = GetMatrixValue(oMatrixPendEmail, i, "clEbizkey").ToString();

                    var chk = (SAPbouiCOM.CheckBox)oMatrixPendEmail.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    if (chk.Checked)
                    {
                        FlgChange = true;
                        PostPendEmail(EdTemplate, i, oMatrixPendEmail, form);
                    }
                    else
                    {

                    }

                }
                if (FlgChange)
                {
                    //   oMatrixBpExport.LoadFromDataSource();
                    oMatrixPendEmail.FlushToDataSource();
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
    private bool ValidPendingEmailSearch()
    {
        bool flag = false;
        try
        {
            var FromValue = GetItemValue(txtFromdt, FormSendPendingEmail);
            var ToDateValue = GetItemValue(txtTodt, FormSendPendingEmail);

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

    private bool ValidPendingEmailSearch(SAPbouiCOM.Form oFrom)
    {
        bool flag = false;
        try
        {
            var FromValue = GetItemValue(txtFromdt, oFrom);
            var ToDateValue = GetItemValue(txtTodt, oFrom);

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
    private void SendPendingEmailInit()
    {
        try
        {
            SendPendingEmailflgDocInitiallization = true;
            //  ShowDataSendPendingEmail();
            fillDateFields();
        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
    }
    private void fillDateFields()
    {
        try
        {

            ed_fromDate = ((SAPbouiCOM.EditText)(FormSendPendingEmail.Items.Item(txtFromdt).Specific));
            FormSendPendingEmail.DataSources.UserDataSources.Add(txtFromdt, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_fromDate.DataBind.SetBound(true, "", txtFromdt);
            ed_fromDate.Value = DatetimeTosapFormatDate(DateTime.Now);

            ed_ToDate = ((SAPbouiCOM.EditText)(FormSendPendingEmail.Items.Item(txtTodt).Specific));
            FormSendPendingEmail.DataSources.UserDataSources.Add(txtTodt, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_ToDate.DataBind.SetBound(true, "", txtTodt);
            ed_ToDate.Value = DatetimeTosapFormatDate(DateTime.Now);

            Cmb_Menu = ((SAPbouiCOM.ComboBox)(FormSendPendingEmail.Items.Item(cmbMenuString).Specific));
            Cmb_Menu.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    private void ShowDataSendPendingEmail()
    {
        try
        {
            tProcess = new Thread(PendingEmailThreadProc);
            tProcess.Start();
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }

    public void PendingEmailThreadProc()
    {
        try
        {
            if (formSelTarget == null)
            {
                if (FormSendPendingEmail != null)
                    formSelTarget = FormSendPendingEmail;
                else
                    return;
            }
            if (formSelTarget.TypeEx != formEmailPaymentPending)
                return;

            if (!SendPendingEmailflgDocInitiallization) return;

            ShowSendPendEmail(FormSendPendingEmail);

        }
        catch (Exception ex)
        {
        }

    }
    private void ShowSendPendEmail(SAPbouiCOM.Form form)
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxSendPendingEmail) == null)
                return;

            DateTime FromdtValue = getdate(ed_fromDate);
            DateTime TodtValue = getdate(ed_ToDate).Add(DateTime.MaxValue.TimeOfDay);
            var PendingPaymentsList = ebiz.SearchEbizWebFormPendingPayments(getToken(), "", FromdtValue, TodtValue, null, 0, 100, "");

            //  var PaymentResp = ebiz.DeleteEbizWebFormPayment(getToken(), "");
            // var ReminderResp = ebiz.ResendEbizWebFormEmail(getToken(), "");


            form.Items.Item(MtxSendPendingEmail).AffectsFormMode = false;
            oMatrixPendEmail = (SAPbouiCOM.Matrix)form.Items.Item(MtxSendPendingEmail).Specific;
            oMatrixPendEmail.Clear();
            odtSendPendingEmail = getDataTablePendEmail(form);

            sql = String.Format("select  top 1'N' as [clSel],ROW_NUMBER() OVER (ORDER BY [DocEntry]) AS [ID],[DocNum],[DocDate],[DocTotal],[CardCode],[CardName],IFNULL([U_EbizChargeID],'') as [EBiz],CASE WHEN IFNULL([U_EbizChargeID], '') = '' THEN 'False' ELSE 'True'END as [Sync] from ORDR where  [DocStatus]='100' ").Replace("]", "\"").Replace("[", "\"");
            odtSendPendingEmail.ExecuteQuery(sql);
            BindMatrixPendEmail(oMatrixPendEmail, "clSel", "clSel", dtItem, true);
            BindMatrixPendEmail(oMatrixPendEmail, "clNo", "ID");
            BindMatrixPendEmail(oMatrixPendEmail, "DocNum", "DocNum");
            BindMatrixPendEmail(oMatrixPendEmail, "DocDate", "DocDate");
            BindMatrixPendEmail(oMatrixPendEmail, "CardCode", "CardCode");
            BindMatrixPendEmail(oMatrixPendEmail, "Email", "CardName");
            BindMatrixPendEmail(oMatrixPendEmail, "DocTotal", "DocTotal");
            BindMatrixPendEmail(oMatrixPendEmail, "clEbizkey", "EBiz");
            BindMatrixPendEmail(oMatrixPendEmail, "clSync", "Sync");

            oMatrixPendEmail.LoadFromDataSource();
            oMatrixPendEmail.Clear();
            foreach (var itm in PendingPaymentsList)
            {
                oMatrixPendEmail.AddRow(1);
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatrixPendEmail.Columns.Item("CardCode").Cells.Item(oMatrixPendEmail.RowCount).Specific;
                oEdit.Value = itm.CustomerId;
                oEdit = (SAPbouiCOM.EditText)oMatrixPendEmail.Columns.Item("DocTotal").Cells.Item(oMatrixPendEmail.RowCount).Specific;
                oEdit.Value = itm.AmountDue;
                oEdit = (SAPbouiCOM.EditText)oMatrixPendEmail.Columns.Item("DocNum").Cells.Item(oMatrixPendEmail.RowCount).Specific;
                oEdit.Value = itm.InvoiceNumber;
                oEdit = (SAPbouiCOM.EditText)oMatrixPendEmail.Columns.Item("clEbizkey").Cells.Item(oMatrixPendEmail.RowCount).Specific;
                oEdit.Value = itm.PaymentInternalId;
                oEdit = (SAPbouiCOM.EditText)oMatrixPendEmail.Columns.Item("Email").Cells.Item(oMatrixPendEmail.RowCount).Specific;
                oEdit.Value = itm.CustomerEmailAddress;
                oEdit = (SAPbouiCOM.EditText)oMatrixPendEmail.Columns.Item("DocDate").Cells.Item(oMatrixPendEmail.RowCount).Specific;
                var date = DatetimeTosapFormatDate(Convert.ToDateTime(itm.PaymentRequestDateTime));
                oEdit.Value = date;
            }
            oMatrixPendEmail.AutoResizeColumns();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }

    private SAPbouiCOM.DataTable getDataTablePendEmail(SAPbouiCOM.Form form)
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = form.DataSources.DataTables.Item(MtxSendPendingEmail);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
            oDT = form.DataSources.DataTables.Add(MtxSendPendingEmail);

        }

        return oDT;
    }
    private SAPbouiCOM.DataTable getDataTableReceivedPay(SAPbouiCOM.Form form)
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = form.DataSources.DataTables.Item(MtxReceivedPaymentLog);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
            oDT = form.DataSources.DataTables.Add(MtxReceivedPaymentLog);

        }

        return oDT;
    }
    private void BindMatrixPendEmail(SAPbouiCOM.Matrix oMatrix1, string mtCol, string dtCol, string dt = null, bool editable = false)
    {
        try
        {
            SAPbouiCOM.Column oColumn = oMatrix1.Columns.Item(mtCol);
            oColumn.Editable = editable;
            if (dt == null)
            {
                oColumn.DataBind.Bind(MtxSendPendingEmail, dtCol);
            }
            else
                oColumn.DataBind.Bind(dt, dtCol);
        }
        catch (Exception ex)
        {
            errorLog("Can not bind " + mtCol + " to " + dtCol + ".  error: " + ex.Message);
        }

    }
    private void BindMatrixReceivedEmail(SAPbouiCOM.Matrix oMatrix1, string mtCol, string dtCol, string dt = null, bool editable = false)
    {
        try
        {
            SAPbouiCOM.Column oColumn = oMatrix1.Columns.Item(mtCol);
            oColumn.Editable = editable;
            if (dt == null)
            {
                oColumn.DataBind.Bind(MtxReceivedPaymentLog, dtCol);
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

    private void PostPendEmail(String TemplateID, int i, SAPbouiCOM.Matrix oMatx, SAPbouiCOM.Form Oform)
    {
        String EbizKey = "";
        bool PostStatus = false;
        try
        {
            var Menu = GetComboValue(Oform, cmbMenuString);
            PaymentResponse oPaymentResponse;

            if (Menu.Equals("Delete"))
            {
                oPaymentResponse = ebiz.DeleteEbizWebFormPayment(getToken(), TemplateID);
                if (oPaymentResponse.Status.Equals("Success"))
                {
                    PostStatus = true;
                }
                else
                {
                    PostStatus = false;
                }
            }
            else if(Menu.Equals("Send Email Reminder"))
            {//Delete
                var ReminderResp = ebiz.ResendEbizWebFormEmail(getToken(), TemplateID);
                if (ReminderResp.Equals("1"))
                {
                    PostStatus = true;
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
                ShowSystemLog(String.Format("Done: Pending Email Row:{0} Selected Menu:{1}",i+1,Menu));
            }
            else
            {
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
                oEdit.Value = "False";
                ShowSystemLog(String.Format("Issue: Pending Email Row:{0} Selected Menu:{1}", i+1, Menu));
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }

}
