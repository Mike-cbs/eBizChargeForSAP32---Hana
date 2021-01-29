using eBizChargeForSAP;
using eBizChargeForSAP.ServiceReference1;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

partial class SAP
{
    #region Data Member
    SAPbouiCOM.Form theEmailPaytmentSO;
    string MtxCreatePaymentRequest = "mtCPR1";//mtxlog
    string MtxSOLog = "mtxlog";//
    string MtxReceivedPaymentLog = "mtRP";//
    SAPbouiCOM.Matrix oMatrixCreatePayReq, oMatrixCreatePayReqSOLog, oMtxRPLog;
    SAPbouiCOM.DataTable odtEmailPayTab1, odtEmailPaySOLOGTab1, odataRecPayLog;
    String SelectAlltab1 = "SelAllTab1";
    SAPbouiCOM.CheckBox Cmb_EmailPaymentSelectAllTab1;
    string MtxPendingPaymentRequest = "mtCPR2";
    SAPbouiCOM.Matrix oMatrixPendingPayReq;
    SAPbouiCOM.DataTable odtPendingPayTab2;
    String SelectAlltab2 = "SelAllTab2";
    SAPbouiCOM.CheckBox Cmb_EmailPaymentSelectAllTab2;
    String RadioDeposit = "radDeposit", RadioPreAuth = "RedPreAuth";
    SAPbouiCOM.OptionBtn radDeposit, radPreAuth;
    SAPbouiCOM.EditText ed_fromDate2, ed_ToDate2;
    String txtFromdt2 = "txtFromdt2", txtTodt2 = "txtTodt2";
    string MtxReceivedPayment = "mtCPR3";
    SAPbouiCOM.Matrix oMatrixReceivedPayment;
    SAPbouiCOM.DataTable odtReceivedPayment;
    String SelectAlltab3 = "SelAlTab3";
    SAPbouiCOM.CheckBox Cmb_EmailPaymentSelectAllTab3;
    ///
    //mtxSO1
    string dtCreatePaymentRequestTAB1 = "dtemp";
    SAPbouiCOM.DataTable oDTCreatePaymentStatus;
    bool FlgCreatePayReqDocInitiallization = false;

    bool flStateCreatePay = true;

    #endregion

    #region Form Event
    private void CreateMenuCreatePaymentForSO()
    {
        try
        {
            FlgCreatePayReqDocInitiallization = false;
            AddXMLForm("Presentation_Layer.SendEmailPaySO.xml");
            theEmailPaytmentSO.Freeze(true);
            InitForm();
            theEmailPaytmentSO.Freeze(false);
            flStateCreatePay = true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void SyncEmailPaymentSOFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            theEmailPaytmentSO = form;
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

                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        {
                            if (pVal.ItemUID.Equals("cmbEmail"))
                            {
                                SetEmailSendNewOnChange();
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        //Tab1
                        if (pVal.ItemUID == "btnFind" || pVal.ItemUID == "btnRef")
                        {
                            if (ValidFindEmailPaymentSO())
                            {
                                ShowDataCreatePaymentRequest();
                            }
                        }
                        if (pVal.ItemUID == "btnPost")
                        {
                            SendEmail_CreatePaymentRequest(form);
                        }
                        if (pVal.ItemUID.Equals(SelectAlltab1))
                        {
                            SelectAllcheckbox(theEmailPaytmentSO, MtxCreatePaymentRequest, odtEmailPayTab1, SelectAlltab1);
                        }
                        //Tab2
                        if (pVal.ItemUID == "btnShowAll")
                        {
                            if (ValidPendingEmailSearch(theEmailPaytmentSO))
                            {
                                ShowDataSendPendingEmail1();
                            }
                        }
                        if (pVal.ItemUID.Equals(SelectAlltab2))
                        {
                            SelectAllcheckbox(theEmailPaytmentSO, MtxPendingPaymentRequest, odtPendingPayTab2, SelectAlltab2);
                        }
                        if (pVal.ItemUID == "btnRemove2")
                        {
                            SendPendingFunctinality(form, "Delete");
                        }
                        if (pVal.ItemUID == "btnResnd2")
                        {
                            SendPendingFunctinality(form, "Send Email Reminder");
                        }
                        ///TAB3

                        if (pVal.ItemUID == "btnShowA2")
                        {
                            if (ValidDateRecPay1(form))
                            {
                                ShowDataRecPay1();
                            }
                        }
                        if (pVal.ItemUID.Equals("Item_4") || pVal.ItemUID.Equals("Item_18"))//BP List Tab
                        {
                            if (flgReceivedfolder)
                            {
                                flgReceivedfolder = false;
                                SetFolder(form, "Item_4");
                                SetFolder(form, "Item_18");
                            }
                        }
                        if (pVal.ItemUID.Equals(SelectAlltab3))
                        {
                            SelectAllcheckbox2(theEmailPaytmentSO, MtxReceivedPayment, odtReceivedPayment, SelectAlltab3);
                        }
                        if (pVal.ItemUID.Equals("btnRem"))
                        {
                            ReceivedPayment(theEmailPaytmentSO, "Delete");
                        }
                        if (pVal.ItemUID.Equals("btnexprt"))
                        {
                            FileBrowser("ReceivedPayment");
                        }
                        if (pVal.ItemUID.Equals("btnsapb1"))
                        {
                            SetAllIncomingPaymentEmailPAymentForSO();
                        }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                        try
                        {
                            if (pVal.ItemUID == MtxCreatePaymentRequest & pVal.ColUID == "DocNum")
                            {//SO
                                SAPbouiCOM.EditText DocNum = (SAPbouiCOM.EditText)oMatrixCreatePayReq.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific;
                                DoOpenLinkedObjectForm("139", "2050", "8", DocNum.Value);
                            }
                            if (pVal.ItemUID == MtxPendingPaymentRequest & pVal.ColUID == "DocNum")
                            {//SO
                                SAPbouiCOM.EditText DocNum = (SAPbouiCOM.EditText)oMatrixPendingPayReq.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific;
                                DoOpenLinkedObjectForm("139", "2050", "8", DocNum.Value);
                            }
                            if (pVal.ItemUID == MtxReceivedPayment & pVal.ColUID == "DocNum")
                            {//SO
                                SAPbouiCOM.EditText DocNum = (SAPbouiCOM.EditText)oMatrixReceivedPayment.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific;
                                DoOpenLinkedObjectForm("139", "2050", "8", DocNum.Value);
                            }
                            if (pVal.ItemUID == MtxCreatePaymentRequest & pVal.ColUID == "clEbizkey")
                            {//SO
                                SAPbouiCOM.EditText DocNum = (SAPbouiCOM.EditText)oMatrixCreatePayReq.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific;
                                String WebUrl = getEbizURL(DocNum.Value);// EbizUrl.Value;
                                System.Diagnostics.Process.Start(WebUrl);
                            }
                        }
                        catch (Exception ex)
                        {
                            var error = ex.Message;

                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        // handleItemPress(form, pVal);
                        if (pVal.ItemUID.Equals("radDeposit") || pVal.ItemUID.Equals("RedPreAuth"))
                        {
                            if (ValidFindEmailPaymentSO())
                            {
                                ShowDataCreatePaymentRequest();
                            }
                        }
                        if (pVal.ItemUID.Equals("logREP"))
                        {
                            //  ShowEmailPayForSoLogTab1Matrix();
                            ShowReceivedPayLog(theEmailPaytmentSO);
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
    #endregion

    #region Initialize Form
    private void InitForm()
    {
        try
        {
            // tProcess = new Thread(CreateEmailPaymentShowData);
            //tProcess.Start();
            CreateEmailPaymentShowData();
            flgReceivedfolder = true;
            ColorItem("");
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void ColorItem(String ItemID)
    {
        try
        {
            //    SAPbouiCOM.Button ResendButton = ((SAPbouiCOM.Button)(theEmailPaytmentSO.Items.Item("btnResnd2").Specific));
            //    int blueBackColor = Color.Blue.R | (Color.Blue.G << 8) | (Color.Blue.B << 16);
            //   ResendButton.Item.ForeColor = 0;// System.Drawing.Color.Green.ToArgb();
            //ResendButton.Item.BackColor = System.Drawing.Color.Blue.ToArgb();

            theEmailPaytmentSO.Items.Item("2").ForeColor = System.Drawing.Color.Blue.ToArgb();

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void CreateEmailPaymentShowData()
    {
        try
        {
            FillEmailComboEmailPaySO();
            initCheckBox();
            initDateField();
            SetFolderEmailPay("Item_2");
            radDeposit.Selected = true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void FillEmailComboEmailPaySO()
    {
        if (formSelTarget == null)
        {
            if (theEmailPaytmentSO != null)
                formSelTarget = theEmailPaytmentSO;
            else
                return;
        }
        if (formSelTarget.TypeEx != formEmailPaymentSO)
            return;

        //    if (!FlgCreatePayReqDocInitiallization) return;

        try
        {// Email template Code
            SAPbouiCOM.ComboBox cmb_EmailTemplate = (SAPbouiCOM.ComboBox)theEmailPaytmentSO.Items.Item(ComBoEmail).Specific;
            var eTempList = ebiz.GetEmailTemplates(getToken(), "", "");
            var count = eTempList.Length;
            // cmb_EmailTemplate.ValidValues.Add("-1", "Choose a Template");
            foreach (var itm in eTempList)
            {
                ListEmailTemplateCust.Add(new EmailTemplateCust(itm.TemplateName, itm.TemplateInternalId, itm.FromEmail, itm.TemplateSubject, itm.TemplateDescription));
                cmb_EmailTemplate.ValidValues.Add(Convert.ToString(itm.TemplateName), Convert.ToString(itm.TemplateName));
            }
            cmb_EmailTemplate.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            if (count > 0)
            {
                if (!eTempList.ElementAt(0).FromEmail.Equals(""))
                {
                    SetItemValue(theEmailPaytmentSO, txtFrom, eTempList.ElementAt(0).FromEmail);
                }
                else
                {
                    //support@ebizcharge.com
                    SetItemValue(theEmailPaytmentSO, txtFrom, "support@ebizcharge.com");
                }
                SetItemValue(theEmailPaytmentSO, txtSubject, eTempList.ElementAt(0).TemplateSubject);
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        try
        {
            //PaymentMethod
            SAPbouiCOM.ComboBox cmb_PaymentMethod = (SAPbouiCOM.ComboBox)theEmailPaytmentSO.Items.Item(ComBoPaymentMethod).Specific;
            var PreAuthFlag = getPreAuthStatus();

            if (PreAuthFlag.Equals("Y"))
            {
                cmb_PaymentMethod.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            else if (PreAuthFlag.Equals("N"))
            {
                cmb_PaymentMethod.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void initCheckBox()
    {
        try
        {
            Cmb_EmailPaymentSelectAllTab1 = ((SAPbouiCOM.CheckBox)(theEmailPaytmentSO.Items.Item(SelectAlltab1).Specific));
            theEmailPaytmentSO.DataSources.UserDataSources.Add(SelectAlltab1, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            Cmb_EmailPaymentSelectAllTab1.DataBind.SetBound(true, "", SelectAlltab1);
            theEmailPaytmentSO.DataSources.UserDataSources.Item(SelectAlltab1).Value = "N";

            Cmb_EmailPaymentSelectAllTab3 = ((SAPbouiCOM.CheckBox)(theEmailPaytmentSO.Items.Item(SelectAlltab3).Specific));
            theEmailPaytmentSO.DataSources.UserDataSources.Add(SelectAlltab3, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            Cmb_EmailPaymentSelectAllTab3.DataBind.SetBound(true, "", SelectAlltab3);
            theEmailPaytmentSO.DataSources.UserDataSources.Item(SelectAlltab3).Value = "N";


            //String RadioDeposit = "", RadioPreAuth = "";
            radDeposit = ((SAPbouiCOM.OptionBtn)(theEmailPaytmentSO.Items.Item(RadioDeposit).Specific));
            theEmailPaytmentSO.DataSources.UserDataSources.Add(RadioDeposit, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            radDeposit.DataBind.SetBound(true, "", RadioDeposit);
            theEmailPaytmentSO.DataSources.UserDataSources.Item(RadioDeposit).Value = "Y";

            radPreAuth = ((SAPbouiCOM.OptionBtn)(theEmailPaytmentSO.Items.Item(RadioPreAuth).Specific));
            theEmailPaytmentSO.DataSources.UserDataSources.Add(RadioPreAuth, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            radPreAuth.DataBind.SetBound(true, "", RadioPreAuth);
            theEmailPaytmentSO.DataSources.UserDataSources.Item(RadioPreAuth).Value = "Y";
            radDeposit.GroupWith(RadioPreAuth);



        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    private void initDateField()
    {
        try
        {
            ed_fromDate = ((SAPbouiCOM.EditText)(theEmailPaytmentSO.Items.Item(txtFromdt).Specific));
            theEmailPaytmentSO.DataSources.UserDataSources.Add(txtFromdt, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_fromDate.DataBind.SetBound(true, "", txtFromdt);
            ed_fromDate.Value = DatetimeTosapFormatDate(DateTime.Now);

            ed_ToDate = ((SAPbouiCOM.EditText)(theEmailPaytmentSO.Items.Item(txtTodt).Specific));
            theEmailPaytmentSO.DataSources.UserDataSources.Add(txtTodt, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_ToDate.DataBind.SetBound(true, "", txtTodt);
            ed_ToDate.Value = DatetimeTosapFormatDate(DateTime.Now);

            ed_SendNewfromDate = ((SAPbouiCOM.EditText)(theEmailPaytmentSO.Items.Item(txtSendNewFromdt).Specific));
            theEmailPaytmentSO.DataSources.UserDataSources.Add(txtSendNewFromdt, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_SendNewfromDate.DataBind.SetBound(true, "", txtSendNewFromdt);
            ed_SendNewfromDate.Value = DatetimeTosapFormatDate(DateTime.Now);

            ed_SendNewToDate = ((SAPbouiCOM.EditText)(theEmailPaytmentSO.Items.Item(txtSendNewTodt).Specific));
            theEmailPaytmentSO.DataSources.UserDataSources.Add(txtSendNewTodt, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_SendNewToDate.DataBind.SetBound(true, "", txtSendNewTodt);
            ed_SendNewToDate.Value = DatetimeTosapFormatDate(DateTime.Now);

            ed_fromDate2 = ((SAPbouiCOM.EditText)(theEmailPaytmentSO.Items.Item(txtFromdt2).Specific));
            theEmailPaytmentSO.DataSources.UserDataSources.Add(txtFromdt2, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_fromDate2.DataBind.SetBound(true, "", txtFromdt2);
            ed_fromDate2.Value = DatetimeTosapFormatDate(DateTime.Now);

            ed_ToDate2 = ((SAPbouiCOM.EditText)(theEmailPaytmentSO.Items.Item(txtTodt2).Specific));
            theEmailPaytmentSO.DataSources.UserDataSources.Add(txtTodt2, SAPbouiCOM.BoDataType.dt_DATE, 20);
            ed_ToDate2.DataBind.SetBound(true, "", txtTodt2);
            ed_ToDate2.Value = DatetimeTosapFormatDate(DateTime.Now);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    #endregion

    #region Helping Function 
    private String GetEmailComboSendNewValue()
    {
        var Output = String.Empty;
        try
        {
            SAPbouiCOM.ComboBox cmb = (SAPbouiCOM.ComboBox)theEmailPaytmentSO.Items.Item("cmbEmail").Specific;
            Output = cmb.Selected.Value;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return Output;

    }
    private void SetEmailSendNewOnChange()
    {
        try
        {
            var SelectcmbVAl = GetEmailComboSendNewValue();
            var respEmailTemp = (from a in ListEmailTemplateCust where a.TemplateName == SelectcmbVAl select a).FirstOrDefault();
            SetItemValue(theEmailPaytmentSO, txtFrom, respEmailTemp.FromEmail);
            SetItemValue(theEmailPaytmentSO, txtSubject, respEmailTemp.Subject);
            if (respEmailTemp.FromEmail.Equals(""))
            {
                SetItemValue(theEmailPaytmentSO, txtFrom, "support@ebizcharge.com");
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private bool ValidFindEmailPaymentSO()
    {
        return true;
        bool flag = false;
        try
        {
            var cardCode = GetBpCode(theEmailPaytmentSO, "txtBpCode");
            if (string.IsNullOrEmpty(cardCode))
            {
                //    MsgWarning("Customer can't be left empty.");
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
    private void SelectAllcheckbox(SAPbouiCOM.Form form, String MatxStr, SAPbouiCOM.DataTable oDataTable, string CheckboxID)
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
    private void SelectAllcheckbox2(SAPbouiCOM.Form form, String MatxStr, SAPbouiCOM.DataTable oDataTable, string CheckboxID)
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
    private void SetFolderEmailPay(String oFolderId)
    {
        try
        {
            oFolder = ((SAPbouiCOM.Folder)theEmailPaytmentSO.Items.Item(oFolderId).Specific);
            oFolder.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
            errorLog("EmailPayentForSO SetFolder :" + ex.Message);
        }
    }
    #endregion

    #region Create Payment Request
    #region Show matrix Data
    private void ShowDataCreatePaymentRequest()
    {
        try
        {
            tProcess = new Thread(ShowEmailPayForSoTab1Matrix);
            tProcess.Start();
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void ShowEmailPayForSoTab1Matrix()
    {
        string sql = "";
        try
        {
            if (theEmailPaytmentSO.Items.Item(MtxCreatePaymentRequest) == null)
                return;

            theEmailPaytmentSO.Items.Item(MtxCreatePaymentRequest).AffectsFormMode = false;
            oMatrixCreatePayReq = (SAPbouiCOM.Matrix)theEmailPaytmentSO.Items.Item(MtxCreatePaymentRequest).Specific;
            oMatrixCreatePayReq.Clear();
            odtEmailPayTab1 = getDataTableSendEmailSO(theEmailPaytmentSO);

            var CardCode = GetBpCode(theEmailPaytmentSO, "txtBpCode");
            if (CardCode.Equals(""))
            {
                CardCode = "%";
            }
            var SalesOrder = GetItemValue(theEmailPaytmentSO, "txtSO");
            if (SalesOrder.Equals(""))
            {
                SalesOrder = "%";
            }
            DateTime FromdtValue = DateTime.MinValue;
            DateTime TodtValue = DateTime.MinValue;
            var FromSQl = "";
            var ToSQl = "";
            if (ed_SendNewfromDate.Value != "")
            {
                FromdtValue = getdate(ed_SendNewfromDate);
                FromSQl = String.Format(" and \"DocDate\">=cast('{0}' as date) ", FromdtValue.ToString("yyyy/MM/dd"));
            }
            else
            {
                FromSQl = String.Format(" and \"DocDate\" like '%' ");
            }

            if (ed_SendNewToDate.Value != "")
            {
                TodtValue = getdate(ed_SendNewToDate);
                ToSQl = String.Format(" and \"DocDate\"<=cast('{0}' as date)", TodtValue.ToString("yyyy/MM/dd"));
            }
            else
            {
                ToSQl = String.Format(" and \"DocDate\" like '%' ");
            }


            var emailqtr = string.Format("select \"E_Mail\" from OCRD where \"CardCode\" = ORDR.\"CardCode\" ");
            var subQuery = "IFNULL((SELECT  sum(T0.\"DpmAmnt\") FROM ODPI T0  INNER JOIN DPI1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" inner join RDR1 on RDR1.\"DocEntry\" = T1.\"BaseEntry\" where RDR1.\"DocEntry\" =ORDR.\"DocEntry\" and T0.\"DocStatus\" = 'O' and T0.\"CANCELED\" = 'N'),0)";
            sql = String.Format("Select * from (select 'N' as [clSel],'    ' as [Status],({2}) as Email,ROW_NUMBER() OVER (ORDER BY [DocEntry]) AS [ID],[DocNum],[DocDate],[DocTotal],[CardCode],[CardName],IFNULL([{1}],'') as [EBiz],CASE WHEN IFNULL([{1}], '') = '' THEN 'False' ELSE 'True'END as [Sync],CASE WHEN IFNULL([{1}], '') = '' THEN '' ELSE [{1}]  END as [ViewForm],[DocTotal]-{3} as [AmtDue] from ORDR where  [DocStatus]='O' and [CardCode] like '{0}' and [DocNum] like '{4}'  and [DocTotal]>0 {5} {6})T where T.[AmtDue]>=0   ", CardCode, "U_" + U_EbizChargeURL, emailqtr, subQuery, SalesOrder, FromSQl, ToSQl).Replace("]", "\"").Replace("[", "\"");
            odtEmailPayTab1.ExecuteQuery(sql);

            BindMatrixSendEmailSO(oMatrixCreatePayReq, "clSel", "clSel", dtItem, true);
            BindMatrixSendEmailSO(oMatrixCreatePayReq, "clNo", "ID");
            BindMatrixSendEmailSO(oMatrixCreatePayReq, "DocNum", "DocNum");//Email
            BindMatrixSendEmailSO(oMatrixCreatePayReq, "Email", "Email", null, true);//
            BindMatrixSendEmailSO(oMatrixCreatePayReq, "DocDate", "DocDate");
            BindMatrixSendEmailSO(oMatrixCreatePayReq, "CardCode", "CardCode");
            if (radDeposit.Selected)
            {
                BindMatrixSendEmailSO(oMatrixCreatePayReq, "AmtDue", "AmtDue", null, true);
            }
            else
            {
                BindMatrixSendEmailSO(oMatrixCreatePayReq, "AmtDue", "AmtDue");
            }
            BindMatrixSendEmailSO(oMatrixCreatePayReq, "DocTotal", "DocTotal");
            BindMatrixSendEmailSO(oMatrixCreatePayReq, "clEbizkey", "ViewForm");
            //  BindMatrixSendEmailSO(oMatrixCreatePayReq, "clSync", "Sync");

            oMatrixCreatePayReq.LoadFromDataSource();
            oMatrixCreatePayReq.AutoResizeColumns();


        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }



    #endregion

    #region send Email
    private void SendEmail_CreatePaymentRequest(SAPbouiCOM.Form form)
    {
        bool FlgChange = false;
        try
        {
            if (form.Items.Item(MtxCreatePaymentRequest) == null)
                return;

            if (SBO_Application.MessageBox("Are you sure you want to send email(s)?", 1, "Yes", "No") == 1)
            {
                form.Items.Item(MtxCreatePaymentRequest).AffectsFormMode = false;
                oMatrixCreatePayReq = (SAPbouiCOM.Matrix)form.Items.Item(MtxCreatePaymentRequest).Specific;
                int RowCount = odtEmailPayTab1.Rows.Count;
                if (RowCount > 0)
                {
                    for (int i = odtEmailPayTab1.Rows.Count - 1; i >= 0; i--)
                    {
                        var EdDocNum = GetMatrixValue(oMatrixCreatePayReq, i, "DocNum").ToString();
                        var EdAmtEntered = ConvertDecimal(GetMatrixValue(oMatrixCreatePayReq, i, "AmtDue").ToString());
                        var chk = (SAPbouiCOM.CheckBox)oMatrixCreatePayReq.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                        if (chk.Checked)
                        {
                            FlgChange = true;
                            SalesOrder oSalesOrder = GetSO(EdDocNum);
                            oSalesOrder.AmountDue = EdAmtEntered;

                            POSTSendEmail(oSalesOrder, i, oMatrixCreatePayReq, form);
                        }
                        else
                        {

                        }

                    }
                    if (FlgChange)
                    {
                        //   oMatrixBpExport.LoadFromDataSource();
                        oMatrixCreatePayReq.FlushToDataSource();
                    }
                }
            }
        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }
    private void POSTSendEmail(SalesOrder oSalesOrder, int i, SAPbouiCOM.Matrix oMatx, SAPbouiCOM.Form oForm)
    {
        String EbizKey = "";
        bool Flag = false;
        try
        {
            if (CheckBP(oSalesOrder.CustomerId))
            {
                if (CheckSO(oSalesOrder))
                {
                    var ToEmail = GetEmail(i, oMatx);
                    GetEbizWebFormURLCustom(ToEmail, oSalesOrder, oForm, ref Flag, ref EbizKey);
                    if (Flag)
                    {
                        ShowSystemLog(String.Format("Email Sent Against Sales Order: {0}", oSalesOrder.SalesOrderNumber));
                    }
                }
            }
            SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clEbizkey").Cells.Item(i + 1).Specific;
            oEdit.Value = EbizKey;
            //oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
            //oEdit.Value = Flag.ToString();
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }
    private bool GetEbizWebFormURLCustom(string ToEmail, SalesOrder oSalesOrder, SAPbouiCOM.Form oForm, ref bool FlagStatus, ref string URL)
    {
        bool Flag = false;
        try
        {
            var PaymentMethod_selected = GetPaymentMethodComboValue(oForm);
            EbizWebForm pf = new EbizWebForm();
            pf.Clerk = oCompany.UserName;
            pf.Date = Convert.ToDateTime(oSalesOrder.Date);
            pf.DueDate = Convert.ToDateTime(oSalesOrder.DueDate);
            pf.TotalAmount = oSalesOrder.Amount;
            pf.AmountDue = oSalesOrder.AmountDue;
            pf.CustFullName = oSalesOrder.CustomerId;
            pf.ProcessingCommand = PaymentMethod_selected;

            pf.EmailAddress = ToEmail;
            pf.SoftwareId = Software;
            string TemplateID = GetEmailTempID(oForm);
            var Descp = ListEmailTemplateCust.Where(a => a.TemplateID == TemplateID).FirstOrDefault().templateDescriptionField;
            pf.EmailNotes = Descp;
            pf.EmailSubject = GetItemValue(txtSubject, oForm);
            pf.FromEmail = GetItemValue(txtFrom, oForm);
            pf.CustomerId = oSalesOrder.CustomerId;
            pf.SoNum = oSalesOrder.SalesOrderNumber;
            pf.InvoiceNumber = oSalesOrder.SalesOrderNumber;

            //   pf.ProcessingCommand = "AuthOnly";
            pf.InvoiceInternalId = getSOInternalID(oSalesOrder.SalesOrderNumber);
            pf.ShowViewInvoiceLink = true;
            pf.ShowViewSalesOrderLink = true;
            pf.Description = oSalesOrder.Description;
            pf.EmailTemplateName = TemplateID;
            pf.SendEmailToCustomer = true;

            string url = ebiz.GetEbizWebFormURL(getToken(), pf);
            URL = url;
            Flag = true;
            UpdateEbizChargeURL("ORDR", URL, oSalesOrder.SalesOrderNumber);
            updateSOPayment(oSalesOrder.SalesOrderNumber, PaymentMethod_selected);
            //SetAllIncomingPayment();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            SBO_Application.MessageBox("Issue:" + ex.Message);
        }
        FlagStatus = Flag;
        return Flag;
    }
    private String GetPaymentMethodComboValue(SAPbouiCOM.Form oForm)
    {
        var Output = String.Empty;
        try
        {
            var result1 = radDeposit.Selected;
            var result2 = radPreAuth.Selected;

            if (result1)
            {
                Output = "sales";
            }
            else
            {
                Output = "AuthOnly";
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return Output;

    }
    #endregion

    #endregion

    #region Pending Payment
    private void ShowDataSendPendingEmail1()
    {
        try
        {
            tProcess = new Thread(PendingEmailThreadProc1);
            tProcess.Start();
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    public void PendingEmailThreadProc1()
    {
        try
        {
            if (formSelTarget == null)
            {
                if (theEmailPaytmentSO != null)
                    formSelTarget = theEmailPaytmentSO;
                else
                    return;
            }
            //if (formSelTarget.TypeEx != formEmailPaymentPending)
            //    return;

            //if (!SendPendingEmailflgDocInitiallization) return;

            ShowSendPendEmail1(theEmailPaytmentSO, MtxPendingPaymentRequest);

        }
        catch (Exception ex)
        {
        }

    }
    private void ShowSendPendEmail1(SAPbouiCOM.Form form, string MtrxStr)
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtrxStr) == null)
                return;

            DateTime FromdtValue = getdate(ed_fromDate);
            DateTime TodtValue = getdate(ed_ToDate).Add(DateTime.MaxValue.TimeOfDay);
            var filters = new SearchFilter[1];
            filters[0] = new SearchFilter();

            filters[0].FieldName = "PaymentSourceId";
            filters[0].ComparisonOperator = "eq";
            filters[0].FieldValue = "SAPB1App";



            var PendingPaymentsList = ebiz.SearchEbizWebFormPendingPayments(getToken(), null, FromdtValue, TodtValue, filters, 0, 100, "");



            //PaymentSourceId             SAPB1App
            //  var PaymentResp = ebiz.DeleteEbizWebFormPayment(getToken(), "");
            // var ReminderResp = ebiz.ResendEbizWebFormEmail(getToken(), "");


            form.Items.Item(MtrxStr).AffectsFormMode = false;
            oMatrixPendingPayReq = (SAPbouiCOM.Matrix)form.Items.Item(MtrxStr).Specific;
            oMatrixPendingPayReq.Clear();
            odtPendingPayTab2 = getDataTablePendEmail(form);

            sql = String.Format("select  top 1'N' as [clSel],ROW_NUMBER() OVER (ORDER BY [DocEntry]) AS [ID],[DocNum],[DocDate],[DocTotal],[CardCode],[CardName],IFNULL([U_EbizChargeID],'') as [EBiz],CASE WHEN IFNULL([U_EbizChargeID], '') = '' THEN 'False' ELSE 'True'END as [Sync] from ORDR where  [DocStatus]='100' ").Replace("]", "\"").Replace("[", "\"");
            odtPendingPayTab2.ExecuteQuery(sql);
            BindMatrixPendEmail(oMatrixPendingPayReq, "clSel", "clSel", dtItem, true);
            BindMatrixPendEmail(oMatrixPendingPayReq, "clNo", "ID");
            BindMatrixPendEmail(oMatrixPendingPayReq, "DocNum", "DocNum");
            BindMatrixPendEmail(oMatrixPendingPayReq, "DocDate", "DocDate");//DocSent
            BindMatrixPendEmail(oMatrixPendingPayReq, "DocSent", "DocDate");//DueDate
            BindMatrixPendEmail(oMatrixPendingPayReq, "DueDate", "DocDate");//DueDate
            BindMatrixPendEmail(oMatrixPendingPayReq, "CardCode", "CardCode");
            BindMatrixPendEmail(oMatrixPendingPayReq, "Email", "CardName");
            BindMatrixPendEmail(oMatrixPendingPayReq, "DocTotal", "DocTotal");//AmtDue
            BindMatrixPendEmail(oMatrixPendingPayReq, "AmtDue", "DocTotal");//AmtDue
            BindMatrixPendEmail(oMatrixPendingPayReq, "clebizUrl", "EBiz");
            BindMatrixPendEmail(oMatrixPendingPayReq, "Type", "EBiz");
            BindMatrixPendEmail(oMatrixPendingPayReq, "clSync", "Sync");

            oMatrixPendingPayReq.LoadFromDataSource();
            oMatrixPendingPayReq.Clear();
            foreach (var itm in PendingPaymentsList)
            {
                if (itm.PaymentSourceId.Contains("SAPB1App"))
                {
                    oMatrixPendingPayReq.AddRow(1);
                    SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatrixPendingPayReq.Columns.Item("CardCode").Cells.Item(oMatrixPendingPayReq.RowCount).Specific;
                    oEdit.Value = itm.CustomerId;
                    oEdit = (SAPbouiCOM.EditText)oMatrixPendingPayReq.Columns.Item("DocTotal").Cells.Item(oMatrixPendingPayReq.RowCount).Specific;
                    oEdit.Value = itm.InvoiceAmount;
                    oEdit = (SAPbouiCOM.EditText)oMatrixPendingPayReq.Columns.Item("AmtDue").Cells.Item(oMatrixPendingPayReq.RowCount).Specific;
                    oEdit.Value = itm.AmountDue;
                    //ConvertString(ConvertDecimal(itm.AmountDue) - getSOAmountDue(itm.InvoiceNumber));
                    oEdit = (SAPbouiCOM.EditText)oMatrixPendingPayReq.Columns.Item("DocNum").Cells.Item(oMatrixPendingPayReq.RowCount).Specific;
                    oEdit.Value = itm.InvoiceNumber;
                    oEdit = (SAPbouiCOM.EditText)oMatrixPendingPayReq.Columns.Item("clebizUrl").Cells.Item(oMatrixPendingPayReq.RowCount).Specific;
                    oEdit.Value = String.Format("https://webforms.ebizcharge.net/EBizSecureForm.aspx?pid={0}", itm.PaymentInternalId);
                    oEdit = (SAPbouiCOM.EditText)oMatrixPendingPayReq.Columns.Item("Email").Cells.Item(oMatrixPendingPayReq.RowCount).Specific;
                    oEdit.Value = itm.CustomerEmailAddress;
                    oEdit = (SAPbouiCOM.EditText)oMatrixPendingPayReq.Columns.Item("Type").Cells.Item(oMatrixPendingPayReq.RowCount).Specific;
                    oEdit.Value = getSOPaymentTypeConnect(itm.InvoiceNumber);
                    oEdit = (SAPbouiCOM.EditText)oMatrixPendingPayReq.Columns.Item("DocSent").Cells.Item(oMatrixPendingPayReq.RowCount).Specific;
                    var date = DatetimeTosapFormatDate(Convert.ToDateTime(itm.PaymentRequestDateTime));
                    oEdit.Value = date;
                    oEdit = (SAPbouiCOM.EditText)oMatrixPendingPayReq.Columns.Item("DueDate").Cells.Item(oMatrixPendingPayReq.RowCount).Specific;
                    date = DatetimeTosapFormatDate(Convert.ToDateTime(itm.PaymentRequestDateTime));
                    oEdit.Value = date;
                }
            }
            oMatrixPendingPayReq.AutoResizeColumns();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void SendPendingFunctinality(SAPbouiCOM.Form form, String ButtonFunction)
    {
        bool FlgChange = false;
        try
        {
            if (form.Items.Item(MtxPendingPaymentRequest) == null)
                return;

            form.Items.Item(MtxPendingPaymentRequest).AffectsFormMode = false;
            oMatrixPendingPayReq = (SAPbouiCOM.Matrix)form.Items.Item(MtxPendingPaymentRequest).Specific;
            int RowCount = oMatrixPendingPayReq.RowCount;
            if (RowCount > 0)
            {
                for (int i = RowCount - 1; i >= 0; i--)
                {

                    var PaymentURL = GetMatrixValue(oMatrixPendingPayReq, i, "clebizUrl").ToString();
                    var EdTemplate = GetPaymentIDURL(PaymentURL);


                    var chk = (SAPbouiCOM.CheckBox)oMatrixPendingPayReq.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    if (chk.Checked)
                    {
                        FlgChange = true;
                        PostPendEmail1(EdTemplate, i, oMatrixPendingPayReq, form, ButtonFunction);
                    }
                    else
                    {

                    }

                }
                if (FlgChange)
                {
                    //   oMatrixBpExport.LoadFromDataSource();
                    //  oMatrixPendingPayReq.FlushToDataSource();
                    ShowSendPendEmail1(theEmailPaytmentSO, MtxPendingPaymentRequest);
                }
            }
        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }
    private void PostPendEmail1(String TemplateID, int i, SAPbouiCOM.Matrix oMatx, SAPbouiCOM.Form Oform, String ButtonFunction)
    {
        String EbizKey = "";
        bool PostStatus = false;
        try
        {
            //   var Menu = GetComboValue(Oform, cmbMenuString);
            PaymentResponse oPaymentResponse;

            if (ButtonFunction.Equals("Delete"))
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
            else if (ButtonFunction.Equals("Send Email Reminder"))
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
                //  SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
                // oEdit.Value = "True";
                ShowSystemLog(String.Format("Done: Pending Email Row:{0} Selected Menu:{1}", i + 1, ButtonFunction));
            }
            else
            {
                //  SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
                // oEdit.Value = "False";
                ShowSystemLog(String.Format("Issue: Pending Email Row:{0} Selected Menu:{1}", i + 1, ButtonFunction));
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }
    #endregion

    #region ReceivedPayment
    private bool ValidDateRecPay1(SAPbouiCOM.Form oForm)
    {
        bool flag = false;
        try
        {
            var FromValue = GetItemValue(txtFromdt2, oForm);
            var ToDateValue = GetItemValue(txtTodt2, oForm);

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
    private void ShowDataRecPay1()
    {
        try
        {
            tProcess = new Thread(RecPayThreadProc1);
            tProcess.Start();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            throw;
        }
    }

    public void RecPayThreadProc1()
    {
        try
        {
            ShowRecPay1(theEmailPaytmentSO);
        }
        catch (Exception ex)
        {
        }

    }

    private void ShowRecPay1(SAPbouiCOM.Form form)
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxReceivedPayment) == null)
                return;

            DateTime FromdtValue = getdate(ed_fromDate2);
            DateTime TodtValue = getdate(ed_ToDate2).Add(DateTime.MaxValue.TimeOfDay);
            var filters = new SearchFilter[1];
            filters[0] = new SearchFilter();

            filters[0].FieldName = "PaymentSourceId";
            filters[0].ComparisonOperator = "eq";
            filters[0].FieldValue = "SAPB1App       ";

            var ReceivedPaymentsList = ebiz.SearchEbizWebFormReceivedPayments(getToken(), "", FromdtValue, TodtValue, filters, 0, 100, "");



            form.Items.Item(MtxReceivedPayment).AffectsFormMode = false;
            oMatrixReceivedPayment = (SAPbouiCOM.Matrix)form.Items.Item(MtxReceivedPayment).Specific;
            oMatrixReceivedPayment.Clear();
            odtReceivedPayment = GetDatatableRecPay(form);

            sql = String.Format("select  top 1'N' as [clSel],ROW_NUMBER() OVER (ORDER BY [DocEntry]) AS [ID],[DocNum],'   ' as [Method],[DocDueDate] as [DueDate],[DocDate],[DocTotal],[CardCode],[CardName],IFNULL([U_EbizChargeID],'') as [EBiz],CASE WHEN IFNULL([U_EbizChargeID], '') = '' THEN 'False' ELSE 'True'END as [Sync] from ORDR where  [DocStatus]='100' ").Replace("]", "\"").Replace("[", "\"");
            odtReceivedPayment.ExecuteQuery(sql);
            BindMtxRecPay(oMatrixReceivedPayment, "clSel", "clSel", dtItem, true);
            BindMtxRecPay(oMatrixReceivedPayment, "clNo", "ID");
            BindMtxRecPay(oMatrixReceivedPayment, "DocNum", "DocNum");
            BindMtxRecPay(oMatrixReceivedPayment, "DocDate", "DocDate");
            BindMtxRecPay(oMatrixReceivedPayment, "DueDate", "DueDate");
            BindMtxRecPay(oMatrixReceivedPayment, "CardCode", "CardCode");
            BindMtxRecPay(oMatrixReceivedPayment, "Email", "CardName");
            BindMtxRecPay(oMatrixReceivedPayment, "DocTotal", "DocTotal");
            BindMtxRecPay(oMatrixReceivedPayment, "AmtPaid", "DocTotal");
            BindMtxRecPay(oMatrixReceivedPayment, "AmtDue", "DocTotal");
            BindMtxRecPay(oMatrixReceivedPayment, "clEbizkey", "EBiz");
            BindMtxRecPay(oMatrixReceivedPayment, "Type", "EBiz");
            BindMtxRecPay(oMatrixReceivedPayment, "clSync", "Sync");
            BindMtxRecPay(oMatrixReceivedPayment, "Method", "CardName");
            BindMtxRecPay(oMatrixReceivedPayment, "Status", "CardName");

            oMatrixReceivedPayment.LoadFromDataSource();
            oMatrixReceivedPayment.Clear();
            foreach (var itm in ReceivedPaymentsList)
            {
                if (itm.PaymentSourceId.Contains("SAPB1App"))
                {
                    //  TransactionObject trasactionObj = GetPaymentStatus(itm.PaymentInternalId);
                    //                    if (trasactionObj != null)
                    {
                        oMatrixReceivedPayment.AddRow(1);

                        SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatrixReceivedPayment.Columns.Item("CardCode").Cells.Item(oMatrixReceivedPayment.RowCount).Specific;
                        oEdit.Value = itm.CustomerId;
                        oEdit = (SAPbouiCOM.EditText)oMatrixReceivedPayment.Columns.Item("clEbizkey").Cells.Item(oMatrixReceivedPayment.RowCount).Specific;
                        oEdit.Value = itm.PaymentInternalId;
                        oEdit = (SAPbouiCOM.EditText)oMatrixReceivedPayment.Columns.Item("DocNum").Cells.Item(oMatrixReceivedPayment.RowCount).Specific;
                        oEdit.Value = itm.InvoiceNumber;
                        if (itm.InvoiceNumber.Equals("87") || itm.InvoiceNumber.Equals("90"))
                        {
                            string err0r = "";
                        }
                        oEdit = (SAPbouiCOM.EditText)oMatrixReceivedPayment.Columns.Item("Email").Cells.Item(oMatrixReceivedPayment.RowCount).Specific;
                        oEdit.Value = itm.CustomerEmailAddress;
                        //var AmountPaid = getSOAmountDue(itm.InvoiceNumber);
                        var PaymentType = getSOPaymentTypeConnect(itm.InvoiceNumber);
                        oEdit = (SAPbouiCOM.EditText)oMatrixReceivedPayment.Columns.Item("Type").Cells.Item(oMatrixReceivedPayment.RowCount).Specific;
                        oEdit.Value = PaymentType;
                        //oEdit = (SAPbouiCOM.EditText)oMatrixReceivedPayment.Columns.Item("AmtDue").Cells.Item(oMatrixReceivedPayment.RowCount).Specific;
                        //oEdit.Value = (Convert.ToDecimal(itm.InvoiceAmount)- AmountPaid).ToString();
                        oEdit = (SAPbouiCOM.EditText)oMatrixReceivedPayment.Columns.Item("AmtPaid").Cells.Item(oMatrixReceivedPayment.RowCount).Specific;
                        oEdit.Value = itm.PaidAmount;
                        String Last4digStr = "";
                        try
                        {
                            Last4digStr = itm.PaymentMethod + " ending in " + itm.Last4;
                            //Last4digStr = itm.Last4;
                            oEdit = (SAPbouiCOM.EditText)oMatrixReceivedPayment.Columns.Item("Method").Cells.Item(oMatrixReceivedPayment.RowCount).Specific;
                            oEdit.Value = Last4digStr.ToString();

                        }
                        catch
                        {
                        }


                        oEdit = (SAPbouiCOM.EditText)oMatrixReceivedPayment.Columns.Item("DocDate").Cells.Item(oMatrixReceivedPayment.RowCount).Specific;
                        var date = DatetimeTosapFormatDate(Convert.ToDateTime(itm.PaymentRequestDateTime));
                        oEdit.Value = date;
                        SalesOrder oSalesOrder = GetSO(itm.InvoiceNumber);
                        oEdit = (SAPbouiCOM.EditText)oMatrixReceivedPayment.Columns.Item("DueDate").Cells.Item(oMatrixReceivedPayment.RowCount).Specific;
                        date = DatetimeTosapFormatDate(Convert.ToDateTime(oSalesOrder.DueDate));
                        oEdit.Value = date;
                    }
                }
            }
            oMatrixReceivedPayment.AutoResizeColumns();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }

    private void ReceivedPayment(SAPbouiCOM.Form form, String Action)
    {
        bool FlgChange = false;
        try
        {
            if (form.Items.Item(MtxReceivedPayment) == null)
                return;

            form.Items.Item(MtxReceivedPayment).AffectsFormMode = false;
            oMatrixReceivedPayment = (SAPbouiCOM.Matrix)form.Items.Item(MtxReceivedPayment).Specific;
            int RowCount = oMatrixReceivedPayment.RowCount;
            if (RowCount > 0)
            {
                for (int i = RowCount - 1; i >= 0; i--)
                {
                    var EdTemplate = GetMatrixValue(oMatrixReceivedPayment, i, "clEbizkey").ToString();//DocNum
                    var EdDocNum = GetMatrixValue(oMatrixReceivedPayment, i, "DocNum").ToString();//

                    var chk = (SAPbouiCOM.CheckBox)oMatrixReceivedPayment.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    if (chk.Checked)
                    {
                        FlgChange = true;
                        PostRecPay1(EdTemplate, i, oMatrixReceivedPayment, form, EdDocNum, Action);
                    }
                    else
                    {

                    }

                }
                if (FlgChange)
                {
                    oMatrixReceivedPayment.FlushToDataSource();
                }
            }
        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }

    private void PostRecPay1(String TemplateID, int i, SAPbouiCOM.Matrix oMatx, SAPbouiCOM.Form Oform, String DocNum, String Action)
    {
        String EbizKey = "";
        bool PostStatus = false;
        try
        {
            //   var Menu = GetComboValue(Oform, cmbMenuStrPyR);
            PaymentResponse oPaymentResponse;

            if (Action.Equals("Mark Payment"))
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

            }
            else if (Action.Equals("Delete"))
            {
                var ReminderResp = ebiz.DeleteEbizWebFormPayment(getToken(), TemplateID);
                if (ReminderResp.Status.Equals("Success"))
                {
                    PostStatus = true;
                }
                else
                {
                    PostStatus = false;
                }

            }
            else if (Action.Equals("Apply"))
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
                //SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
                //oEdit.Value = "True";
                ShowSystemLog(String.Format("Payment Received Action: {1} against Sales Order: {0}  Status:Done", DocNum, Action));
            }
            else
            {
                //SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
                //oEdit.Value = "Fasle";
                ShowSystemLog(String.Format("Payment Received Action: {1} against Sales Order: {0} Status:Failed", DocNum, Action));
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }

    private void ShowReceivedPayLog(SAPbouiCOM.Form form, String Where = "")
    {
        string sql = "";
        try
        {
            if (form.Items.Item("mtRP") == null)
                return;

            form.Items.Item("mtRP").AffectsFormMode = false;
            oMatrixSyncUploadSOList1 = (SAPbouiCOM.Matrix)form.Items.Item("mtRP").Specific;
            oMatrixSyncUploadSOList1.Clear();
            oDTSyncUploadSOList1 = getDataTableSyncUpload(form, "mtRP");


            sql = String.Format("select \"U_DocNum\"as \"DocNum\",* from \"@CCSO\" ");
            oDTSyncUploadSOList1.ExecuteQuery(sql);

            //BindMatrixSyncUpload("mtRP", oMatrixSyncUploadSOList1, "clSel", "clSel", null, true);
            //BindMatrixSyncUpload("mtRP", oMatrixSyncUploadSOList1, "clNo", "ID");
            BindMatrixSyncUpload("mtRP", oMatrixSyncUploadSOList1, "DocNum", "DocNum");
            BindMatrixSyncUpload("mtRP", oMatrixSyncUploadSOList1, "Status", "U_Status");
            BindMatrixSyncUpload("mtRP", oMatrixSyncUploadSOList1, "Type", "U_Type");
            BindMatrixSyncUpload("mtRP", oMatrixSyncUploadSOList1, "DocPaid", "U_DocPaidDate");
            BindMatrixSyncUpload("mtRP", oMatrixSyncUploadSOList1, "Email", "U_Email");
            BindMatrixSyncUpload("mtRP", oMatrixSyncUploadSOList1, "CardCode", "U_CardCode");
            BindMatrixSyncUpload("mtRP", oMatrixSyncUploadSOList1, "AmountPaid", "U_AmountPaid");
            BindMatrixSyncUpload("mtRP", oMatrixSyncUploadSOList1, "Method", "U_MethodDigit");
            //BindMatrixSyncUpload(oMatrixCreatePayReqSOLog, "Status", "U_Status");
            //BindMatrixSendEmailSO(oMatrixCreatePayReqSOLog, "Type", "U_Type");//
            //BindMatrixSendEmailSO(oMatrixCreatePayReqSOLog, "DocPaidDate", "U_DocPaidDate");
            //BindMatrixSendEmailSO(oMatrixCreatePayReqSOLog, "Email", "U_Email");
            //BindMatrixSendEmailSO(oMatrixCreatePayReqSOLog, "CardCode", "U_CardCode");
            //BindMatrixSendEmailSO(oMatrixCreatePayReqSOLog, "AmountPaid", "U_AmountPaid");;

            oMatrixSyncUploadSOList1.LoadFromDataSource();
            oMatrixSyncUploadSOList1.AutoResizeColumns();

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }


    #endregion

    #region ApplyToSAP
    private void SetAllIncomingPaymentEmailPAymentForSO()
    {
        LogTable oLogTable;
        // SAPbouiCOM.EditText oEdit = null;
        try
        {
            if (oMatrixReceivedPayment == null) return;
            //    var flg = GetPaymentStatus(PaymentID);
            int RowCount = oMatrixReceivedPayment.RowCount;
            if (RowCount > 0)
            {

                for (int i = 0; i < RowCount; i++)
                {
                    var chk = (SAPbouiCOM.CheckBox)oMatrixReceivedPayment.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    if (chk.Checked)
                    {
                        //  oEdit = (SAPbouiCOM.EditText)oMatrixReceivedPayment.Columns.Item("Status").Cells.Item(i + 1).Specific;
                        var EdDocNum = GetMatrixValue(oMatrixReceivedPayment, i, "DocNum").ToString();
                        var EbixURL = GetMatrixValue(oMatrixReceivedPayment, i, "clEbizkey").ToString();
                        var DocPaiddate = GetMatrixValue(oMatrixReceivedPayment, i, "DocDate").ToString();
                        var Type = GetMatrixValue(oMatrixReceivedPayment, i, "Type").ToString();
                        if (EbixURL.Equals("View Form") || EbixURL.Count() > 0)
                        {
                            oLogTable = new LogTable();
                            oLogTable.DocNum = EdDocNum;
                            EbixURL = getEbizURL(EdDocNum);
                            if (EdDocNum.Equals("101"))
                            {
                                string t = "";
                            }
                            var PID = GetPaymentID(EbixURL);
                            string TransactionRefNum = "";
                            TransactionObject trasactionObj = GetPaymentStatus(PID, ref TransactionRefNum);
                            if (trasactionObj != null)
                            {
                                var PaymentMethod = getSOpaymentType(EdDocNum);
                                if (PaymentMethod.Equals("sales"))
                                {
                                    //  String IsDownPaymentAlreadyPosted = getSODownPaymentStatus(EdDocNum);
                                    //if (IsDownPaymentAlreadyPosted.Equals("") || IsDownPaymentAlreadyPosted.Equals("0"))
                                    if (true)
                                    {
                                        SalesOrder Osales = GetSO(EdDocNum);
                                        oLogTable.DocSent = "Sent";
                                        oLogTable.DocNum = EdDocNum;
                                        oLogTable.Email = Osales.EmailTemplateID;
                                        oLogTable.CardCode = Osales.CustomerId;
                                        oLogTable.DatePaid = DocPaiddate.ToString();
                                        oLogTable.Type = Type;
                                        var AmountPaid = GetMatrixValue(oMatrixReceivedPayment, i, "AmtPaid").ToString();
                                        Osales.Amount = ConvertDecimal(AmountPaid);
                                        oLogTable.AmountPaid = (decimal)Osales.Amount;
                                        var Method4Digit = GetMatrixValue(oMatrixReceivedPayment, i, "Method").ToString();
                                        oLogTable.Method = Method4Digit;

                                        bool flgPostingStatus = false;
                                        String TransationId = "";
                                        String Msg = "";
                                        PostArDownpaymeny(Osales, ref Msg, ref flgPostingStatus, ref TransationId);
                                        if (flgPostingStatus)
                                        {
                                            oLogTable.Status = "successfully Applied";
                                            ShowSystemLog(String.Format("Payment Received Action: {1} against Sales Order: {0} Status:Done", EdDocNum, "AR DownPayment"));
                                            var oPaymentResponse = ebiz.MarkEbizWebFormPaymentAsApplied(getToken(), PID);
                                            ApplicationTransactionRequest oAppTranReq = new ApplicationTransactionRequest();
                                            oAppTranReq.SoftwareId = Software;
                                            oAppTranReq.LinkedToTypeId = Osales.SalesOrderNumber;
                                            oAppTranReq.LinkedToInternalId = Osales.SalesOrderInternalId;
                                            Customer ocust = ebiz.GetCustomer(getToken(), Osales.CustomerId, null);
                                            oAppTranReq.CustomerInternalId = ocust.CustomerInternalId;
                                            oAppTranReq.TransactionTypeId = trasactionObj.TransactionType;
                                            oAppTranReq.TransactionDate = DateTime.Now.ToShortDateString();
                                            oAppTranReq.TransactionId = TransactionRefNum;
                                            var resp = ebiz.AddApplicationTransaction(getToken(), oAppTranReq);
                                        }
                                        else
                                        {
                                            oLogTable.Status = Msg;

                                        }
                                        AddUploadDataLog(oLogTable, "CCSO");
                                        #region last work
                                        /*
                                        String IsCCTransLogAlreadyPosted = getSOAuthonlyAlreadyPostedStatus(EdDocNum);
                                        if (IsCCTransLogAlreadyPosted.Equals("") || IsCCTransLogAlreadyPosted.Equals("0"))
                                        {
                                            #region CCtrans
                                            Osales = GetSO(EdDocNum);
                                            TransactionResponse resp = trasactionObj.Response;

                                            CCTRAN cctran = new CCTRAN();
                                            cctran.OrderID = getDocEntry(EdDocNum,"ORDR").ToString();
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
                                            if (trasactionObj.TransactionType == "Sale")
                                                cctran.command = "cc:sales";
                                            cctran.amount = trasactionObj.Details.Amount.ToString();

                                            insert(cctran);
                                            #endregion CCtrans
                                        }
                                       */
                                        #endregion
                                    }
                                    else
                                    {


                                        //                                        oEdit.Value = "Paid";
                                    }
                                }
                                else
                                {
                                    String IsCCTransLogAlreadyPosted = getSOAuthonlyAlreadyPostedStatus(getDocEntry(EdDocNum, "ORDR").ToString());
                                    if (IsCCTransLogAlreadyPosted.Equals("") || IsCCTransLogAlreadyPosted.Equals("0"))
                                    {
                                        #region CCtrans
                                        SalesOrder Osales = GetSO(EdDocNum);
                                        TransactionResponse resp = trasactionObj.Response;

                                        CCTRAN cctran = new CCTRAN();
                                        cctran.OrderID = getDocEntry(EdDocNum, "ORDR").ToString();// EdDocNum;
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
                                        #region Log
                                        ShowSystemLog(String.Format("Payment Received Action: {1} against Sales Order: {0} Status:Done", EdDocNum, "AR DownPayment"));
                                        var oPaymentResponse = ebiz.MarkEbizWebFormPaymentAsApplied(getToken(), PID);
                                        ApplicationTransactionRequest oAppTranReq = new ApplicationTransactionRequest();
                                        oAppTranReq.SoftwareId = Software;
                                        oAppTranReq.LinkedToTypeId = Osales.SalesOrderNumber;
                                        oAppTranReq.LinkedToInternalId = Osales.SalesOrderInternalId;
                                        Customer ocust = ebiz.GetCustomer(getToken(), Osales.CustomerId, null);
                                        oAppTranReq.CustomerInternalId = ocust.CustomerInternalId;
                                        oAppTranReq.TransactionTypeId = trasactionObj.TransactionType;
                                        oAppTranReq.TransactionDate = DateTime.Now.ToShortDateString();
                                        oAppTranReq.TransactionId = TransactionRefNum;
                                        var resp1 = ebiz.AddApplicationTransaction(getToken(), oAppTranReq);
                                        oLogTable = new LogTable();
                                        oLogTable.Status = "successfully Applied";
                                        oLogTable.DocNum = EdDocNum;
                                        oLogTable.Email = Osales.EmailTemplateID;
                                        oLogTable.CardCode = Osales.CustomerId;
                                        oLogTable.DatePaid = DocPaiddate.ToString();
                                        oLogTable.Type = Type;
                                        var AmountPaid = GetMatrixValue(oMatrixReceivedPayment, i, "AmtPaid").ToString();
                                        Osales.Amount = ConvertDecimal(AmountPaid);
                                        oLogTable.AmountPaid = (decimal)Osales.Amount;
                                        var Method4Digit = GetMatrixValue(oMatrixReceivedPayment, i, "Method").ToString();
                                        oLogTable.Method = Method4Digit;
                                        AddUploadDataLog(oLogTable, "CCSO");
                                        #endregion
                                        #endregion CCtrans
                                    }
                                    else
                                    {
                                        //  oEdit.Value = "Paid";
                                    }
                                }

                            }
                            else
                            {

                                ShowSystemLog(String.Format("Payment Already Marked against Sales Order: {0}", EdDocNum));
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
    }
    #endregion

    #region Common Function
    private String GetPaymentIDURL(String url)
    {
        string output = "";
        var uri = new Uri(url);
        var query = HttpUtility.ParseQueryString(uri.Query);
        output = query.Get("pid");
        return output;
    }
    #endregion
}
