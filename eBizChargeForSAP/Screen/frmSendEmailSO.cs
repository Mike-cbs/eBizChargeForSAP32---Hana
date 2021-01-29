using eBizChargeForSAP;
using eBizChargeForSAP.ServiceReference1;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

partial class SAP
{


    #region DataMember
    SAPbouiCOM.Form FrmSendEmailSO;
    string MtxSendEmailSO = "mtMain";
    string dtSendEmailSOStr = "dtemp";
    SAPbouiCOM.DataTable odtSendEmailSOStr;
    bool SendEmailSOflgDocInitiallization = false;
    SAPbouiCOM.Matrix oMatrixSendEmailSO;
    bool flStateSendEmailSO = true;
    EditText txtBPcode;
    List<EmailTemplateCust> ListEmailTemplateCust = new List<EmailTemplateCust>();
    String ComBoEmail = "cmbEmail";
    String ComBoPaymentMethod = "cmbPay";
    String txtFrom = "txtFrom";
    String txtSubject = "txtSubject";
    SAPbouiCOM.CheckBox oCheckboxSendEmailSelectAll;
    String oCheckboxSendEmailSelectAllstring = "ckSndE";

    #endregion


    #region FormEvent
    private void CreateSendEmailSO()
    {
        try
        {
            SendEmailSOflgDocInitiallization = false;
            AddXMLForm("Presentation_Layer.SendEmailSO.xml");
            FrmSendEmailSO.Freeze(true);
            InitializeFormSendEmailSO();
            CheckBoxSendEmailSo();
            FrmSendEmailSO.Freeze(false);
            flStateSendEmailSO = true;
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }
    private void SendEmailSOFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            FrmSendEmailSO = form;
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
                            if (pVal.ItemUID == "txtBpCode")
                            {
                                //if (ValidFind())
                                //{
                                //    Thread t = new Thread(ShowDataSendEmailSO);
                                //    t.Start();
                                //}
                            }

                        }
                        catch (Exception)
                        { }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        {
                            if (pVal.ItemUID.Equals(ComBoEmail))
                            {
                                SetEmailTempOnChange();
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                       
                        if (pVal.ItemUID.Equals(oCheckboxSendEmailSelectAllstring))
                        {
                            SelectAllSU(FrmSendEmailSO, MtxSendEmailSO, odtSendEmailSOStr, oCheckboxSendEmailSelectAllstring);
                        }
                        if (pVal.ItemUID == "btnFind")
                        {
                            if (ValidFind())
                            {
                                Thread t = new Thread(ShowDataSendEmailSO);
                                t.Start();
                            }
                        }
                        if (pVal.ItemUID == "btnRef")
                        {
                            SetAllIncomingPayment();

                        }
                        if (pVal.ItemUID == "btnPost")
                        {
                            if (ValidProcess())
                            {
                                HandleItemPressSendEmailSO(form);
                            }
                        }
                        if (pVal.ItemUID == MtxSendEmailSO && pVal.ColUID.Equals("clSel") && pVal.Row == 0 && !pVal.BeforeAction)
                        {
                            HandleSendEmailSOSelectAll(form, MtxSendEmailSO, odtSendEmailSOStr);
                        }
                        if (pVal.ItemUID == MtxSendEmailSO)
                        {
                            if (form.Items.Item(MtxSendEmailSO) == null)
                                return;

                            oMatrixSendEmailSO.FlushToDataSource();
                        }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        // handleItemPress(form, pVal);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                        try
                        {
                            if (pVal.ItemUID == MtxSendEmailSO & pVal.ColUID == "DocNum")
                            {//SO
                                SAPbouiCOM.EditText DocNum = (SAPbouiCOM.EditText)oMatrixSendEmailSO.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific;
                                DoOpenLinkedObjectForm("139", "2050", "8", DocNum.Value);
                            }
                            if (pVal.ItemUID == MtxSendEmailSO & pVal.ColUID == "clEbizkey")
                            {//SO
                             //                                SAPbouiCOM.EditText EbizUrl = (SAPbouiCOM.EditText)oMatrixSendEmailSO.Columns.Item("clEbizkey").Cells.Item(pVal.Row).Specific;
                                SAPbouiCOM.EditText DocNum = (SAPbouiCOM.EditText)oMatrixSendEmailSO.Columns.Item("DocNum").Cells.Item(pVal.Row).Specific;
                                String WebUrl = getEbizURL(DocNum.Value);// EbizUrl.Value;
                                System.Diagnostics.Process.Start(WebUrl);
                            }
                        }
                        catch (Exception ex)
                        {
                            var error = ex.Message;

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

    private void HandleSendEmailSOSelectAll(SAPbouiCOM.Form form, String MatxStr, SAPbouiCOM.DataTable oDataTable)
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
                    if (flStateSendEmailSO)
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
                flStateSendEmailSO = !flStateSendEmailSO;

                //    oMatrixBpExport.FlushToDataSource();
            }

        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }

    private void HandleItemPressSendEmailSO(SAPbouiCOM.Form form)
    {
        bool FlgChange = false;
        try
        {
            if (form.Items.Item(MtxSendEmailSO) == null)
                return;

            if (SBO_Application.MessageBox("Are you sure you want to send email?", 1, "Yes", "No") == 1)
            {
                form.Items.Item(MtxSendEmailSO).AffectsFormMode = false;
                oMatrixSendEmailSO = (SAPbouiCOM.Matrix)form.Items.Item(MtxSendEmailSO).Specific;
                int RowCount = odtSendEmailSOStr.Rows.Count;
                if (RowCount > 0)
                {
                    for (int i = odtSendEmailSOStr.Rows.Count - 1; i >= 0; i--)
                    {
                        var EdDocNum = GetMatrixValue(oMatrixSendEmailSO, i, "DocNum").ToString();

                        var chk = (SAPbouiCOM.CheckBox)oMatrixSendEmailSO.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                        if (chk.Checked)
                        {
                            FlgChange = true;
                            SalesOrder oSalesOrder = GetSO(EdDocNum);
                            POSTSendEmailSO(oSalesOrder, i, oMatrixSendEmailSO);
                        }
                        else
                        {

                        }

                    }
                    if (FlgChange)
                    {
                        //   oMatrixBpExport.LoadFromDataSource();
                        oMatrixSendEmailSO.FlushToDataSource();
                    }
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
    private void SetEmailTempOnChange()
    {
        try
        {
            var SelectcmbVAl = GetEmailComboValue();
            var respEmailTemp = (from a in ListEmailTemplateCust where a.TemplateName == SelectcmbVAl select a).FirstOrDefault();
            SetItemValue(FrmSendEmailSO,txtFrom, respEmailTemp.FromEmail);
            SetItemValue(FrmSendEmailSO,txtSubject, respEmailTemp.Subject);
            if (respEmailTemp.FromEmail.Equals(""))
            {
                SetItemValue(FrmSendEmailSO,txtFrom, "support@ebizcharge.com");
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private string GetEmailTempID()
    {
        string output = "";
        try
        {
            var SelectcmbVAl = GetEmailComboValue();
            var respEmailTemp = (from a in ListEmailTemplateCust where a.TemplateName == SelectcmbVAl select a).FirstOrDefault();
            output = respEmailTemp.TemplateID;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return output;
    }
    private string GetEmailTempID(SAPbouiCOM.Form oForm)
    {
        string output = "";
        try
        {
            var SelectcmbVAl = GetEmailComboValue(oForm);
            var respEmailTemp = (from a in ListEmailTemplateCust where a.TemplateName == SelectcmbVAl select a).FirstOrDefault();
            output = respEmailTemp.TemplateID;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return output;
    }
    private string GetItemValue(String ItemID, SAPbouiCOM.Form oFrom)
    {
        string output = "";
        try
        {
            EditText txtItem = (SAPbouiCOM.EditText)oFrom.Items.Item(ItemID).Specific;
            output = txtItem.Value;
        }
        catch (Exception ex)
        {

        }
        return output;
    }
    private void SetItemValue(SAPbouiCOM.Form oFrom, String ItemID, String Value)
    {

        try
        {
            EditText txtItem = (SAPbouiCOM.EditText)oFrom.Items.Item(ItemID).Specific;
            txtItem.Value = Value;
        }
        catch (Exception ex)
        {

        }

    }
    private void FillEmailCombo()
    {
        if (formSelTarget == null)
        {
            if (FrmSendEmailSO != null)
                formSelTarget = FrmSendEmailSO;
            else
                return;
        }
        if (formSelTarget.TypeEx != formEmailEstimates)
            return;

        if (!SendEmailSOflgDocInitiallization) return;

        try
        {// Email template Code
            SAPbouiCOM.ComboBox cmb_EmailTemplate = (SAPbouiCOM.ComboBox)FrmSendEmailSO.Items.Item(ComBoEmail).Specific;
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
                    SetItemValue(FrmSendEmailSO,txtFrom, eTempList.ElementAt(0).FromEmail);
                }
                else
                {
                    //support@ebizcharge.com
                    SetItemValue(FrmSendEmailSO,txtFrom, "support@ebizcharge.com");
                }
                SetItemValue(FrmSendEmailSO,txtSubject, eTempList.ElementAt(0).TemplateSubject);
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        try
        {
            //PaymentMethod
            SAPbouiCOM.ComboBox cmb_PaymentMethod = (SAPbouiCOM.ComboBox)FrmSendEmailSO.Items.Item(ComBoPaymentMethod).Specific;
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
    private String GetBpCode()
    {
        var CardCode = String.Empty;
        try
        {
            txtBPcode = (SAPbouiCOM.EditText)FrmSendEmailSO.Items.Item("txtBpCode").Specific;
            CardCode = txtBPcode.Value;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return CardCode;

    }
    private String GetBpCode(SAPbouiCOM.Form oform, String fieldId)
    {
        var CardCode = String.Empty;
        try
        {
            txtBPcode = (SAPbouiCOM.EditText)oform.Items.Item(fieldId).Specific;
            CardCode = txtBPcode.Value;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return CardCode;

    }
    private String GetItemValue(SAPbouiCOM.Form oform, String fieldId)
    {
        var OutputValue = String.Empty;
        try
        {
            SAPbouiCOM.EditText oItem = (SAPbouiCOM.EditText)oform.Items.Item(fieldId).Specific;
            OutputValue = oItem.Value;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return OutputValue;

    }
    private String GetEmailComboValue()
    {
        var Output = String.Empty;
        try
        {
            SAPbouiCOM.ComboBox cmb = (SAPbouiCOM.ComboBox)FrmSendEmailSO.Items.Item("cmbEmail").Specific;
            Output = cmb.Selected.Value;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return Output;

    }
    private String GetEmailComboValue(SAPbouiCOM.Form oForm)
    {
        var Output = String.Empty;
        try
        {
            SAPbouiCOM.ComboBox cmb = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbEmail").Specific;
            Output = cmb.Selected.Value;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return Output;

    }
    private String GetPaymentMethodComboValue()
    {
        var Output = String.Empty;
        try
        {

            SAPbouiCOM.ComboBox cmb = (SAPbouiCOM.ComboBox)FrmSendEmailSO.Items.Item(ComBoPaymentMethod).Specific;
            var result = cmb.Selected.Value;
            if (result.Equals("Authorize"))
            {
                Output = "AuthOnly";
            }
            else if (result.Equals("Charge"))
            {
                Output = "sales";
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return Output;

    }
    private bool ValidFind()
    {
        bool flag = false;
        try
        {
            var cardCode = GetBpCode();
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
    private bool ValidProcess()
    {
        bool flag = false;
        try
        {
            var ComboValue = GetEmailComboValue();
            var FromEmail = GetItemValue(txtFrom, FrmSendEmailSO);
            var Subject = GetItemValue(txtSubject, FrmSendEmailSO);

            if (string.IsNullOrEmpty(FromEmail))
            {
                MsgWarning("From Email can't be left empty.");
                return false;
            }
            else if (string.IsNullOrEmpty(Subject))
            {
                MsgWarning("Subject can't be left empty.");
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
    private void InitializeFormSendEmailSO()
    {
        try
        {
            SendEmailSOflgDocInitiallization = true;
            // ShowDataSendEmailSO();
            new Thread(FillEmailCombo).Start();

        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
    }
    private void ShowDataSendEmailSO()
    {
        try
        {
            tProcess = new Thread(SendEmailSOCreateTabThreadProc);
            tProcess.Start();
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    public void SendEmailSOCreateTabThreadProc()
    {
        try
        {
            if (formSelTarget == null)
            {
                if (FrmSendEmailSO != null)
                    formSelTarget = FrmSendEmailSO;
                else
                    return;
            }
            if (formSelTarget.TypeEx != formEmailEstimates)
                return;

            if (!SendEmailSOflgDocInitiallization) return;



            ShowSendEmailSO(FrmSendEmailSO);


        }
        catch (Exception ex)
        {
        }

    }
    private void ShowSendEmailSO(SAPbouiCOM.Form form)
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxSendEmailSO) == null)
                return;

            form.Items.Item(MtxSendEmailSO).AffectsFormMode = false;
            oMatrixSendEmailSO = (SAPbouiCOM.Matrix)form.Items.Item(MtxSendEmailSO).Specific;
            oMatrixSendEmailSO.Clear();
            odtSendEmailSOStr = getDataTableSendEmailSO(form);

            var CardCode = GetBpCode();
            var emailqtr = string.Format("select \"E_Mail\" from OCRD where \"CardCode\" = '{0}'", CardCode);

            sql = String.Format("select 'N' as [clSel],'    ' as [Status],({2}) as Email,ROW_NUMBER() OVER (ORDER BY [DocEntry]) AS [ID],[DocNum],[DocDate],[DocTotal],[CardCode],[CardName],IFNULL([{1}],'') as [EBiz],CASE WHEN IFNULL([{1}], '') = '' THEN 'False' ELSE 'True'END as [Sync],CASE WHEN IFNULL([{1}], '') = '' THEN '' ELSE 'View Form'END as [ViewForm] from ORDR where  [DocStatus]='O' and [CardCode]='{0}'  and IFNULL([U_EbizChargeMarkPayment],'')<>'Paid'  ", CardCode, "U_" + U_EbizChargeURL, emailqtr).Replace("]", "\"").Replace("[", "\"");

            odtSendEmailSOStr.ExecuteQuery(sql);

            BindMatrixSendEmailSO(oMatrixSendEmailSO, "clSel", "clSel", dtItem, true);
            BindMatrixSendEmailSO(oMatrixSendEmailSO, "clNo", "ID");
            BindMatrixSendEmailSO(oMatrixSendEmailSO, "DocNum", "DocNum");//Email
            BindMatrixSendEmailSO(oMatrixSendEmailSO, "Email", "Email", null, true);//

            BindMatrixSendEmailSO(oMatrixSendEmailSO, "Status", "Status", null, true);//
            BindMatrixSendEmailSO(oMatrixSendEmailSO, "DocDate", "DocDate");
            BindMatrixSendEmailSO(oMatrixSendEmailSO, "CardCode", "CardCode");
            BindMatrixSendEmailSO(oMatrixSendEmailSO, "CardName", "CardName");
            BindMatrixSendEmailSO(oMatrixSendEmailSO, "DocTotal", "DocTotal");
            BindMatrixSendEmailSO(oMatrixSendEmailSO, "clEbizkey", "ViewForm");
            BindMatrixSendEmailSO(oMatrixSendEmailSO, "clSync", "Sync");

            oMatrixSendEmailSO.LoadFromDataSource();
            oMatrixSendEmailSO.AutoResizeColumns();

            SetAllIncomingPayment();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private SAPbouiCOM.DataTable getDataTableSendEmailSO(SAPbouiCOM.Form form)
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = form.DataSources.DataTables.Item(MtxSendEmailSO);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
            oDT = form.DataSources.DataTables.Add(MtxSendEmailSO);

        }

        return oDT;
    }

    private void BindMatrixSendEmailSO(SAPbouiCOM.Matrix oMatrix1, string mtCol, string dtCol, string dt = null, bool editable = false)
    {
        try
        {
            SAPbouiCOM.Column oColumn = oMatrix1.Columns.Item(mtCol);
            oColumn.Editable = editable;
            if (dt == null)
            {
                oColumn.DataBind.Bind(MtxSendEmailSO, dtCol);
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

    #region SAP Event
    private void POSTSendEmailSO(SalesOrder oSalesOrder, int i, SAPbouiCOM.Matrix oMatx)
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
                    GetEbizWebFormURL(ToEmail, oSalesOrder, ref Flag, ref EbizKey);
                    if (Flag)
                    {
                        ShowSystemLog(String.Format("Email Sent Against Sales Order: {0}", oSalesOrder.SalesOrderNumber));
                    }
                }
            }

            SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clEbizkey").Cells.Item(i + 1).Specific;
            oEdit.Value = EbizKey;
            oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
            oEdit.Value = Flag.ToString();
            //(SAPbouiCOM.EditText)oMatx.Columns.Item("clEbizkey").Cells.Item(i + 1).Specific = "";
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }
    private String GetEmail(int i, SAPbouiCOM.Matrix oMatx)
    {
        string email = "";
        try
        {
            SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("Email").Cells.Item(i + 1).Specific;
            email = oEdit.Value;

        }
        catch (Exception ex)
        {

        }
        return email;
    }
    private bool GetEbizWebFormURL(string ToEmail, SalesOrder oSalesOrder, ref bool FlagStatus, ref string URL)
    {
        bool Flag = false;
        try
        {
            var PaymentMethod_selected = GetPaymentMethodComboValue();
            EbizWebForm pf = new EbizWebForm();
            pf.Clerk = oCompany.UserName;
            pf.Date = Convert.ToDateTime(oSalesOrder.Date);
            pf.DueDate = Convert.ToDateTime(oSalesOrder.DueDate);
            pf.TotalAmount = oSalesOrder.Amount;
            pf.AmountDue = oSalesOrder.Amount;
            pf.CustFullName = oSalesOrder.CustomerId;
            pf.ProcessingCommand = PaymentMethod_selected;

            pf.EmailAddress = ToEmail;
            pf.SoftwareId = Software;
            string TemplateID = GetEmailTempID();
            var Descp = ListEmailTemplateCust.Where(a => a.TemplateID == TemplateID).FirstOrDefault().templateDescriptionField;
            pf.EmailNotes = Descp;
            pf.EmailSubject = GetItemValue(txtSubject, FrmSendEmailSO);
            pf.FromEmail = GetItemValue(txtFrom, FrmSendEmailSO);
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
    private string GetPaymentID(String URL)
    {
        string paymnetId = "";
        try
        {
            Uri myUri = new Uri(URL);
            paymnetId = HttpUtility.ParseQueryString(myUri.Query).Get("pid");
        }
        catch (Exception ex)
        {

        }
        return paymnetId;
    }
    private void SetAllIncomingPayment()
    {
        SAPbouiCOM.EditText oEdit = null;
        try
        {
            if (odtSendEmailSOStr == null) return;
            //    var flg = GetPaymentStatus(PaymentID);
            int RowCount = odtSendEmailSOStr.Rows.Count;
            if (RowCount > 0)
            {

                for (int i = 0; i < odtSendEmailSOStr.Rows.Count; i++)
                {
                    oEdit = (SAPbouiCOM.EditText)oMatrixSendEmailSO.Columns.Item("Status").Cells.Item(i + 1).Specific;
                    //oEdit.Value = "";
                    var EdDocNum = GetMatrixValue(oMatrixSendEmailSO, i, "DocNum").ToString();
                    if (EdDocNum.Equals("87"))
                    {
                        String debug = "";
                    }

                    var EbixURL = GetMatrixValue(oMatrixSendEmailSO, i, "clEbizkey").ToString();
                    //  if (!EbixURL.Equals("") && !EbixURL.Equals("0"))
                    if (EbixURL.Equals("View Form"))
                    {
                        EbixURL = getEbizURL(EdDocNum);
                        var PID = GetPaymentID(EbixURL);
                        String refNum = "";
                        TransactionObject trasactionObj = GetPaymentStatus(PID, ref refNum);
                        if (trasactionObj != null)
                        {
                            var PaymentMethod = getSOpaymentType(EdDocNum);
                            if (PaymentMethod.Equals("sales"))
                            {
                                String IsDownPaymentAlreadyPosted = getSODownPaymentStatus(EdDocNum);
                                if (IsDownPaymentAlreadyPosted.Equals("") || IsDownPaymentAlreadyPosted.Equals("0"))
                                {
                                    SalesOrder Osales = GetSO(EdDocNum);
                                    bool flgPostingStatus = false;
                                    string DocEntry = "";
                                    string Msg = "";
                                    PostArDownpaymeny(Osales, ref Msg, ref flgPostingStatus, ref DocEntry);
                                    if (flgPostingStatus)
                                    {
                                        oEdit.Value = "Paid";
                                    }

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
                                }
                                else
                                {

                                    oEdit.Value = "Paid";
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
                                    #endregion CCtrans
                                }
                                else
                                {
                                    oEdit.Value = "Paid";
                                }
                            }
                            //PostArDownpaymeny();

                        }

                    }
                }
            }
        }
        catch (Exception ex)
        {

        }
    }
    private TransactionObject GetPaymentStatus(String PaymentID, ref String refNum)
    {
        TransactionObject PayDetail = null;
        try
        {
            //filters[0].FieldName = "PaymentInternalId";
            //filters[0].ComparisonOperator = "eq";
            //filters[0].FieldValue = "7331bcf0-4ea4-4259-a14f-3343233d295a";
            DateTime fromDt = Convert.ToDateTime("2016/01/01");
            DateTime toDt = DateTime.Now.AddDays(1);
            var Tokn = getToken();
            var Payment = ebiz.SearchEbizWebFormReceivedPayments(Tokn, "", fromDt, toDt, null, 0, 100, "").ToList().Where(a => a.PaymentInternalId == PaymentID).FirstOrDefault();
            if (Payment != null)
            {
                PayDetail = ebiz.GetTransactionDetails(Tokn, Payment.RefNum);
                refNum = Payment.RefNum;
                if (PayDetail != null)
                {
                    return PayDetail;
                }
            }

        }
        catch (Exception ex)
        {

        }
        return PayDetail;
    }
    private bool CheckSO(SalesOrder oSalesOrder)
    {
        bool Flag = false;
        try
        {
            SalesOrderResponse resp = ebiz.AddSalesOrder(getToken(), oSalesOrder);
            if (resp.ErrorCode == 2 && resp.Error.Equals("Record already exists") && resp.Status.Equals("Error"))
            {
                // var tt=ebiz.SearchSalesOrders(getToken(), oSalesOrder.CustomerId, "", oSalesOrder.SalesOrderNumber,"", null, 0, 100, "1", true);
                //var t= ebiz.GetSalesOrder(getToken(), "", "", oSalesOrder.SalesOrderNumber, "");
                Flag = true;
            }
            else if (resp.ErrorCode == 0 && resp.Error.Equals("") && resp.Status.Equals("Success"))
            {
                UpdateEbizChargeKey("ORDR", resp.SalesOrderInternalId, oSalesOrder.SalesOrderNumber);
                Flag = true;
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);

        }
        return Flag;

    }
    private bool CheckBP(string cardCode)
    {
        bool Flag = false;
        try
        {
            Customer oCustomer = NewCustomer(cardCode);
            if (oCustomer != null)
            {
                CustomerResponse resp = ebiz.AddCustomer(getToken(), oCustomer);
                if (resp.ErrorCode == 2 && resp.Error.Equals("Record already exists") && resp.Status.Equals("Failed"))
                {
                    Flag = true;
                }
                else if (resp.ErrorCode == 0 && resp.Error.Equals("") && resp.Status.Equals("Success"))
                {
                    Flag = true;
                }
            }
            else
            {

            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return Flag;
    }
    private bool CheckBPFromExternal(string cardCode, ref String Msg)
    {
        bool Flag = false;
        try
        {
            List<ExportBpModel> Bp = GetBusinessPartner(cardCode);
            if (Bp != null && Bp.Count == 0)
            {
                Msg = String.Format("Business Partner: {0} does not exist", cardCode);
                return false;
            }
            Customer oCustomer = NewCustomer(cardCode);
            if (oCustomer != null)
            {
                CustomerResponse resp = ebiz.AddCustomer(getToken(), oCustomer);
                if (resp.ErrorCode == 2 && resp.Error.Equals("Record already exists") && resp.Status.Equals("Failed"))
                {
                    Flag = true;
                }
                else if (resp.ErrorCode == 0 && resp.Error.Equals("") && resp.Status.Equals("Success"))
                {
                    Flag = true;
                }
            }
            else
            {

            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            Msg = ex.Message;
        }
        return Flag;
    }
    private bool CheckItem(string ItemCode)
    {
        bool Flag = false;
        try
        {
            ItemMasterSAP oItem = GetSAPItem(ItemCode);
            if (oItem != null)
            {
                ItemDetails NewItem = new ItemDetails();
                NewItem.ItemId = ItemCode;
                NewItem.Description = oItem.ItemName;
                NewItem.UnitPrice = oItem.Price;
                NewItem.QtyOnHand = oItem.onHand;
                NewItem.SoftwareId = Software;
                ItemDetailsResponse resp = ebiz.AddItem(getToken(), NewItem);
                if (resp.ErrorCode == 2 && resp.Error.Equals("Record already exists") && resp.Status.Equals("Error"))
                {

                }
                else if (resp.ErrorCode == 0 && resp.Error.Equals("") && resp.Status.Equals("Success"))
                {
                    Flag = true;
                }
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return Flag;
    }
    private bool CheckItemForExternal(string ItemCode, ref String Msg)
    {
        bool Flag = false;
        try
        {
            ItemMasterSAP oItem = GetSAPItem(ItemCode);
            if (oItem == null)
            {
                Msg = String.Format("Item {0} does not exist", ItemCode);
                return false;
            }
            oItem = GetSAPItem(ItemCode);
            ItemDetails NewItem = new ItemDetails();
            NewItem.ItemId = ItemCode;
            NewItem.Description = oItem.ItemName;
            NewItem.UnitPrice = oItem.Price;
            NewItem.QtyOnHand = oItem.onHand;
            NewItem.SoftwareId = Software;
            ItemDetailsResponse resp = ebiz.AddItem(getToken(), NewItem);
            if (resp.ErrorCode == 2 && resp.Error.Equals("Record already exists") && resp.Status.Equals("Error"))
            {

            }
            else if (resp.ErrorCode == 0 && resp.Error.Equals("") && resp.Status.Equals("Success"))
            {
                Flag = true;
            }


        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return Flag;
    }

    private Customer NewCustomer(String CardCode1)
    {
        Customer oCustomer = new Customer();
        try
        {
            List<ExportBpModel> Listocrd = GetBusinessPartner(CardCode1);
            if (Listocrd != null && Listocrd.Count > 0)
            {
                oCustomer.CustomerId = Listocrd[0].CardCode;
                oCustomer.FirstName = Listocrd[0].CardName;
                oCustomer.Email = Listocrd[0].Email;
            }
            else
            {
                oCustomer = null;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            throw;
        }
        return oCustomer;
    }
    #endregion

    #region SAP Di APi
    private void PostArDownpaymeny(SalesOrder Osales, ref string Msg, ref bool PaymentStatus, ref string DocEntry)
    {
        try
        {
            //Osales = new SalesOrder();
            //Osales.SalesOrderNumber = "225";
            //Osales.CustomerId = "C000011";
            //Osales.Amount = 27;
            //Osales.Currency = "$";
            //eBizChargeForSAP.ServiceReference1.Item[] Lineitems = new eBizChargeForSAP.ServiceReference1.Item[1];
            //Lineitems[0] = new eBizChargeForSAP.ServiceReference1.Item();
            //Osales.Items = Lineitems;
            //  Osales.Items.

            SAPDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments);
            int sNewObjCode = Convert.ToInt32(getSODocEntry(Osales.SalesOrderNumber.Trim())); ;
            SAPDoc.CardCode = Osales.CustomerId.Trim();
            SAPDoc.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptInvoice;
            SAPDoc.DocTotal = ConvertDouble(Osales.Amount);
            SAPDoc.DocDate = DateTime.Now;
            for (int i = 0; i < Osales.Items.Count(); i++)
            {
                SAPDoc.Lines.BaseType = 17;
                SAPDoc.Lines.BaseEntry = sNewObjCode;
                SAPDoc.Lines.BaseLine = i;
                SAPDoc.Lines.Add();
            }

            int sta = 0;
            sta = SAPDoc.Add();
            var sapDocEntry = oCompany.GetNewObjectKey();
            DocEntry = sapDocEntry;
            try
            {
                SAPbobsCOM.Payments oPmt = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                oPmt.DocCurrency = Osales.Currency.Trim();
                oPmt.Invoices.DocEntry = Convert.ToInt32(DocEntry);
                oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_DownPayment;

                oPmt.CardCode = Osales.CustomerId.Trim();
                oPmt.DocDate = DateTime.Now;
                oPmt.Remarks = "eBizCharge " + sapDocEntry;
                oPmt.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
                oPmt.Invoices.SumApplied = ConvertDouble(Osales.Amount);
                oPmt.CashSum= ConvertDouble(Osales.Amount); ;
                int r = oPmt.Add();
            }
            catch { }


            if (sta != 0)
            {
                sapDocEntry = "";
                int erroCode = 0;
                string errDescr = "";
                oCompany.GetLastError(out erroCode, out errDescr);
                PaymentStatus = false;
                //Flag = false;
                Msg = errDescr;
            }
            else
            {// Posted
             //Flag = true;
                Msg = "";
                PaymentStatus = true;
                //DocEntry = Convert.ToInt32(sapDocEntry);
            }
        }
        catch (Exception ex)
        {
            Msg = ex.Message;
            PaymentStatus = false;

        }


    }

    #endregion

    #region CheckBox Functionality
    private void CheckBoxSendEmailSo()
    {
        try
        {

            oCheckboxSendEmailSelectAll = ((SAPbouiCOM.CheckBox)(FrmSendEmailSO.Items.Item(oCheckboxSendEmailSelectAllstring).Specific));
            FrmSendEmailSO.DataSources.UserDataSources.Add(oCheckboxSendEmailSelectAllstring, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oCheckboxSendEmailSelectAll.DataBind.SetBound(true, "", oCheckboxSendEmailSelectAllstring);
            FrmSendEmailSO.DataSources.UserDataSources.Item(oCheckboxSendEmailSelectAllstring).Value = "N";
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }
    }
    #endregion

}
public class EmailTemplateCust
{

    public String TemplateName { get; set; }
    public String TemplateID { get; set; }
    public String FromEmail { get; set; }
    public String Subject { get; set; }
    public String templateDescriptionField { get; set; }

    public EmailTemplateCust() { }

    public EmailTemplateCust(String TempName, String TempID, String Fromemail, String Subj, string description)
    {
        this.TemplateName = TempName;
        this.TemplateID = TempID;
        this.FromEmail = Fromemail;
        this.Subject = Subj;
        this.templateDescriptionField = description;
    }
}

