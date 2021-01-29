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
using System.Collections.Generic;
using eBizChargeForSAP.ServiceReference1;

partial class SAP
{
    #region DataMember
    SAPbouiCOM.Form theExportBPForm;
    string MtxBPExport = "mtMain";
    string dtBPExport = "dtBP";
    SAPbouiCOM.DataTable oDTBpExport;
    bool BPExportflgDocInitiallization = false;
    SAPbouiCOM.Matrix oMatrixBpExport;
    bool flStateExportBP = true;

    #endregion


    #region FormEvent
    private void CreateExportBp()
    {
        try
        {
            BPExportflgDocInitiallization = false;
            AddXMLForm("Presentation_Layer.ExportBP.xml");
            theExportBPForm.Freeze(true);
            InitializeFormExportBp();
            theExportBPForm.Freeze(false);
            flStateExportBP = true;
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }
    private void ExportBPFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            theExportBPForm = form;
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
                        if (pVal.ItemUID == "btnPost")
                        {
                            HandleItemPressBPExport(form);
                        }
                        if (pVal.ItemUID == MtxBPExport && pVal.ColUID.Equals("clSel") && pVal.Row == 0 && !pVal.BeforeAction)
                        {
                            HandleExportBPSelectAll(form, MtxBPExport, oDTBpExport);
                        }
                        if (pVal.ItemUID == MtxBPExport)
                        {
                            if (form.Items.Item(MtxBPExport) == null)
                                return;
                            oMatrixBpExport.FlushToDataSource();
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

    private void HandleItemPressBPExport(SAPbouiCOM.Form form)
    {
        bool FlgChange = false;
        try
        {
            if (form.Items.Item(MtxBPExport) == null)
                return;

            form.Items.Item(MtxBPExport).AffectsFormMode = false;
            oMatrixBpExport = (SAPbouiCOM.Matrix)form.Items.Item(MtxBPExport).Specific;
            int RowCount = oDTBpExport.Rows.Count;
            if (RowCount > 0)
            {
                for (int i = oDTBpExport.Rows.Count - 1; i >= 0; i--)
                {
                    var EdCardCode = GetMatrixValue(oMatrixBpExport, i, "colCard").ToString();
                    var EdCardName = GetMatrixValue(oMatrixBpExport, i, "CardName").ToString();
                    var EdEmail = GetMatrixValue(oMatrixBpExport, i, "E_Mail").ToString();

                    var chk = (SAPbouiCOM.CheckBox)oMatrixBpExport.Columns.Item("clSel").Cells.Item(i + 1).Specific;
                    if (chk.Checked)
                    {
                        FlgChange = true;
                        Customer oCust = NewCustomer(EdCardCode, EdCardName, EdEmail);
                        PostBP(oCust, i, oMatrixBpExport);
                    }
                    else
                    {

                    }

                }
                if (FlgChange)
                {
                    //   oMatrixBpExport.LoadFromDataSource();
                    oMatrixBpExport.FlushToDataSource();
                }
            }
        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }

    private void HandleExportBPSelectAll(SAPbouiCOM.Form form, String MatxStr, SAPbouiCOM.DataTable oDataTable)
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
                    if (flStateExportBP)
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
                flStateExportBP = !flStateExportBP;

                //    oMatrixBpExport.FlushToDataSource();
            }

        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }

    #endregion
    private void PostBP(Customer oCustomer, int i, SAPbouiCOM.Matrix oMatx)
    {
        String EbizKey = "";
        try
        {
            CustomerResponse resp = ebiz.AddCustomer(getToken(), oCustomer);
            if (resp.ErrorCode == 2 && resp.Error.Equals("Record already exists") && resp.Status.Equals("Failed"))
            {
                Customer OldCustomer = ebiz.GetCustomer(getToken(), oCustomer.CustomerId, "");
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
            //SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clEbizkey").Cells.Item(i + 1).Specific;
            //oEdit.Value = EbizKey;
            oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
            oEdit.Value = "True";
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }
    private void UpdateEbizChargeKey(String table, String EbizChargeKey, String TableKey)
    {
        String sql = String.Empty;
        try
        {
            string Where = "";
            if (table.Equals("OCRD"))
            {
                Where = String.Format("where \"CardCode\"='{0}'", TableKey);
            }
            else if (table.Equals("OITM"))
            {
                Where = String.Format("where \"ItemCode\"='{0}'", TableKey);
            }
            else if (table.Equals("ORDR"))
            {
                Where = String.Format("where \"DocNum\"='{0}'", TableKey);
            }
            else if (table.Equals("OINV"))
            {
                Where = String.Format("where \"DocNum\"='{0}'", TableKey);
            }

            sql = String.Format("update {0} set \"U_EbizChargeID\"='{1}'  {2} ", table, EbizChargeKey, Where);

            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void UpdateEbizChargeURL(String table, String EbizChargeKey, String TableKey)
    {
        String sql = String.Empty;
        try
        {
            string Where = "";

            if (table.Equals("ORDR"))
            {
                Where = String.Format("where \"DocNum\"='{0}'", TableKey);
            }

            sql = String.Format("update {0} set \"{3}\"='{1}'  {2} ", table, EbizChargeKey, Where, "U_" + U_EbizChargeURL);

            execute(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }


    #region Helping Function
    private String GetComboValue(SAPbouiCOM.Form oForm, String ItemId)
    {

        var Output = String.Empty;
        try
        {
            SAPbouiCOM.ComboBox cmb = (SAPbouiCOM.ComboBox)oForm.Items.Item(ItemId).Specific;
            Output = cmb.Selected.Value;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return Output;

    }

    public static String DatetimeTosapFormatDate(DateTime datetime)
    {

        String Result = "";
        try
        {
            if (datetime > new DateTime(1920, 1, 1))
            {
                Result = Convert.ToDateTime(datetime).ToString("yyyyMMdd");
            }
            else
            {
                Result = "";
            }
        }
        catch (Exception ex)
        {

        }
        return Result;
    }
    private DateTime getdate(SAPbouiCOM.EditText ed)
    {
        try
        {

            DateTime dt = DateTime.ParseExact(ed.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None);
            return dt;
        }
        catch (Exception ex)
        {
            errorLog(ex);
            return DateTime.Now;

        }

    }
    private DateTime getdateFor(SAPbouiCOM.EditText ed)
    {
        try
        {

            DateTime dt = DateTime.ParseExact(ed.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None);
            return dt;
        }
        catch (Exception ex)
        {
            errorLog(ex);
            return DateTime.MinValue;

        }

    }
    private Customer NewCustomer(String CardCode1, String CardName1, String Email1)
    {
        Customer oCustomer = new Customer();

        try
        {
            oCustomer.CustomerId = CardCode1;
            oCustomer.FirstName = CardName1;
            oCustomer.Email = Email1;

        }
        catch (Exception ex)
        {
            errorLog(ex);
            throw;
        }
        return oCustomer;
    }

    private Object GetMatrixValue(SAPbouiCOM.Matrix oMatrx, int i, String Col)
    {
        String output = "";
        try
        {
            var edOut = (SAPbouiCOM.EditText)oMatrx.Columns.Item(Col).Cells.Item(i + 1).Specific;
            output = edOut.String;
        }
        catch (Exception ex)
        {
            output = "";
        }
        return output;
    }
    private void InitializeFormExportBp()
    {
        try
        {

            //    GetReference();
            BPExportflgDocInitiallization = true;
            ShowDataExportBp();

        }
        catch (Exception ex)
        {

            errorLog(ex);
        }


    }

    private void ShowDataExportBp()
    {
        try
        {
            tProcess = new Thread(ExportBPCreateTabThreadProc);
            tProcess.Start();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            throw;
        }
    }

    public void ExportBPCreateTabThreadProc()
    {
        try
        {
            if (formSelTarget == null)
            {
                if (theExportBPForm != null)
                    formSelTarget = theExportBPForm;
                else
                    return;
            }
            if (formSelTarget.TypeEx != formExportBP)
                return;

            if (!BPExportflgDocInitiallization) return;

            ShowExportBp(theExportBPForm);

        }
        catch (Exception ex)
        {
        }

    }
    private void ShowExportBp(SAPbouiCOM.Form form, String DefaultCheck = "N")
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxBPExport) == null)
                return;

            form.Items.Item(MtxBPExport).AffectsFormMode = false;
            oMatrixBpExport = (SAPbouiCOM.Matrix)form.Items.Item(MtxBPExport).Specific;
            oMatrixBpExport.Clear();
            oDTBpExport = getDataTableBP(form);
            sql = String.Format("SELECT '{0}' as [sel],ROW_NUMBER() OVER (ORDER BY T0.[CardCode]) AS [ID],T0.[CardCode] as [c], T0.[CardName],T0.[Balance],T0.[E_Mail], IFNULL(T0.[U_EbizChargeID],'') as [EBiz],CASE WHEN IFNULL(T0.[U_EbizChargeID], '') = '' THEN 'False' ELSE 'True'END as [Sync] FROM OCRD T0 WHERE T0.[CardType] = 'C'", DefaultCheck).Replace("]", "\"").Replace("[", "\"");

            oDTBpExport.ExecuteQuery(sql);
            BindMatrixBP(oMatrixBpExport, "clSel", "sel", dtBPExport, true);
            BindMatrixBP(oMatrixBpExport, "clNo", "ID");
            BindMatrixBP(oMatrixBpExport, "colCard", "c");
            BindMatrixBP(oMatrixBpExport, "CardName", "CardName");
            BindMatrixBP(oMatrixBpExport, "Balance", "Balance");
            BindMatrixBP(oMatrixBpExport, "E_Mail", "E_Mail");
            // BindMatrixBP(oMatrixBpExport, "clEbizkey", "EBiz");
            BindMatrixBP(oMatrixBpExport, "clSync", "Sync");

            oMatrixBpExport.LoadFromDataSource();
            oMatrixBpExport.AutoResizeColumns();

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private SAPbouiCOM.DataTable getDataTableBP(SAPbouiCOM.Form form)
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = form.DataSources.DataTables.Item(dtBPExport);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
            oDT = form.DataSources.DataTables.Add(dtBPExport);

        }

        return oDT;
    }
    private void BindMatrixBP(SAPbouiCOM.Matrix oMatrix, string mtCol, string dtCol, string dt = null, bool editable = false)
    {
        try
        {
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item(mtCol);
            oColumn.Editable = editable;
            if (dt == null)
            {
                oColumn.DataBind.Bind(dtBPExport, dtCol);
            }
            else
                oColumn.DataBind.Bind(dt, dtCol);
        }
        catch (Exception ex)
        {
            errorLog("Can not bind " + mtCol + " to " + dtCol + ".  error: " + ex.Message);
        }

    }
    public Decimal ConvertDecimal(object input)
    {
        Decimal output = 0;
        try
        {
            output = Convert.ToDecimal(input);
        }
        catch (Exception ex)
        {

        }
        return output;
    }
    public Double ConvertDouble(object input)
    {
        Double output = 0;
        try
        {
            output = Convert.ToDouble(input);
        }
        catch (Exception ex)
        {

        }
        return output;
    }
    public String ConvertString(object input)
    {
        String output = "";
        try
        {
            output = Convert.ToString(input);
        }
        catch (Exception ex)
        {

        }
        return output;
    }
    #endregion

}

public class ExportBpModel
{

    public int serialNo { get; set; }
    public string CardCode { get; set; }
    public string CardName { get; set; }
    public string Email { get; set; }
    public string Sync { get; set; }
    public string EbizChargeKey { get; set; }
    public decimal Balance { get; set; }

}
public class ItemMasterSAP
{

    public string ItemCode { get; set; }
    public string ItemName { get; set; }
    public decimal Price { get; set; }
    public decimal onHand { get; set; }

}
