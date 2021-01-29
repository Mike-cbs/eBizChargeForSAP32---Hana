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
    SAPbouiCOM.Form theExportItemForm;
    string MtxItem = "mtMain";
    string dtItem = "dtemp";
    SAPbouiCOM.DataTable oDTItemExport;
    bool ItemExportflgDocInitiallization = false;
    SAPbouiCOM.Matrix oMatrixItemExport;
    bool flStateExportItem = true;

    #endregion


    #region FormEvent
    private void CreateExportItem()
    {
        try
        {
            ItemExportflgDocInitiallization = false;
            AddXMLForm("Presentation_Layer.ExportItem.xml");
            theExportItemForm.Freeze(true);
            InitializeFormExportItem();
            theExportItemForm.Freeze(false);
            flStateExportItem = true;
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }
    private void ExportItemFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            theExportItemForm = form;
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
                            HandleItemPressItemExport(form);
                        }
                        if (pVal.ItemUID == MtxItem)
                        {
                            if (form.Items.Item(MtxItem) == null)
                                return;

                            oMatrixItemExport.FlushToDataSource();
                        }

                        if (pVal.ItemUID == MtxItem && pVal.ColUID.Equals("clSel") && pVal.Row == 0 && !pVal.BeforeAction)
                        {
                            HandleExportItemSelectAll(form, MtxItem, oDTItemExport);
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
    private void HandleExportItemSelectAll(SAPbouiCOM.Form form, String MatxStr, SAPbouiCOM.DataTable oDataTable)
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
                    if (flStateExportItem)
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
                flStateExportItem = !flStateExportItem;
            }

        }
        catch (Exception ex)
        {

            errorLog(ex.Message);
        }
    }
    private void HandleItemPressItemExport(SAPbouiCOM.Form form)
    {
        bool FlgChange = false;
        try
        {
            if (form.Items.Item(MtxItem) == null)
                return;

            form.Items.Item(MtxItem).AffectsFormMode = false;
            oMatrixItemExport = (SAPbouiCOM.Matrix)form.Items.Item(MtxItem).Specific;
            int RowCount = oDTItemExport.Rows.Count;
            if (RowCount > 0)
            {
                for (int i = oDTItemExport.Rows.Count - 1; i >= 0; i--)
                {
                    var EdItemCode = GetMatrixValue(oMatrixItemExport, i, "ItemCode").ToString();
                    var EdItemName = GetMatrixValue(oMatrixItemExport, i, "ItemName").ToString();

                    var chk = (SAPbouiCOM.CheckBox)oMatrixItemExport.Columns.Item("clSel").Cells.Item(i + 1).Specific;
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

                        PostItem(NewItem, i, oMatrixItemExport);
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

    #endregion
    #region SAP
    private void PostItem(ItemDetails Oitem, int i, SAPbouiCOM.Matrix oMatx)
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

            //SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clEbizkey").Cells.Item(i + 1).Specific;
            //oEdit.Value = EbizKey;
            SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatx.Columns.Item("clSync").Cells.Item(i + 1).Specific;
            oEdit.Value = "True";
            //(SAPbouiCOM.EditText)oMatx.Columns.Item("clEbizkey").Cells.Item(i + 1).Specific = "";
        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }


    #endregion

    #region Helping Function
    private void ShowSystemLog(String msg)
    {
        try
        {
            SBO_Application.StatusBar.SetText(msg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }

    private void InitializeFormExportItem()
    {
        try
        {
            ItemExportflgDocInitiallization = true;
            ShowDataExportItem();
        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
    }

    private void ShowDataExportItem()
    {
        try
        {
            tProcess = new Thread(ExportItemCreateTabThreadProc);
            tProcess.Start();
        }
        catch (Exception ex)
        {
            errorLog(ex);
            throw;
        }
    }

    public void ExportItemCreateTabThreadProc()
    {
        try
        {
            if (formSelTarget == null)
            {
                if (theExportItemForm != null)
                    formSelTarget = theExportItemForm;
                else
                    return;
            }
            if (formSelTarget.TypeEx != formExportItemMaster)
                return;

            if (!ItemExportflgDocInitiallization) return;

            ShowExportItem(theExportItemForm);

        }
        catch (Exception ex)
        {
        }

    }
    private void ShowExportItem(SAPbouiCOM.Form form)
    {
        string sql = "";
        try
        {
            if (form.Items.Item(MtxItem) == null)
                return;

            form.Items.Item(MtxItem).AffectsFormMode = false;
            oMatrixItemExport = (SAPbouiCOM.Matrix)form.Items.Item(MtxItem).Specific;
            oMatrixItemExport.Clear();
            oDTItemExport = getDataTableItemMaster(form);
            sql = String.Format("select 'N' as [clSel],ROW_NUMBER() OVER (ORDER BY [DocEntry]) AS [ID],[ItemCode],[ItemName],IFNULL([U_EbizChargeID],'') as [EBiz],CASE WHEN IFNULL([U_EbizChargeID], '') = '' THEN 'False' ELSE 'True'END as [Sync] from OITM ").Replace("]", "\"").Replace("[", "\"");

            oDTItemExport.ExecuteQuery(sql);

            BindMatrixItem(oMatrixItemExport, "clSel", "clSel", dtItem, true);
            BindMatrixItem(oMatrixItemExport, "clNo", "ID");
            BindMatrixItem(oMatrixItemExport, "ItemCode", "ItemCode");
            BindMatrixItem(oMatrixItemExport, "ItemName", "ItemName");

           // BindMatrixItem(oMatrixItemExport, "clEbizkey", "EBiz");
            BindMatrixItem(oMatrixItemExport, "clSync", "Sync");

            oMatrixItemExport.LoadFromDataSource();
            oMatrixItemExport.AutoResizeColumns();

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }

    private SAPbouiCOM.DataTable getDataTableItemMaster(SAPbouiCOM.Form form)
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = form.DataSources.DataTables.Item(MtxItem);
        }
        catch (Exception ex)
        {
            var error = ex.Message;
            oDT = form.DataSources.DataTables.Add(MtxItem);

        }

        return oDT;
    }
    private void BindMatrixItem(SAPbouiCOM.Matrix oMatrix1, string mtCol, string dtCol, string dt = null, bool editable = false)
    {
        try
        {
            SAPbouiCOM.Column oColumn = oMatrix1.Columns.Item(mtCol);
            oColumn.Editable = editable;
            if (dt == null)
            {
                oColumn.DataBind.Bind(MtxItem, dtCol);
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

}
public class ItemMasterModel
{
    public int serialNo { get; set; }
    public string ItemCode { get; set; }
    public string ItemName { get; set; }
    public string Sync { get; set; }
    public string EbizChargeKey { get; set; }


}
