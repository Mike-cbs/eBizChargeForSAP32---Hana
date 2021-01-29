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
    const string menuOrderImp = "eBizOrderImpMenu";
    const string formOrderImp = "eBizOrderImpForm";
 
     const string editTransID = "edTranID";
     const string cbOrderItem = "cbOITM";
     const string lbCardHolder = "lbCHldr";
     const string lbTAmount = "lbAmt";
     const string lbTax = "lbTax";
     const string lbShipping = "lbShip";
     const string lbDiscount = "lbDisc";
     const string lbCustID = "lbCustID";


     TransactionObject transObjForImport = null;
    SAPbouiCOM.Form oOrderImpForm;

    private void AddOrderImpMenu()
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
            oCreationPackage.UniqueID = menuOrderImp;
            oCreationPackage.String = "eBizCharge Order Import";
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
    private void CreateOrderImpForm()
    {
        try
        {
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Button oButton = null;

            // add a new form
            SAPbouiCOM.FormCreationParams oCreationParams = null;

            oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.FormType = formOrderImp;

            oCreationParams.UniqueID = formOrderImp;
            try
            {
                oOrderImpForm = SBO_Application.Forms.AddEx(oCreationParams);
            }
            catch (Exception)
            {
                oOrderImpForm = SBO_Application.Forms.Item(formOrderImp);

            }

            // set the form properties
            oOrderImpForm.Title = "eBizCharge Order Import";
            oOrderImpForm.Left = 400;
            oOrderImpForm.Top = 100;
            oOrderImpForm.ClientHeight = 260;
            oOrderImpForm.ClientWidth = 450;






            int nTopGap = 25;
            int left = 6;
            int wBtn = 70;
            int hBtn = 19;
            int span = 80;

            oItem = oOrderImpForm.Items.Add(btnFind, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oOrderImpForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Find";
            left += span;
            oItem = oOrderImpForm.Items.Add(btnImp, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oOrderImpForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Import";
            left += span;
            oItem = oOrderImpForm.Items.Add(btnClose, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oOrderImpForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";

            int margin = 8;
            int top = 15;
            int edL = 150; //oItm.Left + oItm.Width;
            int edW = 200;
            int edH = 15;
            int nGap = 22;

            oItem = addPaneItem(oOrderImpForm, editTransID, edL, top, edW, edH, "Gateway Transaction ID:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 100);


            top = oItem.Top + nGap;
            oItem = addPaneItem(oOrderImpForm, lbCardHolder, edL, top, edW, edH, "Card Holder:", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 200);
            top = oItem.Top + nGap;
            oItem = addPaneItem(oOrderImpForm, lbCustID, edL, top, edW, edH, "Customer ID:", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 201);

        
            top = oItem.Top + nGap;

            oItem = addPaneItem(oOrderImpForm, lbTAmount, edL, top, edW, edH, "Amount:", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 202);

            top = oItem.Top + nGap;
            oItem = addPaneItem(oOrderImpForm, lbTax, edL, top, edW, edH, "Tax:", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 203);

            top = oItem.Top + nGap;
            oItem = addPaneItem(oOrderImpForm, lbShipping, edL, top, edW, edH, "Shipping:", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 204);

            top = oItem.Top + nGap;

            oItem = addPaneItem(oOrderImpForm, lbDiscount, edL, top, edW, edH, "Discount:", SAPbouiCOM.BoFormItemTypes.it_STATIC, 0, 206);

            top = oItem.Top + nGap;


            oItem = addPaneItem(oOrderImpForm, cbOrderItem, edL, top, edW, edH, "Order Items:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 205);

            top = oItem.Top + nGap;

          

            oItem = addPaneItem(oOrderImpForm, cbCustomerID, edL, top, edW, edH, "Import to:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 211);

            List<String> list = FindCustomer("");

            SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)oOrderImpForm.Items.Item(cbCustomerID).Specific;
            try
            {
                while (oCB.ValidValues.Count > 0)
                    oCB.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception)
            { }

            foreach (string c in list)
            {
                ComboAddItem(oCB, c);
            }

        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
        oOrderImpForm.Visible = true;

    }
    private void OrderImpFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            oOrderImpForm = form;
            if (pVal.BeforeAction)
            {

            }
            else
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CLICK:

                        break;

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        {
                            switch (pVal.ItemUID)
                            {
                                case btnClose:
                                    oOrderImpForm.Close();
                                    break;
                                case btnImp:
                                    try
                                    {
                                        if ( transObjForImport == null)
                                        {
                                            form.Items.Item(btnFind).Click();
                                        }
                                        else
                                        {
                                            string cardcode = getFormItemVal(form, cbCustomerID);
                                            if (cardcode == "")
                                            {
                                                showMessage("Please select a customer.");
                                            }
                                            else
                                            {
                                                if ( transObjForImport == null)
                                                    form.Items.Item(btnFind).Click();
                                                else
                                                {
                                                    if (SBO_Application.MessageBox("Import order to " + cardcode + "?", 1, "Yes", "No") == 1)
                                                    {
                                                        SBO_Application.Menus.Item(MENU_SO).Activate();
                                                        SAPbobsCOM.Documents doc = importSalesOrder();
                                                        
                                                        SBO_Application.Menus.Item(MENU_LAST).Activate();
                                                        SBO_Application.SetStatusBarMessage(string.Format("Import order {0} added to {1}.", doc.DocNum, cardcode), SAPbouiCOM.BoMessageTime.bmt_Medium, true);

                                                    }                                                   
                                                }

                                            }
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        SBO_Application.SetStatusBarMessage(string.Format("Import order error: {0}.", ex.Message), SAPbouiCOM.BoMessageTime.bmt_Medium, true);


                                    }
                                    break;

                                case btnFind:
                                    try
                                    {
                                        string transID = getFormItemVal(form, editTransID);
                                        if (transID == "")
                                        {
                                            showMessage("Please enter transaction ID.");
                                        }
                                        else
                                        {
                                            SBO_Application.SetStatusBarMessage("Finding transaction please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                                            SecurityToken token = getToken("");
                                            //transObjForImport = ebiz.getTransaction(token, transID);
                                            setLabelCaption(form, lbCardHolder, transObjForImport.BillingAddress.FirstName + " " + transObjForImport.BillingAddress.LastName);
                                            setLabelCaption(form, lbTAmount, transObjForImport.Details.Amount.ToString());
                                            setLabelCaption(form, lbCustID, transObjForImport.CustomerID);
                                            setLabelCaption(form, lbTax, transObjForImport.Details.Tax.ToString());
                                            setLabelCaption(form, lbShipping, transObjForImport.Details.Shipping.ToString());
                                            setLabelCaption(form, lbDiscount, transObjForImport.Details.Discount.ToString());
      
                                            SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbOrderItem).Specific;
                                            try
                                            {
                                                while (oCB.ValidValues.Count > 0)
                                                    oCB.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                            }
                                            catch (Exception)
                                            { }
                                            ComboAddItem(oCB, "");
                                            string s = "";
                                            int i = 0;
                                            foreach (LineItem item in transObjForImport.LineItems)
                                            {
                                                i++;
                                                if (item.SKU == null || item.SKU == "")
                                                    item.SKU = "SKU-" + i.ToString();
                                               s = string.Format("{0} - {1} {2}x{3}", item.SKU, item.ProductName, item.Qty, item.UnitPrice);
                                                
                                                ComboAddItem(oCB, s);
                                            }
                                            try
                                            {
                                                oCB.Select(s);
                                            }
                                            catch (Exception) { }
                                            SBO_Application.SetStatusBarMessage(string.Format("{0} Order Item(s) ready to import.", oCB.ValidValues.Count - 1), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                            // form.Items.Item(btnImp).Click();
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                         transObjForImport = null;
                                        SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbOrderItem).Specific;
                                        try
                                        {
                                            while (oCB.ValidValues.Count > 0)
                                                oCB.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                        }
                                        catch (Exception)
                                        { }
                                        ComboAddItem(oCB, "");
                                        oCB.Select("");
                                        SBO_Application.SetStatusBarMessage(string.Format("Find transaction error: {0}.", ex.Message), SAPbouiCOM.BoMessageTime.bmt_Medium, true);

                                        errorLog(ex);
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
    public string createItem(string code)
    {
        string rtn = "";
        try
        {
            SAPbobsCOM.Items item = (SAPbobsCOM.Items)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
            item.ItemCode = code;
            item.ForeignName = code;
            item.ItemName = code;
            item.InventoryItem = SAPbobsCOM.BoYesNoEnum.tNO;
            int ret = item.Add();
            if (ret != 0)
            {
                errorLog(oCompany.GetLastErrorDescription());
            }
            else
                rtn = code;

            return rtn;

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return rtn;
    }
    public void updateOrderShipTo(string addr, int id)
    {
        string sql = "";
        try
        {
            sql = string.Format("UPDATE ORDR SET \"Address2\" = '{0}' where \"DocEntry\"='{1}'", addr, id);
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    public SAPbobsCOM.Documents importSalesOrder()
    {


        SAPbobsCOM.Documents oDoc = null;
        try
        {
 
            oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
            oDoc.DocNum = getNextTableNum("ORDR");
            oDoc.CardCode = getFormItemVal(oOrderImpForm, cbCustomerID);
            oDoc.DocDueDate = DateTime.Now.AddDays(5);
            oDoc.DocDate = DateTime.Now;
            if (transObjForImport.Details.Shipping > 0)
            {
                oDoc.Expenses.ExpenseCode = getExpenseCode();
                oDoc.Expenses.TaxCode = getEmptTaxCode();
                oDoc.Expenses.LineTotal = transObjForImport.Details.Shipping;
                oDoc.Expenses.Add();
            }
            oDoc.Comments = "Created by eBizCharge Import base on trans id: " + transObjForImport.Response.RefNum;

            foreach (LineItem item in transObjForImport.LineItems)
            {
                if (getItemCode(item.SKU) == "")
                    createItem(item.SKU);

                oDoc.Lines.ItemCode = item.SKU;
                oDoc.Lines.Quantity = getDoubleValue(item.Qty);
                oDoc.Lines.UnitPrice = getDoubleValue(item.UnitPrice);
                oDoc.Lines.Add();
            }
               int r = oDoc.Add();
            if (r != 0)
            {
                string err = "Add Order failed. " + oCompany.GetLastErrorDescription();
                errorLog(err);
            }
            else
            {
                int id = getNextTableID("ORDR") - 1;
                try{
                    if (transObjForImport.ShippingAddress.Street != "")
                    {
                        string addr = transObjForImport.ShippingAddress.Street + " " + transObjForImport.ShippingAddress.Street2 + "\r\n"
                       + transObjForImport.ShippingAddress.City + " " + transObjForImport.ShippingAddress.State + " " + transObjForImport.ShippingAddress.Zip;
                        updateOrderShipTo(addr, id);
                    }
                }catch(Exception)
                {}
                logCustomTransaction(id.ToString(), oDoc.CardCode);
                return oDoc;
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return null;

    }
    private void logCustomTransaction(string id, string cardcode)
    {

        try
        {
            TransactionResponse resp = transObjForImport.Response;

            CCTRAN cctran = new CCTRAN();
            cctran.OrderID = id;
            cctran.InvoiceID = "";
            cctran.CreditMemoID = "";
            cctran.customerID = cardcode;
            cctran.MethodID = "";
            cctran.custNum = "";
            string currency = getCurrency("ORDR" , id);
            string group = getGroupName(cardcode);
            string cardtype = transObjForImport.CreditCardData.CardType;
            string acctid = "";
            string cardName = getCardName(group, cardtype, currency, ref acctid);
            cctran.CCAccountID = int.Parse(acctid);

            cctran.crCardNum = transObjForImport.CreditCardData.CardNumber;
            cctran.Description = transObjForImport.Details.Description;
            cctran.recID = transObjForImport.Details.Invoice;
            
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
            cctran.command = transObjForImport.TransactionType;
           	if(	transObjForImport.TransactionType =="Auth Only")
                cctran.command = "cc:authonly";
            cctran.amount = transObjForImport.Details.Amount.ToString();
            
            insert(cctran);


        }
        catch (Exception ex)
        {
            errorLog(ex);

        }

    }

}

