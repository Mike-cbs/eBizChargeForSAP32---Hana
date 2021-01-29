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
using System.Linq;

public class SAPCust
{
    public CCCUST cccust {get; set; }
    public CustomerObject custObj {get; set; }
    public string key { get; set; }
    public string cardHolder { get; set; }
    public string cardNum { get; set; }
    public string CCAccount { get; set; }
    public string SaveCard { get; set; }
}

partial class SAP
{
    int paneBP = 125;
    int paneCCLog = 126;
    List<SAPCust> SAPCustList = new List<SAPCust>();
    SAPCust selSAPCust = null;
    const string btnAdd = "Add";
    const string btnDelete = "Delete";
    const string btnUpdate = "Update";
    const string lbDeclined = "lbDecl";
    const string tbeBizCharge = "tbeBiz";
    const string tbPaymentMethod = "tbPayMe";
 

    private void BusinessPartnerFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            if (pVal.BeforeAction)
            {
                switch(pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        AddTab(form, paneBP, "eBizCharge", "156", tbeBizCharge);
                        AddTab(form, paneBP, "Payment Method", "243", tbPaymentMethod, paneBP, false, 50, -100);
                        AddTab(form, paneBP, "Credit Card Log", tbPaymentMethod, tbCreditCard, paneBP, false, 50, -100);
                        AddCCLogMatrix(form, paneCCLog, "276", 300, 10);
                        SAPbouiCOM.Item oItemRef = form.Items.Item(mxtCCLog);
                        SAPbouiCOM.Item oItem = form.Items.Add("rect", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
                        oItem.Top = oItemRef.Top - 9;
                        oItem.Left = oItemRef.Left;
                        oItem.Width = oItemRef.Width;
                        oItem.Height = 1;
                        oItem.FromPane = paneBP;
                        oItem.ToPane = paneCCLog;
                        AddCCFormField(form, paneBP);
                        form.Items.Item(tbeBizCharge).Visible = true;
                        loadCustData(form);
                        break;
                
                }
            }
            else
            {
                switch(pVal.EventType) 
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
                       
                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        {
                            switch (pVal.ItemUID)
                            {
                                case cbPaymentMethod:
                                    HandleSelect(form);
                                    break;
                                case cbActive:
                                   
                                        HandleActiveSelect(form);
                                        
                                   
                                    break;
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        handleItemPress(form, pVal);
                        break;
                }
                
            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
    private void loadCustData(SAPbouiCOM.Form form)
    {
        try
        {
           
            if (getComboBoxVal(form, "40") == "C")
            {
                if (getFormEditVal(form, "5") != "")
                {
                    form.Items.Item(tbeBizCharge).Visible = true;
                    if (getLabelCaption(form, stRecordID) != getFormEditVal(form, "5"))
                    {
                        setLabelCaption(form, stRecordID, getFormEditVal(form, "5"));
                        reload(form, getLabelCaption(form, stRecordID), ref SAPCustList, ref selSAPCust, true);
                       // SBO_Application.SetStatusBarMessage("Credit card for customer: " + getLabelCaption(form, stRecordID) + " loaded");
                    }
                }
                else
                    HideCC(form);
            }else
            {
                HideCC(form);
            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
    private void HideCC(SAPbouiCOM.Form form)
    {
        try
        {
            clearForm(form);
            setLabelCaption(form, stRecordID, "");
            form.Items.Item(tbCreditCard).Visible = false;
            form.PaneLevel = 0;
            form.Items.Item("3").Click();
  
        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
   
    private void handleItemPress(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
            switch (pVal.ItemUID)
            {
                case tbeBizCharge:
                    form.PaneLevel = paneBP;
                    break;
                case tbCreditCard:
                    form.PaneLevel = paneCCLog;
                    strWhere = string.Format(" Where CustomerID='{0}' ", getItemValue(form, "5"));
                    populateLog(form);  
                    break;
                case tbPaymentMethod:
                    form.PaneLevel = paneBP;
                    break;
                case "18":
                    loadCustData(form);
                    break;
                case btnAdd:
                    bAuto = false;
                    SAPCust sapcust = new SAPCust();
                    if(!loadFormData(form, ref sapcust, false))
                        return;
                    sapcust.cccust.CustomerID = getFormEditVal(form, "5");
                    sapcust.custObj.CustomerID = getFormEditVal(form, "5");
                    if (Validate(form, sapcust.custObj))
                    {
                        try
                        {
                           
                            if (getComboBoxVal(form, cbActive) == "Y")
                                sapcust.cccust.active = 'Y';
                            else
                                sapcust.cccust.active = 'N';
                            if (AddCustomer(ref sapcust))
                            {
                                SBO_Application.MessageBox("Credit card added.");
                                reload(form, getLabelCaption(form, stRecordID), ref SAPCustList, ref selSAPCust, true);
                            }else
                                SBO_Application.MessageBox("Failed to added credit card.");
                        }
                        catch (Exception ex)
                        {
                            SBO_Application.MessageBox("Failed to add your card: " + ex.Message);
                        }
                        
                    }
                    break;
                case btnDelete:
                    bAuto = false;
                    if (selSAPCust == null)
                    {                        
                        SBO_Application.MessageBox("Please select a credit card.");
                    }
                    else
                    {
                        bAuto = false;
                        if (SBO_Application.MessageBox("Are you sure you want to remove credit card " + selSAPCust.key + "?", 1, "Yes", "No") == 1)
                        {
                            deleteCustomer(selSAPCust.custObj.CustNum);
                            db.CCCUSTs.DeleteOnSubmit(selSAPCust.cccust);
                            db.SubmitChanges();
                            reload(form, getLabelCaption(form, stRecordID), ref SAPCustList, ref selSAPCust, true);

                            SBO_Application.MessageBox("Credit card removed.");
                        }
                    }
                    break;
                case btnUpdate:
                    {
                        bAuto = false;
                        if (selSAPCust == null)
                        {
                            SBO_Application.MessageBox("Please select a credit card.");
                            return;
                        }
                        if(!loadFormData(form, ref selSAPCust, false))
                            return;
                        if (selSAPCust.custObj.PaymentMethods[0].CardNumber.IndexOf("XXX") != -1 && selSAPCust.custObj.PaymentMethods[0].Account != null)
                        {
                            SBO_Application.MessageBox("Please enter unmasked credit card number.");
                            return;
                        }
                        if (Validate(form, selSAPCust.custObj))
                        {
                            if (updateCustomer(selSAPCust.custObj))
                            {
                                updateCustDeclined(selSAPCust.cccust.recID, "N");
                                selSAPCust.cccust.Declined = 'N';
                                reload(form, getLabelCaption(form, stRecordID), ref SAPCustList, ref selSAPCust, true);
                                SBO_Application.MessageBox("Credit card updated.");
                                
                            }
                            else
                            {
                                SBO_Application.MessageBox("Failed to update credit card.");
                            }
                        }
                    }
                    break;
            }

        }catch(Exception ex)
        {
            errorLog(ex);
        }

    }
    /*
    private void setActive(SAPbouiCOM.Form form,string id)
    {
        try {
            var q = from a in db.CCCUSTs
                    where a.CustomerID == getLabelCaption(form, stRecordID)
                    select a;
            if (q.Count() > 0)
            {
                foreach (CCCUST cust in q)
                {
                    if (cust.recID.ToString() == id) // selSAPCust.cccust.recID)
                        cust.active = 'Y';
                    else
                        cust.active = 'N';
                }
                db.SubmitChanges();
            }
        }
        catch(Exception ex)
        {
            errorLog(ex);
        }
    }*/
    private bool loadFormData(SAPbouiCOM.Form form, ref SAPCust cust, bool bPaymentForm)
    {
        try
        {
            handleSwipe(form);
            if (cust == null)
                cust = new SAPCust();
            if (cust.custObj == null)
            {
                cust.custObj = createCustObj();
            }
            if (cust.cccust == null)
                cust.cccust = new CCCUST();
            cust.custObj.BillingAddress.FirstName = getFirstName( getFormEditVal(form,editHolderName) );
            cust.custObj.BillingAddress.LastName = getLastName(getFormEditVal(form,editHolderName));
            cust.custObj.BillingAddress.Street = getFormEditVal(form,editHolderAddr);
            cust.custObj.BillingAddress.City = getFormEditVal(form,editHolderCity);
            cust.custObj.BillingAddress.State = getFormEditVal(form,editHolderState);
            cust.custObj.BillingAddress.Zip = getFormEditVal(form, editHolderZip);
            if(getFormEditVal(form, editCardNum) != "")
            {
                cust.custObj.PaymentMethods[0].MethodType="cc";
                cust.custObj.PaymentMethods[0].CardNumber = getFormEditVal(form, editCardNum);
                cust.custObj.PaymentMethods[0].CardExpiration = getFormEditVal(form, editCardExpDate);
                cust.custObj.PaymentMethods[0].MethodName = cust.custObj.BillingAddress.FirstName + " " + cust.custObj.BillingAddress.LastName;
           }
            if(getFormEditVal(form, editCheckAccount) != "")
            {
                cust.custObj.PaymentMethods[0].MethodType="check";
                cust.custObj.PaymentMethods[0].Account = getFormEditVal(form, editCheckAccount);
                cust.custObj.PaymentMethods[0].Routing = getFormEditVal(form, editCheckRouting);
            }
            if(bPaymentForm)
            {
                cust.SaveCard = getComboBoxVal(form, cbCCSave);
                if(cfgAutoSelectAccount == "N")
                {
                    cust.CCAccount = getComboBoxVal(form, cbCCAccount);
                    if (cust.CCAccount.IndexOf("-") >= 0)
                        cust.CCAccount = cust.CCAccount.Substring(0, cust.CCAccount.IndexOf("-")).Trim();

                }
                else
                {

                    cust.CCAccount = GetCreditCardCode(cust.custObj.PaymentMethods[0].CardNumber, !(getButtonCaption(form, "1") == "Credit")).ToString();
                    if (cust.CCAccount == "0")
                        return false;
                }
            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    private void reload(SAPbouiCOM.Form form, string customerID, ref List<SAPCust> list, ref SAPCust sel, bool bAll = false)
    {
        try
        {
           
            sel = new SAPCust();
            list = new List<SAPCust>();
            SAPbouiCOM.ComboBox oCB = form.Items.Item(cbPaymentMethod).Specific;
            try
            {
                while (oCB.ValidValues.Count > 0)
                    oCB.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception)
            { }
            oCB.ValidValues.Add("New Card", "New Card");
            oCB.ValidValues.Add("New eCheck", "New eCheck");
            var q = from a in db.CCCUSTs
                    where a.CustomerID == customerID
                    select a;
            if (q.Count() == 0)
            {
                int n = oCB.ValidValues.Count - 2;
                SBO_Application.SetStatusBarMessage("Number of record in combo= " + n.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                
                oCB.Select("New Card");
                populateCustFromDB(form);
                SBO_Application.SetStatusBarMessage("Credit card for customer: " + getLabelCaption(form, stRecordID) + " loaded. No credit card record found", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                return;
            }
            if (q.Count() > 0)
            {
            
                bool hasDefault = false;
                foreach(CCCUST cust in q)
                {
                    if (cust.active == 'Y' || bAll)
                    {
                        SAPCust sapcust = new SAPCust();
                        sapcust.cccust = cust;
                        string id = cust.CustomerID + "_" + cust.recID.ToString();
                        sapcust.custObj = getCustomer(id);
                        if (sapcust.custObj.PaymentMethods.Count() == 0)
                        {
                            db.CCCUSTs.DeleteOnSubmit(cust);
                            db.SubmitChanges();
                        }
                        else
                        {
                            if(sapcust.custObj.PaymentMethods[0].MethodType == "check")
                                sapcust.key = cust.recID.ToString() + "_" + sapcust.custObj.PaymentMethods[0].Routing + " " + sapcust.custObj.PaymentMethods[0].Account + "(" + sapcust.custObj.BillingAddress.FirstName + " " + sapcust.custObj.BillingAddress.LastName + ")";
                            else
                                sapcust.key = cust.recID.ToString() + "_" + sapcust.custObj.PaymentMethods[0].CardNumber + " " + sapcust.custObj.PaymentMethods[0].CardExpiration + "(" + sapcust.custObj.BillingAddress.FirstName + " " + sapcust.custObj.BillingAddress.LastName + ")";
                            if (cust.Declined == 'Y')
                                sapcust.key = sapcust.key + ", Declined";
                            SAPbouiCOM.ValidValue v = oCB.ValidValues.Add(sapcust.key, sapcust.key);
                            list.Add(sapcust);
                            if (!hasDefault)
                                sel = sapcust;
                            if (cust.@default == 'Y')
                            {
                                hasDefault = true;
                                sel = sapcust;
                            }
                            

                        }
                    }
                }
            }
            if (sel.key != null)
                oCB.Select(sel.key);
            populateCCInfo(form, sel);
            SBO_Application.SetStatusBarMessage("Credit card for customer: " + getLabelCaption(form, stRecordID) + " loaded. " + oCB.ValidValues.Count + " credit card record(s) found", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
        }catch(Exception ex)
        {
             errorLog(ex);
        }
   
    }
    private void HandleSelect(SAPbouiCOM.Form form)
    {
        if (getComboBoxVal(form, cbPaymentMethod) == "New Card" || getComboBoxVal(form, cbPaymentMethod) == "New eCheck")
        {
            populateCustFormFromDB(form);
        }
        else
        {}
            selSAPCust = getCCInfoSelect(form, SAPCustList);
            populateCCInfo(form, selSAPCust);
        }
    }
    private void HandleActiveSelect(SAPbouiCOM.Form form)
    {
        try
        {
            if (selSAPCust == null)
                return;
            string active = getComboBoxVal(form, cbActive);
            if (active != selSAPCust.cccust.active.ToString())
            {
                if (updateCustState(selSAPCust.cccust.recID, active))
                {
                    if (active == "Y")
                        SBO_Application.MessageBox("Card activated.");
                    else
                        SBO_Application.MessageBox("Card deactivated.");
                    selSAPCust.cccust.active = active[0];
                    foreach(SAPCust cust in SAPCustList)
                    {
                        if(cust.cccust.recID == selSAPCust.cccust.recID)
                            cust.cccust.active = active[0];
                    }
                }
                else
                {
                    SBO_Application.MessageBox("Failed to update card status.");
                }
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        
    }
    private void populateCustFromDB(SAPbouiCOM.Form form)
    {
        try
        {

            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sql = string.Format("Select a.CardName, b.Street,b.City,b.State,b.ZipCode from ocrd a, CRD1 b where a.CardCode = b.CardCode and a.BillToDef = b.Address and cardcode = '{0}'", getLabelCaption(form, stRecordID));
            oRS.DoQuery(sql);
            if(!oRS.EoF)
            {
                setFormEditVal(form, editHolderName, GetFieldVal(oRS, 0));
                setFormEditVal(form, editHolderAddr, GetFieldVal(oRS, 1));
                setFormEditVal(form, editHolderCity, GetFieldVal(oRS, 2));
                setFormEditVal(form, editHolderState, GetFieldVal(oRS, 3));
                setFormEditVal(form, editHolderZip, GetFieldVal(oRS, 4));
                setFormEditVal(form, editCardNum, "");
                setFormEditVal(form, editCardExpDate, "");
                setFormEditVal(form, editCardCode, "");
                setFormEditVal(form, editCheckAccount, "");
                setFormEditVal(form, editCheckRouting, "");
            }
           
             
        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
    private void populateCCInfo(SAPbouiCOM.Form form, SAPCust cust)
    {
        try {
            if(cust == null)
            {
                populateCustFromDB(form);
                return;
            }
            if (cust.custObj != null)
            {
                cust.cardHolder = cust.custObj.BillingAddress.FirstName + " " + cust.custObj.BillingAddress.LastName;
                cust.cardNum = cust.custObj.PaymentMethods[0].CardNumber;
                setFormEditVal(form, editHolderName, cust.custObj.BillingAddress.FirstName + " " + cust.custObj.BillingAddress.LastName);
                setFormEditVal(form, editHolderAddr, cust.custObj.BillingAddress.Street + " " + cust.custObj.BillingAddress.Street2);
                setFormEditVal(form, editHolderCity, cust.custObj.BillingAddress.City);
                setFormEditVal(form, editHolderState, cust.custObj.BillingAddress.State);
                setFormEditVal(form, editHolderZip, cust.custObj.BillingAddress.Zip);
                setFormEditVal(form, editCardNum, cust.custObj.PaymentMethods[0].CardNumber);
                setFormEditVal(form, editCardExpDate, cust.custObj.PaymentMethods[0].CardExpiration);
                setFormEditVal(form, editCardCode, cust.custObj.PaymentMethods[0].CardCode);
                setFormEditVal(form, editCheckAccount, cust.custObj.PaymentMethods[0].Account);
                setFormEditVal(form, editCheckRouting, cust.custObj.PaymentMethods[0].Routing);

                if (form.TypeEx == "134")
                {
                    setComboValue(form, cbActive, cust.cccust.active.ToString());
                }
                if (cust.cccust.Declined == 'Y')
                    setLabelCaption(form, lbDeclined, "Delined");
                else
                    setLabelCaption(form, lbDeclined, "");
   
                return;
            }
           // populateCustFromDB(form);
            

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void clearForm(SAPbouiCOM.Form form)
    {
        try
        {
            SAPbouiCOM.ComboBox oCB = form.Items.Item(cbPaymentMethod).Specific;
            try
            {
                while (oCB.ValidValues.Count > 0)
                    oCB.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
         
            }catch(Exception)
            {

            }
            oCB.ValidValues.Add("New Card", "New Card");
            setFormEditVal(form, editHolderName, "");
            setFormEditVal(form, editHolderAddr, "");
            setFormEditVal(form, editHolderCity, "");
            setFormEditVal(form, editHolderState, "");
            setFormEditVal(form, editHolderZip, "");
            setFormEditVal(form, editCardNum, "");
            setFormEditVal(form, editCardExpDate, "");
            setComboValue(form, cbActive, "Y");
        }catch(Exception ex)
        {
            errorLog(ex);
        }
    }
    private SAPCust getCCInfoSelect(SAPbouiCOM.Form form, List<SAPCust> l)
    {
        try
        {
            string sel = getComboBoxVal(form, cbPaymentMethod);
            foreach (SAPCust c in l)
            {
                if (c.key == sel)
                    return c;
            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return null;
    }
    private void AddCCFormField(SAPbouiCOM.Form form, int pane)
    {
        try
        {
         
            int edL = 110 + 15 ; //oItm.Left + oItm.Width;
            int edW = 250;
            int edT = 178;
            int edH = 15;
            int nGap = 17;
            SAPbouiCOM.Item oItm = addPaneItem(form, stRecordID, edL, edT, edW, edH, "Customer ID", SAPbouiCOM.BoFormItemTypes.it_STATIC, pane, 0);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, cbPaymentMethod, edL, edT, edW, edH, "Payment Method:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, pane, 1);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCardNum, edL, edT, edW, edH, "*Card Number:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 2);
            oItm = form.Items.Add(lbDeclined, SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItm.FromPane = pane;
            oItm.ToPane = pane;
            oItm.Left = 400;
            oItm.Top = edT;
            oItm.Width = 100;
            oItm.Height = edH;
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCardExpDate, edL, edT, edW, edH, "*Exp. Date(MMYY):", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 3);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCardCode, edL, edT, edW, edH, "Card Code:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 4);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderName, edL, edT, edW, edH, "*Card/Acct. Holder:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 6);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCheckRouting, edL, edT, edW, edH, "*Routing Number:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 101);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editCheckAccount, edL, edT, edW, edH, "*Checking Account:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 100);
            edT = oItm.Top + nGap + 10;
            oItm = addPaneItem(form, editHolderAddr, edL, edT, edW, edH, "Address:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 7);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderCity, edL, edT, edW, edH, "City:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 8);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderState, edL, edT, edW, edH, "State:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 9);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editHolderZip, edL, edT, edW, edH, "Zip:", SAPbouiCOM.BoFormItemTypes.it_EDIT, pane, 10);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, cbActive, edL, edT, 60, edH, "Active:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, pane, 11);
            SAPbouiCOM.ComboBox oCB = form.Items.Item(cbActive).Specific;
            oCB.ValidValues.Add("Y", "Y");
            oCB.ValidValues.Add("N", "N");
           

            edT = oItm.Top + 80;
            edL = oItm.Left - 20;
            edH = 20;
            int btnW = 60;
            nGap = 65;
            oItm = addPaneButton(form, btnAdd, edL, edT, btnW, edH, "Add", pane);
            edL = oItm.Left + nGap;
            oItm = addPaneButton(form, btnDelete, edL, edT, btnW, edH, "Delete", pane);
            edL = oItm.Left + nGap;
            oItm = addPaneButton(form, btnUpdate, edL, edT, btnW, edH, "Update", pane);
           
        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
    }
    
    private CustomerObject getDefaultCustObj(SAPbouiCOM.Form form)
    {
         CustomerObject CustomerData = new CustomerObject();
        CustomerData.BillingAddress = new Address();
      
        try
        {
            
            SAPbobsCOM.Recordset oCustomerRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oCustomerRS.DoQuery(string.Format("Select CardName,Address,City,State1,ZipCode, MailAddres,MailCity,State2,MailZipCod from ocrd where cardcode = '{0}'", getLabelCaption(form, stRecordID)));
            CustomerData.BillingAddress.FirstName = getFirstName(GetFieldVal(oCustomerRS, 0));
            CustomerData.BillingAddress.LastName = getLastName(GetFieldVal(oCustomerRS, 0));
            CustomerData.BillingAddress.Street = GetFieldVal(oCustomerRS, 1);
            CustomerData.BillingAddress.City = GetFieldVal(oCustomerRS, 2);
            CustomerData.BillingAddress.State = GetFieldVal(oCustomerRS, 3);
            CustomerData.BillingAddress.Zip = GetFieldVal(oCustomerRS, 4);
            CustomerData.PaymentMethods = new PaymentMethod[1];
            CustomerData.PaymentMethods[0] = new PaymentMethod();
            CustomerData.PaymentMethods[0].CardNumber = getFormEditVal(form, "154");
            CustomerData.PaymentMethods[0].CardExpiration = getFormEditVal(form, "155");
            CustomerData.PaymentMethods[0].CardCode = "";
          
            
        }catch(Exception)
        {

        }
        return CustomerData;
    }
    private CustomerObject createCustObj()
    {
        CustomerObject CustomerData = new CustomerObject();
        CustomerData.CustNum = "";
        CustomerData.BillingAddress = new Address();
        CustomerData.PaymentMethods = new PaymentMethod[1];
        CustomerData.PaymentMethods[0] = new PaymentMethod();
        return CustomerData;
    }
    private bool Validate(SAPbouiCOM.Form form, CustomerObject cusObj)
    {
        string err = "";
        try
        {
            if (cusObj.PaymentMethods[0].MethodType == "cc" || cusObj.PaymentMethods[0].MethodType == "" || cusObj.PaymentMethods[0].MethodType == null)
            {
                if (cusObj.PaymentMethods[0].CardNumber == "")
                {
                    err += "Credit card number is required.\r\n";

                }
                if (cusObj.PaymentMethods[0].CardExpiration == "")
                {
                    err += "Card expiration date is required.\r\n";
                }
                if (cusObj.BillingAddress.FirstName == "")
                {
                    err += "Card holder name is required.\r\n";
                }

            }
            if (cusObj.PaymentMethods[0].MethodType == "check")
            {
                if (cusObj.PaymentMethods[0].Account == "")
                {
                    err += "Checking account number is required.\r\n";

                }
                if (cusObj.PaymentMethods[0].Routing == "")
                {
                    err += "Routing number is required.\r\n";
                }
                if (cusObj.BillingAddress.FirstName == "")
                {
                    err += "Account holder name is required.\r\n";
                }
            }
            if(err == "")
                return true;
            else
                SBO_Application.MessageBox(err);

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    
}
    
