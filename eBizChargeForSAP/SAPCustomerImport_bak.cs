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

partial class SAP
{
    const string menuCustImp = "eBizCustImpMenu";
    const string formCustImp = "eBizCustImpForm";
    const string matrixCustImp = "mBInv";
    const string btnImp = "btnImp";
    const string editCustNum = "edCustNum";
    CustomerObject custObjForImport = null;
    

    SAPbouiCOM.Form oCustImpForm;
   
    private void AddCustImpMenu()
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
            oCreationPackage.UniqueID = menuCustImp;
            oCreationPackage.String = "eBizCharge Payment Method Import";
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
    private void CreateCustImpForm()
    {
        try
        {
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Button oButton = null;
          
            // add a new form
            SAPbouiCOM.FormCreationParams oCreationParams = null;

            oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.FormType = formCustImp;

            oCreationParams.UniqueID = formCustImp;
            try
            {
                oCustImpForm = SBO_Application.Forms.AddEx(oCreationParams);
            }
            catch (Exception)
            {
                oCustImpForm = SBO_Application.Forms.Item(formCustImp);

            }

            // set the form properties
            oCustImpForm.Title = "eBizCharge Payment Method Import";
            oCustImpForm.Left = 400;
            oCustImpForm.Top = 100;
            oCustImpForm.ClientHeight = 260;
            oCustImpForm.ClientWidth = 450;





          
            int nTopGap = 25;
            int left = 6;
            int wBtn = 70;
            int hBtn = 19;
            int span = 80;
          
            oItem = oCustImpForm.Items.Add(btnFind, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oCustImpForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Find";
            left += span;
            oItem = oCustImpForm.Items.Add(btnImp, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oCustImpForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Import";
            left += span;
            oItem = oCustImpForm.Items.Add(btnClose, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oCustImpForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";

            int margin = 8;
            int top = 15;
            int edL = 150; //oItm.Left + oItm.Width;
            int edW = 200;
            int edH = 15;
            int nGap = 26;

            oItem = addPaneItem(oCustImpForm, editCustNum, edL, top, edW, edH, "Gateway Cardholder:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 100);
           

            top = oItem.Top + nGap;

            oItem = addPaneItem(oCustImpForm, cbPaymentMethod, edL, top, edW, edH, "Payment Method:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 200);

            top = oItem.Top + nGap;

            oItem = addPaneItem(oCustImpForm, cbCustomerID, edL, top, edW, edH, "Import to:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 201);

            List<String> list = FindCustomer("");
           
                SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)oCustImpForm.Items.Item(cbCustomerID).Specific;
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
        oCustImpForm.Visible = true;

    }
    private void CustImpFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {
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
                                    oCustImpForm.Close();
                                    break;
                                case btnImp:
                                    try
                                    {
                                        if (custObjForImport == null)
                                        {
                                            form.Items.Item(btnFind).Click();
                                        }
                                        else
                                        {
                                            string cardcode = getComboBoxVal(form, cbCustomerID);
                                            if (cardcode == "")
                                            {
                                                showMessage("Please select a customer.");
                                            }
                                            else
                                            {
                                                if (custObjForImport == null)
                                                    form.Items.Item(btnFind).Click();
                                                else
                                                {
                                                    if (SBO_Application.MessageBox("Import payment method(s) to " + cardcode + "?", 1, "Yes", "No") == 1)
                                                    {

                                                        int i = 0;
                                                        foreach (PaymentMethod pm in custObjForImport.PaymentMethods)
                                                        {
                                                            if (!isPaymentMethodExists(cardcode, pm.CardNumber, pm.CardExpiration))
                                                            {
                                                                CCCUST cccust = new CCCUST();
                                                                cccust.CustomerID = cardcode;
                                                                cccust.active = 'Y';
                                                                cccust.email = custObjForImport.BillingAddress.Email;
                                                                cccust.firstName = custObjForImport.BillingAddress.FirstName;
                                                                cccust.lastName = custObjForImport.BillingAddress.LastName;
                                                                cccust.street = custObjForImport.BillingAddress.Street;
                                                                cccust.city = custObjForImport.BillingAddress.City;
                                                                cccust.state = custObjForImport.BillingAddress.State;
                                                                cccust.zip = custObjForImport.BillingAddress.Zip;
                                                                cccust.expDate = pm.CardExpiration;
                                                                cccust.cardCode = pm.CardCode;

                                                                cccust.routingNumber = pm.Routing;
                                                                cccust.checkingAccount = pm.Account;
                                                                cccust.GroupName = getGroupName(cardcode);
                                                                cccust.CustNum = custObjForImport.CustNum;
                                                                cccust.MethodID = pm.MethodID;
                                                                try
                                                                {
                                                                    if (pm.MethodName != null && pm.MethodName != "")
                                                                    {
                                                                        if (pm.MethodName.IndexOf(" ") > 0)
                                                                        {
                                                                            string name = pm.MethodName;
                                                                            string[] s = name.Split(' ');

                                                                            cccust.firstName = s[0];
                                                                            cccust.lastName = s[s.Length - 1];
                                                                        }
                                                                    }
                                                                }
                                                                catch (Exception) { }
                                                                int id = getNextTableID("@CCCUST");
                                                                string desc = cccust.firstName + " " + cccust.lastName;
                                                                if (pm.MethodName != null && pm.MethodName != "")
                                                                    desc = pm.MethodName;
                                                                if (pm.MethodType == "check")
                                                                {
                                                                    cccust.methodDescription = id.ToString() + "_" + pm.Routing + " " + pm.Account + "(" + desc + ")";
                                                                    cccust.checkingAccount = pm.Account;
                                                                    cccust.routingNumber = pm.Routing;
                                                                }
                                                                else
                                                                {
                                                                    cccust.cardNum = pm.CardNumber;
                                                                    cccust.methodDescription = id.ToString() + "_" + pm.CardNumber + " " + pm.CardExpiration + "(" + desc + ")";
                                                                    cccust.checkingAccount = pm.Account;
                                                                    cccust.cardType = pm.CardType;
                                                                }
                                                                string cardid = "";
                                                                cccust.CardName = getCardName(cccust.GroupName, cccust.cardType, "", ref cardid);
                                                                cccust.CCAccountID = cardid;
                                                                cccust.methodName = pm.MethodName;
                                                                insert(cccust);
                                                                SBO_Application.SetStatusBarMessage(string.Format("Import payment method {0} to {1}.", cccust.cardNum, cardcode), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                                                            }
                                                            i++;
                                                        }
                                                    }
                                                }

                                            }
                                        }

                                    }catch(Exception ex)
                                    {
                                        SBO_Application.SetStatusBarMessage(string.Format("Import payment method error: {0}.", ex.Message), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
 
                                
                                    }
                                    break;
                              
                                case btnFind:
                                    try
                                    {
                                        string custNum = getFormEditVal(form, editCustNum);
                                        if(custNum == "")
                                        {
                                            showMessage("Please enter customer number.");
                                        }else
                                        {
                                            SBO_Application.SetStatusBarMessage("Finding customer please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                               
                                            ueSecurityToken token = getToken("");
                                            custObjForImport = ebiz.getCustomer(token, custNum);
                                            SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbPaymentMethod).Specific;
                                            try
                                            {
                                                while (oCB.ValidValues.Count > 0)
                                                    oCB.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                            }
                                            catch (Exception)
                                            { }
                                            ComboAddItem(oCB, "");
                                            foreach (PaymentMethod pm in custObjForImport.PaymentMethods)
                                            {
                                                string s = "";
                                                string name = "";
                                                name = custObjForImport.BillingAddress.FirstName + " " + custObjForImport.BillingAddress.LastName;
                                                if (pm.MethodName != null && pm.MethodName != "")
                                                    if(pm.MethodName.IndexOf(" ") > 0)
                                                        name = pm.MethodName;
                                                if (pm.MethodType == "check")
                                                {
                                                    s = pm.Routing + " " + pm.Account + "(" + name + ")";
                                                }
                                                else
                                                {
                                                    s = pm.CardNumber + " " + pm.CardExpiration + "(" + name + ")";
                                                }
                                                ComboAddItem(oCB, s);
                                            }
                                            oCB.Select("");
                                            SBO_Application.SetStatusBarMessage(string.Format("{0} payment method(s) ready to import.", oCB.ValidValues.Count - 1), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                           // form.Items.Item(btnImp).Click();
                                        }
                                    }catch(Exception ex)
                                    {
                                        custObjForImport = null;
                                        SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)form.Items.Item(cbPaymentMethod).Specific;
                                        try
                                        {
                                            while (oCB.ValidValues.Count > 0)
                                                oCB.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                        }
                                        catch (Exception)
                                        { }
                                        ComboAddItem(oCB, "");
                                        oCB.Select("");
                                        SBO_Application.SetStatusBarMessage(string.Format("Find customer error: {0}.", ex.Message), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
 
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
    
}

