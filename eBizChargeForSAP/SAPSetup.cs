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
using EBizEMV;
using Newtonsoft.Json;
using eBizChargeForSAP.ServiceReference1;

partial class SAP
{

    SAPbouiCOM.Form oSetupForm;
    const string formSetup = "eBizSetupForm";
 

    const string editDBUser = "edDBUsr";
    const string editDBPasswd = "edDBPWD";
    const string editAppUser = "edAPUsr";
    const string editAppPasswd = "edAPPWD";
    const string menuConnSetup = "mConSet";
    const string passwdMask = "******";
    private void AddConnectionSetupMenu()
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
            oCreationPackage.UniqueID = menuConnSetup;
            oCreationPackage.String = "eBizCharge Connection setup";
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
    private void CreateSetupForm()
    {
        try
        {

            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Button oButton = null;
            //SAPbouiCOM.StaticText oStaticText = null;
            // SAPbouiCOM.EditText oEditText = null;
            //SAPbouiCOM.ComboBox oComboBox = null;

            // add a new form
            SAPbouiCOM.FormCreationParams oCreationParams = null;

            oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.FormType = formSetup;

            oCreationParams.UniqueID = formSetup;
            try
            {
                oSetupForm = SBO_Application.Forms.AddEx(oCreationParams);
            }
            catch (Exception)
            {
                oSetupForm = SBO_Application.Forms.Item(formSetup);

                oSetupForm.Visible = true;

            }
            trace("Create Setup Form. ");

            // set the form properties
            oSetupForm.Title = "eBizCharge Setup";
            oSetupForm.Left = 400;
            oSetupForm.Top = 100;
            oSetupForm.ClientHeight = 460;
            oSetupForm.ClientWidth = 450;

            //************************
            // Adding a Rectangle
            //***********************
            int margin = 5;
            oItem = oSetupForm.Items.Add("Rect1", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = margin;
            oItem.Width = oSetupForm.ClientWidth - 2 * margin;
            oItem.Top = margin;
            oItem.Height = oSetupForm.ClientHeight - 40;

            int nTopGap = 25;
            int left = 6;
            int wBtn = 70;
            int hBtn = 19;
            int span = 80;

            oItem = oSetupForm.Items.Add(btnProcess, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oSetupForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));
            oButton.Caption = "Update";
            left = oItem.Left + wBtn + 10;
            oItem = oSetupForm.Items.Add(btnClose, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = wBtn;
            oItem.Top = oSetupForm.ClientHeight - nTopGap;
            oItem.Height = hBtn;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";

            int edL = 110 + 15; //oItm.Left + oItm.Width;
            int edW = 250;
            int edT = 30;
            int edH = 15;
            int nGap = 22;
            int lbw = 100;
         
            oItem = addPaneItem(oSetupForm, editDBUser, edL, edT, edW, edH, "DB user name:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oSetupForm, editDBPasswd, edL, edT, edW, edH, "DB user password:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oSetupForm, editAppUser, edL, edT, edW, edH, "B1 user name:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 3, lbw);
            edT = oItem.Top + nGap;
            oItem = addPaneItem(oSetupForm, editAppPasswd, edL, edT, edW, edH, "B1 user password:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 4, lbw);
            

            oSetupForm.Visible = true;
            if (B1Info.AppUser == null || B1Info.AppUser == "")
                B1Info.AppUser = oCompany.UserName;
            setFormEditVal(oSetupForm, editAppUser, B1Info.AppUser);
            setFormEditVal(oSetupForm, editAppPasswd, passwdMask);
            setFormEditVal(oSetupForm, editDBUser, B1Info.DBUser);
            setFormEditVal(oSetupForm, editDBPasswd, passwdMask);
        }
        catch (Exception ex)
        {

            errorLog(ex);
        }

        oSetupForm.Visible = true;
    }
   
    private void SetupFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
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

                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        {
                            switch (pVal.ItemUID)
                            {
                              
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        {
                            switch (pVal.ItemUID)
                            {
                                case btnClose:
                                    
                                    oSetupForm.Close();
                                  
                                    break;
                                case btnProcess:
                                    HandleSetupUpdate(form);
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

    private void HandleSetupUpdate(SAPbouiCOM.Form form)
    {
        //errorLog("Handle select");
        try
        {
            string dbPassword = B1Info.DBPassWD;
            string appPassword = B1Info.AppPassWD;
            FormFreeze(form, true);
            if(getFormItemVal(form, editDBPasswd) != passwdMask)
                dbPassword = getFormItemVal(form, editDBPasswd);
            if (getFormItemVal(form, editAppPasswd) != passwdMask)
                appPassword = getFormItemVal(form, editAppPasswd);
            upsertSetupInfo(getFormItemVal(form, editDBUser), dbPassword, getFormItemVal(form, editAppUser), appPassword);
             
            tProcess = new Thread(appInit);
            tProcess.Start();
        }
        catch (Exception ex)
        {
            showMessage(ex.Message);
            errorLog(ex);
        }
        finally
        {
            try
            {

               FormFreeze(form,false);
                form.Close();
            }
            catch (Exception) { }
        }
    }
  
}

