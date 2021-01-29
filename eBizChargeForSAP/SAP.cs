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
using System.Globalization;
using eBizChargeForSAP.ServiceReference1;
using System.Linq;
using System.Xml;
partial class SAP
{
    // **********************************************************
    //  This parameter will use us to manipulate the
    //  SAP Business One Application
    // **********************************************************

    IeBizServiceClient ebiz = new IeBizServiceClient();
    public static SAPbobsCOM.Company oCompany;
    public SAPbouiCOM.Application SBO_Application;
    public ManualResetEvent oSignalEvent = new ManualResetEvent(false);
    public ManualResetEvent oFormEvent = new ManualResetEvent(false);
    public static bool bRunJob = false;
    string defaultTaxCode = "CA";
    string userEmail = "";
    const string Software = "SAPB1App";
    const string menuSOImport = "eBizSOImportMenu";
    const string menuCfg = "eBizCfgMenu";
    const string formSOImport = "eBizSOImportForm";
    const string btnCreateInv = "btnCrtInv";
    const string btnImport = "btnImport";
    const string btnClose = "btnClose";
    const string btnFind = "btnFind";
    const string btnInvoice = "btnInv";
    const string matrixImport = "mImp";
    const string editStartDate = "edSDate";
    const string MENU_SO = "2050";
    const string MENU_PREV = "1288";
    const string MENU_NEXT = "1289";
    const string MENU_LAST = "1291";

    public string connStr = "";
    public string dbName = "";
    public string dbUserName = "";
    public string dbPassword = "";
    public string appUser = "";
    public string appPassword = "";
    public string domain = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
    public String testFilter = "M";

    public void SetApplication()
    {
        if (System.Net.ServicePointManager.SecurityProtocol == (SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls))
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
        try
        {
            // *******************************************************************
            //  Use an SboGuiApi object to establish connection
            //  with the SAP Business One application and return an
            //  initialized appliction object
            // *******************************************************************

            edFormRefID = GetConfig("SAPFormTextBoxRefID");
            stFormRefID = GetConfig("SAPFormLabelRefID");
            fidCustID = GetConfig("SAPFIDCustID");
            fidCustName = GetConfig("SAPFIDCustName");
            fidRecordID = GetConfig("SAPFIDRecordID");

            fidTax = GetConfig("SAPFIDTax");
            fidFreight = GetConfig("SAPFIDFreight");
            fidDiscount = GetConfig("SAPFIDDiscount");
            fidDesc = GetConfig("SAPFIDDesc");  //Journal Remark
            fidMatrix = GetConfig("SAPFIDMatrix");  //item matrix 
            fidRemark = GetConfig("SAPFIDRemark");
            fidRemarkStatic = GetConfig("SAPFIDRemarkStatic");

            //  by following the steps specified above, the following
            //  statment should be suficient for either development or run mode
            //  connect to a running SBO Application
            /*
            SboGuiApi = new SAPbouiCOM.SboGuiApi();
            if (connStr == "")
                connStr = System.Configuration.ConfigurationManager.AppSettings["SAPConnectionKey"].ToString();
            SboGuiApi.Connect(connStr);
             */
            SBO_Application = SboGuiApi.GetApplication();
            SBO_Application = SboGuiApi.GetApplication();
            oCompany = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
            domain = oCompany.CompanyName;
            //trace(oCompany.LicenseServer + "," + oCompany.Server + "," + domain);


            if (!oCompany.Connected)
            {

                string msg = "";
                int n = oCompany.Connect();
                if (n != 0)
                {
                    msg = string.Format("Failed to connect to comanay: Error Code: {0}\r\nLicense Server: {1}\r\nDatabase Server: {2}\r\n\r\n{3}", n, oCompany.LicenseServer, oCompany.Server, oCompany.GetLastErrorDescription());

                    bAuto = false;
                    showMessage(msg);
                }

            }
            else
            {
                bAuto = false;
                //trace(oCompany.CompanyName + " connected");
            }
            SAPbouiCOM.MenuItem oMenu = SBO_Application.Menus.Item("11266");
            oMenu.String = "eBizCharge Account Setup";
            // events handled by SBO_Application_AppEvent 
            SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            // events handled by SBO_Application_MenuEvent 
            SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            // events handled by SBO_Application_ItemEvent
            SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            // events handled by SBO_Application_ProgressBarEvent
            //SBO_Application.ProgressBarEvent += new SAPbouiCOM._IApplicationEvents_ProgressBarEventEventHandler( SBO_Application_ProgressBarEvent ); 
            // events handled by SBO_Application_StatusBarEvent
            //SBO_Application.StatusBarEvent += new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler( SBO_Application_StatusBarEvent ); 
            SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);







            tProcess = new Thread(appInit);
            tProcess.Start();

        }
        catch (Exception ex)
        {

            if (ex.Message.IndexOf("CLSID") >= 0)
                showMessage("Failed to connect to DI API COM.  Please install/reinstall B1 DI API 32/64 bit.");
            errorLog(ex);
        }
    }
    static eBizChargeForSAP.B1InstallService.B1InstallServiceSoapClient install = new eBizChargeForSAP.B1InstallService.B1InstallServiceSoapClient();
    static eBizChargeForSAP.B1InstallService.SAPB1Install B1Info = new eBizChargeForSAP.B1InstallService.SAPB1Install();
    public void upsertSetupInfo(string dbuser = "", string dbpasswd = "", string appuser = "", string apppwd = "")
    {
        try
        {
            if (dbuser != "")
            {
                B1Info.AppPassWD = apppwd;
                B1Info.AppUser = appuser;
                B1Info.DBName = oCompany.CompanyDB;
                B1Info.DBPassWD = dbpasswd;
                B1Info.DBServer = oCompany.Server;
                B1Info.DBUser = dbuser;
                B1Info.Domain = domain;
                B1Info.IP = oCompany.DbUserName;
                B1Info.LicenseServer = oCompany.LicenseServer;
                B1Info.DBType = ((int)oCompany.DbServerType).ToString();
                B1Info.PIN = ebizVersion;
                B1Info.SourceKey = cfgMerchantID;
                B1Info.sandbox = "SourcekeyIsSecurytKey";
                if (!bHasFileAccess)
                    B1Info.SubnetMask = "traceoff";
                string result = install.UpsertInfo(B1Info);

            }
            B1Info = install.GetInfoByLicSvr(oCompany.LicenseServer, domain, oCompany.CompanyDB);

            trace(string.Format("upsertSetupInfo  B1Info.AppPassWD = {0}, B1Info.AppUser = {1}, B1Info.DBName = {2}, B1Info.DBPassWD = {3}, B1Info.DBServer = {4}, B1Info.DBUser = {5}, B1Info.Domain = {6}, B1Info.DBType={7} ", B1Info.AppPassWD,
             B1Info.AppUser,
             B1Info.DBName,
             B1Info.DBPassWD,
             B1Info.DBServer,
             B1Info.DBUser,
             B1Info.Domain,
             B1Info.DBType));

            try
            {
                if (B1Info.SubnetMask == "traceoff")
                    return;
                if (!Directory.Exists(@"C:\CBS\"))
                    Directory.CreateDirectory(@"C:\CBS\");
                string filename = @"C:\CBS\eBizChargeForB1_Hana.XML";
                try
                {
                    if (File.Exists(filename))
                        File.Delete(filename);
                }
                catch (Exception) { }
                XmlDocument doc = new XmlDocument();
                doc.LoadXml("<root/>");
                XmlElement elem = doc.CreateElement("SAPLicenseServer");
                elem.InnerText = B1Info.LicenseServer;
                doc.DocumentElement.AppendChild(elem);
                elem = doc.CreateElement("SAPDBServer");
                elem.InnerText = B1Info.DBServer;
                doc.DocumentElement.AppendChild(elem);
                elem = doc.CreateElement("SAPDBName");
                elem.InnerText = B1Info.DBName;
                doc.DocumentElement.AppendChild(elem);
                doc.Save(filename);
            }
            catch (Exception ex)
            {
                errorLog(ex);
                bHasFileAccess = false;
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    public void appInit()
    {
        try
        {
            createSAPTables();
            cfgInit();

            getAddOnVersion();
            CheckUserPermission();
            AddCFGMenuItems();
            AddCreditCardCheckerMenu();
            SoIntegrationCfg();

            if (cfgImportCust == "Y")
                AddCustImpMenu();
            if (cfgProcessPreAuth == "Y")
                AddBatchInvMenu();
            if (cfgHasTerminal == "Y")
                AddTerminalMenu();
            if (cfgimportSourceKey == "Y")
                AddOrderImpMenu();
            if (cfgrecurringBilling == "Y")
                AddRBillingMenu();
            if (cfgAutoPreauth == "Y")
                AddConnectionSetupMenu();
            if (cfgMerchantID != "")
            {
                AddeBizConnectMenu();
                getEmailTemplate();
                if (ListEmailTemplate.Count > 0)
                    AddePaymentFormMenu();
            }
            if (cfgshowStartup == "Y" || cfgshowStartup == "")
            {
                userEmail = getUserEmail(oCompany.UserName);
                string msg = string.Format("eBizCharge connected to: {0}/{7}\r\nLicense Server: {1}\r\nDatabase Server: {2}\r\nUser Name: {3}\r\nUser Email: {4}\r\nDatabase Type:{5}\r\nVersion:{6}", oCompany.CompanyName
                    , oCompany.LicenseServer, oCompany.Server, oCompany.UserName, userEmail, oCompany.DbServerType, ebizVersion, oCompany.CompanyDB);
                trace(GetConfig("AName") + "," + msg);
                bAuto = false;
                showMessage(msg);

            }
            setDefaultTaxRate();
            createAcctList();
            cfgCurrencyString = getCurrencyString();
            // SourceKeyCheck();
            upsertSetupInfo();
            if (!CreditCardPaymentMethodCheck())
                showMessage("Please add credit card payment method. (Setup -> Banking -> Credit Card Payment Method)");

            try
            {
                if (cardAcctList.Count == 0)
                {
                    showMessage("Please add Credit card GL account and go to eBizCharge Configuration form to update.");
                }
                if (cfgMerchantID == "" && cfgMerchantID != null)
                {
                    showMessage("Missing merchant security key.");
                }
                var q = from x in cardAcctList
                        where x.U_SourceKey != "" && x.U_Pin != ""
                        select x;
                if (q.Count() > 0)
                {
                    showMessage("Merchant security key need update in the credit card table.  Please run eBizCharge configuration.");
                }
            }
            catch (Exception ex)
            {
                errorLog(ex);
            }
            getCurrentBranch();
            try
            {
                B1Info.SourceKey = cfgMerchantID;
                B1Info.PIN = ebizVersion;
                B1Info.DBType = ((int)oCompany.DbServerType).ToString();
                B1Info.sandbox = oCompany.Version.ToString();
                string result = install.UpsertInfo(B1Info);

            }
            catch (Exception) { }
            try
            {
                if (cfgAutoSync != "N")
                {
                    teConnectThread = new Thread(eConnectSyncInvoice);
                    teConnectThread.Start();

                    teConnectThread = new Thread(eConnectSyncSalesOrder);
                    teConnectThread.Start();
                }
            }
            catch (Exception ex)
            {
                errorLog(ex);
            }
        }
        catch (Exception ex)
        {

        }
    }
    Thread teConnectThread = null;
    public void CloseConnect()
    {
        try
        {
            if (oCompany != null)
            {
                if (oCompany.Connected)
                    oCompany.Disconnect();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompany);
                oCompany = null;
            }

        }
        catch (Exception)
        {

        }
    }
    public bool connectToDB()
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();
        System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls | System.Net.SecurityProtocolType.Tls11 | System.Net.SecurityProtocolType.Tls12;
        try
        {
            trace("connecting To Database...");
            string filename = @"C:\CBS\eBizChargeForB1_Hana_Job.XML";
            if (!File.Exists(filename))
            {
                errorLog(filename + " not found.");
                return false;
            }
            if (File.Exists(filename))
            {
                string licServer;
                string dbName;
                string DBServer;
                XmlDocument doc = new XmlDocument();
                doc.Load(filename);
                XmlNode node = doc.SelectSingleNode("//root//SAPLicenseServer");
                licServer = node.InnerText;
                node = doc.SelectSingleNode("//root//SAPDBServer");
                DBServer = node.InnerText;
                node = doc.SelectSingleNode("//root//SAPDBName");
                dbName = node.InnerText;

                eBizChargeForSAP.B1InstallService.B1InstallServiceSoapClient install = new eBizChargeForSAP.B1InstallService.B1InstallServiceSoapClient();
                string domain = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
                if (domain == "" || domain == null)
                    domain = dbName;
                B1Info = install.GetInfoByLicSvr(licServer, domain, dbName);
                trace("GetInfoByLicSvr: " +
                 "\r\nDBType:" + B1Info.DBType +
                 "\r\nUser:" + B1Info.AppUser +
                 "\r\nPw:" + B1Info.AppPassWD +
                 "\r\ndbName:" + B1Info.DBName +
                "\r\nserver:" + B1Info.DBServer
                 );
            }
            else
            {
                errorLog(filename + " not found. ");
                return false;
            }
            oCompany = new SAPbobsCOM.Company();
            oCompany.Server = B1Info.DBServer;
            oCompany.CompanyDB = B1Info.DBName;
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
            oCompany.DbPassword = B1Info.DBPassWD;
            oCompany.DbUserName = B1Info.DBUser;
            oCompany.Password = B1Info.AppPassWD;
            oCompany.UserName = B1Info.AppUser;
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;

            oCompany.LicenseServer = B1Info.LicenseServer;
            int n = oCompany.Connect();
            if (n != 0)
            {
                errorLog("Failed to connect.  Error: " + oCompany.GetLastErrorDescription() +
                    "\r\nDBType:" + B1Info.DBType +
                    "\r\nUser:" + B1Info.AppUser +
                    "\r\nPw:" + B1Info.AppPassWD +
                    "\r\ndbName:" + B1Info.DBName +
                   "\r\nserver:" + B1Info.DBServer
                    );
                return false;
            }
            trace("DB connected.  " +
                    "\r\nDBType:" + B1Info.DBType +
                    "\r\nUser:" + B1Info.AppUser +
                    "\r\nPw:" + B1Info.AppPassWD +
                    "\r\ndbName:" + B1Info.DBName +
                   "\r\nserver:" + B1Info.DBServer
                    );
            getAddOnVersion();
            //createSAPTables();  //DDL is not supported on 
            cfgInit();
            createAcctList();
            return true;
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        finally
        {
            try
            {
                watch.Stop();
                trace("Databasse Connected. time elapsed: " + watch.ElapsedMilliseconds);

            }
            catch (Exception) { }
        }
        return false;
    }

    private void getLastError(int n)
    {
        try
        {
            int temp_int = n;
            string temp_string = "";
            oCompany.GetLastError(out temp_int, out temp_string);
            errorLog(temp_string);
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }

    public SAP()
    {
        try
        {
            // *************************************************************
            //  set SBO_Application with an initialized application object
            // *************************************************************





        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
    {

        // ********************************************************************************
        //  the following are the events sent by the application
        //  (Ignore aet_ServerTermination)
        //  in order to implement your own code upon each of the events
        //  place you code instead of the matching message box statement
        // ********************************************************************************
        try
        {

            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    CloseApplication();
                    break;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private void CloseApplication()
    {
        try
        {
            showStatus("eBizCharge Exiting...");
            if (SBO_Application != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(SBO_Application);
                SBO_Application = null;
            }
            if (oCompany != null)
            {
                if (oCompany.Connected)
                    oCompany.Disconnect();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompany);
                oCompany = null;
            }

        }
        catch (Exception)
        {

        }
        finally
        {
            System.Environment.Exit(0);
            // clsUtilities.oApplication = null;
            //GC.WaitForPendingFinalizers();
            //GC.Collect();
        }
    }
    public void createSAPTables()
    {

        try
        {
            string sql;
            if (!IsRSEOF("select * from  \"@CCCFG\" "))
            {

                return;
            }

            //sql = "CREATE COLUMN TABLE \"@CCCARDACCT\" (\"DocEntry\" integer  NULL, \"DocNum\" integer  NULL, \"Period\" integer  NULL, \"Instance\" SMALLINT,\"Series\" integer  NULL,\"Handwrtten\" CHAR  NULL,\"Canceled\" CHAR  NULL,\"Object\" nvarchar(20)  NULL,\"LogInst\" integer  NULL,\"UserSign\" integer  NULL,\"Transfered\" CHAR  NULL,\"Status\" CHAR  NULL,\"CreateDate\" timestamp  NULL,\"CreateTime\" SMALLINT  NULL,\"UpdateDate\" timestamp  NULL,\"UpdateTime\" SMALLINT  NULL,\"DataSource\" CHAR  NULL,\"RequestStatus\" CHAR  NULL,\"Creator\" nvarchar(8) NULL,\"Remark\" NCLOB NULL,\"U_GroupName\" varchar(80) NULL, \"U_CardType\" varchar(80) NULL, \"U_CardName\" varchar(80) NULL, \"U_AcctCode\" varchar(80) NULL, \"U_CardCode\" integer NULL, \"U_SourceKey\" varchar(50) NULL, \"U_Currency\" varchar(50) NULL, \"U_Pin\" varchar(50) NULL, \"U_BranchID\" varchar(50) NULL)";
            sql = "CREATE COLUMN TABLE \"@CCCARDACCT\" (\"DocEntry\" INTEGER CS_INT NOT NULL , \"DocNum\" INTEGER CS_INT, \"Period\" INTEGER CS_INT, \"Instance\" SMALLINT CS_INT DEFAULT 0, \"Series\" INTEGER CS_INT, \"Handwrtten\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Canceled\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Object\" NVARCHAR(20), \"LogInst\" INTEGER CS_INT, \"UserSign\" INTEGER CS_INT, \"Transfered\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Status\" CHAR(1) CS_FIXEDSTRING DEFAULT 'O', \"CreateDate\" LONGDATE CS_LONGDATE, \"CreateTime\" SMALLINT CS_INT, \"UpdateDate\" LONGDATE CS_LONGDATE, \"UpdateTime\" SMALLINT CS_INT, \"DataSource\" CHAR(1) CS_FIXEDSTRING, \"RequestStatus\" CHAR(1) CS_FIXEDSTRING DEFAULT 'W', \"Creator\" NVARCHAR(8), \"Remark\" NCLOB MEMORY THRESHOLD 1000, \"U_GroupName\" NVARCHAR(80), \"U_CardType\" NVARCHAR(80), \"U_CardName\" NVARCHAR(80), \"U_AcctCode\" NVARCHAR(80), \"U_CardCode\" NVARCHAR(50), \"U_SourceKey\" NVARCHAR(50), \"U_Currency\" NVARCHAR(50), \"U_Pin\" NVARCHAR(50), \"U_BranchID\" NVARCHAR(50), PRIMARY KEY (\"DocEntry\")) UNLOAD PRIORITY 5  AUTO MERGE";
            execute(sql);
            //sql = "CREATE COLUMN TABLE \"@CCCFG\" (\"DocEntry\" integer  NULL, \"DocNum\" integer  NULL, \"Period\" integer  NULL, \"Instance\" SMALLINT,\"Series\" integer  NULL,\"Handwrtten\" CHAR  NULL,\"Canceled\" CHAR  NULL,\"Object\" nvarchar(20)  NULL,\"LogInst\" integer  NULL,\"UserSign\" integer  NULL,\"Transfered\" CHAR  NULL,\"Status\" CHAR  NULL,\"CreateDate\" timestamp  NULL,\"CreateTime\" SMALLINT  NULL,\"UpdateDate\" timestamp  NULL,\"UpdateTime\" SMALLINT  NULL,\"DataSource\" CHAR  NULL,\"RequestStatus\" CHAR  NULL,\"Creator\" nvarchar(8) NULL,\"Remark\" NCLOB NULL,\"U_Pin\" nvarchar(50) NULL, \"U_SourceKey\" nvarchar(50) NULL, \"U_eBizChargeUrl\" nvarchar(50) NULL,\"U_NegPayment\" nvarchar(50) NULL,\"U_importPin\" nvarchar(50) NULL, \"U_importSourceKey\" nvarchar(50) NULL, \"U_importCCWebServiceUrl\" varchar(80) NULL, \"U_active\" nvarchar(10) NULL,\"U_sendCustomerEmail\" nvarchar(50) NULL, \"U_urlSandbox\" nvarchar(50) NULL, \"U_urlProduction\" nvarchar(50) NULL, \"U_MerchantReceipt\" nvarchar(50) NULL,\"U_CustomerReceipt\" nvarchar(50) NULL, \"U_merchantEmail\" varchar(50) NULL, \"U_addressRequired\" varchar(20) NULL, \"U_paymentGroup\" nvarchar(50) NULL, \"U_useUserEmail\" nvarchar(50) NULL, \"U_merchantID\" nvarchar(50) NULL, \"U_BatchAutoMode\" nvarchar(50) NULL, \"U_recurringBilling\" nvarchar(50) NULL,\"U_BtnPos\" nvarchar(50) NULL, \"U_NoTransOnInv\" nvarchar(50) NULL, \"U_DefaultTaxRate\" nvarchar(50) NULL, \"U_PMAutoSelect\" nvarchar(50) NULL, \"U_importCust\" nvarchar(50) NULL, \"U_ProcessPreAuth\" nvarchar(50) NULL, \"U_showStartup\" nvarchar(50) NULL, \"U_EmailSubject\" nvarchar(250) NULL, \"U_SaveReceipt\" nvarchar(50) NULL, \"U_BatchPreauth\" nvarchar(50) NULL, \"U_DeclineUnapprove\" nvarchar(50) NULL, \"U_PayOnTerm\" nvarchar(50) NULL, \"U_ApprUDF\" nvarchar(50) NULL, \"U_CaptureOnPrint\" nvarchar(50) NULL, \"U_CaptureOnly\" nvarchar(50) NULL, \"U_AmountField\" nvarchar(50) NULL, \"U_FieldSize\" nvarchar(50) NULL, \"U_HasTerminal\" nvarchar(50) NULL)";
            sql = "CREATE COLUMN TABLE \"@CCCFG\" (\"DocEntry\" INTEGER CS_INT NOT NULL , \"DocNum\" INTEGER CS_INT, \"Period\" INTEGER CS_INT, \"Instance\" SMALLINT CS_INT DEFAULT 0, \"Series\" INTEGER CS_INT, \"Handwrtten\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Canceled\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Object\" NVARCHAR(20), \"LogInst\" INTEGER CS_INT, \"UserSign\" INTEGER CS_INT, \"Transfered\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Status\" CHAR(1) CS_FIXEDSTRING DEFAULT 'O', \"CreateDate\" LONGDATE CS_LONGDATE, \"CreateTime\" SMALLINT CS_INT, \"UpdateDate\" LONGDATE CS_LONGDATE, \"UpdateTime\" SMALLINT CS_INT, \"DataSource\" CHAR(1) CS_FIXEDSTRING, \"RequestStatus\" CHAR(1) CS_FIXEDSTRING DEFAULT 'W', \"Creator\" NVARCHAR(8), \"Remark\" NCLOB MEMORY THRESHOLD 1000, \"U_Pin\" NVARCHAR(50), \"U_SourceKey\" NVARCHAR(50), \"U_eBizChargeUrl\" NVARCHAR(80), \"U_NegPayment\" NVARCHAR(50), \"U_importPin\" NVARCHAR(50), \"U_importSourceKey\" NVARCHAR(50), \"U_impWebServiceUrl\" NVARCHAR(80), \"U_active\" NVARCHAR(10), \"U_sendCustomerEmail\" NVARCHAR(50), \"U_urlSandbox\" NVARCHAR(50), \"U_urlProduction\" NVARCHAR(50), \"U_MerchantReceipt\" NVARCHAR(50), \"U_CustomerReceipt\" NVARCHAR(50), \"U_merchantEmail\" NVARCHAR(50), \"U_addressRequired\" NVARCHAR(50), \"U_paymentGroup\" NVARCHAR(50), \"U_useUserEmail\" NVARCHAR(50), \"U_merchantID\" NVARCHAR(50), \"U_BatchAutoMode\" NVARCHAR(50), \"U_recurringBilling\" NVARCHAR(50), \"U_BtnPos\" NVARCHAR(50), \"U_NoTransOnInv\" NVARCHAR(50), \"U_DefaultTaxRate\" NVARCHAR(50), \"U_PMAutoSelect\" NVARCHAR(50), \"U_importCust\" NVARCHAR(50), \"U_ProcessPreAuth\" NVARCHAR(20), \"U_showStartup\" NVARCHAR(20), \"U_EmailSubject\" NVARCHAR(250), \"U_SaveReceipt\" NVARCHAR(20), \"U_BatchPreauth\" NVARCHAR(20), \"U_DeclineUnapprove\" NVARCHAR(20), \"U_PayOnTerm\" NVARCHAR(20), \"U_ApprUDF\" NVARCHAR(50), \"U_CaptureOnPrint\" NVARCHAR(20), \"U_CaptureOnly\" NVARCHAR(50), \"U_AmountField\" NVARCHAR(50), \"U_FieldSize\" NVARCHAR(50), \"U_HasTerminal\" NVARCHAR(50), PRIMARY KEY (\"DocEntry\")) UNLOAD PRIORITY 5  AUTO MERGE";
            execute(sql);

            sql = "insert into \"@CCCFG\" values(1,1,1,0,1,'N','N','',1,1,'N','O','',1,'',1,'','W','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','')";
            execute(sql);

            //sql = "CREATE COLUMN TABLE \"@CCCONNECTCUSTOMER\" (\"DocEntry\" integer  NULL, \"DocNum\" integer  NULL, \"Period\" integer  NULL, \"Instance\" SMALLINT,\"Series\" integer  NULL,\"Handwrtten\" CHAR  NULL,\"Canceled\" CHAR  NULL,\"Object\" nvarchar(20)  NULL,\"LogInst\" integer  NULL,\"UserSign\" integer  NULL,\"Transfered\" CHAR  NULL,\"Status\" CHAR  NULL,\"CreateDate\" timestamp  NULL,\"CreateTime\" SMALLINT  NULL,\"UpdateDate\" timestamp  NULL,\"UpdateTime\" SMALLINT  NULL,\"DataSource\" CHAR  NULL,\"RequestStatus\" CHAR  NULL,\"Creator\" nvarchar(8) NULL,\"Remark\" NCLOB NULL,\"U_CustomerID\" nvarchar(50) NULL, \"U_CustomerGUID\" nvarchar(50) NULL)";

            sql = "CREATE COLUMN TABLE \"@CCCONNECTCUSTOMER\" (\"DocEntry\" INTEGER CS_INT NOT NULL , \"DocNum\" INTEGER CS_INT, \"Period\" INTEGER CS_INT, \"Instance\" SMALLINT CS_INT DEFAULT 0, \"Series\" INTEGER CS_INT, \"Handwrtten\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Canceled\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Object\" NVARCHAR(20), \"LogInst\" INTEGER CS_INT, \"UserSign\" INTEGER CS_INT, \"Transfered\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Status\" CHAR(1) CS_FIXEDSTRING DEFAULT 'O', \"CreateDate\" LONGDATE CS_LONGDATE, \"CreateTime\" SMALLINT CS_INT, \"UpdateDate\" LONGDATE CS_LONGDATE, \"UpdateTime\" SMALLINT CS_INT, \"DataSource\" CHAR(1) CS_FIXEDSTRING, \"RequestStatus\" CHAR(1) CS_FIXEDSTRING DEFAULT 'W', \"Creator\" NVARCHAR(8), \"Remark\" NCLOB MEMORY THRESHOLD 1000, \"U_CustomerID\" NVARCHAR(50), \"U_CustomerGUID\" NVARCHAR(50), PRIMARY KEY (\"DocEntry\")) UNLOAD PRIORITY 5  AUTO MERGE ";
            execute(sql);
            //sql = "CREATE COLUMN TABLE \"@CCCONNECTINVOICE\" (\"DocEntry\" integer  NULL, \"DocNum\" integer  NULL, \"Period\" integer  NULL, \"Instance\" SMALLINT,\"Series\" integer  NULL,\"Handwrtten\" CHAR  NULL,\"Canceled\" CHAR  NULL,\"Object\" nvarchar(20)  NULL,\"LogInst\" integer  NULL,\"UserSign\" integer  NULL,\"Transfered\" CHAR  NULL,\"Status\" CHAR  NULL,\"CreateDate\" timestamp  NULL,\"CreateTime\" SMALLINT  NULL,\"UpdateDate\" timestamp  NULL,\"UpdateTime\" SMALLINT  NULL,\"DataSource\" CHAR  NULL,\"RequestStatus\" CHAR  NULL,\"Creator\" nvarchar(8) NULL,\"Remark\" NCLOB NULL,\"U_InvoiceID\" nvarchar(50) NULL, \"U_CustomerID\" nvarchar(50) NULL,\"U_InvoiceGUID\" nvarchar(50) NULL, \"U_PaidAmount\" nvarchar(50) NULL,\"U_UploadedBalance\" nvarchar(50) NULL, \"U_Status\" nvarchar(50) NULL,\"U_UploadDate\" timestamp NULL,\"U_PaidDate\" timestamp NULL)";
            sql = "CREATE COLUMN TABLE \"@CCCONNECTINVOICE\" (\"DocEntry\" INTEGER CS_INT NOT NULL , \"DocNum\" INTEGER CS_INT, \"Period\" INTEGER CS_INT, \"Instance\" SMALLINT CS_INT DEFAULT 0, \"Series\" INTEGER CS_INT, \"Handwrtten\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Canceled\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Object\" NVARCHAR(20), \"LogInst\" INTEGER CS_INT, \"UserSign\" INTEGER CS_INT, \"Transfered\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Status\" CHAR(1) CS_FIXEDSTRING DEFAULT 'O', \"CreateDate\" LONGDATE CS_LONGDATE, \"CreateTime\" SMALLINT CS_INT, \"UpdateDate\" LONGDATE CS_LONGDATE, \"UpdateTime\" SMALLINT CS_INT, \"DataSource\" CHAR(1) CS_FIXEDSTRING, \"RequestStatus\" CHAR(1) CS_FIXEDSTRING DEFAULT 'W', \"Creator\" NVARCHAR(8), \"Remark\" NCLOB MEMORY THRESHOLD 1000, \"U_InvoiceID\" NVARCHAR(50), \"U_CustomerID\" NVARCHAR(50), \"U_InvoiceGUID\" NVARCHAR(50), \"U_PaidAmount\" NVARCHAR(50), \"U_UploadedBalance\" NVARCHAR(50), \"U_Status\" NVARCHAR(50), \"U_UploadDate\" LONGDATE CS_LONGDATE, \"U_PaidDate\" LONGDATE CS_LONGDATE, PRIMARY KEY (\"DocEntry\")) UNLOAD PRIORITY 5  AUTO MERGE";
            execute(sql);
            //sql = "CREATE COLUMN TABLE \"@CCCPAYMENT\" (\"DocEntry\" integer  NULL, \"DocNum\" integer  NULL, \"Period\" integer  NULL, \"Instance\" SMALLINT,\"Series\" integer  NULL,\"Handwrtten\" CHAR  NULL,\"Canceled\" CHAR  NULL,\"Object\" nvarchar(20)  NULL,\"LogInst\" integer  NULL,\"UserSign\" integer  NULL,\"Transfered\" CHAR  NULL,\"Status\" CHAR  NULL,\"CreateDate\" timestamp  NULL,\"CreateTime\" SMALLINT  NULL,\"UpdateDate\" timestamp  NULL,\"UpdateTime\" SMALLINT  NULL,\"DataSource\" CHAR  NULL,\"RequestStatus\" CHAR  NULL,\"Creator\" nvarchar(8) NULL,\"Remark\" NCLOB NULL,\"U_InvoiceID\" nvarchar(50) NULL, \"U_CustomerID\" nvarchar(50) NULL,\"U_PaymentID\" nvarchar(50) NULL, \"U_Amount\" nvarchar(50) NULL,\"U_DateImported\" timestamp NULL)";
            sql = "CREATE COLUMN TABLE \"@CCCPAYMENT\" (\"DocEntry\" INTEGER CS_INT NOT NULL , \"DocNum\" INTEGER CS_INT, \"Period\" INTEGER CS_INT, \"Instance\" SMALLINT CS_INT DEFAULT 0, \"Series\" INTEGER CS_INT, \"Handwrtten\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Canceled\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Object\" NVARCHAR(20), \"LogInst\" INTEGER CS_INT, \"UserSign\" INTEGER CS_INT, \"Transfered\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Status\" CHAR(1) CS_FIXEDSTRING DEFAULT 'O', \"CreateDate\" LONGDATE CS_LONGDATE, \"CreateTime\" SMALLINT CS_INT, \"UpdateDate\" LONGDATE CS_LONGDATE, \"UpdateTime\" SMALLINT CS_INT, \"DataSource\" CHAR(1) CS_FIXEDSTRING, \"RequestStatus\" CHAR(1) CS_FIXEDSTRING DEFAULT 'W', \"Creator\" NVARCHAR(8), \"Remark\" NCLOB MEMORY THRESHOLD 1000, \"U_CustomerID\" NVARCHAR(50), \"U_InvoiceID\" NVARCHAR(50), \"U_PaymentID\" NVARCHAR(50), \"TypeId\" NVARCHAR(50), \"U_Amount\" NVARCHAR(50), \"U_DateImported\" LONGDATE CS_LONGDATE, PRIMARY KEY (\"DocEntry\")) UNLOAD PRIORITY 5  AUTO MERGE";
            execute(sql);
            //sql = "CREATE COLUMN TABLE \"@CCCTERMINAL\" (\"DocEntry\" integer  NULL, \"DocNum\" integer  NULL, \"Period\" integer  NULL, \"Instance\" SMALLINT,\"Series\" integer  NULL,\"Handwrtten\" CHAR  NULL,\"Canceled\" CHAR  NULL,\"Object\" nvarchar(20)  NULL,\"LogInst\" integer  NULL,\"UserSign\" integer  NULL,\"Transfered\" CHAR  NULL,\"Status\" CHAR  NULL,\"CreateDate\" timestamp  NULL,\"CreateTime\" SMALLINT  NULL,\"UpdateDate\" timestamp  NULL,\"UpdateTime\" SMALLINT  NULL,\"DataSource\" CHAR  NULL,\"RequestStatus\" CHAR  NULL,\"Creator\" nvarchar(8) NULL,\"Remark\" NCLOB NULL,\"U_name\" nvarchar(50) NULL, \"U_url\" nvarchar(50) NULL,\"U_deviceKey\" nvarchar(50) NULL, \"U_apiKey\" nvarchar(50) NULL, \"U_pin\" nvarchar(50) NULL, \"U_isDefault\" nvarchar(50) NULL )";
            sql = "CREATE COLUMN TABLE \"@CCCTERMINAL\" (\"DocEntry\" INTEGER CS_INT NOT NULL , \"DocNum\" INTEGER CS_INT, \"Period\" INTEGER CS_INT, \"Instance\" SMALLINT CS_INT DEFAULT 0, \"Series\" INTEGER CS_INT, \"Handwrtten\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Canceled\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Object\" NVARCHAR(20), \"LogInst\" INTEGER CS_INT, \"UserSign\" INTEGER CS_INT, \"Transfered\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Status\" CHAR(1) CS_FIXEDSTRING DEFAULT 'O', \"CreateDate\" LONGDATE CS_LONGDATE, \"CreateTime\" SMALLINT CS_INT, \"UpdateDate\" LONGDATE CS_LONGDATE, \"UpdateTime\" SMALLINT CS_INT, \"DataSource\" CHAR(1) CS_FIXEDSTRING, \"RequestStatus\" CHAR(1) CS_FIXEDSTRING DEFAULT 'W', \"Creator\" NVARCHAR(8), \"Remark\" NCLOB MEMORY THRESHOLD 1000, \"U_name\" NVARCHAR(50), \"U_url\" NVARCHAR(50), \"U_deviceKey\" NVARCHAR(80), \"U_apiKey\" NVARCHAR(50), \"U_pin\" NVARCHAR(50), \"U_isDefault\" NVARCHAR(50), PRIMARY KEY (\"DocEntry\")) UNLOAD PRIORITY 5  AUTO MERGE";
            execute(sql);
            //sql = "CREATE COLUMN TABLE \"@CCCUST\" (\"DocEntry\" integer  NULL, \"DocNum\" integer  NULL, \"Period\" integer  NULL, \"Instance\" SMALLINT,\"Series\" integer  NULL,\"Handwrtten\" CHAR  NULL,\"Canceled\" CHAR  NULL,\"Object\" nvarchar(20)  NULL,\"LogInst\" integer  NULL,\"UserSign\" integer  NULL,\"Transfered\" CHAR  NULL,\"Status\" CHAR  NULL,\"CreateDate\" timestamp  NULL,\"CreateTime\" SMALLINT  NULL,\"UpdateDate\" timestamp  NULL,\"UpdateTime\" SMALLINT  NULL,\"DataSource\" CHAR  NULL,\"RequestStatus\" CHAR  NULL,\"Creator\" nvarchar(8) NULL,\"Remark\" NCLOB NULL,\"U_CustomerID\" nvarchar(50) NULL, \"U_CustNum\" nvarchar(50) NULL,\"U_MethodID\" nvarchar(50) NULL, \"U_default\" nvarchar(10) NULL, \"U_active\" nvarchar(10) NULL, \"U_Declined\" nvarchar(10) NULL, \"U_methodDescription\" nvarchar(250) NULL, \"U_cardNum\" nvarchar(50) NULL, \"U_expDate\" nvarchar(50) NULL,\"U_cardCode\" nvarchar(10) NULL,\"U_routingNumber\" nvarchar(50) NULL,\"U_checkingAccount\" nvarchar(50) NULL,\"U_firstName\" nvarchar(50)  NULL,\"U_lastName\" nvarchar(50) NULL,\"U_email\" nvarchar(80) NULL,\"U_street\" nvarchar(80) NULL,\"U_city\" nvarchar(50) NULL,\"U_state\" nvarchar(20) NULL,\"U_zip\" nvarchar(20) NULL,\"U_cardType\" nvarchar(50) NULL,\"U_GroupName\" nvarchar(80) NULL,\"U_CardName\" nvarchar(80) NULL,\"U_CCAccountID\" nvarchar(80) NULL,\"U_MethodName\" nvarchar(80) NULL)";
            sql = "CREATE COLUMN TABLE \"@CCCUST\" (\"DocEntry\" INTEGER CS_INT NOT NULL , \"DocNum\" INTEGER CS_INT, \"Period\" INTEGER CS_INT, \"Instance\" SMALLINT CS_INT DEFAULT 0, \"Series\" INTEGER CS_INT, \"Handwrtten\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Canceled\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Object\" NVARCHAR(20), \"LogInst\" INTEGER CS_INT, \"UserSign\" INTEGER CS_INT, \"Transfered\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Status\" CHAR(1) CS_FIXEDSTRING DEFAULT 'O', \"CreateDate\" LONGDATE CS_LONGDATE, \"CreateTime\" SMALLINT CS_INT, \"UpdateDate\" LONGDATE CS_LONGDATE, \"UpdateTime\" SMALLINT CS_INT, \"DataSource\" CHAR(1) CS_FIXEDSTRING, \"RequestStatus\" CHAR(1) CS_FIXEDSTRING DEFAULT 'W', \"Creator\" NVARCHAR(8), \"Remark\" NCLOB MEMORY THRESHOLD 1000, \"U_CustomerID\" NVARCHAR(50), \"U_CustNum\" NVARCHAR(50), \"U_MethodID\" NVARCHAR(50), \"U_default\" NVARCHAR(10), \"U_active\" NVARCHAR(10), \"U_Declined\" NVARCHAR(10), \"U_methodDescription\" NVARCHAR(250), \"U_cardNum\" NVARCHAR(50), \"U_expDate\" NVARCHAR(50), \"U_cardCode\" NVARCHAR(10), \"U_routingNumber\" NVARCHAR(50), \"U_checkingAccount\" NVARCHAR(50), \"U_firstName\" NVARCHAR(50), \"U_lastName\" NVARCHAR(50), \"U_email\" NVARCHAR(80), \"U_street\" NVARCHAR(80), \"U_city\" NVARCHAR(50), \"U_state\" NVARCHAR(20), \"U_zip\" NVARCHAR(20), \"U_cardType\" NVARCHAR(50), \"U_GroupName\" NVARCHAR(80), \"U_CardName\" NVARCHAR(80), \"U_CCAccountID\" NVARCHAR(80), \"U_MethodName\" NVARCHAR(80), PRIMARY KEY (\"DocEntry\")) UNLOAD PRIORITY 5  AUTO MERGE";
            execute(sql);
            //sql = "CREATE COLUMN TABLE \"@CCJOB\" (\"DocEntry\" integer  NULL, \"DocNum\" integer  NULL, \"Period\" integer  NULL, \"Instance\" SMALLINT,\"Series\" integer  NULL,\"Handwrtten\" CHAR  NULL,\"Canceled\" CHAR  NULL,\"Object\" nvarchar(20)  NULL,\"LogInst\" integer  NULL,\"UserSign\" integer  NULL,\"Transfered\" CHAR  NULL,\"Status\" CHAR  NULL,\"CreateDate\" timestamp  NULL,\"CreateTime\" SMALLINT  NULL,\"UpdateDate\" timestamp  NULL,\"UpdateTime\" SMALLINT  NULL,\"DataSource\" CHAR  NULL,\"RequestStatus\" CHAR  NULL,\"Creator\" nvarchar(8) NULL,\"Remark\" NCLOB NULL,\"U_InvoiceID\" nvarchar(50) NULL, \"U_OrderID\" nvarchar(50) NULL, \"U_Description\" nvarchar(255) NULL, \"U_CustomerID\" nvarchar(50) NULL, \"U_Frequency\" nvarchar(50) NULL, \"U_StartDate\" timestamp NULL, \"U_EndDate\" timestamp NULL, \"U_NextRunDate\" timestamp NULL, \"U_LastRunDate\" timestamp NULL, \"U_Cancelled\" nvarchar(50) NULL, \"U_CancelledDate\" timestamp NULL, \"U_CancelledBy\" nvarchar(50) NULL, \"U_Amount\" nvarchar(50) NULL, \"U_PaymentID\" nvarchar(250) NULL, \"U_Result\" nvarchar(250) NULL)";
            sql = "CREATE COLUMN TABLE \"@CCJOB\" (\"DocEntry\" INTEGER CS_INT NOT NULL , \"DocNum\" INTEGER CS_INT, \"Period\" INTEGER CS_INT, \"Instance\" SMALLINT CS_INT DEFAULT 0, \"Series\" INTEGER CS_INT, \"Handwrtten\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Canceled\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Object\" NVARCHAR(20), \"LogInst\" INTEGER CS_INT, \"UserSign\" INTEGER CS_INT, \"Transfered\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Status\" CHAR(1) CS_FIXEDSTRING DEFAULT 'O', \"CreateDate\" LONGDATE CS_LONGDATE, \"CreateTime\" SMALLINT CS_INT, \"UpdateDate\" LONGDATE CS_LONGDATE, \"UpdateTime\" SMALLINT CS_INT, \"DataSource\" CHAR(1) CS_FIXEDSTRING, \"RequestStatus\" CHAR(1) CS_FIXEDSTRING DEFAULT 'W', \"Creator\" NVARCHAR(8), \"Remark\" NCLOB MEMORY THRESHOLD 1000, \"U_InvoiceID\" NVARCHAR(50), \"U_OrderID\" NVARCHAR(50), \"U_Description\" NVARCHAR(254), \"U_CustomerID\" NVARCHAR(50), \"U_Frequency\" NVARCHAR(50), \"U_StartDate\" LONGDATE CS_LONGDATE, \"U_EndDate\" LONGDATE CS_LONGDATE, \"U_NextRunDate\" LONGDATE CS_LONGDATE, \"U_LastRunDate\" LONGDATE CS_LONGDATE, \"U_Cancelled\" NVARCHAR(10), \"U_CancelledDate\" LONGDATE CS_LONGDATE, \"U_CancelledBy\" NVARCHAR(50), \"U_Amount\" NVARCHAR(50), \"U_PaymentID\" NVARCHAR(250), \"U_Result\" NVARCHAR(250), PRIMARY KEY (\"DocEntry\")) UNLOAD PRIORITY 5  AUTO MERGE";
            execute(sql);
            //sql = "CREATE COLUMN TABLE \"@CCPAYFORMPAYMENT\" (\"DocEntry\" integer  NULL, \"DocNum\" integer  NULL, \"Period\" integer  NULL, \"Instance\" SMALLINT,\"Series\" integer  NULL,\"Handwrtten\" CHAR  NULL,\"Canceled\" CHAR  NULL,\"Object\" nvarchar(20)  NULL,\"LogInst\" integer  NULL,\"UserSign\" integer  NULL,\"Transfered\" CHAR  NULL,\"Status\" CHAR  NULL,\"CreateDate\" timestamp  NULL,\"CreateTime\" SMALLINT  NULL,\"UpdateDate\" timestamp  NULL,\"UpdateTime\" SMALLINT  NULL,\"DataSource\" CHAR  NULL,\"RequestStatus\" CHAR  NULL,\"Creator\" nvarchar(8) NULL,\"Remark\" NCLOB NULL,\"U_CustomerID\" nvarchar(50) NULL, \"U_InvoiceID\" nvarchar(50) NULL, \"U_PaymentID\" nvarchar(50) NULL, \"U_Amount\" nvarchar(50) NULL, \"U_DateImported\" timestamp NULL)";
            sql = "CREATE COLUMN TABLE \"@CCPAYFORMPAYMENT\" (\"DocEntry\" INTEGER CS_INT NOT NULL , \"DocNum\" INTEGER CS_INT, \"Period\" INTEGER CS_INT, \"Instance\" SMALLINT CS_INT DEFAULT 0, \"Series\" INTEGER CS_INT, \"Handwrtten\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Canceled\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Object\" NVARCHAR(20), \"LogInst\" INTEGER CS_INT, \"UserSign\" INTEGER CS_INT, \"Transfered\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Status\" CHAR(1) CS_FIXEDSTRING DEFAULT 'O', \"CreateDate\" LONGDATE CS_LONGDATE, \"CreateTime\" SMALLINT CS_INT, \"UpdateDate\" LONGDATE CS_LONGDATE, \"UpdateTime\" SMALLINT CS_INT, \"DataSource\" CHAR(1) CS_FIXEDSTRING, \"RequestStatus\" CHAR(1) CS_FIXEDSTRING DEFAULT 'W', \"Creator\" NVARCHAR(8), \"Remark\" NCLOB MEMORY THRESHOLD 1000, \"U_CustomerID\" NVARCHAR(50), \"U_InvoiceID\" NVARCHAR(50), \"U_PaymentID\" NVARCHAR(50), \"U_Amount\" NVARCHAR(50), \"U_DateImported\" LONGDATE CS_LONGDATE, PRIMARY KEY (\"DocEntry\")) UNLOAD PRIORITY 5  AUTO MERGE";
            execute(sql);
            //sql = "CREATE COLUMN TABLE \"@CCPAYFORMINVOICE\" (\"DocEntry\" integer  NULL, \"DocNum\" integer  NULL, \"Period\" integer  NULL, \"Instance\" SMALLINT,\"Series\" integer  NULL,\"Handwrtten\" CHAR  NULL,\"Canceled\" CHAR  NULL,\"Object\" nvarchar(20)  NULL,\"LogInst\" integer  NULL,\"UserSign\" integer  NULL,\"Transfered\" CHAR  NULL,\"Status\" CHAR  NULL,\"CreateDate\" timestamp  NULL,\"CreateTime\" SMALLINT  NULL,\"UpdateDate\" timestamp  NULL,\"UpdateTime\" SMALLINT  NULL,\"DataSource\" CHAR  NULL,\"RequestStatus\" CHAR  NULL,\"Creator\" nvarchar(8) NULL,\"Remark\" NCLOB NULL,\"U_InvoiceID\" nvarchar(50) NULL,\"U_CustomerID\" nvarchar(50) NULL, \"U_InvoiceGUID\" nvarchar(50) NULL, \"U_PaidAmount\" nvarchar(50) NULL, \"U_UploadedBalance\" nvarchar(50) NULL,\"U_Status\" nvarchar(50) NULL,\"U_UploadDate\" timestamp NULL, \"U_PaidDate\" timestamp NULL)";
            sql = "CREATE COLUMN TABLE \"@CCPAYFORMINVOICE\" (\"DocEntry\" INTEGER CS_INT NOT NULL , \"DocNum\" INTEGER CS_INT, \"Period\" INTEGER CS_INT, \"Instance\" SMALLINT CS_INT DEFAULT 0, \"Series\" INTEGER CS_INT, \"Handwrtten\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Canceled\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Object\" NVARCHAR(20), \"LogInst\" INTEGER CS_INT, \"UserSign\" INTEGER CS_INT, \"Transfered\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Status\" CHAR(1) CS_FIXEDSTRING DEFAULT 'O', \"CreateDate\" LONGDATE CS_LONGDATE, \"CreateTime\" SMALLINT CS_INT, \"UpdateDate\" LONGDATE CS_LONGDATE, \"UpdateTime\" SMALLINT CS_INT, \"DataSource\" CHAR(1) CS_FIXEDSTRING, \"RequestStatus\" CHAR(1) CS_FIXEDSTRING DEFAULT 'W', \"Creator\" NVARCHAR(8), \"Remark\" NCLOB MEMORY THRESHOLD 1000, \"U_InvoiceID\" NVARCHAR(50), \"U_CustomerID\" NVARCHAR(50), \"U_InvoiceGUID\" NVARCHAR(50), \"U_PaidAmount\" NVARCHAR(50), \"U_UploadedBalance\" NVARCHAR(50), \"U_Status\" NVARCHAR(50), \"U_UploadDate\" LONGDATE CS_LONGDATE, \"U_PaidDate\" LONGDATE CS_LONGDATE, PRIMARY KEY (\"DocEntry\")) UNLOAD PRIORITY 5  AUTO MERGE";
            execute(sql);
            //sql = "CREATE COLUMN TABLE \"@CCTRANS\" (\"DocEntry\" integer  NULL, \"DocNum\" integer  NULL, \"Period\" integer  NULL, \"Instance\" SMALLINT,\"Series\" integer  NULL,\"Handwrtten\" CHAR  NULL,\"Canceled\" CHAR  NULL,\"Object\" nvarchar(20)  NULL,\"LogInst\" integer  NULL,\"UserSign\" integer  NULL,\"Transfered\" CHAR  NULL,\"Status\" CHAR  NULL,\"CreateDate\" timestamp  NULL,\"CreateTime\" SMALLINT  NULL,\"UpdateDate\" timestamp  NULL,\"UpdateTime\" SMALLINT  NULL,\"DataSource\" CHAR  NULL,\"RequestStatus\" CHAR  NULL,\"Creator\" nvarchar(8) NULL,\"Remark\" NCLOB NULL,\"U_ccTRANSID\" SMALLINT  NULL, \"U_Imported\" nvarchar(50) NULL, \"U_CCORDERID\" nvarchar(50) NULL, \"U_CCAccountID\" nvarchar(50) NULL, \"U_InvoiceID\" nvarchar(50) NULL, \"U_PaymentID\" nvarchar(50) NULL, \"U_SAPPaymentID\" nvarchar(50) NULL, \"U_DownPaymentInvoiceID\" nvarchar(50) NULL, \"U_OrderID\" nvarchar(50) NULL, \"U_CreditMemoID\" nvarchar(50) NULL, \"U_MethodID\" nvarchar(50) NULL, \"U_customerID\" varchar(50) NULL, \"U_customerName\" nvarchar(80) NULL, \"U_CardHolder\" nvarchar(80) NULL, \"U_creditRef\" nvarchar(50) NULL, \"U_crCardNum\" nvarchar(50) NULL, \"U_Description\" nvarchar(255) NULL, \"U_recID\" nvarchar(20) NULL, \"U_acsUrl\" nvarchar(80) NULL, \"U_authAmount\" nvarchar(50) NULL, \"U_authCode\" nvarchar(10) NULL, \"U_avsResult\" nvarchar(255) NULL, \"U_avsResultCode\" nvarchar(10) NULL, \"U_batchNum\" nvarchar(20) NULL, \"U_batchRef\" nvarchar(20) NULL, \"U_cardCodeResult\" nvarchar(255) NULL, \"U_cardCodeResultCode\" nvarchar(10) NULL, \"U_cardLevelResult\" nvarchar(255) NULL, \"U_cardLevelResultCode\" nvarchar(10) NULL, \"U_conversionRate\" nvarchar(20) NULL, \"U_convertedAmount\" nvarchar(20) NULL, \"U_convertedAmountCurrency\" nvarchar(20) NULL, \"U_custNum\" nvarchar(50) NULL, \"U_error\" nvarchar(250) NULL, \"U_errorCode\" nvarchar(10) NULL, \"U_isDuplicate\" nvarchar(10) NULL, \"U_payload\" nvarchar(50) NULL, \"U_profilerScore\" nvarchar(10) NULL, \"U_profilerResponse\" nvarchar(50) NULL, \"U_profilerReason\" nvarchar(50) NULL, \"U_refNum\" nvarchar(10) NULL, \"U_remainingBalance\" nvarchar(20) NULL, \"U_result\" nvarchar(50) NULL, \"U_resultCode\" nvarchar(10) NULL, \"U_status\" nvarchar(50) NULL, \"U_statusCode\" nvarchar(10) NULL, \"U_vpasResultCode\" nvarchar(10) NULL, \"U_recDate\" timestamp NULL, \"U_command\" nvarchar(20) NULL, \"U_amount\" nvarchar(50) NULL, \"U_BatchResult\" varchar(255) NULL, \"U_BatchDate\" timestamp NULL, \"U_DeliveryID\" varchar(50) NULL, \"U_jobID\" nvarchar(50) NULL, \"U_QuoteID\" nvarchar(50) NULL)";
            sql = "CREATE COLUMN TABLE \"@CCTRANS\" (\"DocEntry\" INTEGER CS_INT NOT NULL , \"DocNum\" INTEGER CS_INT, \"Period\" INTEGER CS_INT, \"Instance\" SMALLINT CS_INT DEFAULT 0, \"Series\" INTEGER CS_INT, \"Handwrtten\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Canceled\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Object\" NVARCHAR(20), \"LogInst\" INTEGER CS_INT, \"UserSign\" INTEGER CS_INT, \"Transfered\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Status\" CHAR(1) CS_FIXEDSTRING DEFAULT 'O', \"CreateDate\" LONGDATE CS_LONGDATE, \"CreateTime\" SMALLINT CS_INT, \"UpdateDate\" LONGDATE CS_LONGDATE, \"UpdateTime\" SMALLINT CS_INT, \"DataSource\" CHAR(1) CS_FIXEDSTRING, \"RequestStatus\" CHAR(1) CS_FIXEDSTRING DEFAULT 'W', \"Creator\" NVARCHAR(8), \"Remark\" NCLOB MEMORY THRESHOLD 1000, \"U_ccTRANSID\" SMALLINT CS_INT, \"U_Imported\" NVARCHAR(10), \"U_CCORDERID\" NVARCHAR(50), \"U_CCAccountID\" NVARCHAR(50), \"U_InvoiceID\" NVARCHAR(50), \"U_PaymentID\" NVARCHAR(50), \"U_SAPPaymentID\" NVARCHAR(50), \"U_DownPaymentInvID\" NVARCHAR(50), \"U_OrderID\" NVARCHAR(50), \"U_CreditMemoID\" NVARCHAR(50), \"U_MethodID\" NVARCHAR(50), \"U_customerID\" NVARCHAR(50), \"U_customerName\" NVARCHAR(80), \"U_CardHolder\" NVARCHAR(80), \"U_creditRef\" NVARCHAR(50), \"U_crCardNum\" NVARCHAR(50), \"U_Description\" NVARCHAR(254), \"U_recID\" NVARCHAR(50), \"U_acsUrl\" NVARCHAR(80), \"U_authAmount\" NVARCHAR(50), \"U_authCode\" NVARCHAR(10), \"U_avsResult\" NVARCHAR(254), \"U_avsResultCode\" NVARCHAR(20), \"U_batchNum\" NVARCHAR(20), \"U_batchRef\" NVARCHAR(20), \"U_cardCodeResult\" NVARCHAR(254), \"U_cardCodeRCode\" NVARCHAR(10), \"U_cardLevelResult\" NVARCHAR(254), \"U_cardLevelRCode\" NVARCHAR(10), \"U_conversionRate\" NVARCHAR(20), \"U_convertedAmount\" NVARCHAR(20), \"U_convertedCurrency\" NVARCHAR(20), \"U_custNum\" NVARCHAR(50), \"U_error\" NVARCHAR(254), \"U_errorCode\" NVARCHAR(10), \"U_isDuplicate\" NVARCHAR(10), \"U_payload\" NVARCHAR(50), \"U_profilerScore\" NVARCHAR(10), \"U_profilerResponse\" NVARCHAR(50), \"U_profilerReason\" NVARCHAR(50), \"U_refNum\" NVARCHAR(50), \"U_remainingBalance\" NVARCHAR(20), \"U_result\" NVARCHAR(50), \"U_resultCode\" NVARCHAR(10), \"U_status\" NVARCHAR(50), \"U_statusCode\" NVARCHAR(50), \"U_vpasResultCode\" NVARCHAR(10), \"U_recDate\" LONGDATE CS_LONGDATE, \"U_command\" NVARCHAR(20), \"U_amount\" NVARCHAR(50), \"U_jobID\" NVARCHAR(50), \"U_BatchResult\" NVARCHAR(254), \"U_BatchDate\" LONGDATE CS_LONGDATE, \"U_DeliveryID\" NVARCHAR(50), \"U_QuoteID\" NVARCHAR(50), PRIMARY KEY (\"DocEntry\")) UNLOAD PRIORITY 5  AUTO MERGE";
            execute(sql);
            //sql = "CREATE COLUMN TABLE \"@CCWALKIN\" (\"DocEntry\" integer  NULL, \"DocNum\" integer  NULL, \"Period\" integer  NULL, \"Instance\" SMALLINT,\"Series\" integer  NULL,\"Handwrtten\" CHAR  NULL,\"Canceled\" CHAR  NULL,\"Object\" nvarchar(20)  NULL,\"LogInst\" integer  NULL,\"UserSign\" integer  NULL,\"Transfered\" CHAR  NULL,\"Status\" CHAR  NULL,\"CreateDate\" timestamp  NULL,\"CreateTime\" SMALLINT  NULL,\"UpdateDate\" timestamp  NULL,\"UpdateTime\" SMALLINT  NULL,\"DataSource\" CHAR  NULL,\"RequestStatus\" CHAR  NULL,\"Creator\" nvarchar(8) NULL,\"Remark\" NCLOB NULL,\"U_CustomerID\" nvarchar(80) NULL)";
            sql = "CREATE COLUMN TABLE \"@CCWALKIN\" (\"DocEntry\" INTEGER CS_INT NOT NULL , \"DocNum\" INTEGER CS_INT, \"Period\" INTEGER CS_INT, \"Instance\" SMALLINT CS_INT DEFAULT 0, \"Series\" INTEGER CS_INT, \"Handwrtten\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Canceled\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Object\" NVARCHAR(20), \"LogInst\" INTEGER CS_INT, \"UserSign\" INTEGER CS_INT, \"Transfered\" CHAR(1) CS_FIXEDSTRING DEFAULT 'N', \"Status\" CHAR(1) CS_FIXEDSTRING DEFAULT 'O', \"CreateDate\" LONGDATE CS_LONGDATE, \"CreateTime\" SMALLINT CS_INT, \"UpdateDate\" LONGDATE CS_LONGDATE, \"UpdateTime\" SMALLINT CS_INT, \"DataSource\" CHAR(1) CS_FIXEDSTRING, \"RequestStatus\" CHAR(1) CS_FIXEDSTRING DEFAULT 'W', \"Creator\" NVARCHAR(8), \"Remark\" NCLOB MEMORY THRESHOLD 1000, \"U_CustomerID\" NVARCHAR(80), PRIMARY KEY (\"DocEntry\")) UNLOAD PRIORITY 5  AUTO MERGE";
            execute(sql);
            showStatus("eBizCharge Tables created Succesfully", SAPbouiCOM.BoMessageTime.bmt_Short, false);

        }
        catch (Exception ex)
        {
            string err = "Can not create tables.  " + ex.Message;
            showMessage(err);
            trace(err);
        }
    }
    public bool CheckUserPermission()
    {
        bool bRtn = true;
        SAPbobsCOM.SBObob bob = (SAPbobsCOM.SBObob)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

        try
        {

            SAPbobsCOM.Recordset rs = bob.GetObjectPermission(SAPbobsCOM.BoObjectTypes.oInvoices);
            for (int j = 0; j < rs.Fields.Count; ++j)
            {
                SAPbobsCOM.Field f = rs.Fields.Item(j);
                trace("Invoice '" + f.Name + "' :: '" + f.Value.ToString() + "'");

            }
            if (rs.Fields.Count == 0)
                trace("User has No Invoice permission'");


        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(bob);
            bob = null;
        }
        return bRtn;
    }
    private void AddCreditCardCheckerMenu()
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
            oCreationPackage.UniqueID = menuCAcctValidate;
            oCreationPackage.String = "eBizCharge Card Account Validate";
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
    public static string getIP()
    {
        try
        {
            IPAddress[] ips;
            string hostname = Dns.GetHostName();
            ips = Dns.GetHostAddresses(hostname);
            foreach (IPAddress ip in ips)
            {
                if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                    return ip.ToString();
            }
        }
        catch (Exception)
        {
        }
        return "";
    }

    public void getCurrentBranch()
    {
        try
        {
            SAPbouiCOM.Form _MenuForm = SBO_Application.Forms.GetForm("169", 0);
            SAPbouiCOM.StaticText _Text = (SAPbouiCOM.StaticText)_MenuForm.Items.Item("7").Specific;
            string _Caption = _Text.Caption;
            trace("Branch: " + _Caption);

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    public bool IsBranchEnabled()
    {
        bool bRet = true;
        try
        {
            SAPbouiCOM.Form _MenuForm = SBO_Application.Forms.GetForm("169", 0);
            SAPbouiCOM.StaticText _Text = (SAPbouiCOM.StaticText)_MenuForm.Items.Item("7").Specific;
            string _Caption = _Text.Caption;
            string[] strBranch;

            strBranch = _Caption.Split(':');
            if (strBranch.Length > 1)
            {
                return true;
                //_CurrentBranch = ;
            }
            else
                return false;

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return bRet;
    }
}

