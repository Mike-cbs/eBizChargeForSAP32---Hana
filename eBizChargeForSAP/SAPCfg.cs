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


partial class SAP
{
    #region DataMember
    const string U_EbizChargeKey = "EbizChargeID";
    const string U_EbizChargeURL = "EbizChargeURL";
    const string U_EbizChargeMarkPayment = "EbizChargeMarkPayment";
    const string editPin = "pin";
    const string editSourceKey = "skey";
    const string editSandbox = "edSndx";
    // const string editIncomingFilter = "fincom";
    // const string editOutgoingFilter = "foutgo";
    //  const string editAccountForOutgoingPayment = "edAFOP";
    const string editNegativeIncomingPayment = "edNIP";
    // const string editAutoSelectAccount = "edASA";
    const string editVisaIncomingAccount = "edVIA";
    const string editMasterIncomingAccount = "edMIA";
    const string editAMXIncomingAccount = "edAMXIA";
    const string editOtherIncomingAccount = "edOIA";
    const string editVisaOutgoingAccount = "edVOA";
    const string editMasterOutgoingAccount = "edMOA";
    const string editAMXOutgoingAccount = "edAMXOA";
    const string editOtherOutgoingAccount = "edOOA";
    const string editDiscIncomingAccount = "edDiscIA";
    const string editDiscOutgoingAccount = "edDiscOA";
    const string editDinnIncomingAccount = "edDinnIA";
    const string editDinnOutgoingAccount = "edDinnOA";
    const string editJCBIncomingAccount = "edJCBIA";
    const string editJCBOutgoingAccount = "edJCBOA";
    const string editeCheckIncomingAccount = "edeChkIA";
    const string editeCheckOutgoingAccount = "edeChkOA";
    const string editsendCustomerEmail = "edSndEM";
    const string editmerchantEmail = "edMercEM";
    const string editEmailSubject = "edEMSUB";
    const string editMerchantReceipt = "edMerRcp";
    const string editCustomerReceipt = "edCusRcp";
    //  const string editRecurringBillingReceipt = "edRBRcp";
    //  const string editDeclineReceipt = "edDclRcp";
    //  const string editDeclineMerchantReceipt = "edDclMer";
    //  const string editimportBPGroupName= "edImpBPG";
    const string editimportPin = "edImpPin";
    const string editimportSourceKey = "edImpSK";
    const string editimportCCWebServiceUrl = "edImpURL";
    const string editimportBPHeaderCode = "edImpBPHC";
    const string editimportItemHeaderCode = "edImpIHC";
    const string editimportTransStatus = "edImpTS";
    const string editimportTransType = "edImpTT";
    const string editimportInvoiceType = "edImpIT";
    const string editProcessPreAuth = "edProcPA";
    const string editimportUseZipForTax = "edimpUZ";
    const string editimportBPNextCode = "edimpBNC";
    const string editaddressRequired = "edAddrR";
    const string mxCardAcct = "mxCrdAcc";
    const string dtCardAcct = "dtCrdAcc";
    const string editMerchantID = "edMID";
    const string editUseUserEmail = "edUsrEM";
    const string editSaveReceipt = "edSAVERT";
    const string editBatchAutoMode = "edBAM";
    const string editWalkIn = "edWalkIn";
    const string cbWalkIn = "cbWalkIn";
    const string editrecurringBilling = "edRBill";
    const string editBtnPos = "edBtnPos";
    const string editDefaultTaxRate = "edTaxRT";
    const string editNoTransOnInv = "edNTOI";
    const string editPMAutoSel = "edPMSEL";
    const string editImportCust = "edImpCut";
    const string editshowStartup = "edSwUp";
    const string editBatchPreauth = "edBPreA";
    const string editDeclineUnapprove = "edDeclUA";
    const string editCaptureOnPrint = "edCapPrt";
    const string editApprUDF = "edAppUDF";
    const string editCaptureOnly = "edCptOn";
    const string editAmountField = "edAMFD";
    const string editFieldSize = "edfSize";
    const string editHasTerminal = "edHasT";
    const string editBINoDelivery = "edBICO";
    const string editPreAuthEmail = "edPAEM";
    const string editAutoPreauth = "edAPre";
    const string editCardEMTemplate = "edCEMT";
    string cfgDeclineUnapprove = "";
    string cfgCurrencyString = "";
    string cfgPin = "";
    string cfgSourceKey = "";
    string cfgUrl = "";
    string cfgIsSandbox = "";
    string cfgurlSandbox = "";
    string cfgurlProduction = "";
    string cfgBINoDelivery = "";  //Use this as BI Invoice Onlly
    string cfgOutgoingFilter = "";
    string cfgAccountForOutgoingPayment;
    string cfgNegativeIncomingPayment;
    string cfgAutoSelectAccount;
    string cfgVisaIncomingAccount;
    string cfgVisaOutgoingAccount;
    string cfgBatchQuery;
    string cfgMasterIncomingAccount;
    string cfgMasterOutgoingAccount;

    string cfgAMXIncomingAccount;
    string cfgAMXOutgoingAccount;

    string cfgDiscIncomingAccount;
    string cfgDiscOutgoingAccount;

    string cfgDinnIncomingAccount;
    string cfgDinnOutgoingAccount;

    string cfgJCBIncomingAccount;
    string cfgJCBOutgoingAccount;

    string cfgeCheckIncomingAccount;
    string cfgeCheckOutgoingAccount;


    string cfgOtherIncomingAccount;
    string cfgOtherOutgoingAccount;
    string cfgsendCustomerEmail;
    string cfgimportAccount;
    // string cfgimportPin;
    string cfgimportSourceKey;
    //string cfgimportCCWebServiceUrl;
    string cfgimportBPGroupName;
    string cfgimportBPHeaderCode;
    string cfgimportItemHeaderCode;
    string cfgimportTransStatus;
    string cfgimportTransType;
    string cfgimportInvoiceType;
    string cfgimportUseZipForTax;
    string cfgimportBPNextCode;
    string cfgProcessPreAuth;
    string cfgMerchantReceipt;
    string cfgEmailSubject;
    string cfgCustomerReceipt;

    string cfgRecurringBillingReceipt;
    string cfgDeclineReceipt;
    string cfgDeclineMerchantReceipt;
    string cfgmerchantEmail;
    string cfgaddressRequired;
    string cfgMerchantID;
    string cfgUseUserEmail;
    string cfgSaveReceipt;
    string cfgBatchAutoMode;
    string cfgrecurringBilling;
    string cfgBtnPos;
    string cfgDefaultTaxRate;
    string cfgNoTransOnInv;
    string cfgPMAutoSel;
    string cfgImportCust;
    string cfgshowStartup = "Y";
    string cfgBatchPreauth;
    string cfgCaptureOnPrint;
    string cfgApprUDF;
    string cfgCaptureOnly;
    string cfgAmountField;
    string cfgFieldSize;
    string cfgHasTerminal;
    string cfgCardNumUDF;
    string cfgPreAuthEmail;
    string cfgAutoPreauth;
    string cfgCardEMTemplate;
    string cfgAutoSync;
    string cfgAcctNameField = "";
    string cfgConnectUseDP;
    string cfgDefaultContactID = "";
    string cfgReceiptEmailUDF = "";
    #endregion
    private void SoIntegrationCfg()
    {
        try
        {
            AddXMLForm("Presentation_Layer.Menus.xml");
            new Thread(AddConfiguration).Start();
        }
        catch (Exception ex)
        {
            var Error = ex.Message;
            errorLog(ex);
        }
    }
    private void AddConfiguration()
    {
        try
        {
            AddColumns("OCRD", U_EbizChargeKey, "EbizCharge Key", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("OITM", U_EbizChargeKey, "EbizCharge Key", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("ORDR", U_EbizChargeKey, "EbizCharge Key", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("ORDR", U_EbizChargeURL, "EbizCharge URL", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("ORDR", "ConnectDocNum", "Connect DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("ORDR", "RefNum", "Connect RefNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("OINV", "RefNum", "Connect RefNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("ORDR", U_EbizChargeMarkPayment, "EbizCharge Payment Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, SAPbobsCOM.BoFldSubTypes.st_None, "");

            addQryCategor("EbizCharge");
            var Qsql = String.Format("SELECT T0.[CardCode], T0.[CardName],T0.[Balance],T0.[E_Mail] FROM OCRD T0 WHERE T0.[CardType] = 'C'").Replace("]", "\"").Replace("[", "\"").Replace("'", "''");
            addFms(formEmailEstimates, "txtBpCode", "-1", Qsql);
            addFms(formEmailPaymentSO, "txtBpCode", "-1", Qsql);
            addFms(formExportSO, "txtBp", "-1", Qsql);
            addFms(formSyncUpload, "edBPId", "-1", Qsql);
            addFms(formSyncUpload, "edBp2", "-1", Qsql);
            var SOsql = String.Format("Select * from (select [DocNum[ as [Sales Order#[,[DocDate[,[DocTotal[,[CardCode[,[CardName[,(select[E_Mail[from OCRD where [CardCode[ = ORDR.[CardCode[) as Email,[DocTotal[-IFNULL((SELECT  sum(T0.[DpmAmnt[) FROM ODPI T0  INNER JOIN DPI1 T1 ON T0.[DocEntry[ = T1.[DocEntry[ inner join RDR1 on RDR1.[DocEntry[ = T1.[BaseEntry[   where RDR1.[DocEntry[ = ORDR.[DocEntry[and T0.[DocStatus[ = 'O' and T0.[CANCELED[ = 'N'), 0) as [AmtDue[    from ORDR where  [DocStatus[= 'O'  and[DocTotal[> 0 )T where T.[AmtDue[> 0 order by 1 ").Replace("[","\"").Replace("'", "''");
            addFms(formEmailPaymentSO, "txtSO", "-1", SOsql);
            String sql = "";
            if (!CheckTable(String.Format("select * from  \"{0}\".\"@CCSOLOG\" ", oCompany.CompanyDB)))
            {
                sql = "CREATE COLUMN TABLE [@CCSOLOG[ ([Code[ NVARCHAR(50) NOT NULL ,     [Name[NVARCHAR(100) NOT NULL,[U_DocNum[NVARCHAR(50),[U_DocType[NVARCHAR(50),                [U_Status[ NVARCHAR(50),     [U_PaymentStatus[NVARCHAR(50),     [U_CreateDt[LONGDATE CS_LONGDATE,     [U_UpdateDt[LONGDATE CS_LONGDATE,     [U_RefNum[NVARCHAR(50),                   [U_CustomerID[NVARCHAR(50),     [U_DocEntry[NVARCHAR(50),     [U_PaidAmt[INTEGER CS_INT,[U_DocState[ NVARCHAR(30),     [U_Balance[INTEGER CS_INT,    PRIMARY KEY([Code[)) UNLOAD PRIORITY 5 AUTO MERGE;            CREATE UNIQUE INDEX[KCCSOLOG_NAME[ON[@CCSOLOG[( [Name[ASC )".Replace("[", "\"");
                execute(sql);
                sql = " CREATE UNIQUE INDEX[KCCSOLOG_NAME[ON[@CCSOLOG[( [Name[ASC )".Replace("[", "\"");
                execute(sql);
                sql = "ALTER TABLE [@CCSOLOG[ ADD ([U_DocState[ NVARCHAR(30))".Replace("[", "\"");
                execute(sql);
            }


            String Table = "CCINVLOG";
            if (!CheckTable(String.Format("select * from  \"{0}\".\"@{1}\" ", oCompany.CompanyDB, Table)))
            {
                sql = "CREATE COLUMN TABLE [@CCINVLOG[ ([Code[ NVARCHAR(50) NOT NULL ,	 [Name[ NVARCHAR(100) NOT NULL ,	 [U_DocNum[ NVARCHAR(50),     [U_DocEntry[NVARCHAR(50),  [U_CustomerID[NVARCHAR(50),    [U_Balance[INTEGER CS_INT,     [U_PaidAmt[INTEGER CS_INT,     [U_RefNum[NVARCHAR(50),    [U_DocType[NVARCHAR(50),   [U_Status[NVARCHAR(50),    [U_PaymentStatus[NVARCHAR(50),     [U_CreateDt[LONGDATE CS_LONGDATE,  [U_UpdateDt[LONGDATE CS_LONGDATE, PRIMARY KEY([Code[)) UNLOAD PRIORITY 5 AUTO MERGE;                CREATE UNIQUE INDEX[KCCINVLOG_NAME[ON[@CCINVLOG[( [Name[ASC ); ".Replace("[", "\"");
                execute(sql);
                sql = "CREATE UNIQUE INDEX[KCCINVLOG_NAME[ON[@CCINVLOG[( [Name[ASC ); ".Replace("[", "\"");
                execute(sql);
            }
            Table = "CCSO";
            if (!CheckTable(String.Format("select * from  \"{0}\".\"@{1}\" ", oCompany.CompanyDB, Table)))
            {
                sql = "CREATE COLUMN TABLE [@CCSO[ ([Code[ NVARCHAR(50) NOT NULL ,[Name[ NVARCHAR(100) NOT NULL ,	 [U_DocNum[ NVARCHAR(50),	 [U_DocType[ NVARCHAR(50),	 [U_Status[ NVARCHAR(50),	 [U_Type[ NVARCHAR(50),	 [U_DocPaidDate[ NVARCHAR(50),	 [U_Email[ NVARCHAR(150),	 [U_MethodDigit[ NVARCHAR(30),	 [U_CardCode[ NVARCHAR(30),	 [U_AmountPaid[ NVARCHAR(30),	 [U_CreateDt[ LONGDATE CS_LONGDATE,	 [U_UpdateDt[ LONGDATE CS_LONGDATE,	 PRIMARY KEY ([Code[)) UNLOAD PRIORITY 5 AUTO MERGE ;CREATE UNIQUE INDEX [KCCSO_NAME[ ON [@CCSO[ ( [Name[ ASC );".Replace("[", "\"");
                execute(sql);
                sql = "CREATE UNIQUE INDEX [KCCSO_NAME[ ON [@CCSO[ ( [Name[ ASC );".Replace("[", "\"");
                execute(sql);
            }
            Table = "CCITEMLOGTAB";
            if (!CheckTable(String.Format("select * from  \"{0}\".\"@{1}\" ", oCompany.CompanyDB, Table)))
            {
                sql = "CREATE COLUMN TABLE [@CCITEMLOGTAB[ ([Code[ NVARCHAR(50) NOT NULL ,[Name[ NVARCHAR(100) NOT NULL ,[U_ItemCode[ NVARCHAR(50),[U_CreateDt[ TIMESTAMP,[U_UpdateDt[ TIMESTAMP,[U_EbizKey[ NVARCHAR(120),[U_Status[ NVARCHAR(254),PRIMARY KEY ([Code[)) UNLOAD PRIORITY 5 AUTO MERGE ;".Replace("[", "\"");
                execute(sql);
                sql = "CREATE UNIQUE INDEX [KCCITEMLOGTAB_NAME[ ON [@CCITEMLOGTAB[ ( [Name[ ASC );".Replace("[", "\"");
                execute(sql);
            }


            Table = "CCBPIMP";
            if (!CheckTable(String.Format("select * from  \"{0}\".\"@{1}\" ", oCompany.CompanyDB, Table)))
            {
                sql = "CREATE COLUMN TABLE [@CCBPIMP[ ([Code[ NVARCHAR(50) NOT NULL ,[Name[ NVARCHAR(100) NOT NULL ,[U_CardCode[ NVARCHAR(50),[U_CardName[ NVARCHAR(254),[U_Email[ NVARCHAR(254),[U_Balance[ INTEGER CS_INT,[U_Software[ NVARCHAR(254),[U_Status[ NVARCHAR(254),[U_CreateDt[ LONGDATE CS_LONGDATE,[U_UpdateDt[ LONGDATE CS_LONGDATE,PRIMARY KEY ([Code[)) UNLOAD PRIORITY 5 AUTO MERGE; ".Replace("[", "\"");
                execute(sql);
                sql = "CREATE UNIQUE INDEX [KCCBPIMP_NAME[ ON [@CCBPIMP[([Name[ ASC);".Replace("[", "\"");
                execute(sql);
            }
            Table = "CCITMIMP";
            if (!CheckTable(String.Format("select * from  \"{0}\".\"@{1}\" ", oCompany.CompanyDB, Table)))
            {
                sql = "CREATE COLUMN TABLE [@CCITMIMP[ ([Code[ NVARCHAR(50) NOT NULL ,	 [Name[ NVARCHAR(100) NOT NULL ,	 [U_ItemCode[ NVARCHAR(50),	 [U_ItemName[ NVARCHAR(254),	 [U_Quantity[ INTEGER CS_INT,	 [U_UnitPrice[ INTEGER CS_INT,	 [U_Software[ NVARCHAR(254),	 [U_Status[ NVARCHAR(254),	 [U_CreateDt[ LONGDATE CS_LONGDATE,	 [U_UpdateDt[ LONGDATE CS_LONGDATE,	 PRIMARY KEY ([Code[)) UNLOAD PRIORITY 5 AUTO MERGE".Replace("[", "\"");
                execute(sql);
                sql = "CREATE UNIQUE INDEX [KCCITMIMP_NAME[ ON [@CCITMIMP[ ( [Name[ ASC )".Replace("[", "\"");
                execute(sql);
            }
            Table = "CCSOIMP";
            if (!CheckTable(String.Format("select * from  \"{0}\".\"@{1}\" ", oCompany.CompanyDB, Table)))
            {
                sql = "CREATE COLUMN TABLE [@CCSOIMP[ ([Code[ NVARCHAR(50) NOT NULL ,	 [Name[ NVARCHAR(100) NOT NULL ,	 [U_DocNum[ NVARCHAR(50),	 [U_DocEntry[ NVARCHAR(50),	 [U_CardCode[ NVARCHAR(254),	 [U_Customer[ NVARCHAR(254),	 [U_DatePaid[ LONGDATE CS_LONGDATE,	 [U_Software[ NVARCHAR(254),	 [U_Status[ NVARCHAR(254),	 [U_CreateDt[ LONGDATE CS_LONGDATE,	 [U_UpdateDt[ LONGDATE CS_LONGDATE,	 PRIMARY KEY ([Code[)) UNLOAD PRIORITY 5 AUTO MERGE ;".Replace("[", "\"");
                execute(sql);
                sql = "CREATE UNIQUE INDEX [KCCSOIMP_NAME[ ON [@CCSOIMP[ ( [Name[ ASC )".Replace("[", "\"");
                execute(sql);
            }
            Table = "CCBPLOGTAB";
            if (!CheckTable(String.Format("select * from  \"{0}\".\"@{1}\" ", oCompany.CompanyDB, Table)))
            {
                sql = "CREATE COLUMN TABLE [@CCBPLOGTAB[ ([Code[ NVARCHAR(50) NOT NULL ,	 [Name[ NVARCHAR(100) NOT NULL ,	 [U_CardCode[ NVARCHAR(50),	 [U_Status[ NVARCHAR(254),	 [U_EbizKey[ NVARCHAR(100),	 [U_CreateDt[ LONGDATE CS_LONGDATE,	 [U_UpdateDt[ LONGDATE CS_LONGDATE,	 PRIMARY KEY ([Code[)) UNLOAD PRIORITY 5 AUTO MERGE ;".Replace("[", "\"");
                execute(sql);
                sql = "CREATE UNIQUE INDEX [KCCBPLOGTAB_NAME[ ON [@CCBPLOGTAB[ ( [Name[ ASC )".Replace("[", "\"");
                execute(sql);
            }
            /*
                      AddTable("CCSOLOG", "Ebiz Sales Order Log", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                      AddColumns("@CCSOLOG", "DocNum", "DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
                      AddColumns("@CCSOLOG", "DocType", "DocType", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
                      AddColumns("@CCSOLOG", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
                      AddColumns("@CCSOLOG", "PaymentStatus", "PaymentStatus", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
                      AddColumns("@CCSOLOG", "CreateDt", "CreateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
                      AddColumns("@CCSOLOG", "UpdateDt", "UpdateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
                      AddColumns("@CCSOLOG", "RefNum", "RefNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
                      AddColumns("@CCSOLOG", "CustomerID", "CustomerID", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
                      AddColumns("@CCSOLOG", "DocEntry", "DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
                      AddColumns("@CCSOLOG", "PaidAmt", "PaidAmt", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, SAPbobsCOM.BoFldSubTypes.st_None, "");
                      AddColumns("@CCSOLOG", "Balance", "Balance", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, SAPbobsCOM.BoFldSubTypes.st_None, "");
          */
            /*
           AddTable("CCINVLOG", "Ebiz Invoice Log", SAPbobsCOM.BoUTBTableType.bott_NoObject);
           AddColumns("@CCINVLOG", "DocNum", "DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");

           AddColumns("@CCINVLOG", "DocEntry", "DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
           AddColumns("@CCINVLOG", "CustomerID", "CustomerID", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
           AddColumns("@CCINVLOG", "Balance", "Balance", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, SAPbobsCOM.BoFldSubTypes.st_None, "");
           AddColumns("@CCINVLOG", "PaidAmt", "PaidAmt", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, SAPbobsCOM.BoFldSubTypes.st_None, "");
           AddColumns("@CCINVLOG", "RefNum", "RefNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");

           AddColumns("@CCINVLOG", "DocType", "DocType", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
           AddColumns("@CCINVLOG", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
           AddColumns("@CCINVLOG", "PaymentStatus", "PaymentStatus", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
           AddColumns("@CCINVLOG", "CreateDt", "CreateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
           AddColumns("@CCINVLOG", "UpdateDt", "UpdateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
       */
            /*
            AddTable("CCSO", "Ebiz SO Connect Log", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            AddColumns("@CCSO", "DocNum", "DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCSO", "DocType", "DocType", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCSO", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCSO", "Type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCSO", "DocPaidDate", "DocPaidDate", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCSO", "Email", "Email", SAPbobsCOM.BoFieldTypes.db_Alpha, 150, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCSO", "MethodDigit", "MethodDigit", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCSO", "CardCode", "CardCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCSO", "AmountPaid", "AmountPaid", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCSO", "CreateDt", "CreateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCSO", "UpdateDt", "UpdateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");

            AddTable("CCITEMLOGTAB", "Ebiz Item Log ", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            AddColumns("@CCITEMLOGTAB", "ItemCode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCITEMLOGTAB", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCITEMLOGTAB", "EbizKey", "EbizKey", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCITEMLOGTAB", "CreateDt", "CreateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCITEMLOGTAB", "UpdateDt", "UpdateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");

            AddTable("CCBPLOGTAB", "Ebiz Business Partner Log ", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            AddColumns("@CCBPLOGTAB", "CardCode", "CardCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCBPLOGTAB", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCBPLOGTAB", "EbizKey", "EbizKey", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCBPLOGTAB", "CreateDt", "CreateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns("@CCBPLOGTAB", "UpdateDt", "UpdateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");

            var Table = "CCBPIMP";
            var TableUDT = "@" + "CCBPIMP";
            AddTable(Table, "Ebiz BP Import", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            AddColumns(TableUDT, "CardCode", "CardCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "CardName", "CardName", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "Email", "Email", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "Balance", "Balance", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "Software", "Software", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "CreateDt", "CreateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "UpdateDt", "UpdateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");

            Table = "CCITMIMP";
            TableUDT = "@" + "CCITMIMP";
            AddTable(Table, "Ebiz Item Import", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            AddColumns(TableUDT, "ItemCode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "ItemName", "ItemName", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "UnitPrice", "UnitPrice", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "Software", "Software", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "CreateDt", "CreateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "UpdateDt", "UpdateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");

            Table = "CCSOIMP";
            TableUDT = "@" + "CCSOIMP";//[U_DocNum],[U_DocEntry],[U_CardCode],[U_Customer],[U_DatePaid],[U_Software],[U_Status],[U_CreateDt]
            AddTable(Table, "Ebiz SO Import", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            AddColumns(TableUDT, "DocNum", "DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "DocEntry", "DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "CardCode", "CardCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "Customer", "Customer", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "DatePaid", "DatePaid", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "Software", "Software", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "CreateDt", "CreateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "UpdateDt", "UpdateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");

            Table = "CCSOPIMP";
            TableUDT = "@" + "CCSOPIMP";//[U_DocNum],[U_DocEntry],[U_CardCode],[U_Customer],[U_PaymentType],[U_AmtPaid],[U_DatePaid],[U_Software],[U_Status],[U_CreateDt]
            AddTable(Table, "Ebiz SO payments Import", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            AddColumns(TableUDT, "DocNum", "DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "DocEntry", "DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "CardCode", "CardCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "Customer", "Customer", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "PaymentType", "PaymentType", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "AmtPaid", "AmtPaid", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "DatePaid", "DatePaid", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "Software", "Software", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "CreateDt", "CreateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            AddColumns(TableUDT, "UpdateDt", "UpdateDt", SAPbobsCOM.BoFieldTypes.db_Date, 50, SAPbobsCOM.BoFldSubTypes.st_None, "");
            */
        }
        catch (Exception ex)
        {
            var Error = ex.Message;
            errorLog(ex);
        }
    }
    private void AddCFGMenuItems()
    {
        try
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = SBO_Application.Menus.Item("43525"); // credit Card Setup
            oMenus = oMenuItem.SubMenus;
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
            oCreationPackage.UniqueID = menuCfg;
            oCreationPackage.String = "eBizCharge Configuration";
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
    private void CreateCfgForm()
    {
        try
        {



            // add a new form
            SAPbouiCOM.FormCreationParams oCreationParams = null;

            oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.FormType = formCfg;
            oCreationParams.UniqueID = formCfg;
            try
            {
                oCfgForm = SBO_Application.Forms.AddEx(oCreationParams);
            }
            catch (Exception ex)
            {
                errorLog(ex);
                oCfgForm = SBO_Application.Forms.Item(formCfg);

            }
            // set the form properties
            oCfgForm.Title = "eBizCharge Configuration";
            oCfgForm.Left = 400;
            oCfgForm.Top = 100;
            oCfgForm.ClientHeight = 680;
            oCfgForm.ClientWidth = 580;

            AddCFGFormField(oCfgForm);

            oCfgForm.Visible = true;

            setFormEditVal(oCfgForm, editPin, cfgPin);
            setFormEditVal(oCfgForm, editSourceKey, cfgSourceKey);
            setFormEditVal(oCfgForm, editSandbox, cfgIsSandbox);
            setFormEditVal(oCfgForm, editNegativeIncomingPayment, cfgNegativeIncomingPayment);

            setFormEditVal(oCfgForm, editProcessPreAuth, cfgProcessPreAuth);
            setFormEditVal(oCfgForm, editBatchAutoMode, cfgBatchAutoMode);

            setFormEditVal(oCfgForm, editimportSourceKey, cfgimportSourceKey);
            setFormEditVal(oCfgForm, editsendCustomerEmail, cfgsendCustomerEmail);
            setFormEditVal(oCfgForm, editmerchantEmail, cfgmerchantEmail);
            setFormEditVal(oCfgForm, editMerchantReceipt, cfgMerchantReceipt);
            setFormEditVal(oCfgForm, editEmailSubject, cfgEmailSubject);
            setFormEditVal(oCfgForm, editCustomerReceipt, cfgCustomerReceipt);
            // setFormEditVal(oCfgForm, editRecurringBillingReceipt, cfgRecurringBillingReceipt);
            // setFormEditVal(oCfgForm, editDeclineReceipt, cfgDeclineReceipt);
            //  setFormEditVal(oCfgForm, editDeclineMerchantReceipt, cfgDeclineMerchantReceipt);
            setFormEditVal(oCfgForm, editaddressRequired, cfgaddressRequired);
            setFormEditVal(oCfgForm, editMerchantID, cfgMerchantID);
            setFormEditVal(oCfgForm, editUseUserEmail, cfgUseUserEmail);
            setFormEditVal(oCfgForm, editSaveReceipt, cfgSaveReceipt);
            setFormEditVal(oCfgForm, editrecurringBilling, cfgrecurringBilling);
            setFormEditVal(oCfgForm, editBtnPos, cfgBtnPos);
            setFormEditVal(oCfgForm, editDefaultTaxRate, cfgDefaultTaxRate);
            setFormEditVal(oCfgForm, editNoTransOnInv, cfgNoTransOnInv);
            setFormEditVal(oCfgForm, editImportCust, cfgImportCust);
            setFormEditVal(oCfgForm, editshowStartup, cfgshowStartup);
            setFormEditVal(oCfgForm, editBatchPreauth, cfgBatchPreauth);
            setFormEditVal(oCfgForm, editDeclineUnapprove, cfgDeclineUnapprove);
            setFormEditVal(oCfgForm, editCaptureOnPrint, cfgCaptureOnPrint);
            setFormEditVal(oCfgForm, editApprUDF, cfgApprUDF);
            setFormEditVal(oCfgForm, editCaptureOnly, cfgCaptureOnly);
            setFormEditVal(oCfgForm, editAmountField, cfgAmountField);
            setFormEditVal(oCfgForm, editFieldSize, cfgFieldSize);
            setFormEditVal(oCfgForm, editHasTerminal, cfgHasTerminal);
            setFormEditVal(oCfgForm, editBINoDelivery, cfgBINoDelivery);
            setFormEditVal(oCfgForm, editPreAuthEmail, cfgPreAuthEmail);
            setFormEditVal(oCfgForm, editAutoPreauth, cfgAutoPreauth);
            setFormEditVal(oCfgForm, editCardEMTemplate, cfgCardEMTemplate);

        }
        catch (Exception ex)
        {

            errorLog(ex);
        }

    }
    private void validDefaultTax()
    {
        try
        {
            if (decimal.Parse(cfgDefaultTaxRate) < 1 && decimal.Parse(cfgDefaultTaxRate) != 0)
                cfgDefaultTaxRate = "8";
        }
        catch (Exception)
        {
            cfgDefaultTaxRate = "8";
        }
    }
    private void CFGFormHandler(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;
        try
        {
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        break;
                }
            }
            else
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        initSetupCardAcct();
                        break;

                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:

                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:

                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        HandleCFGPress(form, pVal);
                        break;
                }

            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }

    }
    private bool Validate(SAPbouiCOM.Form form)
    {
        try
        {
            string errStr = "";
            if (!(getFormItemVal(form, editNegativeIncomingPayment) == "Y" || getFormItemVal(form, editNegativeIncomingPayment) == "N"))
                errStr += "Please enter Y or N for Negative Incoming Payment option.\r\n";

            if (errStr != "")
            {
                SBO_Application.MessageBox(errStr);
                return false;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return true;
    }
    private void HandleCFGPress(SAPbouiCOM.Form form, SAPbouiCOM.ItemEvent pVal)
    {
        try
        {

            switch (pVal.ItemUID)
            {
                case btnUpdate:
                    if (!Validate(form))
                    {
                        return;
                    }
                    if (SBO_Application.MessageBox("Update eBizCharge configuration?", 1, "Yes", "No") == 1)
                    {

                        cfgPin = getFormItemVal(form, editPin);
                        cfgSourceKey = getFormItemVal(form, editSourceKey);
                        cfgIsSandbox = getFormItemVal(form, editSandbox);
                        if (cfgIsSandbox == "Y")
                            cfgUrl = cfgurlSandbox;
                        else
                            cfgUrl = cfgurlProduction;
                        cfgNegativeIncomingPayment = getFormItemVal(form, editNegativeIncomingPayment);
                        cfgAutoSelectAccount = "Y";  // getFormItemVal(form, editAutoSelectAccount);
                        cfgimportSourceKey = getFormItemVal(form, editimportSourceKey);
                        cfgProcessPreAuth = getFormItemVal(form, editProcessPreAuth);
                        cfgsendCustomerEmail = getFormItemVal(form, editsendCustomerEmail);
                        cfgmerchantEmail = getFormItemVal(form, editmerchantEmail);
                        cfgEmailSubject = getFormItemVal(form, editEmailSubject);
                        cfgMerchantReceipt = getFormItemVal(form, editMerchantReceipt);
                        cfgCustomerReceipt = getFormItemVal(form, editCustomerReceipt);
                        //    cfgRecurringBillingReceipt = getFormItemVal(form, editRecurringBillingReceipt);
                        //   cfgDeclineReceipt = getFormItemVal(form, editDeclineReceipt);
                        //    cfgDeclineMerchantReceipt = getFormItemVal(form, editDeclineMerchantReceipt);
                        cfgMerchantID = getFormItemVal(form, editMerchantID);
                        cfgUseUserEmail = getFormItemVal(form, editUseUserEmail);
                        cfgSaveReceipt = getFormItemVal(form, editSaveReceipt);
                        cfgaddressRequired = getFormItemVal(form, editaddressRequired);
                        cfgBatchAutoMode = getFormItemVal(form, editBatchAutoMode);
                        cfgDefaultTaxRate = getFormItemVal(form, editDefaultTaxRate);
                        cfgNoTransOnInv = getFormItemVal(form, editNoTransOnInv);
                        cfgImportCust = getFormItemVal(form, editImportCust);
                        cfgshowStartup = getFormItemVal(form, editshowStartup);
                        cfgBatchPreauth = getFormItemVal(form, editBatchPreauth);
                        cfgDeclineUnapprove = getFormItemVal(form, editDeclineUnapprove);
                        cfgCaptureOnPrint = getFormItemVal(form, editCaptureOnPrint);
                        cfgrecurringBilling = getFormItemVal(form, editrecurringBilling);
                        cfgBtnPos = getFormItemVal(form, editBtnPos);
                        cfgApprUDF = getFormItemVal(form, editApprUDF);
                        cfgCaptureOnly = getFormItemVal(form, editCaptureOnly);
                        cfgAmountField = getFormItemVal(form, editAmountField);
                        cfgFieldSize = getFormItemVal(form, editFieldSize);
                        cfgHasTerminal = getFormItemVal(form, editHasTerminal);
                        cfgBINoDelivery = getFormItemVal(form, editBINoDelivery);
                        cfgPreAuthEmail = getFormItemVal(form, editPreAuthEmail);
                        cfgAutoPreauth = getFormItemVal(form, editAutoPreauth);
                        cfgCardEMTemplate = getFormItemVal(form, editCardEMTemplate);
                        validDefaultTax();
                        updateConfig();
                        if (!updateAcctAcct())
                            return;
                        showMessage("eBizCharge configuration updated.");
                        form.Close();
                        if (cfgProcessPreAuth == "Y")
                            AddBatchInvMenu();
                        //if (cfgPayOnTerm == "Y")
                        //AddPayOnTermMenu();

                        if (cfgimportSourceKey == "Y")
                            AddOrderImpMenu();
                        if (cfgrecurringBilling == "Y")
                            AddRBillingMenu();
                        if (cfgMerchantID != "")
                            AddeBizConnectMenu();
                        if (cfgImportCust == "Y")
                            AddCustImpMenu();
                        createAcctList();
                        getEmailTemplate();
                        if (ListEmailTemplate.Count > 0)
                            AddePaymentFormMenu();
                        try
                        {

                            B1Info.SourceKey = cfgMerchantID;
                            B1Info.PIN = ebizVersion;
                            B1Info.DBType = ((int)oCompany.DbServerType).ToString();
                            B1Info.sandbox = oCompany.Version.ToString();
                            string result = install.UpsertInfo(B1Info);
                        }
                        catch (Exception) { }
                    }
                    break;
                case btnClose:
                    if (oCfgForm != null)
                        oCfgForm.Close();
                    break;
                case btnAdd:
                    AddToWalkInList();
                    LoadWalkInList();
                    break;
                case btnDelete:
                    DeleteFromWalkInList();
                    LoadWalkInList();
                    break;
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    private bool updateAcctAcct()
    {
        bool bRet = true;
        try
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oCfgForm.Items.Item(mxCardAcct).Specific;
            if (oMatrix.RowCount >= 2)
            {
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    string id = getMatrixItem(oMatrix, "AcctID", i);
                    string group = getMatrixItem(oMatrix, "GroupName", i);
                    string cardtype = getMatrixItem(oMatrix, "CardType", i);
                    string pin = ""; // getMatrixItem(oMatrix, "Pin", i);
                    string sourceKey = getMatrixItem(oMatrix, "SourceKey", i);
                    string currency = getMatrixItem(oMatrix, "Currency", i);
                    string BranchID = getMatrixItem(oMatrix, "BranchID", i);
                    updateCardAcct(id, group, cardtype, pin, sourceKey, currency, BranchID);
                }

            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return bRet;
    }
    SAPbouiCOM.Item otheEditRefItem = null;
    SAPbouiCOM.Item otheButtonRefItem = null;
    private void AddCFGFormField(SAPbouiCOM.Form form)
    {
        try
        {
            int edL = 140; //oItm.Left + oItm.Width;
            int edW = 120;
            int edT = 20;
            int edH = 15;
            int nGap = 17;
            int nSpan = 250;

            if (otheEditRefItem == null)
            {
                try
                {
                    if (theActiveForm.TypeEx == FORMSALESORDER)
                    {
                        otheEditRefItem = theActiveForm.Items.Item("7");
                        otheButtonRefItem = theActiveForm.Items.Item("1");
                    }

                }
                catch (Exception) { }
            }
            if (otheEditRefItem != null)
            {
                edW = otheEditRefItem.Width;
                edL = otheEditRefItem.Left + otheEditRefItem.Width;
                edH = otheEditRefItem.Height;
                cfgFieldSize = string.Format("{0},{1},{2}", edW, edL, edH);
                setFormEditVal(form, editFieldSize, cfgFieldSize);
                updateConfig();
            }
            if (cfgFieldSize != null && cfgFieldSize != "")
            {
                try
                {
                    string[] wlh = cfgFieldSize.Split(',');
                    edW = int.Parse(wlh[0]);
                    edL = int.Parse(wlh[1]);
                    edH = int.Parse(wlh[2]);
                    nGap = edH + 2;
                    nSpan = edW * 2;
                    form.ClientWidth = edW * 5;
                    form.ClientHeight = edH * 45;
                }
                catch (Exception) { }
            }

            SAPbouiCOM.Item oItm = null;
            SAPbouiCOM.Button oButton = null;
            oItm = addPaneItem(form, editMerchantID, edL, edT, edW * 2, edH, "Security Key:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1002, edW);
            oItm = addPaneItem(form, editPin, edL, edT, edW * 2, edH, "Pin:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2, edW);
            oItm.Visible = false;
            form.Items.Item("LB2").Visible = false;
            oItm = addPaneItem(form, editSourceKey, edL, edT, edW * 2, edH, "SourceKey:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 3, edW);
            oItm.Visible = false;
            form.Items.Item("LB3").Visible = false;
            oItm = addPaneItem(form, editSandbox, edL, edT, edW * 2, edH, "Sandbox(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 4, edW);
            oItm.Visible = false;
            form.Items.Item("LB4").Visible = false;

            edT = oItm.Top + nGap + 2;
            oItm = addPaneItem(form, editmerchantEmail, edL, edT, edW * 2, edH, "Merchant Email:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 110, edW);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editEmailSubject, edL, edT, edW * 2, edH, "Email Subject:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 11290, edW);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editsendCustomerEmail, edL, edT, edW, edH, "Send Receipt(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 109, edW);
            oItm = addPaneItem(form, editCustomerReceipt, edL + nSpan, edT, edW, edH, "Declined Receipt:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 203, edW);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editMerchantReceipt, edL, edT, edW, edH, "Merchant Receipt:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 201, edW);
            oItm = addPaneItem(form, editUseUserEmail, edL + nSpan, edT, edW, edH, "Void Notify(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1001, edW);


            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editNegativeIncomingPayment, edL, edT, edW, edH, "Neg. Payment(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 5, edW);
            oItm = addPaneItem(form, editProcessPreAuth, edL + nSpan, edT, edW, edH, "Proc. PreAuth(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 8, edW);
            edT = oItm.Top + nGap;


            oItm = addPaneItem(form, editBatchAutoMode, edL, edT, edW, edH, "Batch Auto(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1251, edW);
            oItm = addPaneItem(form, editaddressRequired, edL + nSpan, edT, edW, edH, "Addr. Required(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 159, edW);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editrecurringBilling, edL, edT, edW, edH, "Recur. Billing(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2001, edW);
            oItm = addPaneItem(form, editBtnPos, edL + nSpan, edT, edW, edH, "Button Position:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2002, edW);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editDefaultTaxRate, edL, edT, edW, edH, "Def. Tax Rate:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2006, edW);
            oItm = addPaneItem(form, editNoTransOnInv, edL + nSpan, edT, edW, edH, "No Trans On Inv(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2007, edW);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editImportCust, edL, edT, edW, edH, "Import method(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2236, edW);
            oItm = addPaneItem(form, editshowStartup, edL + nSpan, edT, edW, edH, "Show Startup(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2087, edW);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editimportSourceKey, edL, edT, edW, edH, "Import Order(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2256, edW);
            oItm = addPaneItem(form, editBatchPreauth, edL + nSpan, edT, edW, edH, "Batch Preauth(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2257, edW);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editDeclineUnapprove, edL, edT, edW, edH, "Decline Unappr.(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2259, edW);
            oItm = addPaneItem(form, editCaptureOnPrint, edL + nSpan, edT, edW, edH, "Capture on Add(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2288, edW);

            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editApprUDF, edL, edT, edW, edH, "Approved UDF:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2280, edW);
            oItm = addPaneItem(form, editCaptureOnly, edL + nSpan, edT, edW, edH, "BI Capture Only(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2291, edW);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editAmountField, edL, edT, edW, edH, "Amount Field:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2285, edW);
            oItm = addPaneItem(form, editFieldSize, edL + nSpan, edT, edW, edH, "Field Size:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2293, edW);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editSaveReceipt, edL, edT, edW, edH, "Print Receipt(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 500, edW);

            oItm = addPaneItem(form, editHasTerminal, edL + nSpan, edT, edW, edH, "Has Terminal(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 2286, edW);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editBINoDelivery, edL, edT, edW, edH, "BI No Delivery(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 50022, edW);
            oItm = addPaneItem(form, editPreAuthEmail, edL + nSpan, edT, edW, edH, "Pre-auth email(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 26486, edW);
            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, editAutoPreauth, edL, edT, edW, edH, "Auto Preauth(Y/N):", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 50233, edW);
            oItm = addPaneItem(form, editCardEMTemplate, edL + nSpan, edT, edW, edH, "Card EM Template:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 36486, edW);

            edT = oItm.Top + nGap + 2;
            oItm = addPaneItem(form, editWalkIn, edL, edT, edW, edH, "Walk-In:", SAPbouiCOM.BoFormItemTypes.it_EDIT, 0, 1258);
            oItm = addPaneItem(form, btnAdd, edL + edW, edT, 60, edH, "Add", SAPbouiCOM.BoFormItemTypes.it_BUTTON, 0, 0);

            edT = oItm.Top + nGap;
            oItm = addPaneItem(form, cbWalkIn, edL, edT, edW, edH, "Walk-In List:", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 0, 1259);
            oItm = addPaneItem(form, btnDelete, edL + edW, edT, 60, edH, "Delete", SAPbouiCOM.BoFormItemTypes.it_BUTTON, 0, 0);
            edT = oItm.Top + nGap + 10;



            int lMargin = 25;
            edT = oItm.Top + nGap;
            oItm = form.Items.Add("lbAcct", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItm.Left = lMargin;
            oItm.Top = edT;
            oItm.Width = 200;
            oItm.Height = 16;
            setLabelCaption(oCfgForm, "lbAcct", "Credit Card Setup:");
            edT = oItm.Top + nGap;
            oItm = oCfgForm.Items.Add(mxCardAcct, SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oItm.FromPane = 0;
            oItm.ToPane = 0;
            oItm.Left = lMargin;
            oItm.Width = oCfgForm.ClientWidth - 2 * lMargin;
            oItm.Top = edT;
            oItm.Height = oCfgForm.ClientHeight / 3;
            edT = edT + oItm.Height;
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oCfgForm.Items.Item(mxCardAcct).Specific;
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Add("AcctID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "";
            oColumn.Width = 40;

            oColumn = oMatrix.Columns.Add("CardName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Card Name";
            oColumn.Width = 80;

            oColumn = oMatrix.Columns.Add("GroupName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Customer Group";
            oColumn.Width = 60;

            oColumn = oMatrix.Columns.Add("CardType", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Card Type";
            oColumn.Width = 40;
            oColumn = oMatrix.Columns.Add("Currency", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Currency";
            oColumn.Width = 40;
            oColumn = oMatrix.Columns.Add("BranchID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "BranchID";
            oColumn.Width = 40;
            oColumn = oMatrix.Columns.Add("Pin", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Pin";
            oColumn.Width = 0;
            oColumn = oMatrix.Columns.Add("SourceKey", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Security Key";
            oColumn.Width = oItm.Width - 300;

            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
            populateCardAcctMatrix();

            edT = oItm.Top + oItm.Height;

            oItm = form.Items.Add("lbInst", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItm.Left = lMargin;
            oItm.Top = edT;
            oItm.Width = 600;
            oItm.Height = 16;
            cfgCurrencyString = getCurrencyString();
            setLabelCaption(oCfgForm, "lbInst", "*Valid card type are v, m, a, ds, eCheck. " + "Valid currency are " + cfgCurrencyString);





            int buttonTop = oCfgForm.ClientHeight - 28;
            int buttonLeft = 60;
            int buttonWidth = 65;
            int buttonHeight = 19;

            oItm = oCfgForm.Items.Add(btnUpdate, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItm.Left = buttonLeft;
            oItm.Width = buttonWidth;
            oItm.Top = buttonTop;
            oItm.Height = buttonHeight;

            oButton = ((SAPbouiCOM.Button)(oItm.Specific));

            oButton.Caption = "Update";

            //************************
            // Adding a Cancel button
            //***********************

            oItm = oCfgForm.Items.Add(btnClose, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItm.Left = buttonLeft + buttonWidth + 5;
            oItm.Width = buttonWidth;
            oItm.Top = buttonTop;
            oItm.Height = buttonHeight;

            oButton = ((SAPbouiCOM.Button)(oItm.Specific));

            oButton.Caption = "Cancel";
            LoadWalkInList();

        }
        catch (Exception ex)
        {

            errorLog(ex);
        }
    }
    private void populateCardAcctMatrix()
    {
        string sql = "";
        try
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oCfgForm.Items.Item(mxCardAcct).Specific;
            oMatrix.Clear();

            sql = string.Format("select \"DocEntry\" , \"U_GroupName\", \"U_CardName\", \"U_CardCode\", \"U_CardType\", \"U_Pin\", \"U_SourceKey\", \"U_Currency\",\"U_BranchID\" from \"@CCCARDACCT\" Order By \"U_CardCode\" ");
            SAPbouiCOM.DataTable oDts = getCardAcctDataTable();
            oDts.ExecuteQuery(sql);
            BindCardAcctMatrix(oMatrix, "AcctID", "DocEntry");
            BindCardAcctMatrix(oMatrix, "CardName", "U_CardName");
            BindCardAcctMatrix(oMatrix, "CardType", "U_CardType");
            BindCardAcctMatrix(oMatrix, "Currency", "U_Currency");
            BindCardAcctMatrix(oMatrix, "GroupName", "U_GroupName");
            BindCardAcctMatrix(oMatrix, "Pin", "U_Pin");
            BindCardAcctMatrix(oMatrix, "SourceKey", "U_SourceKey");
            BindCardAcctMatrix(oMatrix, "BranchID", "U_BranchID");
            oMatrix.LoadFromDataSource();


        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void LoadWalkInList()
    {
        tProcess = new Thread(LoadWalkInThread);
        tProcess.Start();
    }
    private void LoadWalkInThread()
    {
        string sql = "";
        try
        {
            SAPbouiCOM.ComboBox oCB = (SAPbouiCOM.ComboBox)oCfgForm.Items.Item(cbWalkIn).Specific;
            try
            {
                while (oCB.ValidValues.Count > 0)
                    oCB.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception)
            { }
            ComboAddItem(oCB, "");
            sql = string.Format("select \"U_CustomerID\" from \"@CCWALKIN\" Order By \"U_CustomerID\" ");
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);
            if (!oRS.EoF)
            {
                string id = "";
                while (!oRS.EoF)
                {
                    id = (string)oRS.Fields.Item(0).Value;
                    ComboAddItem(oCB, id);
                    oRS.MoveNext();
                }
                oCB.Select(id);

            }
            oCB.Select("");
            // refresh();

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private void AddToWalkInList()
    {
        bAuto = false;
        int id = getNextTableID("@CCWALKIN");
        string w = getFormItemVal(oCfgForm, editWalkIn);
        if (SBO_Application.MessageBox("Add " + w + " to walk-in list?", 1, "Yes", "No") != 1)
            return;
        string sql = "";
        string sql2 = "";
        try
        {
            sql = string.Format("DELETE FROM \"@CCWALKIN\" Where \"U_CustomerID\"='{0}'", w);
            sql2 = string.Format("INSERT INTO \"@CCWALKIN\" (\"DocEntry\",\"U_CustomerID\") VALUES('{0}','{1}')", id, w);
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);
            oRS.DoQuery(sql2);

            SBO_Application.MessageBox(w + " added to walk-in list");
            if (SBO_Application.MessageBox("Remove all payment methods from customer: " + w + "?", 1, "Yes", "No") == 1)
            {
                removePaymentMethods(w);
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            errorLog(sql2);

        }
    }
    private void DeleteFromWalkInList()
    {
        string w = getFormItemVal(oCfgForm, cbWalkIn);
        if (w == "")
        {
            SBO_Application.MessageBox("Please select from walk-in list");
            return;
        }
        if (SBO_Application.MessageBox("Delete " + w + " from walk-in list?", 1, "Yes", "No") != 1)
            return;
        string sql = "";
        try
        {
            sql = string.Format("DELETE FROM \"@CCWALKIN\" Where \"U_CustomerID\"='{0}'", w);
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);

            SBO_Application.MessageBox(w + " deleted from walk-in list");
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
    }
    private SAPbouiCOM.DataTable getCardAcctDataTable()
    {
        SAPbouiCOM.DataTable oDT;
        try
        {
            oDT = oCfgForm.DataSources.DataTables.Item(dtCardAcct);
        }
        catch (Exception)
        {
            oDT = oCfgForm.DataSources.DataTables.Add(dtCardAcct);

        }

        return oDT;
    }
    private void BindCardAcctMatrix(SAPbouiCOM.Matrix oMatrix, string mtCol, string dtCol)
    {
        try
        {
            SAPbouiCOM.DataTable dt = getCardAcctDataTable();

            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item(mtCol);
            oColumn.DataBind.Bind(dtCardAcct, dtCol);


        }
        catch (Exception ex)
        {
            errorLog("Can not bind " + mtCol + " to " + dtCol + ".  error: " + ex.Message);
        }

    }
}


