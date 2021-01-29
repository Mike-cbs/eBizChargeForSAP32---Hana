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
using System.Globalization;

partial class SAP
{
    public SAPbobsCOM.Company oSetupCompany;
    //  Error handling variables
    public string sErrMsg;
    public int lErrCode;
    public int lRetCode;
    public SAPbobsCOM.BoDataServerTypes dbtype;
    public static bool bSetup = false;
    public string txtLicenseServer = "";
    public string txtServer = "";
    public string txtDBUser = "";
    public string txtSBOUser = "";
    public string txtDBPass = "";
    public string txtSBOPass = "";
    public string txtDBName = "";
    public string txtMerchantID = "";
    public string lastError = "";
    public void SAPSetup()
    {
        try
        {
            bSetup = true;
  
            string tname = "CCCFG"; 
            AddUserTable(tname , tname, SAPbobsCOM.BoUTBTableType.bott_Document);
            AddTableField(tname, "Pin", 50);
            AddTableField(tname, "SourceKey", 50);
            AddTableField(tname, "eBizChargeUrl", 80);
            AddTableField(tname, "NegPayment", 50);
            AddTableField(tname, "importPin", 50);
            AddTableField(tname, "importSourceKey", 50);
            AddTableField(tname, "impWebServiceUrl", 80);
            AddTableField(tname, "active", 10);
            AddTableField(tname, "sendCustomerEmail", 50);
            AddTableField(tname, "urlSandbox", 50);
            AddTableField(tname, "urlProduction", 50);
            AddTableField(tname, "MerchantReceipt", 50);
            AddTableField(tname, "CustomerReceipt", 50);
            AddTableField(tname, "merchantEmail", 50);
            AddTableField(tname, "addressRequired", 50);
            AddTableField(tname, "paymentGroup", 50);
            AddTableField(tname, "useUserEmail", 50);
            AddTableField(tname, "merchantID", 50);
            AddTableField(tname, "BatchAutoMode", 50);
            AddTableField(tname, "recurringBilling", 50);
            AddTableField(tname, "BtnPos", 50);
            AddTableField(tname, "NoTransOnInv", 50);
            AddTableField(tname, "DefaultTaxRate", 50);
            AddTableField(tname, "PMAutoSelect", 50);
            AddTableField(tname, "importCust", 50);
            AddTableField(tname, "ProcessPreAuth", 20);
            AddTableField(tname, "showStartup", 20);
            AddTableField(tname, "EmailSubject", 250);
            AddTableField(tname, "SaveReceipt", 20);
            AddTableField(tname, "BatchPreauth", 20);
            AddTableField(tname, "DeclineUnapprove", 20);
            AddTableField(tname, "PayOnTerm", 20);
            AddTableField(tname, "CaptureOnPrint", 20);
            AddTableField(tname, "ApprUDF", 50);
            AddTableField(tname, "CaptureOnly", 50);
            AddTableField(tname, "AmountField", 50);
            AddTableField(tname, "FieldSize", 50);
            AddTableField(tname, "HasTerminal", 50);

            tname = "CCCTerminal";
            AddUserTable(tname, tname, SAPbobsCOM.BoUTBTableType.bott_Document);
            AddTableField(tname, "name", 50);
            AddTableField(tname, "url", 50);
            AddTableField(tname, "deviceKey", 80);
            AddTableField(tname, "apiKey", 50);
            AddTableField(tname, "pin", 50);
            AddTableField(tname, "isDefault", 50);

         

            tname = "CCTRANS";
            AddUserTable(tname, tname, SAPbobsCOM.BoUTBTableType.bott_Document);
            AddTableField(tname, "ccTRANSID", 0, SAPbobsCOM.BoFieldTypes.db_Numeric);
            AddTableField(tname, "Imported", 10);
            AddTableField(tname, "CCORDERID", 50);
            AddTableField(tname, "CCAccountID", 50);
            AddTableField(tname, "InvoiceID", 50);
            AddTableField(tname, "PaymentID", 50);
            AddTableField(tname, "SAPPaymentID", 50);
            AddTableField(tname, "DownPaymentInvID", 50);
            AddTableField(tname, "OrderID", 50);
            AddTableField(tname, "QuoteID", 50);
            AddTableField(tname, "CreditMemoID", 50);
            AddTableField(tname, "MethodID", 50);
            AddTableField(tname, "customerID", 50);
            AddTableField(tname, "customerName", 80);
            AddTableField(tname, "CardHolder", 80);
            AddTableField(tname, "creditRef", 50);
            AddTableField(tname, "crCardNum", 50);
            AddTableField(tname, "Description", 254);
            AddTableField(tname, "recID", 50);
            AddTableField(tname, "acsUrl", 80);
            AddTableField(tname, "authAmount", 50);
            AddTableField(tname, "authCode", 10);
            AddTableField(tname, "avsResult", 254);
            AddTableField(tname, "avsResultCode", 20);
            AddTableField(tname, "batchNum", 20);
            AddTableField(tname, "batchRef", 20);
            AddTableField(tname, "cardCodeResult", 254);
            AddTableField(tname, "cardCodeRCode", 10);
            AddTableField(tname, "cardLevelResult", 254);
            AddTableField(tname, "cardLevelRCode", 10);
            AddTableField(tname, "conversionRate", 20);
            AddTableField(tname, "convertedAmount", 20);
            AddTableField(tname, "convertedCurrency", 20);
            AddTableField(tname, "custNum", 50);
            AddTableField(tname, "error", 254);
            AddTableField(tname, "errorCode", 10);
            AddTableField(tname, "isDuplicate", 10);
            AddTableField(tname, "payload", 50);
            AddTableField(tname, "profilerScore", 10);
            AddTableField(tname, "profilerResponse", 50);
            AddTableField(tname, "profilerReason", 50);
            AddTableField(tname, "refNum", 50);
            AddTableField(tname, "remainingBalance", 20);
            AddTableField(tname, "result", 50);
            AddTableField(tname, "resultCode", 10);
            AddTableField(tname, "status", 50);
            AddTableField(tname, "statusCode", 50);
            AddTableField(tname, "vpasResultCode", 10);
            AddTableField(tname, "recDate", 0, SAPbobsCOM.BoFieldTypes.db_Date);
            AddTableField(tname, "command", 20);
            AddTableField(tname, "amount", 50);
            AddTableField(tname, "jobID", 50);
            AddTableField(tname, "DeliveryID", 50);
            AddTableField(tname, "BatchResult", 254);
            AddTableField(tname, "BatchDate", 0, SAPbobsCOM.BoFieldTypes.db_Date);

      
            tname = "CCWALKIN";
            AddUserTable(tname, tname, SAPbobsCOM.BoUTBTableType.bott_Document);
            AddTableField(tname, "CustomerID", 80);


            tname = "CCJOB";
            AddUserTable(tname, tname, SAPbobsCOM.BoUTBTableType.bott_Document);
            AddTableField(tname, "InvoiceID", 50);
	        AddTableField(tname, "OrderID",50);
	        AddTableField(tname, "Description",254);
            AddTableField(tname, "CustomerID",50);
            AddTableField(tname, "Frequency",50);
            AddTableField(tname, "StartDate", 0, SAPbobsCOM.BoFieldTypes.db_Date);
            AddTableField(tname, "EndDate",0, SAPbobsCOM.BoFieldTypes.db_Date);
            AddTableField(tname, "NextRunDate", 0, SAPbobsCOM.BoFieldTypes.db_Date);
            AddTableField(tname, "LastRunDate", 0, SAPbobsCOM.BoFieldTypes.db_Date);
            AddTableField(tname, "Cancelled", 10);
            AddTableField(tname, "CancelledDate",0, SAPbobsCOM.BoFieldTypes.db_Date);
            AddTableField(tname, "CancelledBy",50);
            AddTableField(tname, "Amount",50);
            AddTableField(tname, "PaymentID",250);
            AddTableField(tname, "Result", 250);

     
            tname = "CCCUST";
            AddUserTable(tname,tname, SAPbobsCOM.BoUTBTableType.bott_Document);
	        AddTableField(tname, "CustomerID",50);
	        AddTableField(tname, "CustNum",50);
            AddTableField(tname, "MethodID",50);
            AddTableField(tname, "default",10);
            AddTableField(tname, "active",10);
            AddTableField(tname, "Declined",10);
            AddTableField(tname, "methodDescription",250);
            AddTableField(tname, "cardNum",50);
            AddTableField(tname, "expDate",50);
            AddTableField(tname, "cardCode",10);
            AddTableField(tname, "routingNumber",50);
            AddTableField(tname, "checkingAccount",50);
            AddTableField(tname, "firstName",50);
            AddTableField(tname, "lastName",50);
            AddTableField(tname, "email",80);
            AddTableField(tname, "street",80);
            AddTableField(tname, "city",50);
            AddTableField(tname, "state",20);
            AddTableField(tname, "zip",20);
            AddTableField(tname, "cardType",50);
            AddTableField(tname, "GroupName",80);
            AddTableField(tname, "CardName",80);
            AddTableField(tname, "CCAccountID", 80);
            AddTableField(tname, "MethodName", 80);
         
            tname = "CCConnectInvoice";
            AddUserTable(tname,tname, SAPbobsCOM.BoUTBTableType.bott_Document);
      
	        AddTableField(tname, "InvoiceID",50);
	        AddTableField(tname, "CustomerID",50);
            AddTableField(tname, "InvoiceGUID",50);
            AddTableField(tname, "PaidAmount",50);
            AddTableField(tname, "UploadedBalance",50);
            AddTableField(tname, "Status",50);
            AddTableField(tname, "UploadDate", 0 , SAPbobsCOM.BoFieldTypes.db_Date);
            AddTableField(tname, "PaidDate", 0, SAPbobsCOM.BoFieldTypes.db_Date);

       
            tname = "CCCPayment";
            AddUserTable(tname,tname, SAPbobsCOM.BoUTBTableType.bott_Document);
           
	        AddTableField(tname,"CustomerID",50);
	        AddTableField(tname,"InvoiceID",50);
	        AddTableField(tname,"PaymentID", 50);
            AddTableField(tname,"Amount",50);
            AddTableField(tname, "DateImported", 0, SAPbobsCOM.BoFieldTypes.db_Date);

            tname = "CCPayFormInvoice";
            AddUserTable(tname, tname, SAPbobsCOM.BoUTBTableType.bott_Document);

            AddTableField(tname, "InvoiceID", 50);
            AddTableField(tname, "CustomerID", 50);
            AddTableField(tname, "InvoiceGUID", 50);
            AddTableField(tname, "PaidAmount", 50);
            AddTableField(tname, "UploadedBalance", 50);
            AddTableField(tname, "Status", 50);
            AddTableField(tname, "UploadDate", 0, SAPbobsCOM.BoFieldTypes.db_Date);
            AddTableField(tname, "PaidDate", 0, SAPbobsCOM.BoFieldTypes.db_Date);


            tname = "CCPayFormPayment";
            AddUserTable(tname, tname, SAPbobsCOM.BoUTBTableType.bott_Document);

            AddTableField(tname, "CustomerID", 50);
            AddTableField(tname, "InvoiceID", 50);
            AddTableField(tname, "PaymentID", 50);
            AddTableField(tname, "Amount", 50);
            AddTableField(tname, "DateImported", 0, SAPbobsCOM.BoFieldTypes.db_Date);


            tname = "CCConnectCustomer";
            AddUserTable(tname,tname, SAPbobsCOM.BoUTBTableType.bott_Document);
           AddTableField(tname,"CustomerID", 50);
           AddTableField(tname, "CustomerGUID", 50);

            tname = "CCCardAcct";
            AddUserTable(tname, tname, SAPbobsCOM.BoUTBTableType.bott_Document);
            AddTableField(tname, "GroupName",80);
	        AddTableField(tname, "CardType",80);
	        AddTableField(tname, "CardName",80);
            AddTableField(tname, "AcctCode",80);
            AddTableField(tname, "CardCode",50);
            AddTableField(tname, "SourceKey",50);
            AddTableField(tname, "Currency", 50);
            AddTableField(tname, "Pin", 50);
            AddTableField(tname, "BranchID", 50);
            SetupInitCardAcct();
            SetupCfgInit();
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    public bool connect()
    {
        try
        {
            oSetupCompany = new SAPbobsCOM.Company();
            /*
            txtServer.Text = @"SIMON-PC\MSSQL2014";
            txtUser.Text = "sa";
            txtPass.Text = "Hsyhmin1";
            txtSBOPass.Text = "Hsyhmin1";
            */
           
          
            // once the Server property of the Company is set
            // we may query for a list of companies to choos from
            // this method returns a Recordset object

            oSetupCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
            oSetupCompany.Server = txtServer;
            oSetupCompany.CompanyDB = txtDBName;
            oSetupCompany.DbUserName =  txtDBUser;
            oSetupCompany.DbPassword = txtDBPass;
            oSetupCompany.UserName = txtSBOUser;
            oSetupCompany.Password = txtSBOPass; 
            oSetupCompany.LicenseServer = txtLicenseServer;
            oSetupCompany.DbServerType =SAPbobsCOM.BoDataServerTypes.dst_HANADB;
            lRetCode = oSetupCompany.Connect();
            if (lRetCode != 0)
            {
                oSetupCompany.GetLastError(out lErrCode, out sErrMsg);
                lastError = string.Format("Cannot connect to the company DB.\r\n\r\nLicense Server: {0}\r\nDB Server: {1}\r\nB1 User:{3}\r\nB1 User Pwd:{4}\r\nDB User:{5}\r\nDB User Pwd{6}\r\nError:{2}"
                    , oSetupCompany.LicenseServer, oSetupCompany.Server, sErrMsg
                    , txtSBOUser, txtSBOPass, txtDBUser, txtDBPass);
                errorLog(lastError);
                return false;
            }
            else
            {
                errorLog("Connected to " + oSetupCompany.CompanyName);
                return true;
            }
        }catch(Exception ex)
        {
            errorLog(ex);
        }
        return false;
    }
    public void SetupInitCardAcct()
    {

        string sql = "delete from \"@CCCARDACCT\" where \"U_CardCode\" Not in (select \"CreditCard\" from OCRC) or \"U_CardCode\" is null";
        string sql2 = "insert into \"@CCCARDACCT\"(\"DocEntry\", \"U_CardCode\", \"U_CardName\")\r\n" +
"Select \"CreditCard\", \"CreditCard\", \"CardName\" from OCRC a where a.\"CreditCard\" not in (select \"U_CardCode\" from \"@CCCARDACCT\")";
string sql3 = "Update \"@CCCARDACCT\" set \"U_CardName\" = a.\"CardName\" from OCRC a, \"@CCCARDACCT\" b where a.\"CreditCard\" = b.\"U_CardCode\"";
SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oSetupCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
  
        try
        {
            oRS.DoQuery(sql);
            oRS.DoQuery(sql2);
            oRS.DoQuery(sql3);

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            errorLog(sql2);
            errorLog(sql3);

        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
        }
        SetupinitUpdateCardAcctWithOCRC("v", "Visa");
        SetupinitUpdateCardAcctWithOCRC("m", "Ma");
        SetupinitUpdateCardAcctWithOCRC("a", "Am");
        SetupinitUpdateCardAcctWithOCRC("ds", "Disc");
        SetupinitUpdateCardAcctWithOCRC("eCheck", "eCheck");
    }
    public void SetupinitUpdateCardAcctWithOCRC(string field, string like)
    {

        string sql = string.Format("Update \"@CCCARDACCT\" set \"U_CardType\" = '{0}' from \"@CCCARDACCT\" a, OCRC b where  a.\"U_CardCode\" = b.\"CreditCard\" and b.\"CardName\" like '%{1}%' and a.\"U_CardType\" is null", field, like);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oSetupCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
            
            oRS.DoQuery(sql);
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
        }
    }
    public void SetupCfgInit()
    {
      
        string sql = string.Format("select 1 from \"@CCCFG\"");


        string sql2 = string.Format("INSERT INTO \"@CCCFG\"(\"DocEntry\", \"U_Pin\", \"U_SourceKey\", \"U_eBizChargeUrl\", \"U_NegPayment\", \"U_ProcessPreAuth\", \"U_sendCustomerEmail\" , \"U_urlSandbox\" , \"U_urlProduction\",\"U_merchantID\") " +
            " values(1, '{0}','{1}','{2}', 'Y', 'Y', 'Y','https://sandbox.ebizcharge.com/soap/gate/CCEBDC0A','https://secure.ebizcharge.com/soap/gate/CCEBDC0A','{3}')", "", "", "", txtMerchantID);
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oSetupCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        try
        {
          oRS.DoQuery(sql);
          if(oRS.EoF)
          {
              oRS.DoQuery(sql2);
          }
 
        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
            errorLog(sql2);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
        }
    }
     public void errorLog(Exception ex)
    {
        try
        {
            errorLog(ex.Message + "\r\n\r\n" + ex.StackTrace);
        }
        catch (Exception)
        {
        }
    }
    public void errorLog(string msg)
    {
        try
        {

            string logFileName = @"c:\CBS\eBizChargeForSAP_Setup.log";
        
            msg = string.Format("{0}:\r\n\t{1}\r\n\r\n", DateTime.Now, msg);
            if (!Directory.Exists(@"c:\CBS"))
                Directory.CreateDirectory(@"c:\CBS");
            if (File.Exists(logFileName))
            {
                FileInfo fi = new FileInfo(logFileName);
                if (fi.Length > 100000)
                {
                    File.Copy(logFileName, logFileName + DateTime.Now.ToString("yyyyMMddhhmmss"));
                    File.Delete(logFileName);
                }
            }
            FileStream fs = new FileStream(logFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamReader sr = new StreamReader(fs);

            fs.Seek(0, SeekOrigin.End);
            StreamWriter sw = new StreamWriter(fs);
            sw.Write(msg);
            sw.Close();
            fs.Close();
        }
        catch (Exception)
        {
        }
    }
    public void AddTableField(string tablename, string fieldname, int size, SAPbobsCOM.BoFieldTypes type = SAPbobsCOM.BoFieldTypes.db_Alpha)
    {
        try
        {
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;
            oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(oSetupCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));
            // ************************************
            //  Adding "Name" field
            // ************************************
            // Setting the Field's properties
            oUserFieldsMD.TableName = "@" + tablename;
            oUserFieldsMD.Name = fieldname;
            oUserFieldsMD.Description = fieldname;
            oUserFieldsMD.Type = type;
            if (type == SAPbobsCOM.BoFieldTypes.db_Alpha)
                oUserFieldsMD.EditSize = size;
            // Adding the Field to the Table
            lRetCode = oUserFieldsMD.Add();
            // Check for errors
            if (lRetCode != 0)
            {
                if (lRetCode == -1)
                {
                    errorLog(fieldname + " already added");
                }
                else
                {
                    oSetupCompany.GetLastError(out lRetCode, out sErrMsg);
                    errorLog("Add field error. Field: '" + oUserFieldsMD.Name + "', table: " + oUserFieldsMD.TableName + ", " + sErrMsg);
                }
            }
            else
            {
                errorLog("Field: '" + oUserFieldsMD.Name + "' was added successfuly to " + oUserFieldsMD.TableName + " Table");
            }
            GC.Collect();
        }
        catch (Exception ex)
        {
            string err = string.Format("Add Table field: {0}/{1}, Exception: {2}", tablename, fieldname, ex.Message);
            errorLog(err);
            if (ex.Message.IndexOf("Ref count") >= 0)
                GC.Collect();
        }
        finally
        {
            GC.Collect(); // Release the handle to the User Fields

        }
    }
    public void AddUserTable(string Name, string Description, SAPbobsCOM.BoUTBTableType Type)
    {
        try
        {
            if(!connect())
                return;
            //****************************************************************************
            // The UserTablesMD represents a meta-data object which allows us
            // to add\remove tables, change a table name etc.
            //****************************************************************************
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;
            //****************************************************************************
            // In any meta-data operation there should be no other object "alive"
            // but the meta-data object, otherwise the operation will fail.
            // This restriction is intended to prevent a collisions
            //****************************************************************************
            // the meta-data object needs to be initialized with a
            // regular UserTables object
            oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(oSetupCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));
            //**************************************************
            // when adding user tables or fields to the SBO DB
            // use a prefix identifying your partner name space
            // this will prevent collisions between different
            // partners add-ons
            // SAP's name space prefix is "BE_"
            //**************************************************		
            // set the table parameters
            oUserTablesMD.TableName = Name;
            oUserTablesMD.TableDescription = Description;
            oUserTablesMD.TableType = Type;
            // Add the table
            // This action add an empty table with 2 default fields
            // 'Code' and 'Name' which serve as the key
            // in order to add your own User Fields
            // see the AddUserFields.frm in this project
            // a privat, user defined, key may be added
            // see AddPrivateKey.frm in this project
            lRetCode = oUserTablesMD.Add();
            // check for errors in the process
            if (lRetCode != 0)
            {
                oSetupCompany.GetLastError(out lRetCode, out sErrMsg);
                errorLog("Add table error. Table: " + oUserTablesMD.TableName + ", " + sErrMsg);

            }
            else
            {
                errorLog("Table: " + oUserTablesMD.TableName + " was added successfully");
            }
            oUserTablesMD = null;
            GC.Collect();
        }
        catch (Exception ex)
        {

            string err = string.Format("Add Table: {0}, Exeption: {1}", Name, ex.Message);
            errorLog(err);
            if (ex.Message.IndexOf("Ref count") >= 0)
                GC.Collect();
        }
        finally
        {
            connect();
            GC.Collect(); // Release the handle to the table

        }
    }
    
}
