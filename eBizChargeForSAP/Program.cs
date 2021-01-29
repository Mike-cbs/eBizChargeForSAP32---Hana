using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using System.Data.SqlClient;
using System.Xml;

namespace eBizChargeForSAP
{

    class Program
    {
        //static Mutex mutex = new Mutex(true, "{8F6F0AC4-B9A1-45fd-A8CF-72F04E6BDE8F}");
        [STAThread]

        static void Main()
        {
            try
            {
                string[] arg = Environment.GetCommandLineArgs();
                string conStr = "";

                if (arg.Length >= 2)
                {
                    conStr = arg[1];

                }

                //conStr = "runjob";
                //conStr = "-capture";
                
                if (conStr.ToLower() == "runjob")
                {
                    #region runjob
                    try
                    {
                        string connStr = "";
                        string dbName = "";

                        SAP sap = new SAP();
                        SAP.bRunJob = true;

                        if (sap.connectToDB())
                        {
                            sap.ProcessJob();
                            SAP.errorLog("eBizCharge job run completed.");
                            sap.CloseConnect();
                        }

                    }
                    catch (Exception ex)
                    {
                        SAP.errorLog(ex);
                    }
                    finally
                    {
                        System.Environment.Exit(0);
                    }
                    #endregion
                }
                if (conStr.ToLower() == "-capture")
                {
                    #region -capture
                    try
                    {
                        SAP sap = new SAP();
                        SAP.bRunJob = true;
                        if (sap.connectToDB())
                        {

                            sap.captureService();

                            sap.CloseConnect();
                        }

                    }
                    catch (Exception ex)
                    {
                        SAP.errorLog(ex);
                    }
                    finally
                    {
                        System.Environment.Exit(0);
                    }
                    #endregion
                }
                if (conStr.ToLower() == "-nodeliverycapture")
                {
                    #region nodeliverycapture
                    try
                    {
                        SAP sap = new SAP();
                        SAP.bRunJob = true;
                        if (sap.connectToDB())
                        {
                            while (true)
                            {
                                sap.captureService("nodelivery");
                                Thread.Sleep(1000);
                            }
                            sap.CloseConnect();
                        }

                    }
                    catch (Exception ex)
                    {
                        SAP.errorLog(ex);
                    }
                    finally
                    {
                        System.Environment.Exit(0);
                    }
                    #endregion
                }

                if (conStr == "")
                    conStr = System.Configuration.ConfigurationManager.AppSettings["SAPConnectionKey"].ToString();
                if (conStr != "")
                {
                    #region SAP
                    SAP sap = new SAP();
                    string ver = "(v1.00.91)";
                    sap.SboGuiApi = new SAPbouiCOM.SboGuiApi();

                    sap.SboGuiApi.Connect(conStr);
                    sap.SetApplication();
                    SAP.trace("Hana conStr = " + conStr);
                    if (conStr == "")
                        sap.showStatus("eBizCharge System Started. " + ver, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    else
                        sap.showStatus("eBizCharge System Started. " + ver + "  session id: " + conStr, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    //SaveConfig(sap.dbName, sap.connStr);
                    sap.oSignalEvent.WaitOne();

                    #endregion
                }
            
                //mutex.ReleaseMutex();
                /*
            }
            else
            {
                //"only one instance at a time";
            }
            */
            }
            catch (Exception ex)
            {
                SAP.errorLog(ex);

            }

        }


    }


}
