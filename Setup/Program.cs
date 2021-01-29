using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Reflection.Context;
using System.Windows.Forms;
using Microsoft.Win32.SafeHandles;
using System.Xml;


namespace Setup
{
    
    static class NativeMethods
    {
        [DllImport("kernel32.dll")]
        public static extern IntPtr LoadLibrary(string dllToLoad);

        [DllImport("kernel32.dll")]
        public static extern IntPtr GetProcAddress(IntPtr hModule, string procedureName);



        [DllImport("kernel32.dll")]
        public static extern bool FreeLibrary(IntPtr hModule);

        [DllImport("kernel32.dll")]
        internal static extern Boolean AllocConsole();

        [DllImport("kernel32.dll")]
        public static extern IntPtr GetStdHandle(int nStdHandle);
       

    }

    class Program
    {
        private const int STD_OUTPUT_HANDLE = -11;
        private const int MY_CODE_PAGE = 437;   

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int EndInstall();

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int SetAddOnFolder(string path);

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int RestartNeeded();

        static public ManualResetEvent oSignalEvent = new ManualResetEvent(false);
        [STAThread]
        static int Main(string[] args)
        {
            try
            {
                string folder = getFolder();
                errorLog("install to path:" + folder);
                if (args.Length == 0)
                {
                    Application.EnableVisualStyles();
                    Application.Run(new Form1());
                }
                else
                {
                    string msg = "";
                    
                   
                                       

                    msg = string.Format("Argument: {0}", args[0]);
                    errorLog(msg);
                    string[] s = args[0].Split('|');
                    if (s.Length != 2)
                    {

                        errorLog("Wrong argument");
                        return -1;
                    }
                    string sDll = s[1];
                    sDll = sDll.Replace("AddOnInstallAPI.dll", "AddOnInstallAPI_x64.dll");

                    IntPtr pDll = NativeMethods.LoadLibrary(sDll);
                    IntPtr pAddressOfFunctionToCall = NativeMethods.GetProcAddress(pDll, "SetAddOnFolder");
                    //if(pAddressOfFunctionToCall == IntPtr.Zero) 

                    SetAddOnFolder setAddOnFolder = (SetAddOnFolder)Marshal.GetDelegateForFunctionPointer(
                                                                                            pAddressOfFunctionToCall,
                                                                                            typeof(SetAddOnFolder));

                    int theResult = setAddOnFolder(folder);

                    errorLog("setAddOnFolder return " + theResult.ToString() + "\r\nFolder:" + folder);

                   
                    
                    
                    
                    pAddressOfFunctionToCall = NativeMethods.GetProcAddress(pDll, "EndInstall");
                    //if(pAddressOfFunctionToCall == IntPtr.Zero) 

                    EndInstall endInstall = (EndInstall)Marshal.GetDelegateForFunctionPointer(
                                                                                            pAddressOfFunctionToCall,
                                                                                            typeof(EndInstall));

                    theResult = endInstall();
                    errorLog("EndInstall return " + theResult.ToString());


                    bool result = NativeMethods.FreeLibrary(pDll);
                    errorLog("Installation completed.");
                   // MessageBox.Show("eBizCharge Install Completed.");
                   
                }

            }catch(Exception ex)
            {
                errorLog(ex);
            }
            return 0;
        }
        public static string getFolder()
        {
            string folder = "";
            try
            {
                folder = ProgramFiles() + @"\Century Business Solutions\eBizChargeForSAP\";
                errorLog("Program Folder: " + folder);
                string f = System.Configuration.ConfigurationManager.AppSettings["InstalledFolder"].ToString();
                // if (f != "")
                //     folder = f;
                errorLog("Add On Folder: " + folder);

            }catch(Exception ex)
            {
                errorLog(ex);
            }
            return folder;
        }

        public static string ProgramFiles()
        {
            
            return Environment.GetEnvironmentVariable("ProgramFiles");
        }
        public static void errorLog(Exception ex)
        {
            try
            {
                errorLog(ex.Message + "\r\n\r\n" + ex.StackTrace);
            }
            catch (Exception)
            {
            }
        }
        public static void errorLog(string msg)
        {
            try
            {
               // string logFileName = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\eBizChargeForSAPSetup.log";
              //  Console.Write("\r\n\r\n" + msg + "\r\n\r\n");
                string logFileName = @"c:\CBS\CBSForSAPAddOnInstall.log";
                if (!Directory.Exists(@"c:\CBS"))
                    Directory.CreateDirectory(@"c:\CBS");
                msg = string.Format("{0}:\r\n\t{1}\r\n\r\n", DateTime.Now, msg);
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
      
    }
    
}
