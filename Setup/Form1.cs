using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Win32.TaskScheduler;
using System.Diagnostics;
using System.Net;
using System.Threading;
using System.IO;

namespace Setup
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try{
                    string filename = @"C:\CBS\eBizChargeForB1_Hana.XML";
                    if (File.Exists(filename))
                    {
                        XmlDocument doc = new XmlDocument();
                        doc.Load(filename);
                        XmlNode node = doc.SelectSingleNode("//root//SAPLicenseServer");
                        textLicense.Text = node.InnerText;
                        node = doc.SelectSingleNode("//root//SAPDBServer");
                        textDatabase.Text  = node.InnerText;
                        node = doc.SelectSingleNode("//root//SAPDBName");
                        txtCompanyDBName.Text = node.InnerText;
                          
                        B1InstallService.B1InstallServiceSoapClient install = new B1InstallService.B1InstallServiceSoapClient();
                        string domain = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;

                        B1InstallService.SAPB1Install info = install.GetInfoByLicSvr(textLicense.Text, domain, txtCompanyDBName.Text);
                        if (info != null)
                        {

                            textDBUser.Text = info.DBUser;
                            textDBPassword.Text = info.DBPassWD;
                            txtSiteUser.Text = info.AppUser;
                            txtSiteUserPassword.Text = info.AppPassWD;
                            textPin.Text =info.PIN;
                            textSourceKey.Text = info.SourceKey;
                            string val = info.sandbox;
                            if (val.IndexOf("sandbox") >= 0)
                                cbURL.SelectedIndex = 0;
                            else
                                cbURL.SelectedIndex = 1;
                        }
      
                    }
                
            }catch(Exception ex)
            {
                Program.errorLog(ex);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                TextBox.CheckForIllegalCrossThreadCalls = false;
                PictureBox.CheckForIllegalCrossThreadCalls = false;
                this.pictureBox1.Visible = true;
                Thread tProcess = new Thread(SaveAction);
                tProcess.Start();
            }
            catch (Exception) { }
        }
        public void SaveAction()
        {
            bool bError = false;
            try
            {
                SAP oCompany = new SAP();

                oCompany.txtServer = textDatabase.Text;
                oCompany.txtDBName = txtCompanyDBName.Text;
                oCompany.txtDBPass = textDBPassword.Text;
                oCompany.txtDBUser = textDBUser.Text;
                oCompany.txtSBOPass = txtSiteUserPassword.Text;
                oCompany.txtSBOUser = txtSiteUser.Text;
                oCompany.txtLicenseServer = textLicense.Text;
                oCompany.txtMerchantID = textSourceKey.Text;
                if (!oCompany.connect())
                {
                    MessageBox.Show(oCompany.lastError);
                    bError = true;
                    return;
                }
                try
                {
                    if(!Directory.Exists(@"C:\CBS\"))
                        Directory.CreateDirectory(@"C:\CBS\");
                    string filename = @"C:\CBS\eBizChargeForB1_Hana.XML";
                    try
                    {
                        if(File.Exists(filename))
                            File.Delete(filename);
                    }
                    catch (Exception) { }
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml("<root/>"); 
                    XmlElement elem = doc.CreateElement("SAPLicenseServer");
                    elem.InnerText = textLicense.Text;
                    doc.DocumentElement.AppendChild(elem);
                    elem = doc.CreateElement("SAPDBServer");
                    elem.InnerText = textDatabase.Text;
                    doc.DocumentElement.AppendChild(elem);
                    elem = doc.CreateElement("SAPDBName");
                    elem.InnerText = txtCompanyDBName.Text;
                    doc.DocumentElement.AppendChild(elem);
                    doc.Save(filename);
                }catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                try
                {
                    
                    MessageBox.Show("eBizCharge for Hana installation may take more then 30 sec to completed.  Please wait for process to complete");
                  
                    oCompany.SAPSetup();
                }
                catch (Exception ex2)
                {
                    Program.errorLog(ex2.Message);
                }
                try
                {
                    B1InstallService.B1InstallServiceSoapClient install = new B1InstallService.B1InstallServiceSoapClient();
                    B1InstallService.SAPB1Install info = new B1InstallService.SAPB1Install();
                    info.AppPassWD = txtSiteUserPassword.Text;
                    info.AppUser = txtSiteUser.Text;
                    info.DBName = txtCompanyDBName.Text;
                    info.DBPassWD = textDBPassword.Text;
                    info.DBServer = textDatabase.Text;
                    info.DBUser = textDBUser.Text;
                    info.Domain = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
                    if (info.Domain == "")
                        info.Domain = txtCompanyDBName.Text;
                    
                    info.IP = getIP();
                    info.LicenseServer = textLicense.Text;
                    info.PIN = textPin.Text;
                    info.SourceKey = textSourceKey.Text;
                    install.UpsertInfo(info);

                }
                catch (Exception) { }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                bError = true;
                Program.errorLog(ex);
            }
            finally
            {
                CopyFiles();
                this.pictureBox1.Visible = false;
                if (!bError)
                {
                    MessageBox.Show("Setup completed.  Please restart B1.");
                    this.Close();
                }
            }
        }
        public void CopyFiles()
        {
            string targetDir = textInstalledFolder.Text;
            string sourceDir = Program.ProgramFiles() + @"\Century Business Solutions\eBizChargeForSAP\";
            try
            {
                try
                {
                    Directory.CreateDirectory(targetDir);
                }
                catch (Exception) { }
                foreach (var file in Directory.GetFiles(sourceDir))
                    File.Copy(file, Path.Combine(targetDir, Path.GetFileName(file)), true);



            }
            catch (Exception)
            {

            }

        }
        public string getLicseServer()
        {
            try
            {
                string filename = Program.ProgramFiles() + @"\SAP\SAP Business One\Conf\b1-local-machine.xml";
                XmlDocument doc = new XmlDocument();
                doc.Load(filename);
                XmlNodeList nodelist = doc.SelectNodes("//configuration//node[@name='REGISTRY']//node[@name='HKEY_LOCAL_MACHINE']//node[@name='SOFTWARE']//node[@name='SAP']//node[@name='SAP Manage']//node[@name='SAP Business One']");
                foreach (XmlNode n in nodelist)
                {
                    string s = n.FirstChild.FirstChild.InnerText;

                    if (s.IndexOf("localhost") >= 0)
                    {
                        return Dns.GetHostName();
                    }
                    else
                        return s;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return "";
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
        public static string Base64Encode(string plainText)
        {
            try
            {
                var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
                return System.Convert.ToBase64String(plainTextBytes);
            }catch(Exception)
            { }
            return plainText;
        }
        public static string Base64Decode(string base64EncodedData)
        {
            try
            {
                var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
                return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
            }catch(Exception)
            {

            }
            return base64EncodedData;
        }
        private const string TaskName = @"eBizCharge Job Task";
        private void addTask()
        {
            try
            {
                using (TaskService ts = new TaskService())
                {
                    var task = ts.RootFolder.GetTasks().Where(a => a.Name.ToLower() == TaskName.ToLower()).FirstOrDefault();
                    if(task != null)
                    {
                        return;
                    }

                    string folder = Program.getFolder();
                    string filename = folder + "eBizChargeForSAP.exe";
                    // Create a new task definition and assign properties
                    TaskDefinition td = ts.NewTask();
                    td.RegistrationInfo.Description = "eBizCharge Job Task";

                    // Create a trigger that will fire the task at this time every other day
                    td.Triggers.Add(new DailyTrigger { DaysInterval = 1 });

                    // Create an action that will launch Notepad whenever the trigger fires
                    td.Actions.Add(new ExecAction(filename, "runjob", null));

                    // Register the task in the root folder
                    ts.RootFolder.RegisterTaskDefinition(TaskName, td);

                    // Remove the task we just created
                    //ts.RootFolder.DeleteTask("Test");
                }

            }catch(Exception)
            {

            }
        }
    }

}
