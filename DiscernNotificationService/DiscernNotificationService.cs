using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.Net;
using Newtonsoft.Json;
using System.Timers;
using System.Globalization;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Configuration;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Discern_Notification
{
    public class Order
    {
        public string FIN;
        public string OrderDescription;
        public string OrderDate;
        public string CPT4;
        public string DischargeDate;
        public string Comments;
        public string Organization;
    }
    public partial class DiscernNotificationService : ServiceBase
    {
        static private Thread workerThread;
        static private EventWaitHandle waitHandle;
        static private ManualResetEvent _shutdownEvent = new ManualResetEvent(false);
        string StartTime = System.Configuration.ConfigurationManager.AppSettings["StartTime"].ToString();
        int VerificationInterval1 = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["VerificationInterval1"].ToString());
        int VerificationInterval2 = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["VerificationInterval2"].ToString());
        int VerificationInterval3 = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["VerificationInterval3"].ToString());
        int VerificationInterval4 = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["VerificationInterval4"].ToString());
        int VerificationInterval5 = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["VerificationInterval5"].ToString());
        int VerificationInterval6 = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["VerificationInterval6"].ToString());
        string WaitTimeAccountValidation = System.Configuration.ConfigurationManager.AppSettings["WaitTimeForAccountValidation"].ToString();
        string EmailSending = System.Configuration.ConfigurationManager.AppSettings["EmailSendTime"].ToString();
        DateTime timeOfStart = DateTime.Now;
        DateTime timeOfFirstStart = DateTime.Now;
        System.Timers.Timer timer = new System.Timers.Timer();
        public DiscernNotificationService()
        {
            InitializeComponent();
        }
        //public void StartTimer(DateTime time)
        //{

        //}
        protected override void OnStart(string[] args)
        {
            timeOfFirstStart = DateTime.ParseExact(StartTime, "HH:mm:ss", CultureInfo.InvariantCulture);
            Global.WriteLog(string.Format("Service Start At {0}", DateTime.Now.ToString("dd MMMM yyyy hh:mm:ss")));
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["TimerInterval"].ToString()); //number in milisecinds  
            timer.Enabled = true;
        }
        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            timeOfStart = DateTime.ParseExact(StartTime, "HH:mm:ss", CultureInfo.InvariantCulture);
            int timeDiff = Convert.ToInt32(DateTime.Now.Subtract(timeOfStart).TotalMinutes);
            int regTimeDiff = Convert.ToInt32(DateTime.Now.Subtract(timeOfFirstStart).TotalMinutes);
            int Hour = DateTime.Now.Hour;
            int MinutesPassed = (int)DateTime.Now.TimeOfDay.TotalMinutes;
            //Global.WriteLog(MinutesPassed.ToString());

            //Modified by Velan on 21/01/2022 for Flint and Lapeer Golive
            //string facility = System.Configuration.ConfigurationManager.AppSettings["Facility"].ToString();
            String[] facility = System.Configuration.ConfigurationManager.AppSettings["Facility"].ToString().Split(',');

            string receiverEmail = System.Configuration.ConfigurationManager.AppSettings["Email.ReceiverEmail"] != null ? System.Configuration.ConfigurationManager.AppSettings["Email.ReceiverEmail"].ToString() : string.Empty; //"osman.alikhan@atos.net"; //ConfigurationManager.AppSettings["Email.SenderEmail"]; //
            string senderEmail = System.Configuration.ConfigurationManager.AppSettings["Email.SenderEmail"] != null ? System.Configuration.ConfigurationManager.AppSettings["Email.SenderEmail"].ToString() : string.Empty; //"osman.alikhan@atos.net"; //ConfigurationManager.AppSettings["Email.SenderEmail"]; //
            string senderPassword = System.Configuration.ConfigurationManager.AppSettings["Email.SenderPassword"] != null ? System.Configuration.ConfigurationManager.AppSettings["Email.SenderPassword"].ToString() : string.Empty; //ConfigurationManager.AppSettings["Email.Password"];//
            string host = System.Configuration.ConfigurationManager.AppSettings["Email.Host"] != null ? System.Configuration.ConfigurationManager.AppSettings["Email.Host"].ToString() : string.Empty;//"10.14.32.146";// ConfigurationManager.AppSettings["Email.Host"]; //"10.14.32.146";
            int port = System.Configuration.ConfigurationManager.AppSettings["Email.Port"] != null ? Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Email.Port"].ToString()) : 0;//25;//Convert.ToInt32(ConfigurationManager.AppSettings["Email.Port"]); //25;
            try
            {
                if (timeDiff == 0)
                {
                    Global.WriteLog(timeOfStart.ToString());
                    Global.WriteLog(timeOfFirstStart.ToString());
                    Global.WriteLog(timeDiff.ToString());
                    Global.WriteLog(regTimeDiff.ToString());



                }

                int Interval1_archivetime = VerificationInterval1 - 2;
                int Interval2_archivetime = VerificationInterval2 - 2;
                int Interval3_archivetime = VerificationInterval3 - 2;
                int Interval4_archivetime = VerificationInterval4 - 2;
                int Interval5_archivetime = VerificationInterval5 - 2;
                int Interval6_archivetime = VerificationInterval6 - 2;

                if (MinutesPassed == Interval1_archivetime || MinutesPassed == Interval2_archivetime || MinutesPassed == Interval3_archivetime || MinutesPassed == Interval4_archivetime || MinutesPassed == Interval5_archivetime || MinutesPassed == Interval6_archivetime)
                {
                    try
                    {
                        Global.WriteLog("START - Archiving the existing files");
                        ArchiveExistingFiles(true);
                        Global.WriteLog("END - Archiving the existing files");
                    }
                    catch (Exception ex)
                    {

                        throw new Exception("An error occured while Archiving the Existing files. Error:" + ex.Message);
                    }

                    try
                    {
                        Global.WriteLog("START - Moving files from source SFTP folder to root Folder ");
                        //TOBO: Move files from ProdFiles folder to TaskQueue
                        string rootFolder = System.Configuration.ConfigurationManager.AppSettings["RootFolderPath"] != null ? System.Configuration.ConfigurationManager.AppSettings["RootFolderPath"] : string.Empty;
                        string sourceFolder = System.Configuration.ConfigurationManager.AppSettings["SourceFolderPath"] != null ? System.Configuration.ConfigurationManager.AppSettings["SourceFolderPath"] : string.Empty; //@"E:\TaskQueue\";//ConfigurationManager.AppSettings["AccessHIM.MasterExcel.Path"];//@"E:\TaskQueue\Archived\";
                        int numberOfAgents = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["NumberOfAgents"].ToString());
                        string catalogFolder = System.Configuration.ConfigurationManager.AppSettings["CatalogFolder"] != null ? System.Configuration.ConfigurationManager.AppSettings["CatalogFolder"] : string.Empty;
                        DirectoryInfo dir = new DirectoryInfo(sourceFolder);
                        string date = DateTime.Now.ToString("MMddyyyy");


                        // Modified by Velan on 20-01-2022 for Flint and Lapeer Go Live
                        //FileInfo[] toBeProcessedfiles = dir.GetFiles(string.Format("{0}_*ToBeProcessed*", facility), SearchOption.TopDirectoryOnly);

                        FileInfo[] toBeProcessedfiles = new FileInfo[10];
                        foreach (var f in facility)
                        {
                            toBeProcessedfiles = dir.GetFiles(string.Format("{0}_*ToBeProcessed*", f), SearchOption.TopDirectoryOnly);
                            if (toBeProcessedfiles.Length > 0)
                            {
                                Global.WriteLog("To be processed files found for " + f);
                                break;
                            }
                        }





                        if (toBeProcessedfiles.Count() > 0)
                        {
                            foreach (FileInfo file in toBeProcessedfiles)
                            {
                                if (file.Name.Contains(date))
                                {
                                    System.IO.File.Move(file.FullName, string.Format(@"{0}\{1}", rootFolder, file.Name));
                                }
                            }
                        }




                        //Modified by Velan on 01-20-2022, for Flint and Lapeer Go Live

                        //FileInfo[] toBeAllocatedfiles = dir.GetFiles(string.Format("{0}_*ToBeAllocated*", facility), SearchOption.TopDirectoryOnly);

                        FileInfo[] toBeAllocatedfiles = new FileInfo[10];
                        foreach (var f in facility)
                        {
                            toBeAllocatedfiles = dir.GetFiles(string.Format("{0}_*ToBeAllocated*", f), SearchOption.TopDirectoryOnly);
                            if (toBeAllocatedfiles.Length > 0)
                            {
                                Global.WriteLog("To be Allocated files found for " + f);
                                break;
                            }
                        }




                        if (toBeAllocatedfiles.Count() > 0)
                        {
                            foreach (FileInfo file in toBeAllocatedfiles)
                            {
                                if (file.Name.Contains(date))
                                {
                                    System.IO.File.Move(file.FullName, string.Format(@"{0}\{1}", rootFolder, file.Name));
                                }
                            }
                        }

                        string ordersFilePath = string.Empty;

                        //Modified by Velan on 01-20-2022, for Flint and Lapeer Go Live
                        //FileInfo[] ordersfiles = dir.GetFiles(string.Format("{0}_*orders*", facility), SearchOption.TopDirectoryOnly);

                        FileInfo[] ordersfiles = new FileInfo[10];
                        foreach (var f in facility)
                        {
                            ordersfiles = dir.GetFiles(string.Format("{0}_*orders*", f), SearchOption.TopDirectoryOnly);
                            if (ordersfiles.Length > 0)
                            {
                                Global.WriteLog("Orders files found for " + f);
                                break;
                            }
                        }




                        if (ordersfiles.Count() > 0)
                        {
                            foreach (var file in ordersfiles)
                            {
                                if (file.Name.Contains(date))
                                {
                                    ordersFilePath = file.FullName;
                                    for (int i = 1; i <= numberOfAgents; i++)
                                    {
                                        string ordersNewFileName = Path.GetFileNameWithoutExtension(ordersFilePath) + i + ".csv";
                                        System.IO.File.Copy(ordersFilePath, string.Format(@"{0}{1}", rootFolder, ordersNewFileName), true);
                                    }
                                    System.IO.File.Delete(ordersFilePath);
                                }
                            }
                        }

                        string chargesFilePath = string.Empty;

                        //Modified by Velan on 01-20-2022, for Flint and Lapeer Go Live
                        //FileInfo[] chargesfiles = dir.GetFiles(string.Format("{0}_*charges*", facility, SearchOption.TopDirectoryOnly));
                        FileInfo[] chargesfiles = new FileInfo[10];
                        foreach (var f in facility)
                        {
                            chargesfiles = dir.GetFiles(string.Format("{0}_*charges*", f), SearchOption.TopDirectoryOnly);
                            if (chargesfiles.Length > 0)
                            {
                                Global.WriteLog("charges files found for " + f);
                                break;
                            }
                        }



                        if (chargesfiles.Count() > 0)
                        {
                            foreach (var file in chargesfiles)
                            {
                                if (file.Name.Contains(date))
                                {
                                    chargesFilePath = file.FullName;

                                    for (int i = 1; i <= numberOfAgents; i++)
                                    {
                                        string chargesNewFileName = Path.GetFileNameWithoutExtension(chargesFilePath) + i + ".csv";
                                        System.IO.File.Copy(chargesFilePath, string.Format(@"{0}{1}", rootFolder, chargesNewFileName), true);
                                    }
                                    System.IO.File.Delete(chargesFilePath);
                                }
                            }
                        }




                        //Modified by Velan on 20-01-2022, for Flint and Lapeer Go Live
                        //FileInfo[] catalogfiles = dir.GetFiles(string.Format("{0}_*catalog*", facility), SearchOption.TopDirectoryOnly);
                        FileInfo[] catalogfiles = new FileInfo[10];
                        foreach (var f in facility)
                        {
                            catalogfiles = dir.GetFiles(string.Format("{0}_*catalog*", f), SearchOption.TopDirectoryOnly);
                            if (catalogfiles.Length > 0)
                            {
                                Global.WriteLog("catalog files found for " + f);
                                break;
                            }
                        }



                        string newFileName = string.Empty;
                        string catalogFileName = string.Empty;
                        string catalogDirectory = string.Empty;

                        if (catalogfiles.Count() > 0)
                        {

                            catalogFileName = catalogfiles[0].Name;
                            catalogDirectory = catalogfiles[0].DirectoryName;

                            for (int i = 1; i <= numberOfAgents; i++)
                            {
                                newFileName = Path.GetFileNameWithoutExtension(catalogFileName) + i.ToString() + ".csv";

                                System.IO.File.Copy(catalogfiles[0].FullName, string.Format(@"{0}\\{1}", catalogFolder, newFileName), true);
                            }
                            System.IO.File.Delete(catalogfiles[0].DirectoryName + "\\" + catalogFileName);
                        }



                        Global.WriteLog("END - Moving files from source SFTP folder to root Folder ");
                    }
                    catch (Exception ex)
                    {

                        throw new Exception("An error occured while moving files. Error:" + ex.Message);
                    }
                }

                if (MinutesPassed == VerificationInterval1 || MinutesPassed == VerificationInterval2 || MinutesPassed == VerificationInterval3 || MinutesPassed == VerificationInterval4 || MinutesPassed == VerificationInterval5 || MinutesPassed == VerificationInterval6)
                {
                    Global.WriteLog(timeOfStart.ToString());
                    Global.WriteLog(MinutesPassed.ToString());
                    Global.WriteLog(timeDiff.ToString());
                    Global.WriteLog(regTimeDiff.ToString());

                    string rootFolderP = System.Configuration.ConfigurationManager.AppSettings["RootFolderPath"].ToString();
                    string botAgentToBeProcessedFileName = System.Configuration.ConfigurationManager.AppSettings["BotAgentToBeProcessedFileName"];
                    string toBeProcessedfilepath = string.Empty;
                    DirectoryInfo directory = new DirectoryInfo(rootFolderP);
                    FileInfo[] toBeProcessedfilesArr = directory.GetFiles("*ToBeProcessed*", SearchOption.TopDirectoryOnly);
                    if (toBeProcessedfilesArr.Count() > 0)
                    {
                        toBeProcessedfilepath = toBeProcessedfilesArr[0].FullName;

                        var toBeProcessedRecords = File.ReadLines(toBeProcessedfilepath).ToList();

                        bool isEmptyFile = toBeProcessedRecords.Count() > 1 ? false : true;

                        if (isEmptyFile)
                        {
                            MergeFiles(true);
                        }
                        else
                        {
                            string sopUrl = System.Configuration.ConfigurationManager.AppSettings["SOP.Url"].ToString();
                            string sopUsername = System.Configuration.ConfigurationManager.AppSettings["SOP.Username"].ToString();
                            string sopPassword = System.Configuration.ConfigurationManager.AppSettings["SOP.Password"].ToString();
                            string authorization = sopUsername + ":" + sopPassword;
                            int numberOfBotAgents = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["NumberOfAgents"].ToString());
                            
                            string appCode = System.Configuration.ConfigurationManager.AppSettings["SOP.AppCode"].ToString();
                            string sopName = System.Configuration.ConfigurationManager.AppSettings["SOP.Name"].ToString();
                            string botProcess = System.Configuration.ConfigurationManager.AppSettings["SOP.BotProcess"].ToString();//"Verification_DN_MultiBot_" + facility;
                            ResetBotsExecutionStatus();


                            try
                            {

                                Global.WriteLog("START - Splitting the files");
                                SplitFiles();
                                Global.WriteLog("END - Splitting the files");
                            }
                            catch (Exception ex)
                            {

                                throw new Exception("An error occured while splitting files. Error:" + ex.Message);
                            }


                            try
                            {
                                Global.WriteLog("START - Sending SOP request");
                                //sopName = sopName + facility;
                                for (int i = 1; i <= numberOfBotAgents; i++)
                                {
                                    ServicePointManager.Expect100Continue = true;
                                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;

                                    ServicePointManager.ServerCertificateValidationCallback = delegate (object s,
                                        System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                                        System.Security.Cryptography.X509Certificates.X509Chain chain,
                                        System.Net.Security.SslPolicyErrors sslPolicyErrors)
                                    {
                                        return true;
                                    };

                                    var httpWebRequest = (HttpWebRequest)WebRequest.Create(sopUrl);
                                    //req.Connection = "Close";
                                    httpWebRequest.Method = "POST";
                                    httpWebRequest.Headers["Authorization"] = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(authorization));
                                    //req.Credentials = new NetworkCredential("username", "password");
                                    httpWebRequest.ContentType = "application/json";
                                    httpWebRequest.KeepAlive = false;
                                    
                                    
                                      //"_DN_OnHold_VerificationAndDNRaising_SOP_"
                                    var reqHeaderData = new Requestheaderdata()
                                    {
                                        botprocess = botProcess,
                                        timeoutseconds = "1000",
                                        appcode = appCode,
                                        userid = "sbsuperadmin",
                                        stepid = "",
                                        correlationid = "",
                                        userrole = ""
                                    };
                                    var sopexecjson = new Sopexecjson()
                                    {
                                        sopName = i + sopName,
                                        requestHeaderData = reqHeaderData,
                                        requestRulesData = new Requestrulesdata() { maxage = "30" },
                                        execSOPParamsData = new Execsopparamsdata() { }
                                    };


                                    var reqbody = new Rootobject() { sopExecJSON = sopexecjson };
                                    var requestData = JsonConvert.SerializeObject(reqbody);
                                    Global.WriteLog(string.Format("SOP Request data # {0} ", requestData));
                                    var bytes = Encoding.ASCII.GetBytes(requestData);

                                    httpWebRequest.ContentLength = bytes.Length;

                                    using (var outputStream = httpWebRequest.GetRequestStream())
                                    {
                                        outputStream.Write(bytes, 0, bytes.Length);
                                    }

                                    HttpWebResponse resp = httpWebRequest.GetResponse() as HttpWebResponse;
                                    Global.WriteLog(string.Format("SOP Request for Agent # {0} sent.", i));
                                    Global.WriteLog(string.Format("Response from portal on Account Verification {0}", DateTime.Now.ToString("dd MMMM yyyy hh:mm:ss")));
                                    Global.WriteLog(string.Format("Response Code : {0}", resp.StatusCode));
                                    
                                    httpWebRequest.Abort();

                                }
                                Global.WriteLog("END - Sending SOP request");
                            }
                            catch (Exception ex)
                            {

                                throw new Exception("An error occured while sending the SOP request. Error:" + ex.Message);
                            }

                            bool isAllAgentsExecutionCompleted = false;

                            //try
                            //{
                            for (int i = 0; i < 480; i++)  //Check the status for Max bot time out which is 8 hours
                            {

                                string agentsExecutionStatus = CheckBotsExecutionStatus();

                                if (agentsExecutionStatus == "Started")
                                {
                                    Global.WriteLog("Waiting for Bot agents to complete the execution..");
                                    Thread.Sleep(60000);
                                }
                                else if (agentsExecutionStatus == "Completed")
                                {
                                    isAllAgentsExecutionCompleted = true;
                                    Global.WriteLog("All Bot Agents execution completed");
                                    break;
                                }
                                else if (agentsExecutionStatus == "Failed")
                                {
                                    isAllAgentsExecutionCompleted = true;
                                    Global.WriteLog("One of the Bot Agents execution failed. Check all the agents log for more details.");
                                    //throw new Exception("One of the Bot Agents execution failed. Check all the agents log for more details.");
                                }
                                else if (agentsExecutionStatus == "Pending")
                                {
                                    Global.WriteLog("One or more agents not started yet.");
                                    Thread.Sleep(60000);
                                }

                            }
                            //}
                            //catch (Exception ex)
                            //{

                            //    throw new Exception("An error occured while Checking the bot execution status in JSON. Error:" + ex.Message);
                            //}

                            if (isAllAgentsExecutionCompleted)
                            {
                                try
                                {

                                    Global.WriteLog("START - Merging the files");
                                    MergeFiles();
                                    Global.WriteLog("END - Merging the files");
                                }
                                catch (Exception ex)
                                {

                                    throw new Exception("An error occuring while merging the files. Error: " + ex.Message);
                                }

                                ////try
                                ////{
                                ////    Global.WriteLog("START - Archiving the existing files");
                                ////    ArchiveExistingFiles(false);
                                ////    Global.WriteLog("END - Archiving the existing files");
                                ////}
                                ////catch (Exception ex)
                                ////{

                                ////    throw new Exception("An error occured while Archiving the Existing files. Error:" + ex.Message);
                                ////}

                            }
                        }
                    }
                    else
                    {
                        throw new Exception("File not found in the source directory");
                    }

                }

                //if (regTimeDiff % Convert.ToInt32(EmailSending) == 0 && regTimeDiff > 0)
                //{
                //    Global.WriteLog(timeOfStart.ToString());
                //    Global.WriteLog(timeOfFirstStart.ToString());
                //    Global.WriteLog(timeDiff.ToString());
                //    Global.WriteLog(regTimeDiff.ToString());
                //    //Task AccountVerification = Task.Factory.StartNew(CallAccountVerification);
                //    //AccountVerification.Start();
                //    //sendEmail();
                //}
            }
            catch (Exception ex)
            {

                Global.WriteLog("DN Automation Run Failed. Error: " + ex.Message + " " + "Inner Exception:" + ex.InnerException);
                try
                {
                    EmailHelper email = new EmailHelper(senderEmail, senderPassword, host, port);
                    string automationTeamEmail = ConfigurationManager.AppSettings["Email.AutomationTeam"];
                    //string facility = System.Configuration.ConfigurationManager.AppSettings["Facility"].ToString();
                    string messagebody = "Bot Execution failed. Error: " + ex.Message + " " + "Inner Exception:" + ex.InnerException;
                    email.SendEMail(automationTeamEmail, "ALERT : DN Automation Run Failed!! - " + facility, messagebody, "");

                }
                catch (Exception exp)
                {

                    Global.WriteLog("An error occured while sending email." + "Error: " + exp.Message);
                }
            }


        }

        private void CallAccountCategorization()
        {
            try
            {
                Global.WriteLog(string.Format("Called Account Categorization {0}", DateTime.Now.ToString("dd MMMM yyyy hh:mm:ss")));

                for (int i = 0; i < 10000; i++)
                {
                    Item BotResp = BotResponse();
                    Global.WriteLog("Json - " + BotResp.BotTypeToStart + " - " + BotResp.Execution);
                    if (BotResp.Execution == "Completed")
                        break;
                    else
                        Thread.Sleep(60000);
                }
                Global.WriteLog("Out of thread wait");
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;

                ServicePointManager.ServerCertificateValidationCallback = delegate(object s,
                    System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                    System.Security.Cryptography.X509Certificates.X509Chain chain,
                    System.Net.Security.SslPolicyErrors sslPolicyErrors)
                {
                    return true;
                };
                Global.WriteLog("I am Here");
                var req = (HttpWebRequest)WebRequest.Create(@"https://10.14.48.21:8444/syntbotssm/rest/executeSOPBySOPName");
                req.Method = "POST";
                req.KeepAlive = false;
                req.Headers["Authorization"] = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes("sbsuperadmin:Med2O!sChnTn"));
                //req.Credentials = new NetworkCredential("username", "password");
                req.ContentType = "application/json";

                var reqHeaderData = new Requestheaderdata()
                {
                    botprocess = "AccountCategorization",
                    timeoutseconds = "1000",
                    appcode = "DN001",
                    userid = "sbsuperadmin",
                    stepid = "",
                    correlationid = "",
                    userrole = ""
                };
                var sopexecjson = new Sopexecjson()
                {
                    sopName = "DN_OnHold_AccountCategorization_SOP",
                    requestHeaderData = reqHeaderData,
                    requestRulesData = new Requestrulesdata() { maxage = "30" },
                    execSOPParamsData = new Execsopparamsdata() { }
                };

                var reqbody = new Rootobject() { sopExecJSON = sopexecjson };
                var requestData = JsonConvert.SerializeObject(reqbody);
                Global.WriteLog("I am Here2");
                var bytes = Encoding.ASCII.GetBytes(requestData);
                req.ContentLength = bytes.Length;
                using (var outputStream = req.GetRequestStream())
                {
                    outputStream.Write(bytes, 0, bytes.Length);
                }
                Global.WriteLog("I am Here3");
                HttpWebResponse resp = req.GetResponse() as HttpWebResponse;
                //req.Connection = "Close";
                Global.WriteLog(string.Format("Response from portal on Account Categorization {0}", DateTime.Now.ToString("dd MMMM yyyy hh:mm:ss")));
                req.Abort();
            }
            catch (Exception ex)
            {
                Global.WriteLog(ex.InnerException.ToString());
            }
        }

        private void CallAccountVerification()
        {
            try
            {
                Global.WriteLog(string.Format("Called Account Verification {0}", DateTime.Now.ToString("dd MMMM yyyy hh:mm:ss")));
                for (int i = 0; i < 10000; i++)
                {
                    Item BotResp = BotResponse();
                    Global.WriteLog("Json - " + BotResp.BotTypeToStart + " - " + BotResp.Execution);
                    if (BotResp.Execution == "Completed")
                        break;
                    else
                        Thread.Sleep(60000);
                }
                Global.WriteLog("Out of thread wait");
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;

                ServicePointManager.ServerCertificateValidationCallback = delegate(object s,
                    System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                    System.Security.Cryptography.X509Certificates.X509Chain chain,
                    System.Net.Security.SslPolicyErrors sslPolicyErrors)
                {
                    return true;
                };

                Global.WriteLog("I am Here");
                var req = (HttpWebRequest)WebRequest.Create(@"https://10.14.48.21:8444/syntbotssm/rest/executeSOPBySOPName");
                //req.Connection = "Close";
                req.Method = "POST";
                req.Headers["Authorization"] = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes("sbsuperadmin:Med2O!sChnTn"));
                //req.Credentials = new NetworkCredential("username", "password");
                req.ContentType = "application/json";
                req.KeepAlive = false;
                var reqHeaderData = new Requestheaderdata()
                {
                    botprocess = "Verification_DNRaising",
                    timeoutseconds = "1000",
                    appcode = "DN003",
                    userid = "sbsuperadmin",
                    stepid = "",
                    correlationid = "",
                    userrole = ""
                };
                var sopexecjson = new Sopexecjson()
                {
                    sopName = "DN_OnHold_VerificationAndDNRaising_SOP",
                    requestHeaderData = reqHeaderData,
                    requestRulesData = new Requestrulesdata() { maxage = "30" },
                    execSOPParamsData = new Execsopparamsdata() { }
                };

                var reqbody = new Rootobject() { sopExecJSON = sopexecjson };
                var requestData = JsonConvert.SerializeObject(reqbody);
                Global.WriteLog("I am Here2");
                var bytes = Encoding.ASCII.GetBytes(requestData);
                Global.WriteLog("I am Here21");
                req.ContentLength = bytes.Length;
                Global.WriteLog("I am Here22");
                using (var outputStream = req.GetRequestStream())
                {
                    outputStream.Write(bytes, 0, bytes.Length);
                }
                Global.WriteLog("I am Here23");
                HttpWebResponse resp = req.GetResponse() as HttpWebResponse;
                Global.WriteLog("I am Here3");

                Global.WriteLog(string.Format("Response from portal on Account Verification {0}", DateTime.Now.ToString("dd MMMM yyyy hh:mm:ss")));
                req.Abort();
            }
            catch (Exception ex)
            {
                Global.WriteLog(ex.InnerException.ToString());
            }
        }
        protected override void OnStop()
        {

        }
        public Item BotResponse()
        {
            string jsonFilePath = System.Configuration.ConfigurationManager.AppSettings["jsonFilePath"].ToString();
            string json = File.ReadAllText(jsonFilePath);
            var jobj = JObject.Parse(json);

            Item BotResp = new Item
            {
                BotTypeToStart = jobj.Property("BotTypeToStart").Value.ToString(),
                Execution = jobj.Property("Execution").Value.ToString()
            };
            return BotResp;
        }

        public class Item
        {
            public string BotTypeToStart;
            public string Execution;
        }

        public void sendEmail()
        {
            string receiverEmail = System.Configuration.ConfigurationManager.AppSettings["Email.ReceiverEmail"] != null ? System.Configuration.ConfigurationManager.AppSettings["Email.ReceiverEmail"].ToString() : string.Empty; //"osman.alikhan@atos.net"; //ConfigurationManager.AppSettings["Email.SenderEmail"]; //
            string senderEmail = System.Configuration.ConfigurationManager.AppSettings["Email.SenderEmail"] != null ? System.Configuration.ConfigurationManager.AppSettings["Email.SenderEmail"].ToString() : string.Empty; //"osman.alikhan@atos.net"; //ConfigurationManager.AppSettings["Email.SenderEmail"]; //
            string senderPassword = System.Configuration.ConfigurationManager.AppSettings["Email.SenderPassword"] != null ? System.Configuration.ConfigurationManager.AppSettings["Email.SenderPassword"].ToString() : string.Empty; //ConfigurationManager.AppSettings["Email.Password"];//
            string host = System.Configuration.ConfigurationManager.AppSettings["Email.Host"] != null ? System.Configuration.ConfigurationManager.AppSettings["Email.Host"].ToString() : string.Empty;//"10.14.32.146";// ConfigurationManager.AppSettings["Email.Host"]; //"10.14.32.146";
            int port = System.Configuration.ConfigurationManager.AppSettings["Email.Port"] != null ? Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Email.Port"].ToString()) : 0;//25;//Convert.ToInt32(ConfigurationManager.AppSettings["Email.Port"]); //25;
            string rootFolder = System.Configuration.ConfigurationManager.AppSettings["RootFolder"].ToString();
            Global.WriteLog(rootFolder);
            string toBeAllocatedfilepath = string.Empty;
            string toBesentfilepath = string.Empty;
            DirectoryInfo dir = new DirectoryInfo(rootFolder);
            FileInfo[] toBeAllocatedfiles = dir.GetFiles("MasterExcel*ToBeAllocated*", SearchOption.TopDirectoryOnly);
            if (toBeAllocatedfiles.Count() > 0)
            {
                toBeAllocatedfilepath = toBeAllocatedfiles[0].FullName;
                toBesentfilepath = toBeAllocatedfiles[0].DirectoryName + "\\" + DateTime.Now.ToString().Replace(":", "-").Replace("/", "-") + toBeAllocatedfiles[0].Name;
                Global.WriteLog(toBesentfilepath);
            }
            System.IO.File.Copy(toBeAllocatedfilepath, toBesentfilepath);
            var attachments = new List<string>();
            attachments.Add(toBesentfilepath);

            Global.WriteLog(toBesentfilepath);
            try
            {
                string subject = System.Configuration.ConfigurationManager.AppSettings["Email.ToBeAllocated.Subject"].ToString();
                subject = subject + " " + DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                string messageBody = System.Configuration.ConfigurationManager.AppSettings["Email.ToBeAllocated.MessageBody"].ToString();
                Global.WriteLog(subject);
                EmailHelper email = new EmailHelper(senderEmail, senderPassword, host, port);
                //string receiverEmail = "osman.alikhan@atos.net";//ConfigurationManager.AppSettings["Email.ReceiverEmail"];
                Global.WriteLog("Sending email to " + receiverEmail);
                email.SendEMail(receiverEmail, subject, messageBody, attachments);

                Global.WriteLog("Email sent with the attachments");
            }
            catch (Exception ex)
            {
                Global.WriteLog("An error occured while sending email." + "Error: " + ex.Message);
            }
        }

        public void MergeFiles(bool isEmptyFile = false)
        {

            Configuration _configuration = ConfigurationManager.OpenExeConfiguration(this.GetType().Assembly.Location);
            string rootFolder = _configuration.AppSettings.Settings["RootFolderPath"] != null ? _configuration.AppSettings.Settings["RootFolderPath"].Value : string.Empty;
            string archivedFolder = _configuration.AppSettings.Settings["ArchieveFolderPath"] != null ? _configuration.AppSettings.Settings["ArchieveFolderPath"].Value : string.Empty; //@"E:\TaskQueue\";//ConfigurationManager.AppSettings["AccessHIM.MasterExcel.Path"];//@"E:\TaskQueue\Archived\";
            string receiverEmail = _configuration.AppSettings.Settings["Email.ReceiverEmail"] != null ? _configuration.AppSettings.Settings["Email.ReceiverEmail"].Value : string.Empty; //"osman.alikhan@atos.net"; //ConfigurationManager.AppSettings["Email.SenderEmail"]; //
            string senderEmail = _configuration.AppSettings.Settings["Email.SenderEmail"] != null ? _configuration.AppSettings.Settings["Email.SenderEmail"].Value : string.Empty; //"osman.alikhan@atos.net"; //ConfigurationManager.AppSettings["Email.SenderEmail"]; //
            string senderPassword = _configuration.AppSettings.Settings["Email.SenderPassword"] != null ? _configuration.AppSettings.Settings["Email.SenderPassword"].Value : string.Empty; //ConfigurationManager.AppSettings["Email.Password"];//
            string host = _configuration.AppSettings.Settings["Email.Host"] != null ? _configuration.AppSettings.Settings["Email.Host"].Value : string.Empty;//"10.14.32.146";// ConfigurationManager.AppSettings["Email.Host"]; //"10.14.32.146";
            int port = _configuration.AppSettings.Settings["Email.Port"] != null ? Convert.ToInt32(_configuration.AppSettings.Settings["Email.Port"].Value) : 0;//25;//Convert.ToInt32(ConfigurationManager.AppSettings["Email.Port"]); //25;

            //Modified by Velan on 19/01/2022, for Flint and Lapeer Go Live
            //int dnRaiseTime = _configuration.AppSettings.Settings["DNRaiseTime"] != null ? Convert.ToInt32(_configuration.AppSettings.Settings["DNRaiseTime"].Value) : 0;//25;//Convert.ToInt32(ConfigurationManager.AppSettings["Email.Port"]); //25;

            var DNRaiseTimes = _configuration.AppSettings.Settings["DNRaiseTime"].Value.ToString().Split(',');

            

            string facility = _configuration.AppSettings.Settings["Facility"] != null ? _configuration.AppSettings.Settings["Facility"].Value : string.Empty;

            string jsonFilePath = _configuration.AppSettings.Settings["BotResponseJson.FilePath"] != null ? _configuration.AppSettings.Settings["BotResponseJson.FilePath"].Value : string.Empty; //@"E:\Syntbots\BotResponse.json";

            string botAgentFileName = _configuration.AppSettings.Settings["BotAgentToBeProcessedFileName"] != null ? _configuration.AppSettings.Settings["BotAgentToBeProcessedFileName"].Value : string.Empty;

            //Microsoft.Office.Interop.Excel.Application oXL;
            string facilities = _configuration.AppSettings.Settings["ListOfFacilities"] != null ? _configuration.AppSettings.Settings["ListOfFacilities"].Value : string.Empty;

            //Added by Velan on 17/02/2022 for Flint and Lapeer Go Live

            int FEN_FLT_OAK_MAC_BAY_dnRaiseTime = _configuration.AppSettings.Settings["FEN_FLT_OAK_MAC_BAY_DNRaiseTime"] != null ? Convert.ToInt32(_configuration.AppSettings.Settings["FEN_FLT_OAK_MAC_BAY_DNRaiseTime"].Value) : 0;
            int CEN_CLK_LAP_dnRaiseTime = _configuration.AppSettings.Settings["CEN_CLK_LAP_DNRaiseTime"] != null ? Convert.ToInt32(_configuration.AppSettings.Settings["CEN_CLK_LAP_DNRaiseTime"].Value) : 0;

            // For FEN_FLT_OAK_MAC_BAY run
            int currentTime1 = (int)DateTime.Now.TimeOfDay.TotalMinutes;
            int timediff1;
            if (FEN_FLT_OAK_MAC_BAY_dnRaiseTime <= currentTime1)
            {

                //sameday
                Global.WriteLog(" Merge Automation completed on Same Day");
                timediff1 = currentTime1 - FEN_FLT_OAK_MAC_BAY_dnRaiseTime;
            }
            else
            {

                //nextday
                Global.WriteLog("Merge Automation completed on Next Day");
                currentTime1 = currentTime1 + 1440;
                timediff1 = currentTime1 - FEN_FLT_OAK_MAC_BAY_dnRaiseTime;
            }

            if (timediff1 <= 60)
            {
                facilities = _configuration.AppSettings.Settings["FEN_FLT_OAK_MAC_BAY_Facilities"] != null ? _configuration.AppSettings.Settings["FEN_FLT_OAK_MAC_BAY_Facilities"].Value : string.Empty;

            }



            //For CEN_CLK_LAP run

            int currentTime2 = (int)DateTime.Now.TimeOfDay.TotalMinutes;
            int timediff2;
            if (CEN_CLK_LAP_dnRaiseTime <= currentTime2)
            {

                //sameday
                Global.WriteLog(" Merge Automation completed on Same Day");
                timediff2 = currentTime2 - CEN_CLK_LAP_dnRaiseTime;
            }
            else
            {

                //nextday
                Global.WriteLog(" Merge Automation completed on Same Day");
                currentTime2 = currentTime2 + 1440;
                timediff2 = currentTime2 - CEN_CLK_LAP_dnRaiseTime;
            }

            if (timediff2 <= 60)
            {
                facilities = _configuration.AppSettings.Settings["CEN_CLK_LAP_Facilities"] != null ? _configuration.AppSettings.Settings["CEN_CLK_LAP_Facilities"].Value : string.Empty;
            }

            //Global.WriteLog(" current time2 :" + currentTime2);
            //Global.WriteLog(" CEN_CLK_LAP_dnRaiseTime :" + CEN_CLK_LAP_dnRaiseTime);
            //Global.WriteLog(" Time difference :"+timediff2);
            //Global.WriteLog(" Facility :" + facilities);




            var listOfFaclities = facilities.Split(',');
            //var listOfFaclities = new List<string>() { "Lansing" };

            DirectoryInfo dir = new DirectoryInfo(rootFolder);
            FileInfo[] botOutPutFiles = dir.GetFiles(string.Format("*{0}*", botAgentFileName), SearchOption.TopDirectoryOnly);

            foreach (string facility1 in listOfFaclities)
            {

                try
                {
                    Global.WriteLog("Starting merging process");

                    //Microsoft.Office.Interop.Excel.Workbook toBeProcessedWorkBook;
                    //Microsoft.Office.Interop.Excel.Sheets toBeProcessedWorkSheets;
                    //Microsoft.Office.Interop.Excel.Worksheet toBeProcessedWSheet;
                    //Microsoft.Office.Interop.Excel.Workbook toBeAllocatedWorkBook;
                    //Microsoft.Office.Interop.Excel.Sheets toBeAllocatedWorkSheets;
                    //Microsoft.Office.Interop.Excel.Worksheet toBeAllocatedWSheet;

                    //Microsoft.Office.Interop.Excel.Workbook botToBeProcessedWorkBook;
                    //Microsoft.Office.Interop.Excel.Sheets botToBeProcessedWorkSheets;
                    //Microsoft.Office.Interop.Excel.Worksheet botToBeProcessedWSheet;

                    //string toBeProcessedfilepath = string.Empty;//fillePath + "_ToBeProcessed.xls";
                    //string toBeAllocatedfilepath = string.Empty;

                    //"D:\\TaskQueue\\"
                    //FileInfo[] toBeProcessedfiles = dir.GetFiles("*ToBeProcessed*", SearchOption.TopDirectoryOnly);
                    //if (toBeProcessedfiles.Count() > 0)
                    //{
                    //    toBeProcessedfilepath = toBeProcessedfiles[0].FullName;
                    //    List<Order> FINlist = File.ReadLines(toBeProcessedfilepath).Skip(1).Select(line =>
                    //    {
                    //        Order ch = new Order();
                    //        Regex CSVParser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");
                    //        String[] split = CSVParser.Split(line);
                    //        ch.FIN = split[6].Replace("\"", ""); ;
                    //        ch.DischargeDate = split[8].Replace("\"", "");
                    //        return ch;
                    //    }).ToList<Order>();
                    //}
                    //FileInfo toBeProcessedInfo = new FileInfo(toBeProcessedfilepath);

                    //FileInfo[] toBeAllocatedfiles = dir.GetFiles("*ToBeAllocated*", SearchOption.TopDirectoryOnly);
                    //if (toBeAllocatedfiles.Count() > 0)
                    //{
                    //    toBeAllocatedfilepath = toBeAllocatedfiles[0].FullName;
                    //}



                    //oXL = new Microsoft.Office.Interop.Excel.Application();
                    //oXL.Visible = true;
                    //oXL.DisplayAlerts = false;

                    //var _workbooks = oXL.Workbooks;

                    //toBeProcessedWorkBook = oXL.Workbooks.Open(toBeProcessedfilepath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //toBeProcessedWorkSheets = toBeProcessedWorkBook.Worksheets;
                    //toBeProcessedWSheet = (Microsoft.Office.Interop.Excel.Worksheet)toBeProcessedWorkSheets.get_Item(1);
                    //Microsoft.Office.Interop.Excel.Range toBeProcessedRange = toBeProcessedWSheet.UsedRange;
                    //int columnsCount = toBeProcessedRange.Columns.Count;
                    //int rowsCount = toBeProcessedRange.Rows.Count;

                    //toBeAllocatedWorkBook = oXL.Workbooks.Open(toBeAllocatedfilepath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //toBeAllocatedWorkSheets = toBeAllocatedWorkBook.Worksheets;
                    //toBeAllocatedWSheet = (Microsoft.Office.Interop.Excel.Worksheet)toBeAllocatedWorkSheets.get_Item(1);
                    //Microsoft.Office.Interop.Excel.Range toBeAllocatedRange = toBeAllocatedWSheet.UsedRange;
                    //int toBeAllocatedColumnsCount = toBeAllocatedRange.Columns.Count;
                    //int toBeAllocatedRowsCount = 0;//toBeAllocatedRange.Rows.Count;

                    int totalRecords = 0;
                    int totalGoodRecords = 0;
                    int totalDNRaised = 0;
                    int totalErrorRecords = 0;
                    int totalDataErrorRecords = 0;
                    int totalBadRecords = 0;
                    int totalSuspendedRecords = 0;
                    string FINnumbersWithError = "";
                    string FINnumbersWithDataError = "";
                    bool shouldDNBeRaised = false;
                    int MinutesPassed = (int)DateTime.Now.TimeOfDay.TotalMinutes;









                    //Commented by Velan for DN Run Issue. The Time difference is getting confused when it goes to next day
                    //if (MinutesPassed > dnRaiseTime && MinutesPassed < (dnRaiseTime + 60))
                    //{
                    //    shouldDNBeRaised = true;
                    //}



                    //Modified by Velan on 19/01/2022, for Flint and Lapeer Go Live

                    //Added by Velan for DN Run Issue. The Time difference is getting confused when it goes to next day
                    //int currentTime = (int)DateTime.Now.TimeOfDay.TotalMinutes;
                    //int AutomationTime = dnRaiseTime;

                    //int timediff;

                    //if (AutomationTime <= currentTime)
                    //{

                    //    //sameday
                    //    Global.WriteLog(" Automation completed on Same Day");
                    //    timediff = currentTime - AutomationTime;
                    //}
                    //else
                    //{

                    //    //nextday
                    //    Global.WriteLog(" Automation completed on Next Day");
                    //    currentTime = currentTime + 1440;
                    //    timediff = currentTime - AutomationTime;
                    //}

                    //if (timediff <= 60)
                    //{
                    //    Global.WriteLog(" This is DN Run, Time Difference" + timediff);
                    //    shouldDNBeRaised = true;
                    //}



                    foreach (var d in DNRaiseTimes)
                    {
                        int dnRaiseTime = Int32.Parse(d);                     



                        int currentTime = (int)DateTime.Now.TimeOfDay.TotalMinutes;
                        
                        int AutomationTime = dnRaiseTime;

                        int timediff;

                        if (AutomationTime <= currentTime)
                        {

                            //sameday
                            Global.WriteLog(" Automation completed on Same Day");
                            timediff = currentTime - AutomationTime;
                        }
                        else
                        {

                            //nextday
                            Global.WriteLog(" Automation completed on Next Day");
                            currentTime = currentTime + 1440;
                            timediff = currentTime - AutomationTime;
                        }

                        if (timediff <= 60)
                        {
                            Global.WriteLog(" This is DN Run, Time Difference" + timediff);
                            shouldDNBeRaised = true;
                            break;
                        }
                    }









                    if (!isEmptyFile)
                    {
                        for (int i = 0; i < botOutPutFiles.Length; i++)
                        {
                            try
                            {
                                Global.WriteLog("Processing the Bot output file: " + botOutPutFiles[i].FullName);

                                List<Order> FINlist = File.ReadLines(botOutPutFiles[i].FullName).Skip(1).Select(line =>
                                {
                                    Order ch = new Order();
                                    Regex CSVParser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");
                                    String[] split = CSVParser.Split(line);
                                    ch.FIN = split[6].Replace("\"", "");
                                    ch.Comments = split[30].Replace("\"", "");
                                    ch.Organization = split[11].Replace("\"", "");
                                    return ch;
                                }).ToList<Order>();

                                List<Order> result = FINlist.Where(a => a.Organization.Contains(facility1)).ToList(); //Select(a => a.Facility == facility1);

                                totalRecords = totalRecords + result.Count();


                                foreach (var Encounter in result)
                                {
                                    string FIN = string.Empty;
                                    string botStatus = string.Empty;

                                    FIN = Encounter.FIN;
                                    botStatus = Encounter.Comments;

                                    if (string.IsNullOrEmpty(FIN))
                                    {
                                        break;
                                    }
                                    else
                                    {
                                        if (!string.IsNullOrEmpty(botStatus))
                                        {

                                            //If it is a good record append to tobeAllocatedFile
                                            if (botStatus == "Good Record" || botStatus == "Description not found" || botStatus == "Excluded CPT")
                                            {

                                                if (botStatus == "Description not found" || botStatus == "Excluded CPT")
                                                {
                                                    totalDataErrorRecords++;
                                                    FINnumbersWithDataError = FINnumbersWithDataError + "," + FIN;
                                                }
                                                else
                                                {

                                                    totalGoodRecords++;
                                                }

                                            }
                                            else if (botStatus.Contains("DN Raised"))
                                            {
                                                totalDNRaised++;

                                            }
                                            else if (botStatus == "Failed to process")
                                            {

                                                totalErrorRecords++;
                                                FINnumbersWithError = FINnumbersWithError + "," + FIN;

                                            }
                                            else if (botStatus.Contains("Suspended Charge"))
                                            {
                                                totalSuspendedRecords++;
                                            }
                                            else
                                            {
                                                totalBadRecords++;
                                                //toBeProcessedFindRange.Columns[commentsColumnIndex] = botStatus;
                                                //toBeProcessedFindRange.Columns[commentsColumnIndex + 1] = details;

                                            }
                                        }

                                    }


                                }


                                Global.WriteLog("Completed the Processing of the Bot output file: " + botOutPutFiles[i].FullName);

                                //Global.WriteLog("Start - Deleting Bot output file: " + botOutPutFiles[i].FullName);
                                //System.IO.File.Delete(botOutPutFiles[i].FullName);
                                //Global.WriteLog("END - Deleting Bot output file: " + botOutPutFiles[i].FullName);


                            }
                            catch (Exception ex)
                            {
                                Global.WriteLog("An error occured while merging the Bot output file" + botOutPutFiles[i].FullName + ". Error message:" + ex.Message + " Inner exception: " + ex.InnerException);
                            }

                        }


                        Global.WriteLog("Completed the merging process");
                    }

                    if (true)
                    {
                        //totalRecords = totalGoodRecords + totalDNRaised + totalErrorRecords + totalBadRecords;
                        System.Text.StringBuilder sbMessageBody = new System.Text.StringBuilder();
                        sbMessageBody.Append(string.Format("Hi Team, <br><br> Please see below the summary for coding – OP Lab only accounts {0}.", DateTime.Now.ToString("MM-dd-yyyy HH:mm")));

                        sbMessageBody.Append("<br><table align ='center'>");
                        sbMessageBody.Append("<tr><th colspan='2' bgcolor='#BDD7EE'>Status</th></tr>");

                        sbMessageBody.Append("<tr><td>Total Inventory</td>");
                        sbMessageBody.Append(string.Format("<td> {0} </td></tr>", totalRecords));

                        sbMessageBody.Append("<tr><td>Good a/c without discrepancy</td>");
                        sbMessageBody.Append(string.Format("<td> {0} </td></tr>", totalGoodRecords));

                        sbMessageBody.Append("<tr><td>Total a/c with discrepancy</td>");
                        sbMessageBody.Append(string.Format("<td> {0} </td></tr>", totalBadRecords));

                        sbMessageBody.Append("<tr><td>Total a/c with suspended charge</td>");
                        sbMessageBody.Append(string.Format("<td> {0} </td></tr>", totalSuspendedRecords));


                        sbMessageBody.Append("<tr><td>a/c for which DN Sent</td>");
                        sbMessageBody.Append(string.Format("<td> {0} </td></tr>", totalDNRaised));


                        sbMessageBody.Append("<tr><td>a/c which require Manual Intervention </td>");
                        sbMessageBody.Append(string.Format("<td> {0} </td></tr>", totalDataErrorRecords));

                        sbMessageBody.Append("<tr><td>Exceptions a/c (Due to system issue)</td>");
                        sbMessageBody.Append(string.Format("<td> {0} </td></tr>", totalErrorRecords));
                        FINnumbersWithDataError = FINnumbersWithDataError.StartsWith(",") ? FINnumbersWithDataError.Remove(0, 1) : FINnumbersWithDataError;
                        FINnumbersWithError = FINnumbersWithError.StartsWith(",") ? FINnumbersWithError.Remove(0, 1) : FINnumbersWithError;
                        sbMessageBody.Append("</table>");
                        sbMessageBody.Append("<br><b>Accounts for manual interventions: </b>" + FINnumbersWithDataError);
                        sbMessageBody.Append("<br><b>Accounts failed due to system issues: </b>" + FINnumbersWithError);
                        sbMessageBody.Append("<br><br>Note: Manual DN needs to be raised for ‘Exceptions accounts’, details available in ‘to be processed sheet’. <br>Should you have any queries, please feel free to contact Syntbots team.");
                        sbMessageBody.Append("<br><br>Thank you");
                        sbMessageBody.Append("<br>BOT Signature");
                        sbMessageBody.Append("<br>(Need to mention the contact details)");
                        sbMessageBody.Append("<br><br><br>*****This is an auto-generated email, please do not reply. ******");

                        string subject = string.Empty;
                        if (shouldDNBeRaised)
                        {
                            subject = string.Format(facility1 + " = End of the day summary & inventory details {0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm"));//ConfigurationManager.AppSettings["Email.Subject"];
                        }
                        else
                        {
                            subject = string.Format(facility1 + " - Discern Notification Summary & Inventory Details {0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm"));//ConfigurationManager.AppSettings["Email.Subject"];
                        }
                        subject = isEmptyFile ? subject + " - EMPTY FILE" : subject;

                        string messageBody = sbMessageBody.ToString(); //ConfigurationManager.AppSettings["Email.MessageBody"];
                        var attachments = new List<string>();
                        //attachments.Add(toBeAllocatedfilepath);
                        //attachments.Add(toBeProcessedfilepath);

                        try
                        {
                            EmailHelper email = new EmailHelper(senderEmail, senderPassword, host, port);

                            Global.WriteLog("Sending email to " + receiverEmail);
                            email.SendEMail(receiverEmail, subject, messageBody, attachments);
                            Global.WriteLog("Email sent.");
                        }
                        catch (Exception ex)
                        {
                            Global.WriteLog("An error occured while sending email." + "Error: " + ex.Message);
                        }
                    }
                    //else
                    //{
                    //    System.Text.StringBuilder sbMessageBody = new System.Text.StringBuilder();
                    //    sbMessageBody.Append("Hi Team, <br>Please find attached details of FIN# which need to be allocated to SME’s for coding.");
                    //    sbMessageBody.Append("<br><br>Should you have any queries, please feel free to contact Syntbots team.");
                    //    sbMessageBody.Append("<br><br>Thank you");
                    //    sbMessageBody.Append("<br>BOT Signature");
                    //    sbMessageBody.Append("<br>(Need to mention the contact details)");
                    //    sbMessageBody.Append("<br><br><br>*****This is an auto-generated email, please do not reply. ******");

                    //    string subject = string.Format("Allocation file with good records for {0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm"));//ConfigurationManager.AppSettings["Email.Subject"];
                    //    string messageBody = sbMessageBody.ToString(); //ConfigurationManager.AppSettings["Email.MessageBody"];
                    //    var attachments = new List<string>();
                    //    attachments.Add(toBeAllocatedfilepath);
                    //    attachments.Add(toBeProcessedfilepath);
                    //    try
                    //    {
                    //        EmailHelper email = new EmailHelper(senderEmail, senderPassword, host, port);

                    //        Global.WriteLog("Sending email to " + receiverEmail);
                    //        email.SendEMail(receiverEmail, subject, messageBody, attachments);
                    //        Global.WriteLog("Email sent with the attachment to be allocated file " + toBeAllocatedfilepath);

                    //    }

                    //    catch (Exception ex)
                    //    {

                    //        Global.WriteLog("An error occured while sending email." + "Error: " + ex.Message);
                    //        throw new Exception("An error occured while sending email." + "Error: " + ex.Message);
                    //    }
                    //}


                }
                catch (Exception ex)
                {
                    Global.WriteLog("An Error occured during the mering process. Error message: " + ex.Message + "\n Inner exception: " + ex.InnerException);
                    //objJsnHlpr.SetErrorList(ex.Message.ToString(), "Technical", "006");
                    //objJsnHlpr.SetOutCome("Failure");

                }
                finally
                {
                    //Atos.SyntBots.Common.Logger.Log(jsonFilePath);
                    //string jsonString = File.ReadAllText(jsonFilePath);
                    //JObject jObject = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonString) as JObject;
                    //JToken jBotTypeToStart = jObject.SelectToken("BotTypeToStart");
                    //jBotTypeToStart.Replace("Account Categorization");

                    //JToken jExecution = jObject.SelectToken("Execution");
                    //jExecution.Replace("Completed");


                    //string updatedJsonString = jObject.ToString();
                    //File.WriteAllText(jsonFilePath, updatedJsonString);

                }

            }

            for (int i = 0; i < botOutPutFiles.Length; i++)
            {
                System.IO.File.Move(botOutPutFiles[i].FullName, string.Format(@"{0}\{1}", archivedFolder, botOutPutFiles[i].Name));
            }
        }
                    
        private void CallAccountVerification(HttpWebRequest httpWebRequest, Requestheaderdata requestheaderdata, Sopexecjson sopexecjson)
        {
            try
            {
                Global.WriteLog(string.Format("Called Account Verification {0}", DateTime.Now.ToString("dd MMMM yyyy hh:mm:ss")));
                for (int i = 0; i < 10000; i++)
                {
                    Item BotResp = BotResponse();
                    Global.WriteLog("Json - " + BotResp.BotTypeToStart + " - " + BotResp.Execution);
                    if (BotResp.Execution == "Completed")
                        break;
                    else
                        Thread.Sleep(60000);
                }
                Global.WriteLog("Out of thread wait");
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;

                ServicePointManager.ServerCertificateValidationCallback = delegate (object s,
                    System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                    System.Security.Cryptography.X509Certificates.X509Chain chain,
                    System.Net.Security.SslPolicyErrors sslPolicyErrors)
                {
                    return true;
                };

                Global.WriteLog("I am Here");


                var reqbody = new Rootobject() { sopExecJSON = sopexecjson };
                var requestData = JsonConvert.SerializeObject(reqbody);
                Global.WriteLog("I am Here2");
                var bytes = Encoding.ASCII.GetBytes(requestData);
                Global.WriteLog("I am Here21");
                httpWebRequest.ContentLength = bytes.Length;
                Global.WriteLog("I am Here22");
                using (var outputStream = httpWebRequest.GetRequestStream())
                {
                    outputStream.Write(bytes, 0, bytes.Length);
                }
                Global.WriteLog("I am Here23");
                HttpWebResponse resp = httpWebRequest.GetResponse() as HttpWebResponse;
                Global.WriteLog("I am Here3");

                Global.WriteLog(string.Format("Response from portal on Account Verification {0}", DateTime.Now.ToString("dd MMMM yyyy hh:mm:ss")));
                httpWebRequest.Abort();
            }
            catch (Exception ex)
            {
                Global.WriteLog(ex.InnerException.ToString());
            }
        }


        private string CheckBotsExecutionStatus()
        {
            string jsonFilePath = System.Configuration.ConfigurationManager.AppSettings["jsonFilePath"].ToString();
            string json = File.ReadAllText(jsonFilePath);
            dynamic jsonObj = JsonConvert.DeserializeObject(json);

            string executionStatus = string.Empty;

            int numberOfBotAgents = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["NumberOfAgents"].ToString());

            for (int i = 1; i <= numberOfBotAgents; i++)
            {
                //jsonFilePath = System.Configuration.ConfigurationManager.AppSettings["jsonFilePath"].ToString() + i +".JSON";
                //json = File.ReadAllText(jsonFilePath);
                //jsonObj = JsonConvert.DeserializeObject(json);

                string status = jsonObj["ExecutionStatusBot" + i];
                if (status == "Completed")
                {
                    executionStatus = "Completed";
                }
                else if (status == "Started")
                {
                    executionStatus = "Started";
                    break;
                }
                else if (status == "Failed")
                {
                    executionStatus = "Failed";
                    //break;
                }
                else if (status == "Pending")
                {
                    executionStatus = "Pending";
                    break;
                }

            }
            return executionStatus;
        }

        private void ArchiveExistingFiles(bool isFirstRun)
        {
            string rootFolder = System.Configuration.ConfigurationManager.AppSettings["RootFolderPath"].ToString();
            string archivedFolder = ConfigurationManager.AppSettings["ArchieveFolderPath"];

            #region Move the exiting files to archive folder
            try
            {
                if (isFirstRun)
                {
                    if (Directory.Exists(rootFolder))
                    {
                        foreach (var file in new DirectoryInfo(rootFolder).GetFiles())
                        {
                            file.MoveTo(string.Format(@"{0}\{1}", archivedFolder, file.Name));
                        }
                    }
                }
                else
                {
                    if (Directory.Exists(rootFolder))
                    {
                        foreach (var file in new DirectoryInfo(rootFolder).GetFiles("*BotOutput_Agent*", SearchOption.TopDirectoryOnly))
                        {
                            file.MoveTo(string.Format(@"{0}\{1}", archivedFolder, file.Name));
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                //Atos.SyntBots.Common.Logger.Log("An error occured while archieving existing files. Error: " + ex.Message);
                throw new Exception("An error occured while archieving existing files. Error: " + ex.Message);
            }


            #endregion
        }

        public void SplitFiles()
        {
            string rootFolder = System.Configuration.ConfigurationManager.AppSettings["RootFolderPath"].ToString();
            string NumberOfAgents = System.Configuration.ConfigurationManager.AppSettings["NumberOfAgents"];
            int noOfFiles = Convert.ToInt32(NumberOfAgents);

            string botAgentToBeProcessedFileName = System.Configuration.ConfigurationManager.AppSettings["BotAgentToBeProcessedFileName"];
            #region Split to be processed

            object misValue = System.Reflection.Missing.Value;

            string toBeProcessedfilepath = string.Empty;
            string toBeAllocatedfilepath = string.Empty;

            DirectoryInfo dir = new DirectoryInfo(rootFolder);

            string archivedFolder = System.Configuration.ConfigurationManager.AppSettings["ArchieveFolderPath"] != null ? System.Configuration.ConfigurationManager.AppSettings["ArchieveFolderPath"] : string.Empty;
            FileInfo[] botOutPutFiles = dir.GetFiles(string.Format("*{0}*", botAgentToBeProcessedFileName), SearchOption.TopDirectoryOnly);

            if (botOutPutFiles.Length > 0)
            {
                for (int i = 0; i < botOutPutFiles.Length; i++)
                {
                    System.IO.File.Move(botOutPutFiles[i].FullName, string.Format(@"{0}\{1}", archivedFolder, botOutPutFiles[i].Name));
                }
            }

            FileInfo[] toBeProcessedfiles = dir.GetFiles("*ToBeProcessed*", SearchOption.TopDirectoryOnly);
            if (toBeProcessedfiles.Count() > 0)
            {
                toBeProcessedfilepath = toBeProcessedfiles[0].FullName;
            }

            var masterDataList = File.ReadLines(toBeProcessedfilepath).ToList();
            System.Text.RegularExpressions.Regex CSVParser = new System.Text.RegularExpressions.Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");

            var masterDataFins = masterDataList.Select(line =>
            {
                String[] split = CSVParser.Split(line);
                string fin = split[6].Replace("\"", "");
                return fin;
            }).ToList<string>();

            //int finsCount = masterDataList.Count() - 1;
            int toBeProcessedRowCount = masterDataList.Count() - 1;     // Count excludes the header row

            int dataRowsCount = toBeProcessedRowCount;
            int recordsPerFile = (dataRowsCount - (dataRowsCount % noOfFiles)) / noOfFiles;
            int mod = dataRowsCount % noOfFiles;
            int take = 0;
            int beginIndex = 1;

            if (dataRowsCount >= noOfFiles)
            {
                for (int i = 1; i <= noOfFiles; i++)
                {
                    take = i == 1 ? recordsPerFile + mod : recordsPerFile;
                    var list = masterDataList.GetRange(beginIndex, take);
                    string[] headerArrItem = new string[1] { masterDataList.ToList()[0] };
                    list.InsertRange(0, headerArrItem);
                    string csvFilepath = string.Format("{0}{1}{2}_{3}.csv", rootFolder, botAgentToBeProcessedFileName, i, DateTime.Now.ToString("yyyyMMddTHHmmss"));
                    File.WriteAllLines(csvFilepath, list);
                    beginIndex = beginIndex + take;
                }
            }
            else
            {
                for (int i = 1; i <= noOfFiles; i++)
                {
                    List<string> list = new List<string>();

                    string headerArrItem = masterDataList[0];
                    list.Add(headerArrItem);
                    try
                    {
                        string row = masterDataList[i];
                        list.Add(row);

                    }
                    catch (ArgumentOutOfRangeException ex)
                    {


                    }

                    string csvFilepath = string.Format("{0}{1}{2}_{3}.csv", rootFolder, botAgentToBeProcessedFileName, i, DateTime.Now.ToString("yyyyMMddTHHmmss"));
                    File.WriteAllLines(csvFilepath, list);
                    beginIndex = beginIndex + take;
                }
            }

            

            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            #endregion

        }

        private void ResetBotsExecutionStatus()
        {
            string jsonFilePath = System.Configuration.ConfigurationManager.AppSettings["jsonFilePath"].ToString();
            string json = File.ReadAllText(jsonFilePath);
            dynamic jsonObj = JsonConvert.DeserializeObject(json);

            int numberOfBotAgents = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["NumberOfAgents"].ToString());

            for (int i = 1; i <=numberOfBotAgents; i++)
            {
                var jBotExecution = jsonObj.SelectToken("ExecutionStatusBot" + i);
                jBotExecution.Replace("Pending");
            }
            string updatedJsonString = jsonObj.ToString();
            File.WriteAllText(jsonFilePath, updatedJsonString);
        }
    }
}

