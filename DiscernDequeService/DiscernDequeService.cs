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


namespace DiscernDequeService
{
    public class Order
    {
        public string FIN;
        public string OrderDescription;
        public string OrderDate;
        public string CPT4;
        public string DischargeDate;
        public string Comments;
    }
    public partial class DiscernDequeService : ServiceBase
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
        public DiscernDequeService()
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

            string receiverEmail = System.Configuration.ConfigurationManager.AppSettings["Email.ReceiverEmail"] != null ? System.Configuration.ConfigurationManager.AppSettings["Email.ReceiverEmail"].ToString() : string.Empty; //"osman.alikhan@atos.net"; //ConfigurationManager.AppSettings["Email.SenderEmail"]; //
            string senderEmail = System.Configuration.ConfigurationManager.AppSettings["Email.SenderEmail"] != null ? System.Configuration.ConfigurationManager.AppSettings["Email.SenderEmail"].ToString() : string.Empty; //"osman.alikhan@atos.net"; //ConfigurationManager.AppSettings["Email.SenderEmail"]; //
            string senderPassword = System.Configuration.ConfigurationManager.AppSettings["Email.SenderPassword"] != null ? System.Configuration.ConfigurationManager.AppSettings["Email.SenderPassword"].ToString() : string.Empty; //ConfigurationManager.AppSettings["Email.Password"];//
            string host = System.Configuration.ConfigurationManager.AppSettings["Email.Host"] != null ? System.Configuration.ConfigurationManager.AppSettings["Email.Host"].ToString() : string.Empty;//"10.14.32.146";// ConfigurationManager.AppSettings["Email.Host"]; //"10.14.32.146";
            int port = System.Configuration.ConfigurationManager.AppSettings["Email.Port"] != null ? Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Email.Port"].ToString()) : 0;//25;//Convert.ToInt32(ConfigurationManager.AppSettings["Email.Port"]); //25;
            try
            {
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
                        DirectoryInfo dir = new DirectoryInfo(sourceFolder);
                        string date = DateTime.Now.ToString("MMddyyyy");
                        FileInfo[] ChargeCorrectionfiles = dir.GetFiles("Facility_Charge_Correction_Requests*", SearchOption.TopDirectoryOnly);
                        if (ChargeCorrectionfiles.Count() > 0)
                        {
                            foreach (FileInfo file in ChargeCorrectionfiles)
                            {
                                if (file.Name.Contains(date))
                                {
                                    System.IO.File.Move(file.FullName, string.Format(@"{0}\{1}", rootFolder, file.Name));
                                }
                            }
                            
                        }
                        Global.WriteLog("END - Moving files from source SFTP folder to root Folder ");
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("An error occured while moving files. Error:" + ex.Message);
                    }
                }
                else
                {
                    if (MinutesPassed == VerificationInterval1 || MinutesPassed == VerificationInterval2 || MinutesPassed == VerificationInterval3 || MinutesPassed == VerificationInterval4 || MinutesPassed == VerificationInterval5 || MinutesPassed == VerificationInterval6)
                    {
                        //Global.WriteLog(timeOfStart.ToString());
                        //Global.WriteLog(MinutesPassed.ToString());
                        //Global.WriteLog(timeDiff.ToString());
                        //Global.WriteLog(regTimeDiff.ToString());

                        string rootFolderP = System.Configuration.ConfigurationManager.AppSettings["RootFolderPath"].ToString();
                        string chargeCorrectionfilepath = string.Empty;
                        DirectoryInfo directory = new DirectoryInfo(rootFolderP);
                        FileInfo[] chargeCorrectionFilesArr = directory.GetFiles("Facility_Charge_Correction_Requests*", SearchOption.TopDirectoryOnly);
                        if (chargeCorrectionFilesArr.Count() > 0)
                        {
                            chargeCorrectionfilepath = chargeCorrectionFilesArr[0].FullName;

                            var toBeProcessedRecords = File.ReadLines(chargeCorrectionfilepath).ToList();

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
                                ResetBotsExecutionStatus();
                                try
                                {
                                    Global.WriteLog("START - Splitting the files");
                                    SplitFacilityChargeFile();
                                    Global.WriteLog("END - Splitting the files");
                                }
                                catch (Exception ex)
                                {
                                    throw new Exception("An error occured while splitting files. Error:" + ex.Message);
                                }
                                try
                                {
                                    Global.WriteLog("START - Sending SOP request");
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

                                        var reqHeaderData = new Requestheaderdata()
                                        {
                                            botprocess = "Deque_Process",
                                            timeoutseconds = "1000",
                                            appcode = "DQ001",
                                            userid = "sbsuperadmin",
                                            stepid = "",
                                            correlationid = "",
                                            userrole = ""
                                        };
                                        var sopexecjson = new Sopexecjson()
                                        {
                                            sopName = i + "_DequeProcess_SOP",
                                            requestHeaderData = reqHeaderData,
                                            requestRulesData = new Requestrulesdata() { maxage = "30" },
                                            execSOPParamsData = new Execsopparamsdata() { }
                                        };


                                        var reqbody = new Rootobject() { sopExecJSON = sopexecjson };
                                        var requestData = JsonConvert.SerializeObject(reqbody);

                                        var bytes = Encoding.ASCII.GetBytes(requestData);

                                        httpWebRequest.ContentLength = bytes.Length;

                                        using (var outputStream = httpWebRequest.GetRequestStream())
                                        {
                                            outputStream.Write(bytes, 0, bytes.Length);
                                        }

                                        HttpWebResponse resp = httpWebRequest.GetResponse() as HttpWebResponse;
                                        Global.WriteLog(string.Format("SOP Request for Agent # {0} sent.", i));
                                        Global.WriteLog(string.Format("Response from portal on Account Verification {0}", DateTime.Now.ToString("dd MMMM yyyy hh:mm:ss")));
                                        httpWebRequest.Abort();

                                    }
                                    Global.WriteLog("END - Sending SOP request");
                                }
                                catch (Exception ex)
                                {
                                    throw new Exception("An error occured while sending the SOP request. Error:" + ex.Message);
                                }

                                bool isAllAgentsExecutionCompleted = false;
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
                                    }
                                    else if (agentsExecutionStatus == "Pending")
                                    {
                                        Global.WriteLog("One or more agents not started yet.");
                                        Thread.Sleep(60000);
                                    }
                                }

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
                                }
                            }

                        }
                        else
                        {
                            throw new Exception("File not found in the source directory");
                        }
                    }
                   
                }
            }
            catch (Exception ex)
            {

                Global.WriteLog("Discern deque run failed. Error: " + ex.Message + " " + "Inner Exception:" + ex.InnerException);
                try
                {
                    EmailHelper email = new EmailHelper(senderEmail, senderPassword, host, port);
                    string automationTeamEmail = ConfigurationManager.AppSettings["Email.AutomationTeam"];
                    string messagebody = "Bot Execution failed. Error: " + ex.Message + " " + "Inner Exception:" + ex.InnerException;
                    email.SendEMail(automationTeamEmail, "ALERT : Discern deque run Failed!! ", messagebody, "");

                }
                catch (Exception exp)
                {
                    Global.WriteLog("An error occured while sending email." + "Error: " + exp.Message);
                }
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

            string jsonFilePath = _configuration.AppSettings.Settings["BotResponseJson.FilePath"] != null ? _configuration.AppSettings.Settings["BotResponseJson.FilePath"].Value : string.Empty; //@"E:\Syntbots\BotResponse.json";

            string botAgentFileName = _configuration.AppSettings.Settings["BotAgentFacilityChargeFileName"] != null ? _configuration.AppSettings.Settings["BotAgentFacilityChargeFileName"].Value : string.Empty;

            try
            {
                Global.WriteLog("Starting merging process");

                DirectoryInfo dir = new DirectoryInfo(rootFolder); //"D:\\TaskQueue\\"
                FileInfo[] botOutPutFiles = dir.GetFiles(string.Format("*{0}*", botAgentFileName), SearchOption.TopDirectoryOnly);
                
                int totalRecords = 0;
                int totalDeque = 0;
                int totalRevCycleErrorRecords = 0;
                int totalAccessHIMErrorRecords = 0;
                int totalBadRecords = 0;
                string FINsWithAccessHIMError = "";
                string FINsWithRevCycleError = "";

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
                                ch.FIN = split[0].Replace("\"", ""); ;
                                ch.Comments = split[8].Replace("\"", "");
                                return ch;
                            }).ToList<Order>();

                            totalRecords = totalRecords + FINlist.Count();

                            foreach (var Encounter in FINlist)
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

                                        if (botStatus.Contains("Dequeued"))
                                        {
                                            totalDeque++;
                                        }
                                        else if (botStatus == "Failed to process - AccessHIM")
                                        {
                                            totalAccessHIMErrorRecords++;
                                            FINsWithAccessHIMError = FINsWithAccessHIMError + "," + FIN;
                                        }
                                        else if (botStatus == "Failed to process - RevCycle")
                                        {
                                            totalRevCycleErrorRecords++;
                                            FINsWithRevCycleError = FINsWithRevCycleError + "," + FIN;
                                        }
                                        else
                                        {
                                            totalBadRecords++;
                                        }
                                    }
                                }
                            }
                            Global.WriteLog("Completed the Processing of the Bot output file: " + botOutPutFiles[i].FullName);
                            System.IO.File.Move(botOutPutFiles[i].FullName, string.Format(@"{0}\{1}", archivedFolder, botOutPutFiles[i].Name));
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

                    sbMessageBody.Append("<tr><td>a/c Dequed Successfully</td>");
                    sbMessageBody.Append(string.Format("<td> {0} </td></tr>", totalDeque));

                    sbMessageBody.Append("<tr><td>Exceptions a/c (require Manual Intervention)</td>");
                    sbMessageBody.Append(string.Format("<td> {0} </td></tr>", totalAccessHIMErrorRecords));

                    sbMessageBody.Append("<tr><td>Exceptions a/c (Will be attempted in next Cycle)</td>");
                    sbMessageBody.Append(string.Format("<td> {0} </td></tr>", totalRevCycleErrorRecords));

                    sbMessageBody.Append("<tr><td>Total UnProcessed a/c</td>");
                    sbMessageBody.Append(string.Format("<td> {0} </td></tr>", totalBadRecords));

                    FINsWithAccessHIMError = FINsWithAccessHIMError.StartsWith(",") ? FINsWithAccessHIMError.Remove(0, 1) : FINsWithAccessHIMError;
                    FINsWithRevCycleError = FINsWithRevCycleError.StartsWith(",") ? FINsWithRevCycleError.Remove(0, 1) : FINsWithRevCycleError;
                    sbMessageBody.Append("</table>");

                    sbMessageBody.Append("<br><b>Accounts for manual interventions: </b>" + FINsWithAccessHIMError);
                    sbMessageBody.Append("<br><b>Accounts to be attempted in next Cycle: </b>" + FINsWithRevCycleError);
                    sbMessageBody.Append("<br><br>Note: Manual processing required for ‘Exceptions accounts’. <br>Should you have any queries, please feel free to contact Syntbots team.");
                    sbMessageBody.Append("<br><br>Thank you");
                    sbMessageBody.Append("<br>BOT Signature");
                    sbMessageBody.Append("<br>(Need to mention the contact details)");
                    sbMessageBody.Append("<br><br><br>*****This is an auto-generated email, please do not reply. ******");

                    string subject = string.Empty;
                    subject = string.Format("Deque Process - Summary & Inventory Details {0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm"));//ConfigurationManager.AppSettings["Email.Subject"];
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
            }
            catch (Exception ex)
            {
                Global.WriteLog("An Error occured during the merging process. Error message: " + ex.Message + "\n Inner exception: " + ex.InnerException);
            }
            finally
            {

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

        public void SplitFacilityChargeFile()
        {
            string rootFolder = System.Configuration.ConfigurationManager.AppSettings["RootFolderPath"].ToString();
            string NumberOfAgents = System.Configuration.ConfigurationManager.AppSettings["NumberOfAgents"];
            int noOfFiles = Convert.ToInt32(NumberOfAgents);

            string botAgentFacilityChargeFileName = System.Configuration.ConfigurationManager.AppSettings["BotAgentFacilityChargeFileName"];
            #region Split Facility Charge file
            object misValue = System.Reflection.Missing.Value;

            string facilityCCfilepath = string.Empty;

            DirectoryInfo dir = new DirectoryInfo(rootFolder);
            FileInfo[] facilityCCfiles = dir.GetFiles("*facility_charge_correction*", SearchOption.TopDirectoryOnly);
            if (facilityCCfiles.Count() > 0)
            {
                facilityCCfilepath = facilityCCfiles[0].FullName;
            }

            var masterDataList = File.ReadLines(facilityCCfilepath).ToList();
            System.Text.RegularExpressions.Regex CSVParser = new System.Text.RegularExpressions.Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");

            var masterDataFins = masterDataList.Select(line =>
            {
                String[] split = CSVParser.Split(line);
                string fin = split[6].Replace("\"", "");
                return fin;
            }).ToList<string>();

            //int finsCount = masterDataList.Count() - 1;
            int facilityChargeRowCount = masterDataList.Count() - 1;     // Count excludes the header row

            int dataRowsCount = facilityChargeRowCount;
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
                    string csvFilepath = string.Format("{0}{1}{2}_{3}.csv", rootFolder, botAgentFacilityChargeFileName, i, DateTime.Now.ToString("yyyyMMddTHHmmss"));
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

                    string csvFilepath = string.Format("{0}{1}{2}_{3}.csv", rootFolder, botAgentFacilityChargeFileName, i, DateTime.Now.ToString("yyyyMMddTHHmmss"));
                    File.WriteAllLines(csvFilepath, list);
                    beginIndex = beginIndex + take;
                }
            }
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
