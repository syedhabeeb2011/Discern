/*
 * Created by SharpDevelop.
 * User: v.chakradhar.ark
 * Date: 9/11/2019
 * Time: 3:01 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using SyntBotsDesktopBase.Utilities;
using SyntBotsDesktopBase.BaseScript;
using TestStack.White;
using System.Reflection;
using SyntBotsUIUtility;
using System.Threading;
using System.Linq;
using System.Configuration;
using System.IO;
using Newtonsoft.Json.Linq;

namespace SyntBotsDesktopImplementation.RPAScripts
{

    /// <summary>
    /// Description of SyntBotsAdapter. 
    /// </summary>

    public partial class SyntBotsAdapter : BaseScripts
    {
        public SyntBotsAdapter()
        {
            //
            // Add constructor logic here
            //
        }

        private TestStack.White.UIItems.WindowItems.Window _mainWindow;
        public override void AutomateScript()
        {
            JsonHelper objJsnHlpr = JsonHelper.Instance;
            var _launcher = new Atos.SyntBots.AccessHIM.Launcher();
            Atos.SyntBots.Common.Logger.Log(this.GetType().Assembly.Location);
            Configuration _configuration = ConfigurationManager.OpenExeConfiguration(this.GetType().Assembly.Location);
            string kk = _configuration.AppSettings.Settings["AccessHIM.MasterExcel.Path"] != null ? _configuration.AppSettings.Settings["AccessHIM.MasterExcel.Path"].Value : string.Empty;

            string AccessHIMApplicationPath = _configuration.AppSettings.Settings["AccessHIM.ApplicationPath"] != null ? _configuration.AppSettings.Settings["AccessHIM.ApplicationPath"].Value : string.Empty;
            string AccessHIMUsername = _configuration.AppSettings.Settings["AccessHIM.Username"] != null ? _configuration.AppSettings.Settings["AccessHIM.Username"].Value : string.Empty;
            string AccessHIMPassword = _configuration.AppSettings.Settings["AccessHIM.Password"] != null ? _configuration.AppSettings.Settings["AccessHIM.Password"].Value : string.Empty;

            string NumberOfAgents = _configuration.AppSettings.Settings["NumberOfAgents"] != null ? _configuration.AppSettings.Settings["NumberOfAgents"].Value : string.Empty;
            int noOfFiles = Convert.ToInt32(NumberOfAgents);    // input this value from settings



            string jsonFilePath = _configuration.AppSettings.Settings["BotResponseJson.FilePath"] != null ? _configuration.AppSettings.Settings["BotResponseJson.FilePath"].Value : string.Empty; //@"E:\Syntbots\BotResponse.json";
            try
            {


                Atos.SyntBots.Common.Logger.Log(jsonFilePath);
                string jsonString = File.ReadAllText(jsonFilePath);
                JObject jObject = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonString) as JObject;
                JToken jBotTypeToStart = jObject.SelectToken("BotTypeToStart");
                jBotTypeToStart.Replace("Account Categorization");

                JToken jExecution = jObject.SelectToken("Execution");
                jExecution.Replace("Started");


                string updatedJsonString = jObject.ToString();
                File.WriteAllText(jsonFilePath, updatedJsonString);


                Atos.SyntBots.Common.Logger.Log("Starting the Account Categorization process");
                string rootFolder = _configuration.AppSettings.Settings["RootFolderPath"] != null ? _configuration.AppSettings.Settings["RootFolderPath"].Value : string.Empty; //@"E:\TaskQueue\";//ConfigurationManager.AppSettings["AccessHIM.MasterExcel.Path"];
                string archivedFolder = _configuration.AppSettings.Settings["ArchieveFolderPath"] != null ? _configuration.AppSettings.Settings["ArchieveFolderPath"].Value : string.Empty; //@"E:\TaskQueue\";//ConfigurationManager.AppSettings["AccessHIM.MasterExcel.Path"];//@"E:\TaskQueue\Archived\";
                string fname = _configuration.AppSettings.Settings["AccessHIM.MasterExcel.FileName"] != null ? _configuration.AppSettings.Settings["AccessHIM.MasterExcel.FileName"].Value : string.Empty;
                string filename = fname + "_" + DateTime.Now.ToString("yyyyMMddTHHmmss");//"MasterExcel132181238700686548";// //ConfigurationManager.AppSettings["AccessHIM.MasterExcel.FileName"] +  System.DateTime.Now.ToFileTime();
                string botAgentToBeProcessedFileName = _configuration.AppSettings.Settings["BotAgentToBeProcessedFileName"] != null ? _configuration.AppSettings.Settings["BotAgentToBeProcessedFileName"].Value : string.Empty; //@"E:\TaskQueue\";//ConfigurationManager.AppSettings["AccessHIM.MasterExcel.Path"];//@"E:\TaskQueue\Archived\";

                string fillePath = rootFolder + filename;

                #region Move the exiting files to archive folder
                try
                {
                    if (Directory.Exists(rootFolder))
                    {
                        foreach (var file in new DirectoryInfo(rootFolder).GetFiles())
                        {
                            file.MoveTo(string.Format(@"{0}\{1}", archivedFolder, file.Name));
                        }
                    }
                }
                catch (Exception ex)
                {

                    Atos.SyntBots.Common.Logger.Log("An error occured while archieving existing files. Error: " + ex.Message);
                    throw new Exception("An error occured while archieving existing files. Error: " + ex.Message);
                }


                #endregion

                SyntBotsUIUtil driver = new SyntBotsUIUtil();

                Atos.SyntBots.Common.CodingQueueAssignment codingQueueAssignment = new Atos.SyntBots.Common.CodingQueueAssignment();


                Microsoft.Office.Interop.Excel.Application oXL;
                Microsoft.Office.Interop.Excel.Workbook mWorkBook;
                Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
                Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
                object misValue = System.Reflection.Missing.Value;


                #region Account Categoriztion
                Atos.SyntBots.Common.Logger.Log("Generating the master excel");

                try
                {

                    _launcher.LaunchAccessHIM(AccessHIMApplicationPath, AccessHIMUsername, AccessHIMPassword);
                    System.Threading.Thread.Sleep(7000);

                    var _generateMasterExcel = new Atos.SyntBots.AccessHIM.GenerateMasterExcel();
                    _generateMasterExcel.Generate(fillePath);
                    Atos.SyntBots.Common.Logger.Log("Generated the master excel file " + fillePath);

                }
                catch (Exception ex)
                {
                    Atos.SyntBots.Common.Logger.Log("An error occured while generating the master excel");
                    throw ex;
                }


                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;
                oXL.DisplayAlerts = false;
                var _workbooks = oXL.Workbooks;
                _workbooks.OpenText(fillePath + ".csv",
                                DataType: Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited,
                                TextQualifier: Microsoft.Office.Interop.Excel.XlTextQualifier.xlTextQualifierNone,
                                ConsecutiveDelimiter: true,
                                Semicolon: true);
                // Convert To Excle 97 / 2003
                _workbooks[1].SaveAs(fillePath + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel5);
                _workbooks.Close();

                try
                {
                    Atos.SyntBots.Common.Logger.Log("Starting Account Categorization");
                    mWorkBook = oXL.Workbooks.Open(fillePath + ".xls", 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //Get all the sheets in the workbook
                    mWorkSheets = mWorkBook.Worksheets;
                    //Get the allready exists sheet
                    mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item(1);
                    Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
                    int colCount = range.Columns.Count;
                    int rowCount = range.Rows.Count;

                    var _search = new Atos.SyntBots.AccessHIM.Search();
                    int totalRows = 0;
                    int toBeAllocatedRows = 0;
                    int toBeProcessedRows = 0;
                    for (int index = 1; index < rowCount; index++)
                    {
                        string a = Convert.ToString(mWSheet1.Cells[index, 1].Value2);

                        if (a == "Task Type")
                        {
                            //Delete the Search creteria from the header
                            Microsoft.Office.Interop.Excel.Range ran = mWSheet1.Range[mWSheet1.Cells[1, 1], mWSheet1.Cells[index - 1, 1]].EntireRow;
                            ran.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);

                            #region convert the MRN and FIN in number format
                            Microsoft.Office.Interop.Excel.Range rangeMRN = (Microsoft.Office.Interop.Excel.Range)mWSheet1.Cells[1, 6];
                            rangeMRN.EntireColumn.NumberFormat = "#####";
                            Microsoft.Office.Interop.Excel.Range rFIN = (Microsoft.Office.Interop.Excel.Range)mWSheet1.Cells[1, 7];
                            rFIN.EntireColumn.NumberFormat = "#####";
                            #endregion convert the MRN and FIN in number format

                            //Add the new header column Status 
                            mWSheet1.Cells[1, colCount + 1] = "Comments";

                            _mainWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("AccessHIM"));
                            _mainWindow.SetForeground();

                            #region Create ToBeAllocated and ToBeProcessed excel files

                            Microsoft.Office.Interop.Excel.Workbook toBeProcessedWorkBook;
                            Microsoft.Office.Interop.Excel.Worksheet toBeProcessedWorkSheet;
                            Microsoft.Office.Interop.Excel.Workbook toBeAllocatedWorkBook;
                            Microsoft.Office.Interop.Excel.Worksheet toBeAllocatedWorkSheet;


                            toBeProcessedWorkBook = oXL.Workbooks.Add(misValue);
                            toBeProcessedWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)toBeProcessedWorkBook.Worksheets.get_Item(1);

                            toBeAllocatedWorkBook = oXL.Workbooks.Add(misValue);
                            toBeAllocatedWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)toBeAllocatedWorkBook.Worksheets.get_Item(1);

                            Microsoft.Office.Interop.Excel.Range rMasterExcel = (Microsoft.Office.Interop.Excel.Range)mWSheet1.Cells[1, colCount].EntireRow;
                            rMasterExcel.Copy(Type.Missing);

                            Microsoft.Office.Interop.Excel.Range rToBeProcessed2 = (Microsoft.Office.Interop.Excel.Range)toBeProcessedWorkSheet.Cells[1, colCount].EntireRow;
                            rToBeProcessed2.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll,
                            Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                            Microsoft.Office.Interop.Excel.Range rToBeAllocated2 = (Microsoft.Office.Interop.Excel.Range)toBeAllocatedWorkSheet.Cells[1, colCount].EntireRow;
                            rToBeAllocated2.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll,
                                Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                            #endregion

                            #region Iterate through the rows 
                            int count1 = 0;
                            int tobeProcessCurrentRow = 2;
                            int tobeAllocatedCurrentRow = 2;
                            string FIN = string.Empty;
                            for (int i = 2; i <= rowCount; i++)
                            {
                                try
                                {
                                    totalRows++;
                                    //                    	if(i==3)
                                    //                    	{
                                    //                    		break;
                                    //                    	}
                                    //                    	var fins = new List<string>();
                                    //                    	fins.Add("71000000629280");//lab 
                                    //                    	fins.Add("71000000630762");
                                    //                    	fins.Add("71000000630514");//lab
                                    //						fins.Add("71000000630747");

                                    FIN = Convert.ToString(mWSheet1.Cells[i, 7].Value2); //fins[count1]; //

                                    count1++;

                                    #region Verify the Lab by checking task type 'Coding OP - Lab’ in the master excel

                                    string taskType = Convert.ToString(mWSheet1.Cells[i, 1].Value2);
                                    if (string.IsNullOrEmpty(taskType))
                                    {
                                        totalRows = totalRows - 1;
                                        break;
                                    }

                                    Microsoft.Office.Interop.Excel.Range rangeMasterExcel = (Microsoft.Office.Interop.Excel.Range)mWSheet1.Cells[i, colCount].EntireRow;
                                    rangeMasterExcel.Copy(Type.Missing);
                                    if (taskType == "Coding OP - Lab")
                                    {
                                        Microsoft.Office.Interop.Excel.Range rangeToBeProcessed = (Microsoft.Office.Interop.Excel.Range)toBeProcessedWorkSheet.Cells[tobeProcessCurrentRow, colCount].EntireRow;
                                        rangeToBeProcessed.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll,
                                        Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                                        mWSheet1.Cells[i, colCount + 1] = "To be Processed";
                                        toBeProcessedRows++;
                                        tobeProcessCurrentRow++;
                                    }
                                    else
                                    {
                                        Microsoft.Office.Interop.Excel.Range rangeToBeAllocated = (Microsoft.Office.Interop.Excel.Range)toBeAllocatedWorkSheet.Cells[tobeAllocatedCurrentRow, colCount].EntireRow;
                                        rangeToBeAllocated.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll,
                                        Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                                        mWSheet1.Cells[i, colCount + 1] = "To be Allocated";
                                        toBeAllocatedRows++;
                                        tobeAllocatedCurrentRow++;
                                    }


                                    #endregion

                                }
                                catch (Exception ex)
                                {

                                    Atos.SyntBots.Common.Logger.Log("An error occurred while while performing account categorization for the FIN:" + FIN + ". Error:  " + ex.InnerException); ;
                                }


                            }
                            #endregion

                            //var line = Environment.NewLine + Environment.NewLine;
                            Atos.SyntBots.Common.Logger.Log("Account Categorization completed");
                            Atos.SyntBots.Common.Logger.Log("Total records in Master excel: " + totalRows + ", 'ToBeAllocated' records : " + toBeAllocatedRows + ", ToBeProcessed records : " + toBeProcessedRows);

                            //Save the ToBeAllocated file
                            toBeAllocatedWorkBook.SaveAs(fillePath + "_ToBeAllocated" + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                            Missing.Value, Missing.Value, Missing.Value,
                            Missing.Value, Missing.Value);
                            toBeAllocatedWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
                            toBeAllocatedWorkSheet = null;
                            toBeAllocatedWorkBook = null;

                            //Save the ToBeProcessed file
                            toBeProcessedWorkBook.SaveAs(fillePath + "_ToBeProcessed" + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                            Missing.Value, Missing.Value, Missing.Value,
                            Missing.Value, Missing.Value);



                            toBeProcessedWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
                            toBeProcessedWorkSheet = null;
                            toBeProcessedWorkBook = null;



                            mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
                            mWSheet1 = null;
                            mWorkBook = null;

                            oXL.Quit();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();


                            //                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mWSheet1);
                            //                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mWorkBook);
                            //                    
                            //                    System.Runtime.InteropServices.Marshal.ReleaseComObject(toBeAllocatedWorkSheet);
                            //                    System.Runtime.InteropServices.Marshal.ReleaseComObject(toBeAllocatedWorkBook);
                            //                    
                            //                    System.Runtime.InteropServices.Marshal.ReleaseComObject(toBeProcessedWorkSheet);
                            //                    System.Runtime.InteropServices.Marshal.ReleaseComObject(toBeProcessedWorkBook);





                            //                     System.Text.StringBuilder sbMessageBody = new System.Text.StringBuilder();
                            //                     sbMessageBody.Append("Hi Team, <br>Please find attached details of FIN# which need to be allocated to SME’s for coding.");
                            //                     sbMessageBody.Append("<br><br>Should you have any queries, please feel free to contact Syntbots team.");
                            //                     sbMessageBody.Append("<br><br>Thank you");
                            //                     sbMessageBody.Append("<br>BOT Signature");
                            //                     sbMessageBody.Append("<br>(Need to mention the contact details)");
                            //                     sbMessageBody.Append("<br><br><br>*****This is an auto-generated email, please do not reply. ******");


                            string subject = _configuration.AppSettings.Settings["Email.ToBeAllicated.Subject"] != null ? _configuration.AppSettings.Settings["Email.ToBeAllicated.Subject"].Value + " " + DateTime.UtcNow.AddHours(-5).ToString("yyyy-MM-dd HH:mm") : string.Empty;
                            string messageBody = _configuration.AppSettings.Settings["Email.ToBeAllicated.MessageBody"] != null ? _configuration.AppSettings.Settings["Email.ToBeAllicated.MessageBody"].Value : string.Empty + " " + DateTime.Now.ToString("yyyy-MM-dd HH:mm");//ConfigurationManager.AppSettings["Email.Subject"]; //sbMessageBody.ToString(); 
                            string attachmentfilePath = fillePath + "_ToBeAllocated" + ".xls";

                            try
                            {

                                string senderEmail = _configuration.AppSettings.Settings["Email.SenderEmail"] != null ? _configuration.AppSettings.Settings["Email.SenderEmail"].Value : string.Empty; //"osman.alikhan@atos.net"; //ConfigurationManager.AppSettings["Email.SenderEmail"]; //
                                string senderPassword = _configuration.AppSettings.Settings["Email.SenderPassword"] != null ? _configuration.AppSettings.Settings["Email.SenderPassword"].Value : string.Empty; //ConfigurationManager.AppSettings["Email.Password"];//
                                string host = _configuration.AppSettings.Settings["Email.Host"] != null ? _configuration.AppSettings.Settings["Email.Host"].Value : string.Empty;//"10.14.32.146";// ConfigurationManager.AppSettings["Email.Host"]; //"10.14.32.146";
                                int port = _configuration.AppSettings.Settings["Email.Port"] != null ? Convert.ToInt32(_configuration.AppSettings.Settings["Email.Port"].Value) : 0;//25;//Convert.ToInt32(ConfigurationManager.AppSettings["Email.Port"]); //25;

                                Atos.SyntBots.Common.EmailHelper email = new Atos.SyntBots.Common.EmailHelper(senderEmail, senderPassword, host, port);
                                string receiverEmail = _configuration.AppSettings.Settings["Email.ReceiverEmail"] != null ? _configuration.AppSettings.Settings["Email.ReceiverEmail"].Value : string.Empty; //ConfigurationManager.AppSettings["Email.ReceiverEmail"];
                                Atos.SyntBots.Common.Logger.Log("Sending email to " + receiverEmail);
                                email.SendEMail(receiverEmail, subject, messageBody, attachmentfilePath);
                                Atos.SyntBots.Common.Logger.Log("Email sent with the attachment to be allocated file " + fillePath + "_ToBeAllocated" + ".xls");

                            }
                            catch (Exception ex)
                            {

                                Atos.SyntBots.Common.Logger.Log("An error occured while sending email." + "Error: " + ex.Message);
                            }

                            _launcher.CloseAccessHIM();
                            oXL.Quit();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();
                            break;

                        }
                    }
                }
                catch (Exception ex)
                {
                    oXL.Quit();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    _launcher.CloseAccessHIM();
                    Atos.SyntBots.Common.Logger.Log("An error occured while performing Account Categorization. Error: " + ex.Message);
                    throw new Exception("An error occured while performing Account Categorization. Error: " + ex.Message);
                }

                #endregion







                jsonString = File.ReadAllText(jsonFilePath);
                jObject = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonString) as JObject;
                jBotTypeToStart = jObject.SelectToken("BotTypeToStart");
                jBotTypeToStart.Replace("Account Verification");

                jExecution = jObject.SelectToken("Execution");
                jExecution.Replace("Completed");


                updatedJsonString = jObject.ToString();
                File.WriteAllText(jsonFilePath, updatedJsonString);

                objJsnHlpr.SetOutCome("Success");
            }
            catch (Exception ex)
            {
                Atos.SyntBots.Common.Logger.Log("Error message: " + ex.Message + "\n Inner exception: " + ex.InnerException);
                objJsnHlpr.SetErrorList(ex.Message.ToString(), "Technical", "006");
                objJsnHlpr.SetOutCome("Failure");

                var _accessHIM = Desktop.Instance.Windows().Find(K => K.Name.Contains("AccessHIM"));
                if (_accessHIM != null)
                {
                    _accessHIM.Close();
                }

                var _cernerMillennium = Desktop.Instance.Windows().Find(K => K.Name == "Millennium Logon");
                if (_cernerMillennium != null)
                {
                    _cernerMillennium.Close();
                }

            }
            finally
            {
                string jsonString = File.ReadAllText(jsonFilePath);
                var jObject = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonString) as JObject;
                var jBotTypeToStart = jObject.SelectToken("BotTypeToStart");
                jBotTypeToStart.Replace("Account Verification");

                var jExecution = jObject.SelectToken("Execution");
                jExecution.Replace("Completed");

                string updatedJsonString = jObject.ToString();
                File.WriteAllText(jsonFilePath, updatedJsonString);

                var _accessHIM = Desktop.Instance.Windows().Find(K => K.Name.Contains("AccessHIM"));
                if (_accessHIM != null)
                {
                    _accessHIM.Close();
                }
                //base.DisposeMainWindow();
                //base.DisposeApplication();


                System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("Excel");
                foreach (System.Diagnostics.Process p in processes)
                {

                    try
                    {
                        p.Kill();
                    }
                    catch { }

                }
                objJsnHlpr.PopulateJsonData();

            }
        }
    }
}