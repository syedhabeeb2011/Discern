/*
 * Created by SharpDevelop.
 * User: v.chakradhar.ark
 * Date: 9/11/2019
 * Time: 3:01 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Windows.Automation;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using SyntBotsDesktopBase.Utilities;
using SyntBotsDesktopBase.BaseScript;
using TestStack.White;
using System.Reflection;
using System.Runtime.InteropServices;
using SyntBotsUIUtility;
using System.Diagnostics;
using TestStack.White.InputDevices;
using TestStack.White.UIItems.Finders;
using System.Threading;
using System.Net;
using System.Net.Mail;
using System.Linq;
using System.Configuration;
using TestStack.White.WindowsAPI;
using System.IO;
using Newtonsoft.Json.Linq;
using Atos.SyntBots.Common.Entity;
using System.Text.RegularExpressions;
using TestStack.White.UIItems;
using System.Data;
using System.Data.OleDb;
using System.Globalization;

namespace SyntBotsDesktopImplementation.RPAScripts
{

    /// <summary>
    /// Description of SyntBotsAdapter. 
    /// </summary>

    public partial class SyntBotsAdapter : BaseScripts
    {
        SyntBotsUIUtil driver;

        public SyntBotsAdapter()
        {
            //
            // Add constructor logic here
            //
        }

        public void CloseExistingApplications()
        {


            var _3MWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("360 Encompass:"));
            if (_3MWindow != null)
            {
                _3MWindow.SetForeground();
                Thread.Sleep(500);
                _3MWindow.Close();
                Thread.Sleep(500);
                Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);

                var _3MError = Desktop.Instance.Windows().Find(K => K.Name.Contains("3M Coding and Reimbursement System"));
                if (_3MError != null)
                {
                    _3MError.SetForeground();
                    Thread.Sleep(500);
                    _3MError.Close();
                }

                var _accessHIMwindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("AccessHIM"));
                if (_accessHIMwindow != null)
                {
                    _accessHIMwindow.SetForeground();
                    Thread.Sleep(1000);
                    _accessHIMwindow.Close();
                    Thread.Sleep(500);
                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
                }
            }

            var _3MErrorWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("3M Coding and Reimbursement System"));
            if (_3MErrorWindow != null)
            {
                _3MErrorWindow.SetForeground();
                Thread.Sleep(500);
                _3MErrorWindow.Close();

                var accessHIM = Desktop.Instance.Windows().Find(K => K.Name.Contains("AccessHIM"));
                if (accessHIM != null)
                {
                    accessHIM.SetForeground();
                    Thread.Sleep(500);
                    accessHIM.Close();
                    Thread.Sleep(500);
                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
                }
            }


            var _chargeViewer = Desktop.Instance.Windows().Find(x => x.Name == "Charge Viewer (Cannot Increase Quantity)");
            if (_chargeViewer != null)
            {
                _chargeViewer.SetForeground();
                Thread.Sleep(500);
                _chargeViewer.Close();
            }

            var _powerchart = Desktop.Instance.Windows().Find(K => K.Name.Contains("Opened by"));
            if (_powerchart != null)
            {
                _powerchart.SetForeground();
                Thread.Sleep(500);
                _powerchart.Close();
            }


            var _accessHIM = Desktop.Instance.Windows().Find(K => K.Name.Contains("AccessHIM"));
            if (_accessHIM != null)
            {
                _accessHIM.SetForeground();
                Thread.Sleep(500);
                _accessHIM.Close();
            }


            var _cernerMillennium = Desktop.Instance.Windows().Find(K => K.Name == "Millennium Logon");
            if (_cernerMillennium != null)
            {
                _cernerMillennium.SetForeground();
                Thread.Sleep(500);
                _cernerMillennium.Close();
            }

        }

        public override void AutomateScript()
        {

            JsonHelper objJsnHlpr = JsonHelper.Instance;
            var _launcher = new Atos.SyntBots.AccessHIM.Launcher();
            Configuration _configuration = ConfigurationManager.OpenExeConfiguration(this.GetType().Assembly.Location);
            string kk = _configuration.AppSettings.Settings["AccessHIM.MasterExcel.Path"] != null ? _configuration.AppSettings.Settings["AccessHIM.MasterExcel.Path"].Value : string.Empty;

            string AccessHIMApplicationPath = _configuration.AppSettings.Settings["AccessHIM.ApplicationPath"] != null ? _configuration.AppSettings.Settings["AccessHIM.ApplicationPath"].Value : string.Empty;
            string AccessHIMUsername = _configuration.AppSettings.Settings["AccessHIM.Username"] != null ? _configuration.AppSettings.Settings["AccessHIM.Username"].Value : string.Empty;
            string AccessHIMPassword = _configuration.AppSettings.Settings["AccessHIM.Password"] != null ? _configuration.AppSettings.Settings["AccessHIM.Password"].Value : string.Empty;

            string receiverEmail = _configuration.AppSettings.Settings["Email.ReceiverEmail"] != null ? _configuration.AppSettings.Settings["Email.ReceiverEmail"].Value : string.Empty; //"osman.alikhan@atos.net"; //ConfigurationManager.AppSettings["Email.SenderEmail"]; //
            string senderEmail = _configuration.AppSettings.Settings["Email.SenderEmail"] != null ? _configuration.AppSettings.Settings["Email.SenderEmail"].Value : string.Empty; //"osman.alikhan@atos.net"; //ConfigurationManager.AppSettings["Email.SenderEmail"]; //
            string senderPassword = _configuration.AppSettings.Settings["Email.SenderPassword"] != null ? _configuration.AppSettings.Settings["Email.SenderPassword"].Value : string.Empty; //ConfigurationManager.AppSettings["Email.Password"];//
            string host = _configuration.AppSettings.Settings["Email.Host"] != null ? _configuration.AppSettings.Settings["Email.Host"].Value : string.Empty;//"10.14.32.146";// ConfigurationManager.AppSettings["Email.Host"]; //"10.14.32.146";
            int port = _configuration.AppSettings.Settings["Email.Port"] != null ? Convert.ToInt32(_configuration.AppSettings.Settings["Email.Port"].Value) : 0;//25;//Convert.ToInt32(ConfigurationManager.AppSettings["Email.Port"]); //25;

            //Modified by Velan on 19/01/2022, for Flint and Lapeer Go Live
            //int dnRaiseTime = _configuration.AppSettings.Settings["DNRaiseTime"] != null ? Convert.ToInt32(_configuration.AppSettings.Settings["DNRaiseTime"].Value) : 0;//25;//Convert.ToInt32(ConfigurationManager.AppSettings["Email.Port"]); //25;
            var DNRaiseTimes = _configuration.AppSettings.Settings["DNRaiseTime"].Value.ToString().Split(',');

            string orderCatalogFilePath = _configuration.AppSettings.Settings["OrderCatalog.FilePath"] != null ? _configuration.AppSettings.Settings["OrderCatalog.FilePath"].Value : string.Empty;//@"E:\Syntbots\OrderCatalog.xlsx";//ConfigurationManager.AppSettings["OrderCatalog.FilePath"];
            string codingQueueAssignmentFilePath = _configuration.AppSettings.Settings["CodingQueueAssignment.FilePath"] != null ? _configuration.AppSettings.Settings["CodingQueueAssignment.FilePath"].Value : string.Empty;//@"E:\Syntbots\Cerner Coding Queue Assignment Search.xlsx";//ConfigurationManager.AppSettings[""];
            string rootFolder = _configuration.AppSettings.Settings["RootFolderPath"] != null ? _configuration.AppSettings.Settings["RootFolderPath"].Value : string.Empty; //@"E:\TaskQueue\";

            string jsonFilePath = _configuration.AppSettings.Settings["BotResponseJson.FilePath"] != null ? _configuration.AppSettings.Settings["BotResponseJson.FilePath"].Value : string.Empty; //@"E:\Syntbots\BotResponse.json";
            string botAgentToBeProcessedFileName = _configuration.AppSettings.Settings["BotAgentToBeProcessedFileName"] != null ? _configuration.AppSettings.Settings["BotAgentToBeProcessedFileName"].Value : string.Empty; //@"E:\TaskQueue\";
            string LogFilePath = _configuration.AppSettings.Settings["Logger.FilePath"] != null ? _configuration.AppSettings.Settings["Logger.FilePath"].Value : string.Empty; //@"E:\TaskQueue\";
            string botAgentName = _configuration.AppSettings.Settings["botAgentName"] != null ? _configuration.AppSettings.Settings["botAgentName"].Value : string.Empty;

            string appendDNCommentValue = _configuration.AppSettings.Settings["AppendDNComment"] != null ? _configuration.AppSettings.Settings["AppendDNComment"].Value : string.Empty; //@"E:\TaskQueue\";
            string debugModeValue = _configuration.AppSettings.Settings["DebugMode"] != null ? _configuration.AppSettings.Settings["DebugMode"].Value : string.Empty; //@"E:\TaskQueue\";
            string explodingChargesFileName = _configuration.AppSettings.Settings["ExplodingChargesFileName"] != null ? _configuration.AppSettings.Settings["ExplodingChargesFileName"].Value : string.Empty; //@"E:\TaskQueue\";
            string addonChargesFileName = _configuration.AppSettings.Settings["AddonChargesFileName"] != null ? _configuration.AppSettings.Settings["AddonChargesFileName"].Value : string.Empty; //@"E:\TaskQueue\";
            string _excludedCPTs = _configuration.AppSettings.Settings["ExcludedCPTs"] != null ? _configuration.AppSettings.Settings["ExcludedCPTs"].Value : string.Empty;
            var excludedCPTs = _excludedCPTs.Split(',');


            bool appendDNComment = appendDNCommentValue == "True" ? true : false;
            bool debugMode = debugModeValue == "True" ? true : false;
            bool isDaylightSave = TimeZoneInfo.Local.IsDaylightSavingTime(DateTime.Now);
            try
            {
                if (!IsFileLocked(jsonFilePath))
                {
                    string jsonString = File.ReadAllText(jsonFilePath);
                    JObject jObject = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonString) as JObject;
                    JToken jBotTypeToStart = jObject.SelectToken("BotTypeToStart");
                    jBotTypeToStart.Replace("Account Verification");

                    JToken jExecution = jObject.SelectToken("Execution");
                    jExecution.Replace("Started");


                    var jBotExecution = jObject.SelectToken("ExecutionStatusBot" + botAgentName);
                    jBotExecution.Replace("Started");

                    string updatedJsonString = jObject.ToString();
                    File.WriteAllText(jsonFilePath, updatedJsonString);

                }
                else
                {
                    throw new Exception("An error occurred while updating the json file.");
                }


                Atos.SyntBots.Common.Logger.Log("Starting the Account Verification process", LogFilePath);
                string catalogfilepath = string.Empty;
                driver = new SyntBotsUIUtil();
                //Atos.SyntBots.Common.ExcelReader excelReader = new Atos.SyntBots.Common.ExcelReader();
                Atos.SyntBots.Common.Logger.Log("Started the reading order catalog", LogFilePath);
                //Microsoft.Office.Interop.Excel.Application oXL;
                DirectoryInfo rdirInfo = new DirectoryInfo(orderCatalogFilePath);
                //oXL = new Microsoft.Office.Interop.Excel.Application();
                //oXL.Visible = false;
                //oXL.DisplayAlerts = false;
                //var _workbooks = oXL.Workbooks;
                // 30-12-2022 Modified by Syed for Domain Changes
                //FileInfo[] catalogfiles = rdirInfo.GetFiles(string.Format("*PROD_OrderCatalog{0}*", botAgentName), SearchOption.TopDirectoryOnly);
                FileInfo[] catalogfiles = rdirInfo.GetFiles(string.Format("*p2082_OrderCatalog{0}*", botAgentName), SearchOption.TopDirectoryOnly);
                if (catalogfiles.Count() > 0)
                {
                    catalogfilepath = catalogfiles[0].FullName;

                }
                //var dsOrderCatalog = excelReader.ReadData(orderCatalogFilePath);
                //var orderCatalogData = dsOrderCatalog.Tables[0];
                //var orderCatalogData = GetDataTableFromCsv(catalogfilepath, true);
                var orderCatalogData = ConvertCSVtoDataTable(catalogfilepath);
                Atos.SyntBots.Common.Logger.Log("competed the reading order catalog", LogFilePath);
                //Atos.SyntBots.Common.Logger.Log("Started the reading coding queue data", LogFilePath);
                //var dscodingQueueAssignment = excelReader.ReadData(codingQueueAssignmentFilePath);
                //var codingQueueAssignmentData = dscodingQueueAssignment.Tables[0];
                //Atos.SyntBots.Common.Logger.Log("completed the reading coding queue data", LogFilePath);
                Atos.SyntBots.Common.OrderCatalog orderCatalog = new Atos.SyntBots.Common.OrderCatalog();
                //Atos.SyntBots.Common.CodingQueueAssignment codingQueueAssignment = new Atos.SyntBots.Common.CodingQueueAssignment();

                #region Process the ToBeProcessed file


                string toBeProcessedfilepath = string.Empty;

                Atos.SyntBots.Common.Logger.Log(rootFolder.ToString(), LogFilePath);

                DirectoryInfo dirInfo = new DirectoryInfo(rootFolder);
                FileInfo[] toBeProcessedfiles = dirInfo.GetFiles(string.Format("*{0}*", botAgentToBeProcessedFileName, SearchOption.TopDirectoryOnly));

                if (toBeProcessedfiles.Count() > 0)
                {
                    toBeProcessedfilepath = toBeProcessedfiles[0].FullName;
                }
                FileInfo toBeProcessedInfo = new FileInfo(toBeProcessedfilepath);

                FileInfo[] ordersfiles = dirInfo.GetFiles(string.Format("*{0}*", "orders" + botAgentName, SearchOption.TopDirectoryOnly));

                string ordersFilePath = string.Empty;
                if (ordersfiles.Count() > 0)
                {
                    ordersFilePath = ordersfiles[0].FullName;
                }

                FileInfo[] chargesfiles = dirInfo.GetFiles(string.Format("*{0}*", "charges" + botAgentName, SearchOption.TopDirectoryOnly));

                string chargesFilePath = string.Empty;
                if (chargesfiles.Count() > 0)
                {
                    chargesFilePath = chargesfiles[0].FullName;
                }

                if (toBeProcessedfiles.Count() > 0)
                {
                    List<Order> FINlist = File.ReadLines(toBeProcessedfilepath).Skip(1).Select(line =>
                    {
                        Order ch = new Order();
                        Regex CSVParser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");
                        String[] split = CSVParser.Split(line);
                        ch.FIN = split[6].Replace("\"", ""); ;
                        ch.DischargeDate = split[8].Replace("\"", "");
                        return ch;
                    }).ToList<Atos.SyntBots.Common.Entity.Order>();

                    toBeProcessedInfo = new FileInfo(toBeProcessedfilepath);

                    int totalRecords = 0;
                    int totalGoodRecords = 0;
                    int totalDNRaised = 0;
                    int totalErrorRecords = 0;
                    int count2 = 0;
                    int tobeAllocatedCurrentRow = 0;
                    bool shouldDNBeRaised = false;
                    int totalDataErrorRecords = 0;

                    int MinutesPassed = (int)DateTime.Now.TimeOfDay.TotalMinutes;


                    //Modified by Velan on 19/01/2022, for Flint and Lapeer GoLive

                    //if (MinutesPassed > dnRaiseTime && MinutesPassed < (dnRaiseTime + 60))
                    //{
                    //    shouldDNBeRaised = true;
                    //}


                    foreach (var d in DNRaiseTimes)
                    {
                        var dnRaiseTime = Int32.Parse(d);
                        if (MinutesPassed > dnRaiseTime && MinutesPassed < (dnRaiseTime + 60))
                        {
                            shouldDNBeRaised = true;
                            break;
                        }
                    }




                    string FIN = string.Empty;
                    /*
                    System.Diagnostics.Process[] accessHIMProcesses = System.Diagnostics.Process.GetProcessesByName("AccessHIM");
                    foreach (System.Diagnostics.Process p in accessHIMProcesses)
                    {

                        try
                        {
                            p.Kill();
                        }
                        catch { }

                    }
                    Atos.SyntBots.Common.Logger.Log("START - Launching AccessHIM", LogFilePath);
                    _launcher.LaunchAccessHIM(AccessHIMApplicationPath, AccessHIMUsername, AccessHIMPassword);

                    var _accessHIMWindow = GetWindowWithWait("AccessHIM", 120);
                    _accessHIMWindow.SetForeground();
                    Atos.SyntBots.Common.Logger.Log("END - Launching AccessHIM", LogFilePath);
                    
                    System.Threading.Thread.Sleep(500);

                    #region Apply Filters Pending and OnHold
                    try
                    {

                        var queryButton = _accessHIMWindow.GetMultiple<TestStack.White.UIItems.Button>(SearchCriteria.ByText("Query"));
                        //System.Windows.Rect r = queryButton[0].AutomationElement.Current.BoundingRectangle;
                        //System.Windows.Point pp = r.TopRight;
                        //Atos.SyntBots.Common.Logger.Log(string.Format("Bounding rectangle X:{0} , Y:{1}", r.X, r.Y), LogFilePath);
                        //driver.MouseClick(pp);
                        driver.AEClick(queryButton[0].AutomationElement);
                        System.Threading.Thread.Sleep(1000);
                        //var popup = _accessHIMWindow.MdiChild(SearchCriteria.ByText("Preferences (Filtered)"));
                        var popup = GetMdiChildWithWait(_accessHIMWindow, SearchCriteria.ByText("Preferences (Filtered)"), 10);
                        popup.Focus();
                        Thread.Sleep(500);
                        var abcde = popup.GetMultiple<TestStack.White.UIItems.Hyperlink>(SearchCriteria.ByControlType(ControlType.Hyperlink).AndByText("None"));
                        abcde[0].Click();
                        abcde[1].Click();
                        var abc = popup.GetMultiple<TestStack.White.UIItems.ListView>(SearchCriteria.ByControlType(ControlType.DataGrid));

                        foreach (ListViewRow row in abc[0].Rows)
                        {
                            if ((String.Compare(row.Name, "OnHold", true) == 0) || (String.Compare(row.Name, "Pending", true) == 0))
                            {
                                //driver.MouseClick(listviewitem.Location);
                                //listviewitem.KeyIn(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.SPACE);
                                SelectionItemPattern pattern = (SelectionItemPattern)(BasePattern)row.AutomationElement.GetCurrentPattern(SelectionItemPattern.Pattern);
                                pattern.Select();
                                row.KeyIn(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.SPACE);
                            }
                        }
                        Thread.Sleep(500);
                        abc = popup.GetMultiple<TestStack.White.UIItems.ListView>(SearchCriteria.ByControlType(ControlType.DataGrid));
                        foreach (ListViewRow row in abc[1].Rows)
                        {
                            //listviewrow.Select();
                            if (String.Compare(row.Name, "Coding OP - LAB", true) == 0)
                            {
                                SelectionItemPattern pattern = (SelectionItemPattern)(BasePattern)row.AutomationElement.GetCurrentPattern(SelectionItemPattern.Pattern);
                                pattern.Select();
                                row.KeyIn(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.SPACE);
                                break;
                            }
                        }
                        System.Threading.Thread.Sleep(500);
                        AutomationElement AE5 = driver.GetElmtByWlker(popup.AutomationElement, "Apply", null, null, "button", null);
                        driver.AEClick(AE5);
                        //	
                        System.Threading.Thread.Sleep(500);
                        AutomationElement AE6 = driver.GetElmtByWlker(popup.AutomationElement, "OK", null, null, "button", null);
                        driver.AEClick(AE6);

                    }
                    catch (Exception ex)
                    {

                        throw new Exception("An error occure while applying filters. Error :" + ex.Message);
                    }

                    #endregion Apply Filters Pending and OnHold
                    */

                    foreach (var Encounter in FINlist)
                    {
                        bool isErrorOccurredAtDNLevel = false;
                        try
                        {

                            FIN = Encounter.FIN;
                            string strAdmitDate = DateTime.Parse(Encounter.DischargeDate).ToString("MM/dd/yyyy");

                            if (string.IsNullOrEmpty(FIN))
                            {
                                break;
                            }
                            Atos.SyntBots.Common.Logger.Log("***START - Processing fin:" + FIN, LogFilePath);

                            var listPowerChartOrders = new List<Atos.SyntBots.Common.Entity.Order>();
                            var chargeList = new List<Atos.SyntBots.Common.Entity.Charge>();
                            var explodingChargeList = new List<Atos.SyntBots.Common.Entity.Charge>();
                            int poowerchartOrdersCount = 0;
                            int chargeListCount = 0;
                            string locationTxt = string.Empty;
                            bool isZeroCharges = false;
                            bool isGoodRecord = false;
                            var sbChargeModification = new System.Text.StringBuilder();
                            var sbChargeDeletion = new System.Text.StringBuilder();
                            var sbMissingCharges = new System.Text.StringBuilder();


                            Atos.SyntBots.Common.Logger.Log("START - Get Orders from Powerchart", LogFilePath);
                            #region step-2 - Get lab orders from PowerChart
                            try
                            {
                                var powerChart_MasterData = File.ReadLines(ordersFilePath).Select(line =>
                                {
                                    var ch = new Atos.SyntBots.Common.Entity.Order();
                                    Regex CSVParser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");
                                    String[] split = CSVParser.Split(line);
                                    ch.FIN = split[0].Replace("\"", "");
                                    ch.OrderDescription = split[3].Replace("\"", "");
                                    ch.OrderDate = split[4].Replace("\"", "");
                                    ch.Comments = split[7].Replace("\"", "");
                                    ch.OrderStatus = split[5].Replace("\"", "");
                                    ch.ResultDate = split[9].Replace("\"", "");
                                    return ch;
                                }).ToList<Atos.SyntBots.Common.Entity.Order>();

                                var ordersByFin = powerChart_MasterData.Where(x => x.FIN.Contains(FIN)).ToList();

                                Atos.SyntBots.PowerChart.OrderGrid ordersGrid = new Atos.SyntBots.PowerChart.OrderGrid();

                                listPowerChartOrders = ordersGrid.GetLabOrders(ordersByFin, orderCatalogData, strAdmitDate, excludedCPTs);

                                Atos.SyntBots.Common.Logger.Log("Powerchart records : " + listPowerChartOrders.Count, LogFilePath);

                            }
                            catch (Exception ex)
                            {


                                throw new Exception("An error occured while getting the orders data from PowerChart for the FIN:" + FIN + ". Error message:" + ex.Message + " Inner exception: " + ex.InnerException);

                            }

                            #endregion
                            Atos.SyntBots.Common.Logger.Log("END - Get orders from Powerchart", LogFilePath);

                            Atos.SyntBots.Common.Logger.Log("START - Get Charges from ChargeViewer", LogFilePath);
                            #region step-3 - Get the Charges from ChargeViewer
                            try
                            {

                                var charges_MasterData = File.ReadLines(chargesFilePath).Select(line =>
                                {
                                    Charge ch = new Charge();
                                    Regex CSVParser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");
                                    String[] split = CSVParser.Split(line);
                                    ch.FIN = split[0].Replace("\"", ""); ;
                                    ch.ChargeDescription = split[3].Replace("\"", "");
                                    ch.ChargeDate = split[4].Replace("\"", "");
                                    ch.CPT4 = split[5].Replace("\"", "");
                                    ch.Price = split[7].Replace("\"", "");
                                    ch.ActivityType = split[8].Replace("\"", "");
                                    ch.ChargeType = split[9].Replace("\"", "");
                                    ch.Status = split[10].Replace("\"", "");
                                    return ch;
                                }).ToList<Atos.SyntBots.Common.Entity.Charge>();

                                var chargeByFin = charges_MasterData.Where(x => x.FIN.Contains(FIN) && x.Status == "Posted").ToList();
                                var creditCharges = chargeByFin.Where(x => x.FIN.Contains(FIN) && x.ChargeType == "CREDIT" && x.Status == "Posted").ToList(); //chargeByFin.Where(s => s.ChargeType == "CREDIT");
                                char[] MyChar = { '-' };
                                foreach (var creditRow in creditCharges)
                                {
                                    var CreditResult = chargeByFin.Where(s => s.ChargeDescription.Contains(creditRow.ChargeDescription) && s.ChargeType.ToUpper() == "CREDIT").FirstOrDefault();

                                    if (CreditResult != null)
                                    {
                                        chargeByFin.Remove(CreditResult);
                                        var DebitResult = chargeByFin.Where(s => s.ChargeDescription.Contains(creditRow.ChargeDescription) && s.ChargeType.ToUpper() == "DEBIT").FirstOrDefault();
                                        if (DebitResult != null)
                                        {
                                            chargeByFin.Remove(DebitResult);
                                        }
                                    }

                                }

                                Atos.SyntBots.AccessHIM.ChargeViewer chargeViewer = new Atos.SyntBots.AccessHIM.ChargeViewer();
                                chargeList = chargeViewer.GetChargesData(strAdmitDate, chargeByFin);
                                Atos.SyntBots.Common.Logger.Log("ChargeViewer records : " + chargeList.Count, LogFilePath);
                            }
                            catch (Exception ex)
                            {

                                if (ex.Message.Contains("Suspended Charge"))
                                {
                                    //try
                                    //{
                                    //    _accessHIMWindow = GetWindowWithWait("AccessHIM", 15);
                                    //    var _grid = _accessHIMWindow.GetMultiple<TestStack.White.UIItems.ListView>(SearchCriteria.ByControlType(ControlType.DataGrid));
                                    //    TreeWalker tWalker = TreeWalker.ControlViewWalker;
                                    //    AutomationElement fin = _grid[0].GetElement(SearchCriteria.ByText(FIN));
                                    //    AutomationElement Row = tWalker.GetParent(fin);
                                    //    SelectionItemPattern pattern1 = (SelectionItemPattern)(BasePattern)Row.GetCurrentPattern(SelectionItemPattern.Pattern);
                                    //    pattern1.Select();
                                    //    _grid[0].KeyIn(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
                                    //    _grid[0].KeyIn(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.UP);


                                    //}
                                    //catch (Exception exec)
                                    //{

                                    //    throw new Exception("Error occured while selecting the FIN. Error: " + exec.Message);
                                    //}

                                    try
                                    {
                                        //AutomationElement Btn = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Pass", null, null, "button", null, -1, 15);
                                        //driver.AEClick(Btn);
                                        //AutomationElement aePassReasonDialog = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Pass on Task", null, null, "Dialog", null, -1, 15);
                                        //AutomationElement aePassReasonTxt = driver.GetElmtByWlker(aePassReasonDialog, "Reason", null, null, "edit", null, -1, 15);
                                        ////aePassReasonTxt.
                                        //driver.AESetText(aePassReasonTxt, "Coding - Suspended Charges");
                                        ////driver.AESetText(aePassReasonTxt, "Coding - Action BOT Document Verified");
                                        //AutomationElement aePassDropDown = driver.GetElmtByWlker(aePassReasonDialog, "Drop Down Button", null, null, "button", null, -1, 15);
                                        //driver.AEClick(aePassDropDown);
                                        //driver.AEClick(aePassDropDown);
                                        //AutomationElement aePassOkBtn = driver.GetElmtByWlker(aePassReasonDialog, "Ok", null, null, "button", null);
                                        //driver.AEClick(aePassOkBtn);
                                        //Atos.SyntBots.Common.Logger.Log("Status changed : Coding - Suspended Charges", LogFilePath);
                                        UpdateFINStatus(toBeProcessedfilepath, FIN, "Suspended Charge");
                                        Atos.SyntBots.Common.Logger.Log("***END - Processing fin:" + FIN, LogFilePath);
                                        Atos.SyntBots.Common.Logger.Log("", LogFilePath);
                                    }
                                    catch (Exception exception)
                                    {

                                        throw new Exception("Error occured while changing the status. Error : " + exception.Message);
                                    }
                                    continue;
                                }
                                else
                                {
                                    throw new Exception("An error occured while getting the charges data from ChargeViewer for the FIN-" + FIN + ". Error message:" + ex.Message + " Inner exception: " + ex.InnerException);
                                }

                            }
                            #endregion
                            Atos.SyntBots.Common.Logger.Log("END - Get Charges from ChargeViewer", LogFilePath);
                            Atos.SyntBots.Common.Logger.Log("START - Get Exploding Charges", LogFilePath);
                            #region Get the Exploding Charges
                            bool boCharges = chargeList.Exists(x => x.ChargeDescription.Trim().StartsWith("BO"));
                            if (boCharges)
                            {
                                string explodeChargesfilepath = string.Empty;
                                string addonChargesfilepath = string.Empty;
                                FileInfo[] explodeChargesfiles = rdirInfo.GetFiles(string.Format("*{0}{1}*", explodingChargesFileName, botAgentName), SearchOption.TopDirectoryOnly);
                                FileInfo[] addonChargesfiles = rdirInfo.GetFiles(string.Format("*{0}{1}*", addonChargesFileName, botAgentName), SearchOption.TopDirectoryOnly);
                                if (addonChargesfiles.Count() > 0)
                                {
                                    addonChargesfilepath = addonChargesfiles[0].FullName;
                                    try
                                    {
                                        var addonCharge_MasterList = File.ReadLines(addonChargesfilepath).Skip(1).Select(line =>
                                        {
                                            Charge ch = new Charge();
                                            Regex CSVParser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");
                                            String[] split = CSVParser.Split(line);
                                            ch.ChargeDescription = split[1].Replace("\"", "");
                                            return ch;
                                        }).ToList<Atos.SyntBots.Common.Entity.Charge>();
                                        foreach (var addonCharge in addonCharge_MasterList)
                                        {
                                            var result = chargeList.Where(s => s.ChargeDescription.ToLower().Trim().Contains(addonCharge.ChargeDescription.ToLower().Trim())).FirstOrDefault();
                                            if (result != null)
                                            {
                                                chargeList.Remove(result);
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("An error occured while getting addon charges data for the FIN-" + FIN + ". Error message:" + ex.Message + " Inner exception: " + ex.InnerException);
                                    }
                                }
                                if (explodeChargesfiles.Count() > 0)
                                {
                                    explodeChargesfilepath = explodeChargesfiles[0].FullName;
                                    try
                                    {
                                        var explodingCharge_MasterList = File.ReadLines(explodeChargesfilepath).Skip(1).Select(line =>
                                        {
                                            Charge ch = new Charge();
                                            Regex CSVParser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");
                                            String[] split = CSVParser.Split(line);
                                            ch.OrderDescription = split[0].Replace("\"", "");
                                            ch.ChargeDescription = split[1].Replace("\"", "");
                                            return ch;
                                        }).ToList<Atos.SyntBots.Common.Entity.Charge>();



                                        foreach (var order in listPowerChartOrders)
                                        {

                                            IList<string> orderNames = (from DataRow dr in orderCatalogData.Rows
                                                                        where (string)dr["CATALOG_DISPLAY"] == order.OrderDescription.Trim()
                                                                        select (string)dr["CATALOG_DISPLAY2"]).ToList();

                                            explodingChargeList = explodingCharge_MasterList.Where(x => x.OrderDescription.ToLower().Trim().Equals(order.OrderDescription.ToLower().Trim())).ToList();
                                            foreach (var explodingCharge in explodingChargeList)
                                            {
                                                var result = chargeList.Where(s => s.ChargeDescription.ToLower().Trim().Contains(explodingCharge.ChargeDescription.ToLower().Trim())).FirstOrDefault();
                                                if (result != null)
                                                {
                                                    chargeList.Remove(result);
                                                }
                                            }
                                            foreach (var orderName in orderNames)
                                            {
                                                explodingChargeList = explodingCharge_MasterList.Where(x => x.OrderDescription.ToLower().Trim().Equals(orderName.ToLower().Trim())).ToList();
                                                foreach (var explodingCharge in explodingChargeList)
                                                {
                                                    var result = chargeList.Where(s => s.ChargeDescription.ToLower().Trim().Contains(explodingCharge.ChargeDescription.ToLower().Trim())).FirstOrDefault();
                                                    if (result != null)
                                                    {
                                                        chargeList.Remove(result);
                                                    }
                                                }
                                            }

                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("An error occured while getting exploding charges data for the FIN-" + FIN + ". Error message:" + ex.Message + " Inner exception: " + ex.InnerException);

                                    }
                                }
                            }
                            #endregion
                            Atos.SyntBots.Common.Logger.Log("END - Get exploding charges", LogFilePath);



                            Atos.SyntBots.Common.Logger.Log("START - Compare Charges", LogFilePath);
                            #region Step-4 - Check for Zero Charges

                            if (listPowerChartOrders.Count() == 0 && chargeList.Count() == 0)
                            {
                                //toBeProcessedWSheet.Cells[i, commentsColumnIndex] = "Zero Charges";//string.Format("{0} - Zero Charges", DateTime.Now.ToString());
                                isZeroCharges = true;

                                //try
                                //{
                                //    _accessHIMWindow = GetWindowWithWait("AccessHIM", 15);
                                //    var _grid = _accessHIMWindow.GetMultiple<TestStack.White.UIItems.ListView>(SearchCriteria.ByControlType(ControlType.DataGrid));
                                //    TreeWalker tWalker = TreeWalker.ControlViewWalker;
                                //    AutomationElement fin = _grid[0].GetElement(SearchCriteria.ByText(FIN));
                                //    AutomationElement Row = tWalker.GetParent(fin);
                                //    SelectionItemPattern pattern1 = (SelectionItemPattern)(BasePattern)Row.GetCurrentPattern(SelectionItemPattern.Pattern);
                                //    pattern1.Select();
                                //    _grid[0].KeyIn(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
                                //    _grid[0].KeyIn(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.UP);


                                //}
                                //catch (Exception exec)
                                //{

                                //    throw new Exception("Error occured while selecting the FIN. Error: " + exec.Message);
                                //}

                                /*
                                if (!debugMode)
                                {
                                    if (!shouldDNBeRaised)
                                    {
                                        try
                                        {
                                            AutomationElement Btn = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Pass", null, null, "button", null, -1, 15);
                                            driver.AEClick(Btn);
                                            AutomationElement aePassReasonDialog = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Pass on Task", null, null, "Dialog", null, -1, 15);
                                            AutomationElement aePassReasonTxt = driver.GetElmtByWlker(aePassReasonDialog, "Reason", null, null, "edit", null, -1, 15);
                                            //aePassReasonTxt.
                                            driver.AESetText(aePassReasonTxt, "Coding - Action BOT Hold");
                                            //driver.AESetText(aePassReasonTxt, "Coding - Action BOT Document Verified");
                                            AutomationElement aePassDropDown = driver.GetElmtByWlker(aePassReasonDialog, "Drop Down Button", null, null, "button", null, -1, 15);
                                            driver.AEClick(aePassDropDown);
                                            driver.AEClick(aePassDropDown);
                                            AutomationElement aePassOkBtn = driver.GetElmtByWlker(aePassReasonDialog, "Ok", null, null, "button", null);
                                            driver.AEClick(aePassOkBtn);
                                            Atos.SyntBots.Common.Logger.Log("Status changed : Coding - Action BOT Hold", LogFilePath);
                                        }
                                        catch (Exception exception)
                                        {

                                            throw new Exception("Error occured while changing the status. Error : " + exception.Message);
                                        }
                                    }

                                }
                                */


                            }

                            #endregion

                            if (!isZeroCharges)
                            {

                                var ordersToBeRemoved = new List<Order>();
                                if ((chargeList.Where(s => s.CPT4 == "81001").Count() > 0) && listPowerChartOrders.Where(s => s.CPT4 == "81003").Count() > 0 && listPowerChartOrders.Where(s => s.CPT4 == "81015").Count() > 0)
                                {
                                    listPowerChartOrders.Remove(listPowerChartOrders.Where(s => s.CPT4 == "81003").FirstOrDefault());
                                    listPowerChartOrders.Remove(listPowerChartOrders.Where(s => s.CPT4 == "81015").FirstOrDefault());
                                    chargeList.Remove(chargeList.Where(s => s.CPT4 == "81001").FirstOrDefault());
                                }
                                if ((chargeList.Where(s => s.CPT4 == "81001").Count() > 0) && listPowerChartOrders.Where(s => s.CPT4 == "81003").Count() > 0 && listPowerChartOrders.Where(s => s.CPT4 == "81015").Count() > 0)
                                {
                                    listPowerChartOrders.Remove(listPowerChartOrders.Where(s => s.CPT4 == "81003").FirstOrDefault());
                                    listPowerChartOrders.Remove(listPowerChartOrders.Where(s => s.CPT4 == "81015").FirstOrDefault());
                                    chargeList.Remove(chargeList.Where(s => s.CPT4 == "81001").FirstOrDefault());
                                }
                                foreach (var order in listPowerChartOrders)
                                {
                                    IList<string> cptCodeList = new List<string>();
                                    IList<string> accessHIMChargeDescriptionsList = new List<string>();

                                    if (order.CPT4 == "MultipleCPT")
                                    {
                                        cptCodeList = orderCatalog.GetAllCPTForOrderDescription(orderCatalogData, order.OrderDescription);
                                    }
                                    else if (order.CPT4 != "MultipleCPT" && !string.IsNullOrEmpty(order.CPT4))
                                    {
                                        cptCodeList.Add(order.CPT4);
                                    }
                                    else if (order.CPT4 == "" && !string.IsNullOrEmpty(order.OrderDescription))
                                    {
                                        var res = orderCatalog.GetAccessHIMChargeDescription(orderCatalogData, order.OrderDescription);
                                        if (res.Length > 0)
                                        {
                                            accessHIMChargeDescriptionsList.Add(res);
                                        }
                                    }
                                    accessHIMChargeDescriptionsList = orderCatalog.GetAllAccessHIMChargeDescriptionsForCPT(orderCatalogData, cptCodeList);

                                    foreach (var item in accessHIMChargeDescriptionsList)
                                    {
                                        // #Specail Case Scenario: 
                                        // If either of CPTs (85651, 85652) is available in Orders then look for either of the CPTs (85651, 85652) is exists in Charges.
                                        if (order.CPT4 == "85651" || order.CPT4 == "85652")
                                        {
                                            var result1 = chargeList.Where(s => s.CPT4.Contains("85651") || s.CPT4.Contains("85652") && s.ChargeDate == strAdmitDate).FirstOrDefault();
                                            if (result1 != null)
                                            {
                                                ordersToBeRemoved.Add(order);
                                                chargeList.Remove(result1);
                                                break;
                                            }
                                        }

                                        var result = chargeList.Where(s => s.ChargeDescription.Contains(item) && s.ChargeDate == strAdmitDate).FirstOrDefault();
                                        if (result != null)
                                        {
                                            ordersToBeRemoved.Add(order);
                                            chargeList.Remove(result);
                                            break;
                                        }
                                    }
                                }

                                if (ordersToBeRemoved.Count > 0)
                                {
                                    foreach (var element in ordersToBeRemoved)
                                    {
                                        listPowerChartOrders.Remove(element);
                                    }
                                }

                                //if(chargeList.Where(s=> s.CPT4 == "81001").Count() > 0)
                                //{
                                //    if (listPowerChartOrders.Where(s => s.CPT4 == "81001").Count() > 0)
                                //    {
                                //        chargeList.Remove(chargeList.Where(s => s.CPT4 == "81001").FirstOrDefault());
                                //        ordersToBeRemoved.Add(listPowerChartOrders.Where(s => s.CPT4 == "81001").FirstOrDefault());
                                //    }
                                //    if (listPowerChartOrders.Where(s => s.CPT4 == "81003").Count() > 0 && listPowerChartOrders.Where(s => s.CPT4 == "81015").Count() > 0)
                                //    {
                                //        chargeList.Remove(chargeList.Where(s => s.CPT4 == "81001").FirstOrDefault());
                                //        ordersToBeRemoved.Add(listPowerChartOrders.Where(s => s.CPT4 == "81003").FirstOrDefault());
                                //        ordersToBeRemoved.Add(listPowerChartOrders.Where(s => s.CPT4 == "81015").FirstOrDefault());
                                //    }
                                //}

                                if ((chargeList.Where(s => s.CPT4 == "81001").Count() == 0) && listPowerChartOrders.Where(s => s.CPT4 == "81003").Count() > 0 && listPowerChartOrders.Where(s => s.CPT4 == "81015").Count() > 0)
                                {
                                    listPowerChartOrders.Remove(listPowerChartOrders.Where(s => s.CPT4 == "81003").FirstOrDefault());
                                    listPowerChartOrders.Remove(listPowerChartOrders.Where(s => s.CPT4 == "81015").FirstOrDefault());
                                    listPowerChartOrders.Add(new Order { OrderDescription = "Urinalysis, Macro and Micro (M)", CPT4 = "81001" });
                                }

                                #region Step-6 - Check for Missing Charges - Addition
                                if (listPowerChartOrders.Count > 0 && chargeList.Count == 0)
                                {
                                    //toBeProcessedWSheet.Cells[i, commentsColumnIndex] = " Missing Charges";//string.Format("{0} - Missing Charges", DateTime.Now.ToString());
                                    foreach (var obj in listPowerChartOrders)
                                    {
                                        sbMissingCharges.Append(obj.OrderDescription + " " + obj.CPT4).Append(",");
                                    }
                                }
                                #endregion

                                #region Step-7 - Check for Charge Modification - Deletion
                                if (chargeList.Count > 0 && listPowerChartOrders.Count == 0)
                                {
                                    //toBeProcessedWSheet.Cells[i, commentsColumnIndex] = "Charge Deletion"; //string.Format("{0} - Charge Deletion", DateTime.Now.ToString());
                                    foreach (var obj in chargeList)
                                    {
                                        sbChargeDeletion.Append(obj.ChargeDescription + " " + obj.CPT4).Append(",");
                                    }
                                }
                                #endregion

                                #region Step-8 - Check for Charge Modification - Correction
                                if (chargeList.Count == listPowerChartOrders.Count && chargeList.Count > 0 && listPowerChartOrders.Count > 0)
                                {
                                    //toBeProcessedWSheet.Cells[i, commentsColumnIndex] = "Charge Modification"; //string.Format("{0} - Charge Modification", DateTime.Now.ToString());
                                    for (int k = 0; k < chargeList.Count; k++)
                                    {
                                        sbChargeModification.Append(chargeList[k].ChargeDescription + " " + chargeList[k].CPT4 + " to " + listPowerChartOrders[k].OrderDescription + " " + listPowerChartOrders[k].CPT4).Append(",");
                                    }
                                }

                                #endregion


                                if ((chargeList.Count > 0 && listPowerChartOrders.Count > 0) && (chargeList.Count != listPowerChartOrders.Count))
                                {
                                    //toBeProcessedWSheet.Cells[i, commentsColumnIndex] = "Charge Addition & Deletion";
                                    foreach (var obj in listPowerChartOrders)
                                    {
                                        sbMissingCharges.Append(obj.OrderDescription + " " + obj.CPT4).Append(",");
                                    }

                                    foreach (var obj in chargeList)
                                    {
                                        sbChargeDeletion.Append(obj.ChargeDescription + " " + obj.CPT4).Append(",");
                                    }
                                }

                                //try
                                //{
                                //    _accessHIMWindow = GetWindowWithWait("AccessHIM", 15);
                                //    var _grid = _accessHIMWindow.GetMultiple<TestStack.White.UIItems.ListView>(SearchCriteria.ByControlType(ControlType.DataGrid));
                                //    TreeWalker tWalker = TreeWalker.ControlViewWalker;
                                //    AutomationElement fin = _grid[0].GetElement(SearchCriteria.ByText(FIN));
                                //    AutomationElement Row = tWalker.GetParent(fin);
                                //    SelectionItemPattern pattern1 = (SelectionItemPattern)(BasePattern)Row.GetCurrentPattern(SelectionItemPattern.Pattern);
                                //    pattern1.Select();
                                //    _grid[0].KeyIn(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
                                //    _grid[0].KeyIn(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.UP);


                                //}
                                //catch (Exception ex)
                                //{

                                //    throw new Exception("Error occured while selecting the FIN. Error: " + ex.Message);
                                //}


                                #region If it is Good Record append to 'ToBeAllocated' file
                                if (listPowerChartOrders.Count == 0 && chargeList.Count == 0)
                                {
                                    isGoodRecord = true;

                                    //if (!debugMode)
                                    //{
                                    //    try
                                    //    {

                                    //        AutomationElement Btn = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Pass", null, null, "button", null, -1, 15);
                                    //        driver.AEClick(Btn);
                                    //        AutomationElement aePassReasonDialog = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Pass on Task", null, null, "Dialog", null, -1, 15);
                                    //        AutomationElement aePassReasonTxt = driver.GetElmtByWlker(aePassReasonDialog, "Reason", null, null, "edit", null, -1, 15);
                                    //        //aePassReasonTxt.
                                    //        //driver.AESetText(aePassReasonTxt, "Coding - Action BOT Hold");
                                    //        driver.AESetText(aePassReasonTxt, "Coding - Action BOT Document Verified");
                                    //        AutomationElement aePassDropDown = driver.GetElmtByWlker(aePassReasonDialog, "Drop Down Button", null, null, "button", null, -1, 15);
                                    //        driver.AEClick(aePassDropDown);
                                    //        driver.AEClick(aePassDropDown);
                                    //        AutomationElement aePassOkBtn = driver.GetElmtByWlker(aePassReasonDialog, "Ok", null, null, "button", null);
                                    //        driver.AEClick(aePassOkBtn);
                                    //        Atos.SyntBots.Common.Logger.Log("Status changed : Coding - Action BOT Document Verified", LogFilePath);
                                    //    }
                                    //    catch (Exception ex)
                                    //    {

                                    //        throw new Exception("Error occured while changing the status. Error : " + ex.Message);
                                    //    }
                                    //}




                                    tobeAllocatedCurrentRow++;
                                    totalGoodRecords++;
                                    UpdateFINStatus(toBeProcessedfilepath, FIN, "Good Record");
                                }
                                //else
                                //{
                                //    if (!debugMode)
                                //    {
                                //        if (!shouldDNBeRaised)
                                //        {
                                //            try
                                //            {
                                //                AutomationElement Btn = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Pass", null, null, "button", null, -1, 15);
                                //                driver.AEClick(Btn);
                                //                AutomationElement aePassReasonDialog = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Pass on Task", null, null, "Dialog", null, -1, 15);
                                //                AutomationElement aePassReasonTxt = driver.GetElmtByWlker(aePassReasonDialog, "Reason", null, null, "edit", null, -1, 15);
                                //                //aePassReasonTxt.
                                //                driver.AESetText(aePassReasonTxt, "Coding - Action BOT Hold");
                                //                //driver.AESetText(aePassReasonTxt, "Coding - Action BOT Document Verified");
                                //                AutomationElement aePassDropDown = driver.GetElmtByWlker(aePassReasonDialog, "Drop Down Button", null, null, "button", null, -1, 15);
                                //                driver.AEClick(aePassDropDown);
                                //                driver.AEClick(aePassDropDown);
                                //                AutomationElement aePassOkBtn = driver.GetElmtByWlker(aePassReasonDialog, "Ok", null, null, "button", null);
                                //                driver.AEClick(aePassOkBtn);
                                //                Atos.SyntBots.Common.Logger.Log("Status changed : Coding - Action BOT Hold", LogFilePath);

                                //            }
                                //            catch (Exception ex)
                                //            {

                                //                throw new Exception("Error occured while changing the status. Error : " + ex.Message);
                                //            }
                                //        }
                                //    }


                                //}
                                #endregion



                            }
                            Atos.SyntBots.Common.Logger.Log("END - Compare Charges", LogFilePath);


                            if (appendDNComment)
                            {
                                string comment = string.Empty;
                                string dn = string.Empty;
                                string resPerson = string.Empty;
                                if (isZeroCharges)
                                {
                                    comment = string.Format("Attention {0} Please provide charges and/or orders and/or reports performed on service date {1}", resPerson, strAdmitDate);
                                    dn = "Zero Charges - " + comment;

                                    //toBeProcessedWSheet.Cells[i, commentsColumnIndex] = dn;
                                    //toBeProcessedWSheet.Cells[i, commentsColumnIndex + 1] = comment;
                                    UpdateFINStatus(toBeProcessedfilepath, FIN, dn);
                                }
                                else if (sbMissingCharges.Length > 0)
                                {
                                    comment = string.Format("Attention {0} Please provide charges for {1} for the date of service {2}", resPerson, sbMissingCharges.ToString(), strAdmitDate);
                                    dn = "Missing Charges - " + comment;

                                    //toBeProcessedWSheet.Cells[i, commentsColumnIndex] = dn;
                                    //toBeProcessedWSheet.Cells[i, commentsColumnIndex + 1] = comment;
                                    UpdateFINStatus(toBeProcessedfilepath, FIN, dn);
                                }
                                else if (sbChargeModification.Length > 0)
                                {

                                    comment = string.Format("Attention {0}, Please modify {1} for the date of service {2}.", resPerson, sbChargeModification.ToString(), strAdmitDate);
                                    dn = "Charge Modification - " + comment;

                                    //toBeProcessedWSheet.Cells[i, commentsColumnIndex] = dn;
                                    //toBeProcessedWSheet.Cells[i, commentsColumnIndex + 1] = comment;
                                    UpdateFINStatus(toBeProcessedfilepath, FIN, dn);

                                }
                                else if (sbChargeDeletion.Length > 0)
                                {
                                    comment = string.Format("Attention {0}, please delete {1} as it is not done on date of service {2}.", resPerson, sbChargeDeletion.ToString(), strAdmitDate);
                                    dn = "Charge Deletion - " + comment;

                                    //toBeProcessedWSheet.Cells[i, commentsColumnIndex] = dn;
                                    //toBeProcessedWSheet.Cells[i, commentsColumnIndex + 1] = comment;
                                    UpdateFINStatus(toBeProcessedfilepath, FIN, dn);
                                }

                                if (sbMissingCharges.Length > 0 && sbChargeDeletion.Length > 0 && sbChargeModification.Length == 0)
                                {
                                    comment = string.Format("Attention {0} Please provide charges for {1} for the date of service {2} and delete {3} as it is not done on date of service {4}", resPerson, sbMissingCharges.ToString(), strAdmitDate, sbChargeDeletion.ToString(), strAdmitDate);
                                    dn = "Charge addition and deletion - " + comment;

                                    //toBeProcessedWSheet.Cells[i, commentsColumnIndex] = dn;
                                    //toBeProcessedWSheet.Cells[i, commentsColumnIndex + 1] = comment;
                                    UpdateFINStatus(toBeProcessedfilepath, FIN, dn);
                                }

                                if (isGoodRecord)
                                {
                                    Atos.SyntBots.Common.Logger.Log("Good Record", LogFilePath);
                                }
                                else
                                {
                                    Atos.SyntBots.Common.Logger.Log(comment, LogFilePath);
                                }

                            }


                            if (shouldDNBeRaised && !isGoodRecord)
                            {
                                // Atos.SyntBots.Common.Logger.Log("START - Raise DN", LogFilePath);

                                //AutomationElement Btn = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Open", null, null, "button", null, -1, 15);
                                //driver.AEClick(Btn);

                                /*
                                Atos.SyntBots.Common.Logger.Log("START - Search FIN in AccessHIM", LogFilePath);
                                #region Step-1 - Search FIN number in AccessHIM
                                try
                                {
                                    _accessHIMWindow.SetForeground();
                                    Thread.Sleep(1000);
                                    //Condition conditions = new AndCondition(new PropertyCondition(AutomationElement.IsEnabledProperty, true), new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane));
                                    //AutomationElementCollection elementCollection = _accessHIMWindow.AutomationElement.FindAll(TreeScope.Children, conditions);
                                    var search = new Atos.SyntBots.AccessHIM.Search();
                                    //if(elementCollection !=null)
                                    //{

                                    //search.SearchFIN(_accessHIMWindow, FIN);
                                    //}

                                    #region Search FIN number in AccessHIM
                                    _accessHIMWindow.WaitWhileBusy();
                                    Keyboard.Instance.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                                    Keyboard.Instance.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                                    Keyboard.Instance.Enter("f");
                                    Keyboard.Instance.Enter("o");

                                    AutomationElement AE40 = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Fin Nbr", null, null, "edit", null, -1, 15);
                                    driver.AEClick(AE40);
                                    Keyboard.Instance.Enter(FIN);
                                    AutomationElement AE44 = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Search", null, null, "button", null, -1, 15);
                                    driver.AEClick(AE44);
                                    System.Threading.Thread.Sleep(3000);
                                    //AutomationElement panePersonSearch = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Person Search", null, null, "Dialog", null, -1, 15);


                                    //Check for no result found..

                                    AutomationElement AE45 = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Select", null, null, "button", null);
                                    //System.Threading.Thread.Sleep(1000);

                                    driver.AEClick(AE45);
                                    //System.Threading.Thread.Sleep(1000);
                                    //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);

                                    #endregion


                                }
                                catch (Exception ex)
                                {

                                    throw new Exception("An error occured seaching the FIN:" + FIN + ". Error message:" + ex.Message + " Inner exception: " + ex.InnerException); ;
                                }

                                #endregion Step-1 - Search FIN number in AccessHIM
                                Atos.SyntBots.Common.Logger.Log("END - Search FIN in AccessHIM", LogFilePath);
                                */
                                #region Step-5 - Raise Discern Notification

                                string dnComment = string.Empty;
                                string codingPassReason = string.Empty;
                                string personResponsible = string.Empty;
                                string dnType = string.Empty;
                                if (isZeroCharges || sbMissingCharges.Length > 0 || sbChargeModification.Length > 0 || sbChargeDeletion.Length > 0)
                                {
                                    //var dr = codingQueueAssignment.GetAssignmentDetailsByLocation(codingQueueAssignmentData, locationTxt);
                                    //
                                    //                                        if (dr != null)
                                    //                                        {
                                    codingPassReason = "Coding - Charge Correction Lab";//dr[3].ToString();
                                    personResponsible = "";//dr[2].ToString();


                                    if (isZeroCharges)
                                    {
                                        dnComment = string.Format("Attention {0} Please provide charges and/or orders and/or reports performed on service date {1}", personResponsible, strAdmitDate);
                                        dnType = "Zero Charges";
                                    }
                                    else if (sbMissingCharges.Length > 0)
                                    {
                                        dnComment = string.Format("Attention {0} Please provide charges for {1} for the date of service {2}", personResponsible, sbMissingCharges.ToString(), strAdmitDate);
                                        dnType = "Missing Charges";
                                    }
                                    else if (sbChargeModification.Length > 0)
                                    {

                                        dnComment = string.Format("Attention {0}, Please modify {1} for the date of service {2}.", personResponsible, sbChargeModification.ToString(), strAdmitDate);
                                        dnType = "Charge Modification";

                                    }
                                    else if (sbChargeDeletion.Length > 0)
                                    {
                                        dnComment = string.Format("Attention {0}, please delete {1} as it is not done on date of service {2}.", personResponsible, sbChargeDeletion.ToString(), strAdmitDate);
                                        dnType = "Charge Deletion";
                                    }

                                    if (sbMissingCharges.Length > 0 && sbChargeDeletion.Length > 0 && sbChargeModification.Length == 0)
                                    {
                                        dnComment = string.Format("Attention {0} Please provide charges for {1} for the date of service {2} and delete {3} as it is not done on date of service {4}", personResponsible, sbMissingCharges.ToString(), strAdmitDate, sbChargeDeletion.ToString(), strAdmitDate);
                                        dnType = "Charge addition and deletion";
                                    }



                                    /*
                                    bool isDefaultcodeAdded = true;
                                    try
                                    {
                                        System.Diagnostics.Process[] _3mProcesses = System.Diagnostics.Process.GetProcessesByName("weblink64");
                                        foreach (System.Diagnostics.Process p in _3mProcesses)
                                        {

                                            try
                                            {
                                                p.Kill();
                                            }
                                            catch { }

                                        }
                                        _accessHIMWindow.SetForeground();

                                        Thread.Sleep(500);
                                        // _launcher.Launch3M(_accessHIMWindow);
                                        //Thread.Sleep(10000);
                                        AutomationElement AEEncoderBtn = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Launch Encoder", null, null, "button", null);
                                        driver.AEClick(AEEncoderBtn);

                                        TestStack.White.UIItems.WindowItems.Window _3MWindow = GetWindowWithWait("360 Encompass: Visit:", 120);//Desktop.Instance.Windows().Find(K => K.Name.Contains("360 Encompass: Visit:"));
                                        _3MWindow.SetForeground();

                                        Thread.Sleep(1000);
                                        //Atos.SyntBots._3M.DefaultCode defaultcode = new Atos.SyntBots._3M.DefaultCode();
                                        //defaultcode.AddDefaultCode(_3MWindow, isZeroCharges);
                                        _3MWindow.Close();
                                        Thread.Sleep(2000);
                                        //Close End Session Popup
                                        AutomationElement AEEndSessionDialog = driver.GetElmtByWlker(_3MWindow.AutomationElement, "End Session", null, "#32770", "Dialog", ControlType.Window, -1, 5);
                                        if (AEEndSessionDialog != null)
                                        {
                                            AutomationElement aeYes = driver.GetElmtByWlker(AEEndSessionDialog, "Yes", null, null, "button", null, -1, 5);
                                            driver.AEClick(aeYes);
                                        }
                                        //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);


                                        //3M Error Popup
                                        AutomationElement AErt1 = driver.GetDesktopChild(false, "3M Coding and Reimbursement System", "", "#32770", "Dialog", null, 5);
                                        if (AErt1 != null)
                                        {
                                            AutomationElement AE1 = driver.GetElmtByWlker(AErt1, null, "TitleBar", null, null, null);
                                            driver.AEClick(AE1);

                                            Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);

                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        isErrorOccurredAtDNLevel = true;
                                        isDefaultcodeAdded = false;
                                        var _3MWindow = GetWindowWithWait("360 Encompass:", 30);//Desktop.Instance.Windows().Find(K => K.Name.Contains("360 Encompass: Visit:"));
                                        if (_3MWindow != null)
                                        {
                                            _3MWindow.Close();
                                            Thread.Sleep(500);
                                            Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
                                        }

                                        throw new Exception("An error occured while getting the adding the default code in 3M for the FIN:" + FIN + ". Error message:" + ex.Message + " Inner exception: " + ex.InnerException);
                                    }

                                    if (isDefaultcodeAdded && !isErrorOccurredAtDNLevel)
                                    {
                                        try
                                        {

                                            Atos.SyntBots.AccessHIM.DiscernNotification dn = new Atos.SyntBots.AccessHIM.DiscernNotification();
                                            _accessHIMWindow.SetForeground();
                                            Thread.Sleep(500);
                                            dn.SendDiscernNotification(dnComment, codingPassReason, personResponsible);
                                            totalDNRaised++;
                                            UpdateFINStatus(toBeProcessedfilepath, FIN, string.Format("{0} DN Raised", dnType));
                                            
                                            Atos.SyntBots.Common.Logger.Log(string.Format("{0} DN Raised", dnType), LogFilePath);
                                        }
                                        catch (Exception ex)
                                        {
                                            isErrorOccurredAtDNLevel = true;
                                            throw new Exception("An error occured while raising the DN for the FIN:" + FIN + ". Error message:" + ex.Message + " Inner exception: " + ex.InnerException);
                                        }
                                    }
                                    else
                                    {
                                        throw new Exception(dnType + " DN cannot be raised for the FIN:" + FIN + ". Default code is not added in 3M.");
                                    }

                                    */

                                    //                                        }
                                    //                                        else
                                    //                                        {
                                    //                                            throw new Exception(dnType + " DN cannot be raised for the FIN:" + FIN + ". Could not find the Resposibile person for the location: " + locationTxt);
                                    //                                        }

                                }

                                #endregion
                                //Atos.SyntBots.Common.Logger.Log("END - Raise DN", LogFilePath);


                                //_accessHIMWindow.SetForeground();
                                //Thread.Sleep(500);
                                //Keyboard.Instance.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                                //Keyboard.Instance.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                                //Keyboard.Instance.Enter("f");
                                //Keyboard.Instance.Enter("c");
                            }

                            Atos.SyntBots.Common.Logger.Log("***END - Processing fin:" + FIN, LogFilePath);
                            Atos.SyntBots.Common.Logger.Log("", LogFilePath);

                        }
                        catch (Exception ex)
                        {
                            if (ex.Message.Contains("Description not found"))
                            {
                                totalDataErrorRecords++;
                                UpdateFINStatus(toBeProcessedfilepath, FIN, "Description not found");
                            }
                            else if (ex.Message.Contains("Excluded CPT"))
                            {
                                totalDataErrorRecords++;
                                UpdateFINStatus(toBeProcessedfilepath, FIN, "Excluded CPT");
                            }
                            else
                            {
                                totalErrorRecords++;
                                UpdateFINStatus(toBeProcessedfilepath, FIN, "Failed to process");
                            }

                            Atos.SyntBots.Common.Logger.Log(ex.Message, LogFilePath);

                            //if (shouldDNBeRaised)
                            //{
                            //    Keyboard.Instance.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                            //    Keyboard.Instance.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                            //    Keyboard.Instance.Enter("f");
                            //    Keyboard.Instance.Enter("c");

                            //    if (isErrorOccurredAtDNLevel)
                            //    {
                            //        Keyboard.Instance.Enter("y");
                            //    }
                            //}
                        }
                    }

                    //_accessHIMWindow.Close();

                }
                else
                {
                    Atos.SyntBots.Common.Logger.Log("To be processed file does not exists. File name: " + toBeProcessedfilepath, LogFilePath);
                    throw new Exception(toBeProcessedfilepath + " file not found");
                }

                #endregion Process the ToBeProcessed file
                Atos.SyntBots.Common.Logger.Log("Completed the Account Verification process", LogFilePath);

                if (!IsFileLocked(jsonFilePath))
                {
                    string jsonString2 = File.ReadAllText(jsonFilePath);
                    JObject jObject2 = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonString2) as JObject;
                    JToken jBotTypeToStart2 = jObject2.SelectToken("BotTypeToStart");
                    jBotTypeToStart2.Replace("Account Categorization");

                    JToken jExecution2 = jObject2.SelectToken("Execution");
                    jExecution2.Replace("Completed");

                    JToken jBotExecution2 = jObject2.SelectToken("ExecutionStatusBot" + botAgentName);
                    jBotExecution2.Replace("Completed");

                    string updatedJsonString2 = jObject2.ToString();
                    File.WriteAllText(jsonFilePath, updatedJsonString2);
                }
                else
                {
                    throw new Exception("An error occurred while updating the json file.");
                }


                objJsnHlpr.SetOutCome("Success");
            }
            catch (Exception ex)
            {
                Atos.SyntBots.Common.Logger.Log("Error message: " + ex.Message + "\n Inner exception: " + ex.InnerException, LogFilePath);
                objJsnHlpr.SetErrorList(ex.Message.ToString(), "Technical", "006");
                objJsnHlpr.SetOutCome("Failure");

                string jsonString = File.ReadAllText(jsonFilePath);
                var jObject = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonString) as JObject;
                var jBotTypeToStart = jObject.SelectToken("BotTypeToStart");
                jBotTypeToStart.Replace("Account Verification");

                var jExecution = jObject.SelectToken("Execution");
                jExecution.Replace("Completed");

                var jBotExecution = jObject.SelectToken("ExecutionStatusBot" + botAgentName);
                jBotExecution.Replace("Failed");

                string updatedJsonString = jObject.ToString();
                File.WriteAllText(jsonFilePath, updatedJsonString);

                //var _accessHIM = Desktop.Instance.Windows().Find(K => K.Name.Contains("AccessHIM"));
                //if (_accessHIM != null)
                //{
                //    _accessHIM.Close();
                //}

                //var _powerchart = Desktop.Instance.Windows().Find(K => K.Name.Contains("Opened by"));
                //if (_powerchart != null)
                //{
                //    _powerchart.Close();
                //}

                //var _cernerMillennium = Desktop.Instance.Windows().Find(K => K.Name == "Cerner Millennium");
                //if (_cernerMillennium != null)
                //{
                //    _cernerMillennium.Close();
                //}

                //var _chargeViewer = Desktop.Instance.Windows().Find(x => x.Name == "Charge Viewer (Cannot Increase Quantity)");
                //if (_chargeViewer != null)
                //{
                //    _chargeViewer.Close();
                //}

                //var _3MWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("360 Encompass: Visit:"));
                //if (_3MWindow != null)
                //{
                //    _3MWindow.Close();
                //    Thread.Sleep(500);
                //    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
                //}


                //try
                //{
                //    Atos.SyntBots.Common.EmailHelper email = new Atos.SyntBots.Common.EmailHelper(senderEmail, senderPassword, host, port);
                //    string automationTeamEmail = _configuration.AppSettings.Settings["Email.AutomationTeam"] != null ? _configuration.AppSettings.Settings["Email.AutomationTeam"].Value : string.Empty;
                //    email.SendEMail(automationTeamEmail, "ALERT : Execution Failed - Bot Agent # " + botAgentName, "Bot Execution Failed. Error: " + ex.Message + ". InnerExeception: " + ex.InnerException, "");

                //}
                //catch (Exception exp)
                //{

                //    Atos.SyntBots.Common.Logger.Log("An error occured while sending email." + "Error: " + ex.Message);
                //}
            }
            finally
            {
                string jsonString = File.ReadAllText(jsonFilePath);
                var jObject = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonString) as JObject;
                var jBotTypeToStart = jObject.SelectToken("BotTypeToStart");
                jBotTypeToStart.Replace("Account Verification");

                var jExecution = jObject.SelectToken("Execution");
                jExecution.Replace("Completed");


                //var jBotExecution = jObject.SelectToken("ExecutionStatusBot" + botAgentName);
                //jExecution.Replace("Completed");

                string updatedJsonString = jObject.ToString();
                File.WriteAllText(jsonFilePath, updatedJsonString);

                //System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("Excel");
                //foreach (System.Diagnostics.Process p in processes)
                //{

                //    try
                //    {
                //        p.Kill();
                //    }
                //    catch { }

                //}

                objJsnHlpr.PopulateJsonData();

            }
        }

        public TestStack.White.UIItems.WindowItems.Window GetWindowWithWait(string windowName, int timeout = 30)
        {
            TestStack.White.UIItems.WindowItems.Window result = null;
            for (int i = 0; i < timeout; i++)
            {
                result = Desktop.Instance.Windows().Find(K => K.Name.Contains(windowName));
                if (result != null)
                {
                    return result;
                }
                Thread.Sleep(TimeSpan.FromSeconds(1));
            }
            return result;
        }

        public AutomationElement GetPopUpWindowWithWait(AutomationElement automationElement, int windowIndex = 0, int timeout = 30)
        {
            AutomationElement cvWindowPopup = null;
            for (int i = 0; i < timeout; i++)
            {
                cvWindowPopup = driver.GetRawChildrenByWlker(automationElement)[windowIndex];

                if (cvWindowPopup != null)
                {
                    return cvWindowPopup;
                }
                Thread.Sleep(TimeSpan.FromSeconds(1));
            }
            return cvWindowPopup;
        }

        public bool IsWindowFocusedWithWait(TestStack.White.UIItems.WindowItems.Window window, int timeout = 30)
        {
            bool result = false;
            for (int i = 0; i < timeout; i++)
            {
                result = window.IsFocussed;
                if (result)
                {
                    return result;
                }
                Thread.Sleep(TimeSpan.FromSeconds(1));
            }
            return result;
        }

        static DataTable GetDataTableFromCsv(string path, bool isFirstRowHeader)
        {
            string header = isFirstRowHeader ? "Yes" : "No";

            string pathOnly = Path.GetDirectoryName(path);
            string fileName = Path.GetFileName(path);

            string sql = @"SELECT * FROM [" + fileName + "]";

            using (OleDbConnection connection = new OleDbConnection(
                      @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathOnly +
                      ";Extended Properties=\"Text;HDR=" + header + "\""))
            using (OleDbCommand command = new OleDbCommand(sql, connection))
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
            {
                DataTable dataTable = new DataTable();
                dataTable.Locale = CultureInfo.CurrentCulture;
                adapter.Fill(dataTable);
                return dataTable;
            }
        }

        public static DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            StreamReader sr = new StreamReader(strFilePath);
            string[] headers = sr.ReadLine().Split(',');
            DataTable dt = new DataTable();
            foreach (string header in headers)
            {
                dt.Columns.Add(header.Replace("\"", ""));
            }
            while (!sr.EndOfStream)
            {
                string[] rows = Regex.Split(sr.ReadLine(), ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
                DataRow dr = dt.NewRow();
                for (int i = 0; i < headers.Length; i++)
                {
                    dr[i] = rows[i].Replace("\"", "");
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public void UpdateFINStatus(string csvFilePath, string FIN, string status)
        {
            List<string> ColumnData = new List<string>() { status };
            List<string> lines = File.ReadAllLines(csvFilePath).ToList();
            System.Text.RegularExpressions.Regex csvP = new System.Text.RegularExpressions.Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");

            lines.Skip(1).ToList().ForEach(line =>
            {
                if (line.Contains(FIN))
                {
                    var indx = lines.IndexOf(line);
                    System.Text.StringBuilder sb = new System.Text.StringBuilder(lines[indx]);

                    var s = csvP.Split(line);
                    s[30] = status;
                    string jstring = string.Join(",", s);
                    lines[indx] = jstring;
                }
            });

            File.WriteAllLines(csvFilePath, lines);

        }

        public UIItemContainer GetMdiChildWithWait(TestStack.White.UIItems.WindowItems.Window window, SearchCriteria searchCriteria, int timeout = 30)
        {
            UIItemContainer cvWindowPopup = null;
            for (int i = 0; i < timeout; i++)
            {
                cvWindowPopup = window.MdiChild(searchCriteria);

                if (cvWindowPopup != null)
                {
                    return cvWindowPopup;
                }
                Thread.Sleep(TimeSpan.FromSeconds(1));
            }
            return cvWindowPopup;
        }

        public bool IsFileLocked(string filePath, int secondsToWait = 3)
        {
            bool isLocked = true;
            int i = 0;

            while (isLocked && ((i < secondsToWait) || (secondsToWait == 0)))
            {
                try
                {
                    using (File.Open(filePath, FileMode.Open, FileAccess.ReadWrite)) { }
                    return false;
                }
                catch (IOException e)
                {
                    var errorCode = Marshal.GetHRForException(e) & ((1 << 16) - 1);
                    isLocked = errorCode == 32 || errorCode == 33;
                    i++;

                    if (secondsToWait != 0)
                        new System.Threading.ManualResetEvent(false).WaitOne(1000);
                }
            }

            return isLocked;
        }

    }
}