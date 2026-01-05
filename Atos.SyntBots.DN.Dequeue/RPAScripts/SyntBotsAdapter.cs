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
using SyntBotsDesktopBase.Utilities;
using SyntBotsDesktopBase.BaseScript;
using TestStack.White;
using SyntBotsUIUtility;
using TestStack.White.InputDevices;
using TestStack.White.UIItems.Finders;
using System.Threading;
using System.Linq;
using System.Configuration;
using System.IO;
using System.Text.RegularExpressions;
using TestStack.White.UIItems;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using Atos.SyntBots.Common.Entity;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices;

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
            driver = new SyntBotsUIUtil();
        }

        public override void AutomateScript()
        {
            #region AppSettings
            Configuration _configuration = ConfigurationManager.OpenExeConfiguration(this.GetType().Assembly.Location);
            string AccessHIMApplicationPath = _configuration.AppSettings.Settings["AccessHIM.ApplicationPath"] != null ? _configuration.AppSettings.Settings["AccessHIM.ApplicationPath"].Value : string.Empty;
            string AccessHIMUsername = _configuration.AppSettings.Settings["AccessHIM.Username"] != null ? _configuration.AppSettings.Settings["AccessHIM.Username"].Value : string.Empty;
            string AccessHIMPassword = _configuration.AppSettings.Settings["AccessHIM.Password"] != null ? _configuration.AppSettings.Settings["AccessHIM.Password"].Value : string.Empty;

            string revenueCycleApplicationPath = _configuration.AppSettings.Settings["RevenueCycle.ApplicationPath"] != null ? _configuration.AppSettings.Settings["RevenueCycle.ApplicationPath"].Value : string.Empty;
            string revenueCycleUsername = _configuration.AppSettings.Settings["RevenueCycle.Username"] != null ? _configuration.AppSettings.Settings["RevenueCycle.Username"].Value : string.Empty;
            string revenueCyclePassword = _configuration.AppSettings.Settings["RevenueCycle.Password"] != null ? _configuration.AppSettings.Settings["RevenueCycle.Password"].Value : string.Empty;

            string receiverEmail = _configuration.AppSettings.Settings["Email.ReceiverEmail"] != null ? _configuration.AppSettings.Settings["Email.ReceiverEmail"].Value : string.Empty; //"osman.alikhan@atos.net"; //ConfigurationManager.AppSettings["Email.SenderEmail"]; //
            string senderEmail = _configuration.AppSettings.Settings["Email.SenderEmail"] != null ? _configuration.AppSettings.Settings["Email.SenderEmail"].Value : string.Empty; //"osman.alikhan@atos.net"; //ConfigurationManager.AppSettings["Email.SenderEmail"]; //
            string senderPassword = _configuration.AppSettings.Settings["Email.SenderPassword"] != null ? _configuration.AppSettings.Settings["Email.SenderPassword"].Value : string.Empty; //ConfigurationManager.AppSettings["Email.Password"];//
            string host = _configuration.AppSettings.Settings["Email.Host"] != null ? _configuration.AppSettings.Settings["Email.Host"].Value : string.Empty;//"10.14.32.146";// ConfigurationManager.AppSettings["Email.Host"]; //"10.14.32.146";
            int port = _configuration.AppSettings.Settings["Email.Port"] != null ? Convert.ToInt32(_configuration.AppSettings.Settings["Email.Port"].Value) : 0;//25;//Convert.ToInt32(ConfigurationManager.AppSettings["Email.Port"]); //25;

            string rootFolderPath = _configuration.AppSettings.Settings["RootFolderPath"] != null ? _configuration.AppSettings.Settings["RootFolderPath"].Value : string.Empty;
            string jsonFilePath = _configuration.AppSettings.Settings["BotResponseJson.FilePath"] != null ? _configuration.AppSettings.Settings["BotResponseJson.FilePath"].Value : string.Empty; //@"E:\Syntbots\BotResponse.json";

            string ChargeCorrectionFileName = _configuration.AppSettings.Settings["ChargeCorrectionFileName"] != null ? _configuration.AppSettings.Settings["ChargeCorrectionFileName"].Value : string.Empty; //@"E:\TaskQueue\";
            string LogFilePath = _configuration.AppSettings.Settings["Logger.FilePath"] != null ? _configuration.AppSettings.Settings["Logger.FilePath"].Value : string.Empty; //@"E:\TaskQueue\";
            string botAgentName = _configuration.AppSettings.Settings["botAgentName"] != null ? _configuration.AppSettings.Settings["botAgentName"].Value : string.Empty;

            string debugModeValue = _configuration.AppSettings.Settings["DebugMode"] != null ? _configuration.AppSettings.Settings["DebugMode"].Value : string.Empty; //@"E:\TaskQueue\";

            bool debugMode = debugModeValue == "True" ? true : false;
            bool isDaylightSave = TimeZoneInfo.Local.IsDaylightSavingTime(DateTime.Now);
            #endregion AppSettings

            JsonHelper objJsnHlpr = JsonHelper.Instance;
            var _AccessHIMLauncher = new Atos.SyntBots.AccessHIM.Launcher();
            var _RevenueCycleLauncher = new Atos.SyntBots.RevenueCycle.Launcher();

            try
            {
                if (!IsFileLocked(jsonFilePath))
                {
                    string jsonString = File.ReadAllText(jsonFilePath);
                    JObject jObject = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonString) as JObject;

                    var jBotExecution = jObject.SelectToken("ExecutionStatusBot" + botAgentName);
                    jBotExecution.Replace("Started");

                    string updatedJsonString = jObject.ToString();
                    File.WriteAllText(jsonFilePath, updatedJsonString);
                }
                else
                {
                    throw new Exception("An error occurred while updating the json file.");
                }

                Atos.SyntBots.Common.Logger.Log("Starting the DN Dequeue process", LogFilePath);

                string facilityChargeCorrectionFilepath = string.Empty;

                DirectoryInfo rdirInfo = new DirectoryInfo(rootFolderPath);

                FileInfo[] facilityChargeCorrectionFiles = rdirInfo.GetFiles(string.Format("*{0}{1}*", ChargeCorrectionFileName, botAgentName), SearchOption.TopDirectoryOnly);

                if (facilityChargeCorrectionFiles.Count() > 0)
                {
                    facilityChargeCorrectionFilepath = facilityChargeCorrectionFiles[0].FullName;
                }
                FileInfo ChargeCorrectionFileInfo = new FileInfo(facilityChargeCorrectionFilepath);

                if (facilityChargeCorrectionFiles.Count() > 0)
                {
                    Atos.SyntBots.Common.Logger.Log("Processing the charge correction file: " + facilityChargeCorrectionFilepath, LogFilePath);
                    List<Order> FINlist = new List<Order>();

                    #region Read the FINs from the Charge correction file
                    Atos.SyntBots.Common.Logger.Log("Started the reading charge correction file", LogFilePath);
                    try
                    {
                        FINlist = File.ReadLines(facilityChargeCorrectionFilepath).Skip(1).Select(line =>
                        {
                            Order ch = new Order();
                            Regex CSVParser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");
                            String[] split = CSVParser.Split(line);
                            ch.FIN = split[0].Replace("\"", "");
                            return ch;
                        }).ToList<Atos.SyntBots.Common.Entity.Order>();

                        ChargeCorrectionFileInfo = new FileInfo(facilityChargeCorrectionFilepath);

                    }
                    catch (Exception ex)
                    {
                        throw new Exception("An error occure while reading the FINs from charge correction file. Error :" + ex.Message);
                    }
                    Atos.SyntBots.Common.Logger.Log("Completed reading charge correction file", LogFilePath);
                    #endregion Read the FINs from the Charge correction file

                    #region Launch RevenueCycle
                    Atos.SyntBots.Common.Logger.Log("START - Launching Revenue Cycle", LogFilePath);
                    _RevenueCycleLauncher.LaunchRevenueCycle(revenueCycleApplicationPath, revenueCycleUsername, revenueCyclePassword);
                    var _revenueCycleWindow = GetWindowWithWait("Revenue Cycle", 180);
                    _revenueCycleWindow.SetForeground();
                    Atos.SyntBots.Common.Logger.Log("END - Launching Revenue Cycle", LogFilePath);
                    #endregion Launch RevenueCycle

                    #region Launch AccessHIM
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
                    _AccessHIMLauncher.LaunchAccessHIM(AccessHIMApplicationPath, AccessHIMUsername, AccessHIMPassword);
                    var _accessHIMWindow = GetWindowWithWait("AccessHIM", 120);
                    //_accessHIMWindow.SetForeground();
                    Atos.SyntBots.Common.Logger.Log("END - Launching AccessHIM", LogFilePath);
                    #endregion Launch AccessHIM

                    #region Apply Filters OnHold
                    try
                    {

                        System.Threading.Thread.Sleep(500);
                        _accessHIMWindow.SetForeground();

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
                            if ((String.Compare(row.Name, "OnHold", true) == 0))
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
                        AutomationElement AE6 = driver.GetElmtByWlker(popup.AutomationElement, "Apply and Close", null, null, "button", null);
                        driver.AEClick(AE6);

                    }
                    catch (Exception ex)
                    {

                        throw new Exception("An error occure while applying filters. Error :" + ex.Message);
                    }

                    #endregion Apply Filters OnHold

                    string FIN = string.Empty;
                    bool isStatusCompleted = false;
                    int totalRecords = 0;
                    int totalErrorRecords = 0;
                    bool isAccessHIMError = false;
                    foreach (var Encounter in FINlist)
                    {
                        try
                        {
                            FIN = Encounter.FIN;

                            Atos.SyntBots.Common.Logger.Log("***START - Processing fin:" + FIN, LogFilePath);

                            #region Search FIN number in RevenueCycle
                            Atos.SyntBots.Common.Logger.Log("START - Search FIN in Revenue Cycle", LogFilePath);
                            try
                            {
                                _revenueCycleWindow = GetWindowWithWait("Revenue Cycle", 180);
                                _revenueCycleWindow.SetForeground();
                                Thread.Sleep(1000);

                                Keyboard.Instance.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                                Keyboard.Instance.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                                Keyboard.Instance.Enter("f");
                                Keyboard.Instance.Enter("o");
                                Thread.Sleep(1000);
                                //var popup = _revenueCycleWindow.MdiChild(SearchCriteria.ByText("Person Search"));
                                var popup = GetMdiChildWithWait(_revenueCycleWindow, SearchCriteria.ByText("Person Search"), 10);


                                TestStack.White.UIItems.TextBox btnQuery = _revenueCycleWindow.Get<TestStack.White.UIItems.TextBox>(SearchCriteria.ByText("Encounter"));
                                btnQuery.Click();
                                Keyboard.Instance.Enter(FIN);

                                //var btnSearch = _revenueCycleWindow.Get<TestStack.White.UIItems.Button>(SearchCriteria.ByText("Search"));
                                var btnSearch = _revenueCycleWindow.Get(SearchCriteria.ByText("Search"));
                                btnSearch.Click();

                                System.Threading.Thread.Sleep(1000);
                                var btnSelect = _revenueCycleWindow.Get(SearchCriteria.ByText("Select"));
                                btnSelect.Click();

                            }
                            catch (Exception ex)
                            {

                                throw new Exception("An error occured seaching the FIN:" + FIN + ". Error message:" + ex.Message + " Inner exception: " + ex.InnerException); ;
                            }
                            Atos.SyntBots.Common.Logger.Log("END - Search FIN in Revenue Cycle", LogFilePath);
                            #endregion Search FIN number in RevenueCycle

                            #region Navigate to Patient Account
                            Atos.SyntBots.Common.Logger.Log("START - Navigate to Patient Account", LogFilePath);
                            try
                            {
                                _revenueCycleWindow = GetWindowWithWait("Revenue Cycle", 180);
                                _revenueCycleWindow.SetForeground();
                                System.Threading.Thread.Sleep(3000);
                                var btnPatientAccount = _revenueCycleWindow.Get(SearchCriteria.ByText("Patient Account"));
                                System.Windows.Rect r = btnPatientAccount.Bounds;
                                System.Windows.Point pp = r.TopLeft;
                                driver.MouseClick(pp);

                            }
                            catch (Exception ex)
                            {
                                throw new Exception("An error occured while navigating to Patient Account FIN:" + FIN + ". Error message:" + ex.Message + " Inner exception: " + ex.InnerException); ;
                            }
                            Atos.SyntBots.Common.Logger.Log("END - Navigate to Patient Account", LogFilePath);
                            #endregion Navigate to Patient Account

                            #region Navigate to Workflow tab
                            Atos.SyntBots.Common.Logger.Log("START - Navigate to Workflow tab", LogFilePath);
                            try
                            {
                                Thread.Sleep(2000);
                                _revenueCycleWindow = GetWindowWithWait("Revenue Cycle", 180);
                                _revenueCycleWindow.SetForeground();
                                System.Threading.Thread.Sleep(2000);
                                _revenueCycleWindow.SetForeground();
                                Keyboard.Instance.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                                Keyboard.Instance.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                                Keyboard.Instance.Enter("v");
                                Keyboard.Instance.Enter("o");
                                Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RIGHT);
                                Keyboard.Instance.Enter("w");


                            }
                            catch (Exception ex)
                            {

                                throw new Exception("An error occured while navigating to Workflow tab FIN:" + FIN + ". Error message:" + ex.Message + " Inner exception: " + ex.InnerException); ;
                            }
                            Atos.SyntBots.Common.Logger.Log("END - Navigate to Workflow tab", LogFilePath);
                            #endregion Navigate to Workflow tab

                            #region Apply Action Code 'A133'
                            Atos.SyntBots.Common.Logger.Log("START - Apply Action code 'A133'", LogFilePath);
                            try
                            {
                                AutomationElement tabWorkflow = driver.GetElmtByWlker(_revenueCycleWindow.AutomationElement, "Workflow", null, null, "tab", null);
                                Atos.SyntBots.Common.Logger.Log("L:1", LogFilePath);
                                Condition conditions = new AndCondition(new PropertyCondition(AutomationElement.IsEnabledProperty, true), new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));
                                bool isStatusVisible = false;
                                var children = tabWorkflow.FindAll(TreeScope.Element | TreeScope.Descendants, conditions);
                                foreach (var child in children)
                                {
                                    var item = child as AutomationElement;
                                    if (item.Current.Name == "Status")
                                    {
                                        isStatusVisible = true;
                                        break;
                                    }
                                }

                                if (!isStatusVisible)
                                {

                                    #region Region modified because of Controls name changed for Status tab in 27-02-2023
                                    //Atos.SyntBots.Common.Logger.Log("L:2", LogFilePath);
                                    //AutomationElement _openMenuBtn = driver.GetElmtByWlker(tabWorkflow, "Open Menu", null, null, "button", null);
                                    //driver.AEClick(_openMenuBtn);
                                    //Atos.SyntBots.Common.Logger.Log("L:3", LogFilePath);
                                    //_revenueCycleWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("Cycle"));
                                    //AutomationElement _dialog = driver.GetElmtByWlker(_revenueCycleWindow.AutomationElement, "", null, null, "Dialog", null);
                                    //AutomationElement _statusDropdownItem = driver.GetElmtByWlker(_dialog, "Status", null, null, "item", null);
                                    ////AutomationElement _statusDropdownItem = driver.GetElmtByWlker(_dialog, "All", null, null, "text", null);
                                    //driver.AEInvoke(_statusDropdownItem);
                                    ////driver.AEClick(_statusDropdownItem);
                                    //Atos.SyntBots.Common.Logger.Log("L:4", LogFilePath);
                                    //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.SPACE); //Selecting the Status item
                                    //driver.AEClick(_openMenuBtn);
                                    //Atos.SyntBots.Common.Logger.Log("L:5", LogFilePath);



                                    AutomationElement _filtersCombo = driver.GetElmtByWlker(tabWorkflow, "Filters", null, null, "combo box", null);
                                    driver.AEEnterText(_filtersCombo, "All");
                                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
                                    Thread.Sleep(1000);

                                    driver.AEEnterText(_filtersCombo, "All");
                                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
                                    Thread.Sleep(1000);
                                    #endregion

                                }



                                _revenueCycleWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("Cycle"));

                                //Modified because of Controls name changed for Status tab in 27-02-2023

                                //AutomationElement _status = driver.GetElmtByWlker(tabWorkflow, "Status", null, null, "text", null);

                                //AutomationElement _statusPane = TreeWalker.ControlViewWalker.GetParent(_status);
                                //AutomationElement _statusEditPane = TreeWalker.ControlViewWalker.GetNextSibling(_statusPane);
                                //conditions = new AndCondition(new PropertyCondition(AutomationElement.IsEnabledProperty, true), new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit)); ;
                                //var _editElements = _statusEditPane.FindAll(TreeScope.Element | TreeScope.Descendants, conditions);
                                //AutomationElement editHimCodingErrors = _editElements[0];
                                //Atos.SyntBots.Common.Logger.Log("L:6", LogFilePath);
                                //driver.AEClick(editHimCodingErrors);

                                AutomationElement _status = driver.GetElmtByWlker(tabWorkflow, "", "1001", "Edit", "edit", null);

                                driver.AEClick(_status);
                                Keyboard.Instance.Enter("HIM Coding Errors");
                                Atos.SyntBots.Common.Logger.Log("L:7", LogFilePath);
                                Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
                                Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);

                                tabWorkflow = driver.GetElmtByWlker(_revenueCycleWindow.AutomationElement, "Workflow", null, null, "tab", null);
                                AutomationElement applyBtn = driver.GetElmtByWlker(tabWorkflow, "Apply", null, "Button", "button", null, -1, 1);
                                driver.AEClick(applyBtn);
                                Atos.SyntBots.Common.Logger.Log("L:8", LogFilePath);
                                Thread.Sleep(1000);
                                var _workflowGrid = _revenueCycleWindow.GetMultiple<TestStack.White.UIItems.TreeItems.Tree>(SearchCriteria.ByControlType(ControlType.Tree));
                                Thread.Sleep(2000);
                                Atos.SyntBots.Common.Logger.Log("L:8.1"+ (_workflowGrid==null), LogFilePath);
                                try
                                {
                                    Atos.SyntBots.Common.Logger.Log("Work flow grid count : " + _workflowGrid.Count(), LogFilePath);
                                    Thread.Sleep(2000);
                                }
                                catch(Exception ex)
                                {
                                     _workflowGrid = _revenueCycleWindow.GetMultiple<TestStack.White.UIItems.TreeItems.Tree>(SearchCriteria.ByControlType(ControlType.Tree));
                                }
                                if (_workflowGrid.Count() > 0)
                                {
                                    Atos.SyntBots.Common.Logger.Log("L:9", LogFilePath);
                                    var _rowHimCodingErrors = driver.GetRawChildrenByWlker(_workflowGrid[0].AutomationElement)[0];
                                    Atos.SyntBots.Common.Logger.Log("L:10", LogFilePath);
                                    driver.AERightClick(_rowHimCodingErrors);
                                    //updated by Velan for Dequeue fix after RHO go live , 25-05-2022
                                    Thread.Sleep(1000);
                                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
                                    Thread.Sleep(1000);
                                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
                                    Thread.Sleep(1000);
                                    Atos.SyntBots.Common.Logger.Log("L:11", LogFilePath);
                                    //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
                                    //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
                                }

                                Atos.SyntBots.Common.Logger.Log("else", LogFilePath);
                                _revenueCycleWindow = GetWindowWithWait("Revenue Cycle", 180);
                                //_revenueCycleWindow.SetForeground();
                                //var _applyActioncodeDialog = driver.GetRawChildrenByWlker(_revenueCycleWindow.AutomationElement)[0];
                                var _applyActioncodeDialog = driver.GetElmtByWlker(_revenueCycleWindow.AutomationElement, "Apply Action Code", null, null, "Dialog", null, -1, 5);
                                _applyActioncodeDialog.SetFocus();
                                Thread.Sleep(1000);

                                //Atos.SyntBots.Common.Logger.Log("L:12", LogFilePath);
                                if (_applyActioncodeDialog != null)
                                {
                                    Atos.SyntBots.Common.Logger.Log("L:13", LogFilePath);
                                    Atos.SyntBots.Common.Logger.Log("L:14", LogFilePath);
                                    AutomationElement _actionCodeText = driver.GetElmtByWlker(_applyActioncodeDialog, "Action Code", null, null, "combo box", null, -1, 5);
                                    Atos.SyntBots.Common.Logger.Log("L:15", LogFilePath);
                                    Thread.Sleep(1000);
                                    _applyActioncodeDialog.SetFocus();
                                    Thread.Sleep(1000);
                                    Keyboard.Instance.Enter("A133");
                                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
                                    Thread.Sleep(2000);
                                    Atos.SyntBots.Common.Logger.Log("L:16", LogFilePath);
                                    _revenueCycleWindow = GetWindowWithWait("Revenue Cycle", 180);
                                    _revenueCycleWindow.SetForeground();
                                    Thread.Sleep(500);
                                    _applyActioncodeDialog = driver.GetElmtByWlker(_revenueCycleWindow.AutomationElement, "Apply Action Code", null, null, "Dialog", null, -1, 5);
                                    Thread.Sleep(1000);
                                    //Atos.SyntBots.Common.Logger.Log("AE.Name:" + _applyActioncodeDialog.Current.Name, LogFilePath);
                                    //Atos.SyntBots.Common.Logger.Log("AE.ID:" + _applyActioncodeDialog.Current.AutomationId, LogFilePath);
                                    //Atos.SyntBots.Common.Logger.Log("AE.ControlType:" + _applyActioncodeDialog.Current.ControlType, LogFilePath);

                                    AutomationElement _actionCodeOk = driver.GetElmtByWlker(_applyActioncodeDialog, "OK", null, null, "button", null, -1, 3);
                                    Thread.Sleep(2000);
                                    
                                    Atos.SyntBots.Common.Logger.Log("L:17", LogFilePath);
                                    try
                                    {
                                        driver.AEClick(_actionCodeOk);
                                    }
                                    catch(Exception ex)
                                    {
                                        Atos.SyntBots.Common.Logger.Log("Error at line number 439. _actionCodeOk : "+(_actionCodeOk==null)+" Exception : "+ex.Message, LogFilePath);
                                    }

                                    _applyActioncodeDialog = driver.GetElmtByWlker(_revenueCycleWindow.AutomationElement, "Apply Action Code", null, null, "Dialog", null, -1, 2);
                                    if (_applyActioncodeDialog != null)
                                    {
                                        AutomationElement _actionCodeClose = driver.GetElmtByWlker(_applyActioncodeDialog, "Close", null, null, "button", null, -1, 2);
                                        if (_actionCodeClose != null)
                                        {
                                            Atos.SyntBots.Common.Logger.Log("closing the popup", LogFilePath);
                                            _applyActioncodeDialog.SetFocus();
                                            Thread.Sleep(1000);
                                            try
                                            {
                                                driver.AEClick(_actionCodeClose);
                                            }
                                            catch(Exception ex)
                                            {
                                                Atos.SyntBots.Common.Logger.Log("Error at line number 457. _actionCodeClose : " + (_actionCodeClose == null) + " Exception : " + ex.Message, LogFilePath);
                                            }
                                           Atos.SyntBots.Common.Logger.Log("L:21", LogFilePath);
                                        }
                                    }

                                    isStatusCompleted = true;
                                }

                                //Atos.SyntBots.Common.Logger.Log("L:18", LogFilePath);
                            }
                            catch (Exception ex)
                            {
                                //Atos.SyntBots.Common.Logger.Log("L:19", LogFilePath);
                                _revenueCycleWindow = GetWindowWithWait("Revenue Cycle", 180);
                                var _applyActioncodeDialog = driver.GetElmtByWlker(_revenueCycleWindow.AutomationElement, "Apply Action Code", null, null, "Dialog", null, -1, 5);
                                _applyActioncodeDialog.SetFocus();
                                if (_applyActioncodeDialog != null)
                                {
                                    //Atos.SyntBots.Common.Logger.Log("AE.Name:" + _applyActioncodeDialog.Current.Name, LogFilePath);
                                    //Atos.SyntBots.Common.Logger.Log("AE.ID:" + _applyActioncodeDialog.Current.AutomationId, LogFilePath);
                                    //Atos.SyntBots.Common.Logger.Log("AE.ControlType:" + _applyActioncodeDialog.Current.ControlType, LogFilePath);

                                    //Atos.SyntBots.Common.Logger.Log("L:20", LogFilePath);
                                    Atos.SyntBots.Common.Logger.Log("Closing the popup", LogFilePath);
                                    AutomationElement _actionCodeClose = driver.GetElmtByWlker(_applyActioncodeDialog, "Close", null, null, "button", null, -1, 3);
                                    if (_actionCodeClose != null)
                                    {
                                        driver.AEClick(_actionCodeClose);
                                        Atos.SyntBots.Common.Logger.Log("L:21", LogFilePath);
                                    }

                                }

                                throw new Exception("An error occured while applying Action code  FIN:" + FIN + ". Error message:" + ex.Message + " Inner exception: " + ex.InnerException); ;
                            }
                            Atos.SyntBots.Common.Logger.Log("END - Apply Action code 'A133'", LogFilePath);
                            #endregion Apply Action Code 'A133'

                            #region Change the Status to 'Charge Correction Complete'
                            Atos.SyntBots.Common.Logger.Log("START - Change the status to 'Charge Correction Complete'", LogFilePath);
                            if (isStatusCompleted)
                            {
                                try
                                {
                                    _accessHIMWindow = GetWindowWithWait("AccessHIM", 15);
                                    _accessHIMWindow.SetForeground();
                                    Thread.Sleep(500);
                                    var _grid = _accessHIMWindow.GetMultiple<TestStack.White.UIItems.ListView>(SearchCriteria.ByControlType(ControlType.DataGrid));
                                    TreeWalker tWalker = TreeWalker.ControlViewWalker;
                                    AutomationElement fin = _grid[0].GetElement(SearchCriteria.ByText(FIN));
                                    AutomationElement Row = tWalker.GetParent(fin);
                                    SelectionItemPattern pattern1 = (SelectionItemPattern)(BasePattern)Row.GetCurrentPattern(SelectionItemPattern.Pattern);
                                    pattern1.Select();
                                    _grid[0].KeyIn(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
                                    _grid[0].KeyIn(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.UP);


                                }
                                catch (Exception exec)
                                {
                                    isAccessHIMError = true;
                                    throw new Exception("An error occured while selecting the FIN:" + FIN + ". Error message:" + exec.Message + " Inner exception: " + exec.InnerException);
                                }

                                try
                                {
                                    AutomationElement Btn = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Pass", null, null, "button", null, -1, 15);
                                    driver.AEClick(Btn);
                                    AutomationElement aePassReasonDialog = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Pass on Task", null, null, "Dialog", null, -1, 15);
                                    AutomationElement aePassReasonTxt = driver.GetElmtByWlker(aePassReasonDialog, "Reason", null, null, "edit", null, -1, 15);

                                    //driver.AESetText(aePassReasonTxt, "Coding - Charge Correction Complete");
                                    driver.AEClick(aePassReasonTxt);
                                    Keyboard.Instance.Enter("Coding - Charge Correction Complete");

                                    AutomationElement aePassDropDown = driver.GetElmtByWlker(aePassReasonDialog, "Drop Down Button", null, null, "button", null, -1, 15);
                                    driver.AEClick(aePassDropDown);
                                    driver.AEClick(aePassDropDown);
                                    AutomationElement aePassOkBtn = driver.GetElmtByWlker(aePassReasonDialog, "OK", null, null, "button", null);
                                    driver.AEClick(aePassOkBtn);
                                    Atos.SyntBots.Common.Logger.Log("Status changed : Coding - Charge correction complete", LogFilePath);

                                }
                                catch (Exception exception)
                                {
                                    isAccessHIMError = true;
                                    throw new Exception("An error occured while while changing the status FIN:" + FIN + ". Error message:" + exception.Message + " Inner exception: " + exception.InnerException);
                                }


                            }
                            Atos.SyntBots.Common.Logger.Log("END - Change the status to 'Charge Correction Complete'", LogFilePath);
                            #endregion Change the Status to 'Charge Corretion Complete'                    

                            UpdateFINStatus(facilityChargeCorrectionFilepath, FIN, "Dequeued");
                            totalRecords++;
                            Atos.SyntBots.Common.Logger.Log("***END - Processing fin:" + FIN, LogFilePath);

                            _accessHIMWindow.SetForeground();
                            Thread.Sleep(500);
                            Keyboard.Instance.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                            Keyboard.Instance.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                            Keyboard.Instance.Enter("f");
                            Keyboard.Instance.Enter("c");

                        }
                        catch (Exception ex)
                        {
                            Keyboard.Instance.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                            Keyboard.Instance.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                            Keyboard.Instance.Enter("f");
                            Keyboard.Instance.Enter("c");
                            string errorStatus = isAccessHIMError ? "Failed to process - AccessHIM" : "Failed to process - RevCycle";
                            UpdateFINStatus(facilityChargeCorrectionFilepath, FIN, errorStatus);
                            Atos.SyntBots.Common.Logger.Log(ex.Message, LogFilePath);
                            totalErrorRecords++;
                        }


                    }

                    _accessHIMWindow.Close();
                    _revenueCycleWindow.Close();

                }
                else
                {
                    Atos.SyntBots.Common.Logger.Log("Facility Charge correction file does not exists. File name: " + facilityChargeCorrectionFilepath, LogFilePath);
                    throw new Exception(facilityChargeCorrectionFilepath + " file not found");
                }

                Atos.SyntBots.Common.Logger.Log("Completed the DN Dequeue process", LogFilePath);

                if (!IsFileLocked(jsonFilePath))
                {
                    string jsonString2 = File.ReadAllText(jsonFilePath);
                    JObject jObject2 = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonString2) as JObject;

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

                var jBotExecution = jObject.SelectToken("ExecutionStatusBot" + botAgentName);
                jBotExecution.Replace("Failed");

                string updatedJsonString = jObject.ToString();
                File.WriteAllText(jsonFilePath, updatedJsonString);

                try
                {
                    Atos.SyntBots.Common.EmailHelper email = new Atos.SyntBots.Common.EmailHelper(senderEmail, senderPassword, host, port);
                    string automationTeamEmail = _configuration.AppSettings.Settings["Email.AutomationTeam"] != null ? _configuration.AppSettings.Settings["Email.AutomationTeam"].Value : string.Empty;
                    email.SendEMail(automationTeamEmail, "ALERT : Dequeue Process Execution Failed - Bot Agent # " + botAgentName, "Bot Execution Failed. Error: " + ex.Message + ". InnerExeception: " + ex.InnerException, "");

                }
                catch (Exception exp)
                {

                    Atos.SyntBots.Common.Logger.Log("An error occured while sending email." + "Error: " + exp.Message);
                }

            }
            finally
            {

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
                    s[8] = status;
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