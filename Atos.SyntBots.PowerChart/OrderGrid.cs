/*
 * Created by SharpDevelop.
 * User: v.chakradhar.ark
 * Date: 11/3/2019
 * Time: 8:13 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Windows.Automation;
using System.Collections.Generic;
using TestStack.White;
using SyntBotsUIUtility;
using TestStack.White.InputDevices;
using System.Linq;
using System.Windows.Forms;
using System.Data;
using Atos.SyntBots.Common.Entity;

namespace Atos.SyntBots.PowerChart
{
    /// <summary>
    /// Description of OrderGrid.
    /// </summary>
    public class OrderGrid
    {
        TestStack.White.UIItems.WindowItems.Window _mainWindow;
        SyntBotsUIUtil driver;

        public OrderGrid()
        {
            driver = new SyntBotsUIUtil();
        }


        public List<Atos.SyntBots.Common.Entity.Charge> GetLabOrders(DataTable ordersCatalogdata, string admitdate)
        {


            List<Atos.SyntBots.Common.Entity.Charge> listPowerChartOrders = new List<Atos.SyntBots.Common.Entity.Charge>();

            bool isLabRecord = false;

            ////System.Windows.Point ppp = new System.Windows.Point(1080,360);
            //System.Windows.Point ppp = new System.Windows.Point(1222,440);
            //driver.MouseClick(ppp);

            int rowToclick = 1;
            int i = 1;

            object patternObj;

            string prevDesc = string.Empty;
            string curDesc = string.Empty;

            string prevDatetime = string.Empty;
            string curDatetime = string.Empty;

            try
            {

                while (i <= rowToclick)
                {
                    System.Threading.Thread.Sleep(2000);
                    if (i > 1)
                    {

                        Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
                    }

                    Keyboard.Instance.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                    Keyboard.Instance.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                    Keyboard.Instance.Enter("u");
                    Keyboard.Instance.Enter("o");


                    _mainWindow = GetWindowWithWait("Opened by", 60); //Desktop.Instance.Windows().Find(x => x.Name.Contains("Opened by"));
                    //_mainWindow.SetForeground();

                    AutomationElement pcWindow = _mainWindow.AutomationElement;
                    //var popup = driver.GetRawChildrenByWlker(pcWindow)[0];
                    var popup = GetPopUpWindowWithWait(pcWindow, 0, 20);
                    if (popup.Current.LocalizedControlType != "Dialog")
                        break;

                    var popupPane1 = GetPopUpWindowWithWait(popup, 1, 20); //driver.GetRawChildrenByWlker(popup)[1];
                    var popupPane12 = GetPopUpWindowWithWait(popupPane1, 2, 20); //driver.GetRawChildrenByWlker(popupPane1)[2];
                    string deptartment = popupPane12.Current.Name;

                    AutomationElement tabItemdetails = driver.GetElmtByWlker(pcWindow, "Details", null, null, "tab item", null, -1, 20);
                    //System.Windows.Rect r = tabItemdetails.Current.BoundingRectangle;
                    //System.Windows.Point pp = r.TopLeft;
                    //driver.MouseClick(pp);

                    var popupPane = GetPopUpWindowWithWait(popup, 3, 20);//driver.GetRawChildrenByWlker(popup)[3];
                    //  AutomationElement AEDiagnosisText = driver.GetElmtByWlker(popupPane, "Diagnosis", null, null, "text", null);
                    // var AEDiagnosisVal = TreeWalker.ControlViewWalker.GetNextSibling(AEDiagnosisText);

                    AutomationElement tabItemAdditionalInfo = driver.GetElmtByWlker(pcWindow, "Additional Info", null, null, "tab item", null, -1, 20);
                    System.Windows.Rect r = tabItemAdditionalInfo.Current.BoundingRectangle;
                    System.Windows.Point pp = r.TopLeft;
                    driver.MouseClick(pp);
                    //System.Threading.Thread.Sleep(1000);
                    AutomationElement AEOrderedAsVal = driver.GetElmtByWlker(popupPane, "Ordered As", null, null, "document", null, -1, 20);
                    AutomationElement AEStartDateTimeVal = driver.GetElmtByWlker(popupPane, "Start Date/Time", null, null, "document", null, -1, 20);
                    AutomationElement AEStopDateTimeVal = driver.GetElmtByWlker(popupPane, "Stop Date/Time", null, null, "document", null, -1, 20);

                    string orderedAsVal = string.Empty;
                    string startDateTimeVal = string.Empty;
                    string stopDateTimeVal = string.Empty;

                    if (AEOrderedAsVal.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
                    {
                        var textPattern = (TextPattern)patternObj;
                        orderedAsVal = textPattern.DocumentRange.GetText(-1).TrimEnd('\r');
                        orderedAsVal = orderedAsVal.Trim();
                        if (orderedAsVal.Contains("Venipuncture") || orderedAsVal.Contains("venipuncture"))
                        {
                            orderedAsVal = "Venipuncture Charge (M)";
                        }

                    }
                    if (AEStartDateTimeVal.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
                    {
                        var textPattern = (TextPattern)patternObj;
                        startDateTimeVal = textPattern.DocumentRange.GetText(-1).TrimEnd('\r');
                        startDateTimeVal = startDateTimeVal.Trim();
                    }
                    if (AEStopDateTimeVal.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
                    {
                        var textPattern = (TextPattern)patternObj;
                        stopDateTimeVal = textPattern.DocumentRange.GetText(-1).TrimEnd('\r');
                        stopDateTimeVal = stopDateTimeVal.Trim();
                    }


                    curDesc = orderedAsVal;
                    curDatetime = startDateTimeVal;
                    if (curDesc == prevDesc && curDatetime == prevDatetime)
                    {
                        AutomationElement AE6 = driver.GetElmtByWlker(popup, "Close", null, null, "button", null, -1, 20);
                        driver.AEClick(AE6);
                        break;
                    }
                    else
                    {
                        string chargedate = DateTime.Parse(startDateTimeVal.Replace("EDT", "").Replace("EST", "").Trim()).ToString("MM/dd/yyyy");


                        prevDesc = orderedAsVal;
                        prevDatetime = startDateTimeVal;

                        if (chargedate == admitdate)
                        {
                            var result = new List<string>();
                            string cptCode = string.Empty;

                            result = FindOrderInCatalog(ordersCatalogdata, orderedAsVal);
                            if (result.Count == 1 && result[0] == "")
                            {
                                result = GetHCPCSByOrder(ordersCatalogdata, orderedAsVal);
                            }

                            if (result.Count == 0 || result == null)
                            {
                                string res = FindOrderInCatalogByDescription(ordersCatalogdata, orderedAsVal);

                                if (res != null)
                                {
                                    cptCode = "";
                                    result.Add(cptCode);
                                }
                                else if (string.IsNullOrEmpty(res))
                                {
                                    throw new Exception("Description not found - " + orderedAsVal);
                                }

                            }
                            if (result.Count == 1)
                            {
                                cptCode = result[0];
                            }
                            else if (result.Count > 1)
                            {

                                cptCode = "MultipleCPT";
                            }

                            if (result.Count > 0)
                            {

                                listPowerChartOrders.Add(new Atos.SyntBots.Common.Entity.Charge { ChargeDescription = orderedAsVal, ChargeDate = chargedate, CPT4 = cptCode });

                                #region Venipuncture

                                tabItemdetails = driver.GetElmtByWlker(pcWindow, "Comments", null, null, "tab item", null, -1, 20);
                                r = tabItemdetails.Current.BoundingRectangle;
                                pp = r.TopLeft;
                                driver.MouseClick(pp);
                                //System.Threading.Thread.Sleep(1000);
                                //popupPane = driver.GetRawChildrenByWlker(popup)[3];
                                popupPane = GetPopUpWindowWithWait(popup, 3, 20);

                                var venipunchureChargeItem = new Atos.SyntBots.Common.Entity.Charge();

                                AutomationElement AEOrderCmnt = driver.GetElmtByWlker(popupPane, null, null, null, "text", null, -1, 20);
                                if (!AEOrderCmnt.Current.Name.Contains("No comments exist for this order."))
                                {
                                    //Atos.SyntBots.Common.Logger.Log("Inside Venipuncture");
                                    AutomationElement pane = null;
                                    //object patternObj;
                                    pane = TreeWalker.ControlViewWalker.GetNextSibling(AEOrderCmnt);

                                    string venipunctureInfo = string.Empty;
                                    var paneChildren = driver.GetRawChildrenByWlker(pane);
                                    AutomationElement AEdocumentCntrl = paneChildren[1];

                                    if (AEdocumentCntrl.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
                                    {
                                        var textPattern = (TextPattern)patternObj;
                                        venipunctureInfo = textPattern.DocumentRange.GetText(-1).TrimEnd('\r');
                                    }
                                    if (venipunctureInfo.Contains("Venipuncture") || venipunctureInfo.Contains("venipuncture"))
                                    {
                                        DateTime dos = Convert.ToDateTime(paneChildren[0].Current.Name.Replace("EDT", "|").Replace("EST", "|").Split('|')[0].Trim());
                                        venipunchureChargeItem = new Atos.SyntBots.Common.Entity.Charge { ChargeDescription = "Venipuncture Charge (M)", ChargeDate = dos.ToString("MM/dd/yyyy"), CPT4 = "36415" };

                                        var catalog_descriptionList = (from Atos.SyntBots.Common.Entity.Charge ChargeItem in listPowerChartOrders
                                                                       where ChargeItem.ChargeDescription.ToString() == "Venipuncture Charge (M)"
                                                                       select (string)ChargeItem.ChargeDescription).ToList();

                                        if (catalog_descriptionList.Count < 1)
                                        {
                                            listPowerChartOrders.Add(venipunchureChargeItem);
                                        }
                                    }


                                    //Atos.SyntBots.Common.Logger.Log("end Venipuncture");
                                }

                                #endregion
                            }

                        }

                        i++;
                        rowToclick++;

                        AutomationElement AE6 = driver.GetElmtByWlker(popup, "Close", null, null, "button", null, -1, 20);
                        driver.AEClick(AE6);

                    }
                }
            }
            catch (Exception ex)
            {

                //var popup = driver.GetRawChildrenByWlker(_mainWindow.AutomationElement)[0];
                var popup = GetPopUpWindowWithWait(_mainWindow.AutomationElement, 0, 20);
                AutomationElement AE6 = driver.GetElmtByWlker(popup, "Close", null, null, "button", null, -1, 20);
                driver.AEClick(AE6);
                throw new Exception("An error ocurrend while reading the order grid popup data. Error: " + ex.Message + ". Inner Exception: " + ex.InnerException);
            }

            return listPowerChartOrders;

        }

        public List<Order> GetLabOrders(List<Order> ordersByFIN, System.Data.DataTable ordersCatalogdata, string admitdate, string[] excludedCPTs)
        {
            List<Order> listPowerChartOrders = new List<Order>();

            bool isLabRecord = false;
            try
            {

                //while (i <= rowToclick)
                foreach (var order in ordersByFIN)
                {

                    string orderedAsVal = order.OrderDescription;
                    string startDateTimeVal = order.OrderDate;
                    string resultDateTimeVal = order.ResultDate;

                    if (orderedAsVal.Contains("Venipuncture") || orderedAsVal.Contains("venipuncture"))
                    {
                        orderedAsVal = "Venipuncture Charge (M)";
                    }
                    //string chargedate = DateTime.Parse(startDateTimeVal.Replace("EDT", "").Replace("EST", "").Trim()).ToString("MM/dd/yyyy");
                    string chargedate = string.IsNullOrEmpty(startDateTimeVal) ? startDateTimeVal : DateTime.Parse(startDateTimeVal).ToString("MM/dd/yyyy");
                    string resultdate = string.IsNullOrEmpty(resultDateTimeVal) ? resultDateTimeVal : DateTime.Parse(resultDateTimeVal).ToString("MM/dd/yyyy");

                    //Modified on 30-03-2023 for In Process orders inclusion

                    //if ((chargedate == admitdate || resultdate == admitdate) && order.OrderStatus.ToLower() == "completed")
                    if ((chargedate == admitdate || resultdate == admitdate) && (order.OrderStatus.ToLower() == "completed" || order.OrderStatus.ToLower() == "inprocess"))
                    {
                        var result = new List<string>();
                        string cptCode = string.Empty;

                        result = FindOrderInCatalog(ordersCatalogdata, orderedAsVal);
                        if (result.Count == 1 && result[0] == "")
                        {
                            result = GetHCPCSByOrder(ordersCatalogdata, orderedAsVal);
                        }
                        if (result.Count == 0 || result == null)
                        {
                            string res = FindOrderInCatalogByDescription(ordersCatalogdata, orderedAsVal);

                            if (res != null)
                            {
                                cptCode = "";
                                result.Add(cptCode);
                            }
                            else if (string.IsNullOrEmpty(res))
                            {
                                throw new Exception("Description not found - " + orderedAsVal);
                            }
                        }
                        if (result.Count == 1)
                        {
                            cptCode = result[0];
                        }
                        else if (result.Count > 1)
                        {
                            cptCode = "MultipleCPT";
                        }

                        bool isExlucdedCPT = false;
                        for (int i = 0; i < result.Count(); i++)
                        {
                            var excluded = excludedCPTs.Where(s => s.Trim().Contains(result[i].Trim())).FirstOrDefault();
                            if (!string.IsNullOrEmpty(excluded))
                            {
                                isExlucdedCPT = true;
                                throw new Exception("Excluded CPT - " + orderedAsVal + "-" + result[i]);
                            }
                        }


                        if (result.Count > 0)
                        {
                            listPowerChartOrders.Add(new Order { OrderDescription = orderedAsVal, OrderDate = chargedate, CPT4 = cptCode, ResultDate = resultdate });
                            string commentsInfo = order.Comments;
                            #region Venipuncture
                            var venipunchureChargeItem = new Order();
                            var COVID19Item = new Order();
                            if (commentsInfo.Contains("Venipuncture") || commentsInfo.Contains("venipuncture"))
                            {
                                //DateTime dos = Convert.ToDateTime(venipunctureInfo.Split(new string[] { "at" }, StringSplitOptions.None)[0].Replace("EDT", "|").Replace("EST", "|").Split('|')[0].Trim());
                                //var dosString = venipunctureInfo.Split(new string[] { "at" }, StringSplitOptions.None)[1];
                                //var dos = Convert.ToDateTime(dosString.Substring(0, 5) + "-" + dosString.Substring(5, 2) + "-" + dosString.Substring(7, 2)).ToString("MM/dd/yyyy");
                                venipunchureChargeItem = new Order { OrderDescription = "Venipuncture Charge (M)", CPT4 = "36415" };

                                var catalog_descriptionList = (from Atos.SyntBots.Common.Entity.Order ChargeItem in listPowerChartOrders
                                                               where ChargeItem.OrderDescription.ToString() == "Venipuncture Charge (M)"
                                                               select (string)ChargeItem.OrderDescription).ToList();

                                if (catalog_descriptionList.Count < 1)
                                {
                                    listPowerChartOrders.Add(venipunchureChargeItem);
                                }
                                //Atos.SyntBots.Common.Logger.Log("end Venipuncture");
                            }
                            #endregion
                            #region COVID19 Outpatient Specimen Collection (M)
                            if (commentsInfo.ToLower().Contains("covid19 outpatient specimen collection"))
                            {
                                COVID19Item = new Order { OrderDescription = "COVID19 Outpatient Specimen Collection (M)", CPT4 = "C9803" };
                                var catalog_descriptionList = (from Atos.SyntBots.Common.Entity.Order ChargeItem in listPowerChartOrders
                                                               where ChargeItem.OrderDescription.ToString() == "COVID19 Outpatient Specimen Collection (M)"
                                                               select (string)ChargeItem.OrderDescription).ToList();
                                if (catalog_descriptionList.Count < 1)
                                {
                                    listPowerChartOrders.Add(COVID19Item);
                                }
                            }
                            #endregion

                            #region Capilary
                            if (commentsInfo.Contains("Capillary") || commentsInfo.Contains("capillary"))
                            {
                                COVID19Item = new Order { OrderDescription = "Capillary Draw Charge (M)", CPT4 = "36416" };
                                var catalog_descriptionList = (from Atos.SyntBots.Common.Entity.Order ChargeItem in listPowerChartOrders
                                                               where ChargeItem.OrderDescription.ToString() == "Capillary Draw Charge (M)"
                                                               select (string)ChargeItem.OrderDescription).ToList();
                                if (catalog_descriptionList.Count < 1)
                                {
                                    listPowerChartOrders.Add(COVID19Item);
                                }
                            }
                            #endregion
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                //var popup = driver.GetRawChildrenByWlker(_mainWindow.AutomationElement)[0];
                //var popup = GetPopUpWindowWithWait(_mainWindow.AutomationElement, 0, 20);
                //AutomationElement AE6 = driver.GetElmtByWlker(popup, "Close", null, null, "button", null, -1, 20);
                //driver.AEClick(AE6);
                throw new Exception("An error ocurrend while reading the order grid popup data. Error: " + ex.Message + ". Inner Exception: " + ex.InnerException);
            }
            return listPowerChartOrders;
        }
        public bool HasLabOrders(DataTable ordersCatalogdata)
        {
            bool isLabRecord = false;

            Atos.SyntBots.PowerChart.Navigation navigation = new Atos.SyntBots.PowerChart.Navigation();
            navigation.ClickOrders();


            //System.Windows.Forms.Cursor.Position = new System.Drawing.Point(860, 340);  //1080, 360                                                              
            //Mouse.Instance.Click();

            //TestStack.White.UIItems.WindowItems.Window _powerChartWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("Opened by"));
            //AutomationElement AEOrdersList = driver.GetElmtByWlker(_powerChartWindow.AutomationElement, "Order List", null, null, "pane", null);

            //System.Windows.Rect r = AEOrdersList.Current.BoundingRectangle;
            System.Windows.Point ppp = new System.Windows.Point(1080, 360);
            driver.MouseClick(ppp);

            int rowToclick = 1;
            int i = 1;

            object patternObj;


            string prevDesc = string.Empty;
            string curDesc = string.Empty;
            while (i <= rowToclick)
            {
                System.Threading.Thread.Sleep(2000);
                if (i > 1)
                {

                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
                }

                Keyboard.Instance.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                Keyboard.Instance.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                Keyboard.Instance.Enter("u");
                Keyboard.Instance.Enter("o");


                _mainWindow = Desktop.Instance.Windows().Find(x => x.Name.Contains("Opened by"));
                //_mainWindow.SetForeground();

                AutomationElement pcWindow = _mainWindow.AutomationElement;
                var popup = driver.GetRawChildrenByWlker(pcWindow)[0];
                if (popup.Current.LocalizedControlType != "Dialog")
                    break;

                var popupPane1 = driver.GetRawChildrenByWlker(popup)[1];
                var popupPane12 = driver.GetRawChildrenByWlker(popupPane1)[2];
                string deptartment = popupPane12.Current.Name;

                //AutomationElement tabItemdetails = driver.GetElmtByWlker(pcWindow, "Details", null, null, "tab item", null);
                //System.Windows.Rect r = tabItemdetails.Current.BoundingRectangle;
                //System.Windows.Point pp = r.TopLeft;
                //driver.MouseClick(pp);

                var popupPane = driver.GetRawChildrenByWlker(popup)[3];
                //  AutomationElement AEDiagnosisText = driver.GetElmtByWlker(popupPane, "Diagnosis", null, null, "text", null);
                // var AEDiagnosisVal = TreeWalker.ControlViewWalker.GetNextSibling(AEDiagnosisText);

                AutomationElement tabItemAdditionalInfo = driver.GetElmtByWlker(pcWindow, "Additional Info", null, null, "tab item", null);
                System.Windows.Rect r = tabItemAdditionalInfo.Current.BoundingRectangle;
                System.Windows.Point pp = r.TopLeft;
                driver.MouseClick(pp);

                AutomationElement AEOrderedAsVal = driver.GetElmtByWlker(popupPane, "Ordered As", null, null, "document", null);
                AutomationElement AEStartDateTimeVal = driver.GetElmtByWlker(popupPane, "Start Date/Time", null, null, "document", null);
                AutomationElement AEStopDateTimeVal = driver.GetElmtByWlker(popupPane, "Stop Date/Time", null, null, "document", null);

                string orderedAsVal = string.Empty;
                string startDateTimeVal = string.Empty;
                string stopDateTimeVal = string.Empty;

                if (AEOrderedAsVal.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
                {
                    var textPattern = (TextPattern)patternObj;
                    orderedAsVal = textPattern.DocumentRange.GetText(-1).TrimEnd('\r');

                }
                if (AEStartDateTimeVal.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
                {
                    var textPattern = (TextPattern)patternObj;
                    startDateTimeVal = textPattern.DocumentRange.GetText(-1).TrimEnd('\r');
                }
                if (AEStopDateTimeVal.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
                {
                    var textPattern = (TextPattern)patternObj;
                    stopDateTimeVal = textPattern.DocumentRange.GetText(-1).TrimEnd('\r');
                }


                //listPowerChartOrders.Add(new { ChargeDescription = orderedAsVal, ChargeDate = DateTime.Parse(startDateTimeVal.Replace("EDT", "").Trim()).ToString("MM/dd/yyyy") });
                i++;
                rowToclick++;

                AutomationElement AE6 = driver.GetElmtByWlker(popup, "Close", null, null, "button", null);
                driver.AEClick(AE6);

                var result = FindOrderInCatalog(ordersCatalogdata, orderedAsVal);

                if (result.Count == 1)
                {
                    isLabRecord = true;
                    break;
                }

                curDesc = orderedAsVal;
                if (curDesc == prevDesc)
                    break;
                prevDesc = orderedAsVal;


            }

            return isLabRecord;

        }

        //		public string FindOrderInCatalog(DataTable dt, string chargedescription)
        //        {
        //            string catalog_display = (from DataRow dr in dt.Rows
        //			                          where (string)dr["CATALOG_DISPLAY"] == chargedescription.Trim()
        //                                      select (string)dr["CPT_CODE"]).FirstOrDefault();
        //
        //            return catalog_display;
        //        }

        public List<string> FindOrderInCatalog(DataTable dt, string chargedescription)
        {

            var cptCodesList = (from DataRow dr in dt.Rows
                                where (string)dr["CATALOG_DISPLAY"] == chargedescription.Trim()
                                select (string)dr["CPT_CODE"]).Distinct().ToList();

            if (cptCodesList.Count == 0)
            {
                cptCodesList = (from DataRow dr in dt.Rows
                                where (string)dr["CATALOG_DISPLAY2"] == chargedescription.Trim()
                                select (string)dr["CPT_CODE"]).Distinct().ToList();
            }

            return cptCodesList;

        }

        public List<string> GetHCPCSByOrder(DataTable dt, string chargedescription)
        {

            var cptCodesList = (from DataRow dr in dt.Rows
                                where (string)dr["CATALOG_DISPLAY"] == chargedescription.Trim()
                                select (string)dr["HCPCS_CODE"]).Distinct().ToList();

            if (cptCodesList.Count == 0)
            {
                cptCodesList = (from DataRow dr in dt.Rows
                                where (string)dr["CATALOG_DISPLAY2"] == chargedescription.Trim()
                                select (string)dr["HCPCS_CODE"]).Distinct().ToList();
            }


            return cptCodesList;

        }

        public string FindOrderInCatalogByDescription(DataTable dt, string chargedescription)
        {

            var description = (from DataRow dr in dt.Rows
                               where (string)dr["CATALOG_DISPLAY"] == chargedescription.Trim()
                               select (string)dr["CATALOG_DISPLAY"]).FirstOrDefault();

            if (string.IsNullOrEmpty(description))
            {
                description = (from DataRow dr in dt.Rows
                               where (string)dr["CATALOG_DISPLAY2"] == chargedescription.Trim()
                               select (string)dr["CATALOG_DISPLAY2"]).FirstOrDefault();
            }

            return description;

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
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(1));
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
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(1));
            }
            return cvWindowPopup;
        }

    }
}
