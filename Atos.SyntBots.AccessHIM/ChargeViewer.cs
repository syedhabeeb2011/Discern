/*
 * Created by SharpDevelop.
 * User: v.chakradhar.ark
 * Date: 11/5/2019
 * Time: 3:30 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Windows.Automation;
using TestStack.White;
using SyntBotsUIUtility;
using System.Diagnostics;
using TestStack.White.UIItems.Finders;
using System.Threading;
using System.Linq;
using TestStack.White.InputDevices;
using TestStack.White.WindowsAPI;
using System.Collections.Generic;
using System.Windows.Forms;
using Atos.SyntBots.Common.Entity;

namespace Atos.SyntBots.AccessHIM
{
    //	public class Charge
    //    {
    //        public string ChargeDescription;
    //        public string ChargeDate;
    //        public string CPT4;
    //        public string Price;
    //        public string ActivityType;
    //        public string ChargeType;
    //        public string Status;
    //    }
    /// <summary>
    /// Description of ChargeViewer.
    /// </summary>
    public class ChargeViewer
    {
        SyntBotsUIUtil driver;

        public ChargeViewer()
        {
            driver = new SyntBotsUIUtil();
        }

        public List<Atos.SyntBots.Common.Entity.Charge> GetChargesData(string admitdate, TestStack.White.UIItems.WindowItems.Window accessHIMWindow)
        {

            var chargeList = new List<Atos.SyntBots.Common.Entity.Charge>();


            AutomationElement AEChargeViewer = driver.GetElmtByWlker(accessHIMWindow.AutomationElement, "Launch Charge Viewer", null, null, "button", null, -1, 30);
            Thread.Sleep(500);

            System.Windows.Rect r = AEChargeViewer.Current.BoundingRectangle;
            System.Windows.Point pp = r.TopLeft;
            driver.MouseClick(pp);
            //driver.AEClick(AEChargeViewer);

            //Atos.SyntBots.Common.Logger.Log("Line 55");
            var chargeViwerWindow = GetWindowWithWait("Charge Viewer (Cannot Increase Quantity)", 60);//Desktop.Instance.Windows().Find(x => x.Name == "Charge Viewer (Cannot Increase Quantity)");
            chargeViwerWindow.SetForeground();

            Thread.Sleep(1000);

            AutomationElement cvWindow = chargeViwerWindow.AutomationElement;

            var cvWindowPopup = GetPopUpWindowWithWait(cvWindow, 0, 65);
            AutomationElement okButton = driver.GetElmtByWlker(cvWindowPopup, "OK", null, null, "button", null, -1, 5);
            Thread.Sleep(500);
            driver.AEClick(okButton);
            //Atos.SyntBots.Common.Logger.Log("Line 67");

            Condition condition = driver.CreateCondition(false, "No charges found using selected filters!", null, null, "text");
            var AENoChargesPopup = driver.GetElmtByWlker(cvWindow, condition, -1, 5);

            if (AENoChargesPopup == null)
            {
                Thread.Sleep(500);
                driver.AEClick(chargeViwerWindow.AutomationElement);
                //Atos.SyntBots.Common.Logger.Log("Line 75");
                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);
                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);
                //                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);
                //                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);
                //                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);
                //                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);
                //                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);
                //                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);

                Thread.Sleep(300);
                if (Console.NumberLock)
                    Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.NUMLOCK);

                //Keyboard.Instance.HoldKey(KeyboardInput.SpecialKeys.CONTROL);
                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.PAGEUP);
                //Keyboard.Instance.LeaveKey(KeyboardInput.SpecialKeys.CONTROL);

                Thread.Sleep(200);
                SendKeys.SendWait("+{END}");
                //                Keyboard.Instance.HoldKey(KeyboardInput.SpecialKeys.SHIFT);
                //                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.END);
                //                Keyboard.Instance.LeaveKey(KeyboardInput.SpecialKeys.SHIFT);
                Thread.Sleep(200);
                Keyboard.Instance.HoldKey(KeyboardInput.SpecialKeys.CONTROL);
                Keyboard.Instance.Enter("c");
                Keyboard.Instance.LeaveKey(KeyboardInput.SpecialKeys.CONTROL);
                Thread.Sleep(300);
                var chargesClipbData = driver.GetClipBoardText();
                //var kk = System.Windows.Forms.Clipboard.GetText();
                List<string> data = new List<string>(chargesClipbData.Split('\n'));


                foreach (string iterationRow in data)
                {
                    string row = iterationRow;
                    if (row.EndsWith("\r"))
                    {
                        row = row.Substring(0, row.Length - "\r".Length);
                        string[] rowData = row.Split(new char[] { '\r', '\x09' });
                        string chargeDesc = rowData[5];


                        if (chargeDesc.Contains("venipuncture") || chargeDesc.Contains("Venipuncture"))
                        {
                            chargeDesc = "Venipuncture Charge (M)";
                        }

                        string chargdate = DateTime.Parse(rowData[21]).ToString("MM/dd/yyyy");
                        if (rowData[26] == "Posted" && rowData[17] != "$0.00")
                        {
                            var charge = new Atos.SyntBots.Common.Entity.Charge()
                            {
                                ChargeDescription = chargeDesc,
                                ChargeDate = chargdate,
                                CPT4 = rowData[9],
                                Price = rowData[17],
                                ActivityType = rowData[19],
                                ChargeType = rowData[22],
                                Status = rowData[26]
                            };

                            if (!chargeList.Contains(charge))
                            {
                                chargeList.Add(charge);
                            }
                        }
                        else if (rowData[26] == "Posted" && rowData[17] != "$0.00" && chargdate == admitdate && rowData[19].Contains("Ambulatory") && chargeDesc.Contains("Venipuncture"))
                        {
                            var charge = new Atos.SyntBots.Common.Entity.Charge()
                            {
                                ChargeDescription = chargeDesc,
                                ChargeDate = chargdate,
                                CPT4 = rowData[9],
                                Price = rowData[17],
                                ActivityType = rowData[19],
                                ChargeType = rowData[22],
                                Status = rowData[26]
                            };

                            if (!chargeList.Contains(charge))
                            {
                                chargeList.Add(charge);
                            }
                        }



                    }
                }
                chargeViwerWindow.Close();

            }
            else
            {
                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.RETURN);
                chargeViwerWindow.Close();
            }

            return chargeList;

        }

        public List<Charge> GetChargesData(string admitdate, List<Charge> data)
        {
            var chargeList = new List<Charge>();

            foreach (var iterationRow in data)
            {
                // string row = iterationRow;
                // row = row.Substring(0, row.Length - "\r".Length);
                //  string[] rowData = row.Split(new char[] { '\r', '\x09' });
                string chargeDesc = iterationRow.ChargeDescription;
                if (chargeDesc.Contains("venipuncture") || chargeDesc.Contains("Venipuncture"))
                {
                    chargeDesc = "Venipuncture Charge (M)";
                }

                //string chargdate = DateTime.Parse(iterationRow.ChargeDate).ToString("MM/dd/yyyy");
                //string chargePrice = string.Format("${0}", iterationRow.Price);
                string chargdate = string.IsNullOrEmpty(iterationRow.ChargeDate) ? iterationRow.ChargeDate : DateTime.Parse(iterationRow.ChargeDate).ToString("MM/dd/yyyy");
                string chargePrice = string.IsNullOrEmpty(iterationRow.Price) ? iterationRow.Price : string.Format("${0}", iterationRow.Price);

                if (iterationRow.Status == "Posted" && chargePrice != "$0.00")
                {                    
                        iterationRow.ChargeDescription = chargeDesc;
                        iterationRow.ChargeDate = chargdate;
                        iterationRow.Price = chargePrice;

                        if (!chargeList.Contains(iterationRow))
                        {
                            chargeList.Add(iterationRow);
                        }         
                }
                else if (iterationRow.Status == "Suspended")
                {
                    throw new Exception("Suspended Charge");
                }
            }
            return chargeList;
        }

        public TestStack.White.UIItems.WindowItems.Window GetWindowWithWait(string windowName, int timeout = 60)
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


        public AutomationElement GetPopUpWindowWithWait(AutomationElement automationElement, int windowIndex = 0, int timeout = 60)
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
    }
}
