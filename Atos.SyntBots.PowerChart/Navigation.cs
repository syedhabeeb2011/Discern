/*
 * Created by SharpDevelop.
 * User: osman.alikhan
 * Date: 10/29/2019
 * Time: 6:23 PM
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
using System.Configuration;

namespace Atos.SyntBots.PowerChart
{
    /// <summary>
    /// Description of Navigation.
    /// </summary>
    public class Navigation
    {
        SyntBotsUIUtil driver;

        public Navigation()
        {
            driver = new SyntBotsUIUtil();
        }

        public void ClickOrders()
        {
            TestStack.White.UIItems.WindowItems.Window _powerChartWindow = GetWindowWithWait("Opened by", 15);//Desktop.Instance.Windows().Find(K => K.Name.Contains("Opened by"));

            var AEPowerChartChildren = GetRawChildrenByWlkerWithWait(_powerChartWindow.AutomationElement, 5);//driver.GetRawChildrenByWlker(_powerChartWindow.AutomationElement);

            var pane1Contents = GetRawChildrenByWlkerWithWait(AEPowerChartChildren[1], 5);//driver.GetRawChildrenByWlker(AEPowerChartChildren[1]);

            var menu = GetRawChildrenByWlkerWithWait(pane1Contents[3], 5);//driver.GetRawChildrenByWlker(pane1Contents[3]);

            var dataGrid = GetRawChildrenByWlkerWithWait(menu[0], 5);//driver.GetRawChildrenByWlker(menu[0]);

            var ordersItem = GetRawChildrenByWlkerWithWait(dataGrid[1], 5);//driver.GetRawChildrenByWlker(dataGrid[1]);

            var AEOrders = ordersItem[0];

            driver.AEClick(AEOrders);
        }

        public void ClickLaboratory()
        {
            TestStack.White.UIItems.WindowItems.Window _powerChartWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("Opened by"));
            SearchCriteria searchCriteria = SearchCriteria.ByText("Laboratory");
            var labAE = _powerChartWindow.Get(searchCriteria);
            labAE.Click();
        }

        public void ClickObservationEvents()
        {
            TestStack.White.UIItems.WindowItems.Window _powerChartWindow = GetWindowWithWait("Opened by", 10); //Desktop.Instance.Windows().Find(K => K.Name.Contains("Opened by"));

            var AEPowerChartChildren = GetRawChildrenByWlkerWithWait(_powerChartWindow.AutomationElement, 10);//driver.GetRawChildrenByWlker(_powerChartWindow.AutomationElement);

            var pane1Contents = GetRawChildrenByWlkerWithWait(AEPowerChartChildren[1], 10); //driver.GetRawChildrenByWlker(AEPowerChartChildren[1]);

            var menu = GetRawChildrenByWlkerWithWait(pane1Contents[3], 10);//driver.GetRawChildrenByWlker(pane1Contents[3]);

            var dataGrid = GetRawChildrenByWlkerWithWait(menu[0], 10); //driver.GetRawChildrenByWlker(menu[0]);

            var observationEventsItem = GetRawChildrenByWlkerWithWait(dataGrid[4], 10); //driver.GetRawChildrenByWlker(dataGrid[4]);

            var AEobservationEvents = observationEventsItem[0];

            driver.AEClick(AEobservationEvents);
        }

        public void ClickNotes()
        {

            TestStack.White.UIItems.WindowItems.Window _powerChartWindow = GetWindowWithWait("Opened by", 10);//Desktop.Instance.Windows().Find(K => K.Name.Contains("Opened by"));

            var AEPowerChartChildren = driver.GetRawChildrenByWlker(_powerChartWindow.AutomationElement);

            var pane1Contents = driver.GetRawChildrenByWlker(AEPowerChartChildren[1]);

            var menu = driver.GetRawChildrenByWlker(pane1Contents[3]);

            var dataGrid = driver.GetRawChildrenByWlker(menu[0]);

            var notesItem = driver.GetRawChildrenByWlker(dataGrid[5]);

            var AENotes = notesItem[0];

            driver.AEClick(AENotes);

        }

        public void ClickResultsReview()
        {

            TestStack.White.UIItems.WindowItems.Window _powerChartWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("Opened by"));

            var AEPowerChartChildren = driver.GetRawChildrenByWlker(_powerChartWindow.AutomationElement);

            var pane1Contents = driver.GetRawChildrenByWlker(AEPowerChartChildren[1]);

            var menu = driver.GetRawChildrenByWlker(pane1Contents[3]);

            var dataGrid = driver.GetRawChildrenByWlker(menu[0]);

            var resultsReviewItem = driver.GetRawChildrenByWlker(dataGrid[2]);

            var AEResultsReview = resultsReviewItem[0];

            driver.AEClick(AEResultsReview);

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

        public System.Collections.Generic.List<AutomationElement> GetRawChildrenByWlkerWithWait(AutomationElement automationElement, int timeout = 30)
        {
            System.Collections.Generic.List<AutomationElement> result = null;
            for (int i = 0; i < timeout; i++)
            {
                result = driver.GetRawChildrenByWlker(automationElement);

                if (result != null)
                {
                    return result;
                }
                Thread.Sleep(TimeSpan.FromSeconds(1));
            }
            return result;
        }
    }
}
