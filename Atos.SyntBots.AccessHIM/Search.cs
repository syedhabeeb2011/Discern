/*
 * Created by SharpDevelop.
 * User: osman.alikhan
 * Date: 10/29/2019
 * Time: 2:24 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Windows.Automation;
using SyntBotsUIUtility;
using System.Linq;
using TestStack.White;
using TestStack.White.InputDevices;

namespace Atos.SyntBots.AccessHIM
{
    /// <summary>
    /// Description of SearchFIN.
    /// </summary>
    public class Search
    {
        SyntBotsUIUtil driver;

        public Search()
        {
            driver = new SyntBotsUIUtil();
        }

        public void SearchFIN(TestStack.White.UIItems.WindowItems.Window _accessHIMWindow, string FIN)
        {
            #region Search FIN number in AccessHIM
            //var _accessHIMWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("AccessHIM -"));
            //_accessHIMWindow.SetForeground();

            Condition conditions = new AndCondition(new PropertyCondition(AutomationElement.IsEnabledProperty, true), new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane));

            AutomationElementCollection elementCollection = _accessHIMWindow.AutomationElement.FindAll(TreeScope.Children, conditions);
            System.Threading.Thread.Sleep(3000);
            Atos.SyntBots.Common.Logger.Log("Line 40");
            AutomationElement pane0 = driver.GetElmtByWlker(elementCollection[0], null, null, null, "pane", null);
            Atos.SyntBots.Common.Logger.Log("Line 42");
            //AutomationElement AErt1 = driver.GetDesktopChild(false, "AccessHIM - Task Queue", "", "SWT_Window0", "window", null);
            var pane00 = driver.GetRawChildrenByWlker(pane0);
            Atos.SyntBots.Common.Logger.Log("Line 45");
            var rebar0 = driver.GetRawChildrenByWlker(pane00[0]);
            Atos.SyntBots.Common.Logger.Log("Line 46"); ;
            var rebar00 = driver.GetRawChildrenByWlker(rebar0[0]);
            Atos.SyntBots.Common.Logger.Log("Line 48"); ;
            var toolbar0 = driver.GetRawChildrenByWlker(rebar00[0]);

            var toolbarPane0 = driver.GetRawChildrenByWlker(toolbar0[0]);

            var toolbarPane00 = driver.GetRawChildrenByWlker(toolbarPane0[0]);

            var searchBtn = toolbarPane00[2];
            Atos.SyntBots.Common.Logger.Log("Line 56"); ;
            driver.AEClick(searchBtn);

            //System.Threading.Thread.Sleep(1000);
            AutomationElement AE40 = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Fin Nbr", null, null, "edit", null, -1, 15);
            driver.AEClick(AE40);
            Keyboard.Instance.Enter(FIN);
            AutomationElement AE44 = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Search", null, null, "button", null, -1, 15);
            driver.AEClick(AE44);
            //System.Threading.Thread.Sleep(3000);
            AutomationElement panePersonSearch = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Person Search", null, null, "Dialog", null, -1, 15);


            //Check for no result found..

            AutomationElement AE45 = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Select", null, null, "button", null);
            System.Threading.Thread.Sleep(1000);

            driver.AEClick(AE45);
            //System.Threading.Thread.Sleep(1000);
            //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
            #endregion
        }
    }
}
