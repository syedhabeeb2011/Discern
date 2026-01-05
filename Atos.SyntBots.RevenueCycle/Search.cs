using SyntBotsUIUtility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using TestStack.White;
using TestStack.White.InputDevices;
using TestStack.White.UIItems.Finders;

namespace Atos.SyntBots.RevenueCycle
{
    public class Search
    {
        SyntBotsUIUtil driver;

        public Search()
        {
            driver = new SyntBotsUIUtil();
        }
        
        public void SearchFIN(string FIN)
        {
            var _revenueCycleWindow = GetWindowWithWait("Revenue Cycle", 180);
            _revenueCycleWindow.SetForeground();
            Thread.Sleep(1000);

            var workFlow = _revenueCycleWindow.GetElement(SearchCriteria.ByText("Workflow"));
            var childElements = driver.GetRawChildrenByWlker(workFlow);
            var aa = childElements[2];

            AutomationElement AESelectEncounter = driver.GetElmtByWlker(aa, "HIM Coding Errors", null, null, "tree item", null, -1, 1);
            driver.AEDoubleClick(AESelectEncounter);

            var _grid = _revenueCycleWindow.GetMultiple<TestStack.White.UIItems.ListView>(SearchCriteria.ByControlType(ControlType.DataGrid));

            foreach (var row in _grid[1].Rows)
            {
                if (row.Cells[6].Text.Contains(FIN))
                {
                    row.Select();
                    row.RightClick();
                    var popupActionCode = _revenueCycleWindow.Popup;

                    var actioncode = driver.GetElmtByWlker(_revenueCycleWindow.Popup.AutomationElement, "Apply Action Code", null, null, "menu item", null, -1, 1);
                    driver.AEClick(actioncode);
                    Keyboard.Instance.Enter("A133");
                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);

                    AutomationElement AESaveCodesDialog = driver.GetElmtByWlker(_revenueCycleWindow.AutomationElement, "Apply Action Code", null, null, "Dialog", null, -1, 1);
                    var btnOK = driver.GetElmtByWlker(AESaveCodesDialog, "OK", null, null, "button", null, -1, 1);
                    driver.AEClick(btnOK);
                    
                }

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
    }

}
