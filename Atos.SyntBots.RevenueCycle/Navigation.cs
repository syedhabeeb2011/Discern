using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestStack.White;
using SyntBotsUIUtility;
using TestStack.White.UIItems.Finders;
using System.Threading;
using TestStack.White.InputDevices;

namespace Atos.SyntBots.RevenueCycle
{
    public class Navigation
    {
        SyntBotsUIUtil driver;

        public Navigation()
        {
            driver = new SyntBotsUIUtil();
        }

        public void ClickPatientAccount()
        {
            var _revenueCycleWindow = GetWindowWithWait("Revenue Cycle", 180);
            var btnPatientAccount = _revenueCycleWindow.Get(SearchCriteria.ByText("Patient Account"));
            
            System.Windows.Rect r = btnPatientAccount.Bounds;
            System.Windows.Point pp = r.TopLeft;
            driver.MouseClick(pp);
        }

        public void NagivateToWorkflowTab()
        {
            var _revenueCycleWindow = GetWindowWithWait("Revenue Cycle", 180);

            _revenueCycleWindow.SetForeground();

            Keyboard.Instance.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
            Keyboard.Instance.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
            Keyboard.Instance.Enter("v");
            Keyboard.Instance.Enter("o");
            Keyboard.Instance.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RIGHT);
            Keyboard.Instance.Enter("w");
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
