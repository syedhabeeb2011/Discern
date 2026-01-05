using System;
using TestStack.White;
using SyntBotsUIUtility;
using System.Diagnostics;
using TestStack.White.UIItems.Finders;
using System.Threading;
using System.Windows.Automation;

namespace Atos.SyntBots.RevenueCycle
{
    /// <summary>
    /// Description of Launcher.
    /// </summary>
    public class Launcher
    {
        TestStack.White.UIItems.WindowItems.Window _mainWindow;
        SyntBotsUIUtil driver;

        public Launcher()
        {
            driver = new SyntBotsUIUtil();
        }

        public void LaunchRevenueCycle(string applicationPath, string username, string password)
        {
            //Process.Start(applicationPath);

            System.Diagnostics.Process[] revenueCycleProcesses = System.Diagnostics.Process.GetProcessesByName("RevenueCycle");
            foreach (System.Diagnostics.Process p in revenueCycleProcesses)
            {

                try
                {
                    p.Kill();
                }
                catch { }

            }

            driver.LaunchApp(applicationPath);

            //System.Threading.Thread.Sleep(4000);
            try
            {
                _mainWindow = GetWindowWithWait("Millennium Logon", 200);
            }
            catch(Exception ex)
            {
                throw new Exception("error at main Window initialization. Error : " + ex.Message);
            }
            //Desktop.Instance.Windows().Find(K => K.Name == "Cerner Millennium");

            //TestStack.White.UIItems.TextBox txtUsername = _mainWindow.Get<TestStack.White.UIItems.TextBox>(SearchCriteria.ByText("Username:"));

            //txtUsername.SetValue(username);

            //TestStack.White.UIItems.TextBox txtPassword = _mainWindow.Get<TestStack.White.UIItems.TextBox>(SearchCriteria.ByText("Password:"));

            //txtPassword.SetValue(password);

            try
            {

                AutomationElement txtUsername = driver.GetElmtByWlker(_mainWindow.AutomationElement, null, "1001", null, null, null);
                driver.AESetText(txtUsername, username);
            }
            catch(Exception ex)
            {
                throw new Exception("Error while enter user name. MainWindow object isNull? "+(_mainWindow==null)+" Error: "+ex.Message);
            }



            Thread.Sleep(500);


            AutomationElement txtPassword = driver.GetElmtByWlker(_mainWindow.AutomationElement, null, "1003", null, null, null);
            Thread.Sleep(500);
            driver.AESetText(txtPassword, password);

            System.Threading.Thread.Sleep(1000);

            TestStack.White.UIItems.Button btnOK = _mainWindow.Get<TestStack.White.UIItems.Button>(SearchCriteria.ByText("OK"));
            System.Threading.Thread.Sleep(1000);
            btnOK.Click();
        }

        public bool IsWindowAvailable(string windowName, int indxStrtFromZero = -1, int maxTimeout = 60)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            long num = 0L;
            bool isWindowExists = false;

            if (windowName != null)
            {
                long elapsedMilliseconds;

                do
                {
                    stopwatch.Start();
                    try
                    {
                        var _window = Desktop.Instance.Windows().Find(K => K.Name.Contains(windowName));
                        isWindowExists = _window != null ? true : false;

                    }
                    catch (Exception ex)
                    {
                    }

                    stopwatch.Stop();
                    elapsedMilliseconds = stopwatch.ElapsedMilliseconds;

                    if (num == 0L)
                    {
                        num = elapsedMilliseconds;
                    }
                }
                while (elapsedMilliseconds < (long)(maxTimeout * 1000) + num && isWindowExists);
            }

            return isWindowExists;
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
