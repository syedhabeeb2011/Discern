/*
 * Created by SharpDevelop.
 * User: osman.alikhan
 * Date: 10/28/2019
 * Time: 5:28 PM
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
using System.Configuration;

namespace Atos.SyntBots.AccessHIM
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

        public void LaunchAccessHIM(string applicationPath, string username, string password)
        {
            //        	string appConfigPath = System.IO.Path.GetDirectoryName(this.GetType().Assembly.Location) + "\\App";
            //        	Atos.SyntBots.Common.Logger.Log("From ACCessHIM dll: " +appConfigPath);
            //        	
            //        	Configuration _configuration = ConfigurationManager.OpenExeConfiguration(appConfigPath);
            //        	string rootFolder = _configuration.AppSettings.Settings["RootFolderPath"] != null ? _configuration.AppSettings.Settings["RootFolderPath"].Value : string.Empty; //@"E:\TaskQueue\";//ConfigurationManager.AppSettings["AccessHIM.MasterExcel.Path"];
            //        	Atos.SyntBots.Common.Logger.Log("ACCessHIM AppPath:" + rootFolder);
            //        	string kk = _configuration.AppSettings.Settings["AccessHIM.ApplicationPath"] != null ? _configuration.AppSettings.Settings["AccessHIM.ApplicationPath"].Value : string.Empty; 
            //        	Atos.SyntBots.Common.Logger.Log("ACCessHIM AppPath:" + kk);
            //        	
            //        	string applicationPath = _configuration.AppSettings.Settings["AccessHIM.ApplicationPath"] != null ? _configuration.AppSettings.Settings["AccessHIM.ApplicationPath"].Value : string.Empty; //"D:\\Cerner\\AccessHIM\\AccessHIM.exe"; //ConfigurationManager.AppSettings["AccessHIM.ApplicationPath"]; //  
            //        	string username = _configuration.AppSettings.Settings["AccessHIM.Username"] != null ? _configuration.AppSettings.Settings["AccessHIM.Username"].Value : string.Empty;//"v.chakradhar.ark"; //ConfigurationManager.AppSettings["AccessHIM.Username"]; //"v.chakradhar.ark"; 
            //        	string password = _configuration.AppSettings.Settings["AccessHIM.Password"] != null ? _configuration.AppSettings.Settings["AccessHIM.Password"].Value : string.Empty; //@"Lm60bj%zEW3/";//ConfigurationManager.AppSettings["AccessHIM.Password"];// @"Xy'UCv?5DKe>";

            //Process.Start(applicationPath);
            driver.LaunchApp(applicationPath);

            //System.Threading.Thread.Sleep(4000);

            _mainWindow = GetWindowWithWait("Millennium Logon", 200); //Desktop.Instance.Windows().Find(K => K.Name == "Cerner Millennium");

            //TestStack.White.UIItems.TextBox txtUsername = _mainWindow.Get<TestStack.White.UIItems.TextBox>(SearchCriteria.ByText("Username:"));
            //Thread.Sleep(500);
            //txtUsername.SetValue(username);

            //TestStack.White.UIItems.TextBox txtPassword = _mainWindow.Get<TestStack.White.UIItems.TextBox>(SearchCriteria.ByText("Password:"));
            //Thread.Sleep(500);
            //txtPassword.SetValue(password);

            AutomationElement txtUsername = driver.GetElmtByWlker(_mainWindow.AutomationElement, null, "1001", null, null, null);
            driver.AESetText(txtUsername, username);



            Thread.Sleep(500);


            AutomationElement txtPassword = driver.GetElmtByWlker(_mainWindow.AutomationElement, null, "1003", null, null, null);
            Thread.Sleep(500);
            driver.AESetText(txtPassword, password);


            System.Threading.Thread.Sleep(1000);

            TestStack.White.UIItems.Button btnOK = _mainWindow.Get<TestStack.White.UIItems.Button>(SearchCriteria.ByText("OK"));
            System.Threading.Thread.Sleep(1000);
            btnOK.Click();
        }

        public void LaunchPowerChart()
        {
            TestStack.White.UIItems.WindowItems.Window _accessHIMWindow;

            _accessHIMWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("AccessHIM"));

            //_accessHIMWindow.SetForeground();

            AutomationElement AEPowerChart = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Launch PowerChart", null, null, "button", null, -1, 15);
            System.Windows.Rect r = AEPowerChart.Current.BoundingRectangle;
            System.Windows.Point pp = r.TopLeft;
            driver.MouseClick(pp);
            //driver.AEClick(AEPowerChart);

            //Thread.Sleep(6000);

            //TestStack.White.UIItems.WindowItems.Window _powerChartWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("Opened by"));
            //if (_powerChartWindow != null)
            //{
            //   _powerChartWindow.SetForeground();
            //}  
        }

        public void Launch3M(TestStack.White.UIItems.WindowItems.Window _accessHIMWindow)
        {
            try
            {
                //TestStack.White.UIItems.WindowItems.Window _accessHIMWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("AccessHIM - Coding"));

                _accessHIMWindow.SetForeground();
                Thread.Sleep(1000);

                AutomationElement AEEncoderBtn = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Launch Encoder", null, null, "button", null);
                driver.AEClick(AEEncoderBtn);

                Thread.Sleep(10000);

                TestStack.White.UIItems.WindowItems.Window _3MWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("360 Encompass: Visit:"));
                if (_3MWindow != null)
                {
                    _3MWindow.SetForeground();
                }

            }
            catch (Exception ex)
            {
                throw ex.InnerException;
            }
        }

        public void LaunchChargeViewer()
        {
            TestStack.White.UIItems.WindowItems.Window _accessHIMWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("AccessHIM - Coding"));

            _accessHIMWindow.SetForeground();
            Thread.Sleep(1000);

            AutomationElement AEChargeViewer = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Launch Charge Viewer", null, null, "button", null);
            System.Windows.Rect r = AEChargeViewer.Current.BoundingRectangle;
            System.Windows.Point pp = r.TopLeft;
            driver.MouseClick(pp);
            //driver.AEClick(AEChargeViewer);

            TestStack.White.UIItems.WindowItems.Window _chargeViewerWindow = Desktop.Instance.Windows().Find(x => x.Name == "Charge Viewer (Cannot Increase Quantity)");
            if (_chargeViewerWindow != null)
            {
                _chargeViewerWindow.SetForeground();
            }



        }

        public void CloseAccessHIM()
        {
            TestStack.White.UIItems.WindowItems.Window _accessHIMWindow = Desktop.Instance.Windows().Find(K => K.Name.Contains("AccessHIM"));
            if (_accessHIMWindow != null)
            {
                _accessHIMWindow.SetForeground();

                WindowPattern windowPattern = (WindowPattern)_accessHIMWindow.AutomationElement.GetCurrentPattern(WindowPattern.Pattern);
                windowPattern.Close();
                Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
            }
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
