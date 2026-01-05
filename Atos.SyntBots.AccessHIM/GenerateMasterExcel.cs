/*
 * Created by SharpDevelop.
 * User: osman.alikhan
 * Date: 10/29/2019
 * Time: 12:46 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Windows.Automation;
using SyntBotsUIUtility;
using TestStack.White.UIItems.Finders;
using System.Linq;
using TestStack.White;
using TestStack.White.InputDevices;
using System.Configuration;

namespace Atos.SyntBots.AccessHIM
{
    /// <summary>
    /// Description of GenerateMasterExcel.
    /// </summary>
    public class GenerateMasterExcel
    {
        TestStack.White.UIItems.WindowItems.Window _mainWindow;
        SyntBotsUIUtil driver;

        public GenerateMasterExcel()
        {
            driver = new SyntBotsUIUtil();
        }


        public void Generate(string filename)
        {
            _mainWindow = GetWindowWithWait("AccessHIM - Task Queue", 15);//Desktop.Instance.Windows().Find(K => K.Name == "AccessHIM - Task Queue");

            TestStack.White.UIItems.Button btnQuery = _mainWindow.Get<TestStack.White.UIItems.Button>(SearchCriteria.ByText("Query"));
            btnQuery.Click();

            System.Threading.Thread.Sleep(1000);

            AutomationElement AErt1 = driver.GetDesktopChild(false, "AccessHIM - Task Queue", "", "SWT_Window0", "window", null);


            //Click on Select None hyperLink

            AutomationElement AE2 = driver.GetElmtByWlker(AErt1, "None", null, null, "hyperlink", null);

            driver.AEClick(AE2);

            var popup = _mainWindow.MdiChild(SearchCriteria.ByText("Preferences (Filtered)"));

            var abc = popup.GetMultiple<TestStack.White.UIItems.ListView>(SearchCriteria.ByControlType(ControlType.DataGrid));

            foreach (var listviewitem in abc[0].Rows)
            {

                if (listviewitem.Cells[0].Text == "Pending")
                {

                    listviewitem.Select();

                    listviewitem.KeyIn(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.SPACE);

                }

            }

            //Click on Select None hyperLink


            System.Threading.Thread.Sleep(1000);

            AutomationElement AE4 = driver.GetElmtByWlker(AErt1, "None", null, null, "hyperlink", null, 2);

            driver.AEClick(AE4);

            //                foreach (var listviewitem in abc[1].Rows)
            //                {
            //
            //                    if ((listviewitem.Cells[0].Text == "Coding OBS") || (listviewitem.Cells[0].Text == "Coding OP - General") || (listviewitem.Cells[0].Text == "Coding OP - in a Bed") ||
            //
            //                   (listviewitem.Cells[0].Text == "Coding OP - Lab") || (listviewitem.Cells[0].Text == "Coding OP - Onc and Infusions") || (listviewitem.Cells[0].Text == "Coding OP - Rad and Cardio") ||
            //
            //                   (listviewitem.Cells[0].Text == "Coding SDS - Cardio and Vascular") || (listviewitem.Cells[0].Text == "Coding SDS - General") || (listviewitem.Cells[0].Text == "Coding SDS - GI"))
            //                    {
            //
            //                        listviewitem.Select();
            //
            //                        listviewitem.KeyIn(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.SPACE);
            //
            //                    }
            //
            //                }


            foreach (var listviewitem in abc[1].Rows)
            {

                if ((listviewitem.Cells[0].Text == "Coding OP - Lab"))
                {

                    listviewitem.Select();

                    listviewitem.KeyIn(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.SPACE);

                }

            }

            System.Threading.Thread.Sleep(3000);
            AutomationElement AE5 = driver.GetElmtByWlker(AErt1, "Apply", null, null, "button", null);
            driver.AEClick(AE5);



            System.Threading.Thread.Sleep(1000);
            AutomationElement AE6 = driver.GetElmtByWlker(AErt1, "Apply and Close", null, null, "button", null);
            driver.AEClick(AE6);


            AutomationElement AE7 = driver.GetElmtByWlker(AErt1, "View Menu", null, null, null, null);
            driver.AEClick(AE7);
            System.Threading.Thread.Sleep(1000);


            AutomationElement AErt2 = driver.GetDesktopChild(false, "Menu", "", "#32768", "menu", null);
            AutomationElement AE22 = driver.GetElmtByWlker(AErt2, null, "Item 2", null, null, null);
            driver.AEClick(AE22);

            System.Threading.Thread.Sleep(1000);
            AutomationElement AE3 = driver.GetElmtByWlker(AErt2, null, "Item 123", null, null, null);
            driver.AEClick(AE3);

            System.Threading.Thread.Sleep(1000);
            AutomationElement AErt3 = driver.GetDesktopChild(false, "AccessHIM - Task Queue", "", "SWT_Window0", "window", null);
            AutomationElement AE8 = driver.GetElmtByWlker(AErt3, "File name:", "1001", null, null, null);

            driver.AEClick(AE8);

            Keyboard.Instance.Enter(filename);

            System.Threading.Thread.Sleep(3000);

            AutomationElement AE9 = driver.GetElmtByWlker(AErt3, "Save", null, null, null, null);

            driver.AEClick(AE9);

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
    }
}
