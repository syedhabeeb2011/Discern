/*
 * Created by SharpDevelop.
 * User: v.chakradhar.ark
 * Date: 10/30/2019
 * Time: 10:23 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Windows.Automation;
using TestStack.White;
using SyntBotsUIUtility;
using TestStack.White.InputDevices;
using System.Threading;
using TestStack.White.WindowsAPI;

namespace Atos.SyntBots._3M
{
    /// <summary>
    /// Description of DefaultCode.
    /// </summary>
    public class DefaultCode
    {
        SyntBotsUIUtil driver;

        public DefaultCode()
        {
            driver = new SyntBotsUIUtil();
        }

        public void AddDefaultCode(TestStack.White.UIItems.WindowItems.Window _3MWindow, bool isZerocharge = false)
        {

            //				if(isZerocharge)
            //				{
            //				Thread.Sleep(1000);
            //				Keyboard.Instance.Enter("1");
            //				Thread.Sleep(1000);
            //				Keyboard.Instance.Enter("Z09");
            //				Thread.Sleep(1000);
            //				Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.RETURN);
            //				//Thread.Sleep(1000);
            //				//Keyboard.Instance.Enter("m");
            //				Thread.Sleep(1000);
            //				Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.RETURN);
            //				Thread.Sleep(1000);
            //				Keyboard.Instance.Enter("m");
            //				Thread.Sleep(1000);
            //				Keyboard.Instance.Enter("1");
            //				Thread.Sleep(1000);
            //				Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.RETURN);
            //				Thread.Sleep(1000);
            //				Keyboard.Instance.Enter("m");
            //				}
            //				else
            //				{
            Thread.Sleep(1000);
            Keyboard.Instance.Enter("1");
            Thread.Sleep(1000);
            Keyboard.Instance.Enter("Z09");
            Thread.Sleep(1000);
            Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
            Thread.Sleep(500);
            Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
            ////Thread.Sleep(1000);
            ////Keyboard.Instance.Enter("m");
            //Thread.Sleep(1000);
            //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
            Thread.Sleep(1000);
            Keyboard.Instance.Enter("m");
            Thread.Sleep(1000);
            Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
            Thread.Sleep(1000);
            Keyboard.Instance.Enter("m");
            //		}

            Thread.Sleep(8000);

            AutomationElement AE19 = driver.GetElmtByWlker(_3MWindow.AutomationElement, null, "TitleBar", null, null, null);
            driver.AEClick(AE19);

            Thread.Sleep(100);
            System.Windows.Forms.Cursor.Position = new System.Drawing.Point(1178, 136);
            Mouse.Instance.Click();
            Thread.Sleep(500);
            Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.TAB);
            Thread.Sleep(200);
            Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.TAB);
            Thread.Sleep(200);
            Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.TAB);
            Thread.Sleep(200);
            Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);

        }
    }
}
