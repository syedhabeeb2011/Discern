/*
 * Created by SharpDevelop.
 * User: v.chakradhar.ark
 * Date: 10/30/2019
 * Time: 10:51 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Windows.Automation;
using SyntBotsUIUtility;
using System.Linq;
using TestStack.White;
using TestStack.White.InputDevices;
using System.Threading;
using TestStack.White.WindowsAPI;
//using TestStack.White.UIItems.Finders;
namespace Atos.SyntBots.AccessHIM
{
    /// <summary>
    /// Description of DiscernNotification.
    /// </summary>
    public class DiscernNotification
    {
        TestStack.White.UIItems.WindowItems.Window _mainWindow;
        SyntBotsUIUtil driver;

        public DiscernNotification()
        {
            driver = new SyntBotsUIUtil();
        }
        #region commented by Velan on 20-06-2022
        //public void SendDiscernNotification(string dnComment, string codingPassReason, string personResponsible)
        //{
        //    try
        //    {
        //        #region Discern Notification

        //        TestStack.White.UIItems.WindowItems.Window _accessHIMWindow;

        //        _accessHIMWindow = GetWindowWithWait("AccessHIM", 15); //Desktop.Instance.Windows().Find(K => K.Name.Contains("AccessHIM - "));
        //        _accessHIMWindow.SetForeground();


        //        Thread.Sleep(500);

        //        // Find Notes Pane
        //        AutomationElement notesPane3 = GetPopUpWindowWithWait(_accessHIMWindow.AutomationElement, 3, 10); //driver.GetRawChildrenByWlker(_accessHIMWindow.AutomationElement)[3];
        //        AutomationElement notesPane30 = GetPopUpWindowWithWait(notesPane3, 0, 10); //driver.GetRawChildrenByWlker(notesPane3)[0];
        //        AutomationElement notesPane300 = GetPopUpWindowWithWait(notesPane30, 0, 10);//driver.GetRawChildrenByWlker(notesPane30)[0];
        //        AutomationElement notesPane308 = GetPopUpWindowWithWait(notesPane30, 8, 10);//driver.GetRawChildrenByWlker(notesPane30)[8];

        //        // Click Notes Tab  
        //        //Thread.Sleep(1000);

        //        //AutomationElement AENotesTab  = GetPopUpWindowWithWait(notesPane308, 1 , 15);//driver.GetRawChildrenByWlker(notesPane308)[3];
        //        //driver.AEClick(AENotesTab);

        //        Keyboard.Instance.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
        //        Keyboard.Instance.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
        //        Keyboard.Instance.Enter("v");
        //        Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
        //        Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RIGHT);
        //        Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
        //        Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
        //        Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
        //        Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);

        //        //AutomationElement AE4 = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "View", null, null, null, null); 
        //        //driver.AEClick(AE4); 
        //        ////-----------------------------------------------------------------------------------
        //        // AutomationElement AE5 = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Open View", null, null, null, null); 
        //        //driver.AEClick(AE5); 
        //        ////-----------------------------------------------------------------------------------
        //        // AutomationElement AE6 = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Notes", null, null, null, null); 
        //        //driver.AEClick(AE6); 

        //        //Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.RIGHT);
        //        //Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.RIGHT);
        //        //System.Windows.Rect r  = AENotesTab.Current.BoundingRectangle;
        //        Thread.Sleep(3000);

        //        //System.Windows.Point pp = r.TopRight; 
        //        //driver.AEClick(AENotesTab);

        //        //            if(AENotesTab.Current.Name=="Notes")
        //        //            {
        //        //             AENotesTab.SetFocus();
        //        //            
        //        //            System.Windows.Rect r  = AENotesTab.Current.BoundingRectangle;
        //        //            Thread.Sleep(3000);
        //        //             
        //        //            System.Windows.Point pp = r.TopRight; 
        //        //            driver.MouseClick(pp);
        //        //            //driver.AEClick(AENotesTab);
        //        //            }
        //        // Click Add Note Button
        //        //Thread.Sleep(2000);         
        //        AutomationElement N2notesPane3 = GetPopUpWindowWithWait(_accessHIMWindow.AutomationElement, 3, 20); //driver.GetRawChildrenByWlker(_accessHIMWindow.AutomationElement)[3];
        //        AutomationElement N2notesPane30 = GetPopUpWindowWithWait(notesPane3, 0, 20); //driver.GetRawChildrenByWlker(notesPane3)[0];
        //        AutomationElement N2notesPane300 = GetPopUpWindowWithWait(notesPane30, 0, 20);//driver.GetRawChildrenByWlker(notesPane30)[0];
        //        AutomationElement AEnotesToolbar = driver.GetElmtByWlker(N2notesPane300, "", null, null, "tool bar", null, -1, 20);

        //        var AEAddNoteBtn = GetPopUpWindowWithWait(AEnotesToolbar, 0, 20); //driver.GetRawChildrenByWlker(AEnotesToolbar)[0];
        //        driver.AEClick(AEAddNoteBtn);

        //        //Thread.Sleep(1000);         
        //        // Get Note Type combo box         
        //        AutomationElement AENoteTypeCB = driver.GetElmtByWlker(N2notesPane300, "Note Type: ", null, null, "combo box", null, -1, 15);

        //        // Click Combo box drop down            
        //        AutomationElement AENoteTypeDD = driver.GetElmtByWlker(AENoteTypeCB, "Drop Down Button", null, null, "button", null, -1, 15);
        //        driver.AEClick(AENoteTypeDD);

        //        //Thread.Sleep(1000);            
        //        // Select Encounter from DropDown
        //        AutomationElement AESelectEncounter = driver.GetElmtByWlker(AENoteTypeCB, "Encounter", null, null, "list item", null, -1, 15);
        //        driver.AEClick(AESelectEncounter);

        //        //Thread.Sleep(1000);            
        //        // Choose the document to enter text
        //        AutomationElement AEDocuments = driver.GetElmtByWlker(N2notesPane300, "", null, null, "document", null, -1, 15);
        //        driver.AEClick(AEDocuments);

        //        Keyboard.Instance.Enter(dnComment);

        //        Thread.Sleep(1000);
        //        // Click Notes Savebutton 
        //        /*
        //        AutomationElement AECbParent = TreeWalker.RawViewWalker.GetParent(AENoteTypeCB);
        //        AutomationElement AECbParentParent = TreeWalker.RawViewWalker.GetParent(AECbParent);
        //        var AECbParentParentChld = driver.GetRawChildrenByWlker(AECbParentParent);
        //        var notesToolbar = driver.GetRawChildrenByWlker(AECbParentParentChld[0]);
        //        var AENotesSaveBtn = notesToolbar[0];
        //        driver.AEClick(AENotesSaveBtn);
        //        */
        //        AutomationElement saveNotesAE = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Save Note", null, null, "button", null, -1, 15);
        //        driver.AEClick(saveNotesAE);

        //        Thread.Sleep(1000);

        //        // Click Savebutton
        //        var toolbarSave = GetPopUpWindowWithWait(N2notesPane30, 1, 10);//driver.GetRawChildrenByWlker(N2notesPane30)[1];
        //        var AESaveBtn = GetPopUpWindowWithWait(toolbarSave, 1, 10);// driver.GetRawChildrenByWlker(toolbarSave)[1];
        //        driver.AEClick(AESaveBtn);

        //        Thread.Sleep(1000);

        //        // Navigate to Save Codes dialog            
        //        AutomationElement AESaveCodesDialog = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Save Codes", null, null, "Dialog", null, -1, 15);

        //        var radiobuttons = _accessHIMWindow.GetMultiple<TestStack.White.UIItems.RadioButton>(TestStack.White.UIItems.Finders.SearchCriteria.ByControlType(ControlType.RadioButton));
        //        bool isFinalSelected = radiobuttons.FirstOrDefault(x => x.Name == "Final").IsSelected;
        //        if (isFinalSelected)
        //        {
        //            Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
        //        }




        //        //AutomationElement AEDraftDD = driver.GetElmtByWlker(AESaveCodesDialog, "Draft", "396598", null, "radio button", null);
        //        //driver.AEClick(AEDraftDD);
        //        AutomationElement AEPassReasonTxt = driver.GetElmtByWlker(AESaveCodesDialog, "Pass Reason:", null, null, "edit", null, -1, 15);
        //        if (AEPassReasonTxt == null)
        //        {
        //            AutomationElement aeSaveCodesOk = driver.GetElmtByWlker(AESaveCodesDialog, "Ok", null, null, "button", null, -1, 15);
        //            driver.AEClick(aeSaveCodesOk);
        //            throw new Exception("Pass Reason dropdown not found");
        //        }
        //        driver.AESetText(AEPassReasonTxt, codingPassReason);
        //        AutomationElement AEPassReasonDD = driver.GetElmtByWlker(AESaveCodesDialog, "Drop Down Button", null, null, "button", null, -1, 15);
        //        driver.AEClick(AEPassReasonDD);

        //        //Thread.Sleep(1000);

        //        // Navigate to Responsible Pane
        //        AutomationElement AEResponsibleGroup = driver.GetElmtByWlker(AESaveCodesDialog, "Responsible", null, null, "group", null, -1, 15);

        //        //var popup = _accessHIMWindow.MdiChild(SearchCriteria.ByText("Save Codes"));
        //        //var ab = popup.Get<TestStack.White.UIItems.GroupBox>(SearchCriteria.ByControlType(ControlType.Group));
        //        //var abc = ab.Get<TestStack.White.UIItems.TextBox>(SearchCriteria.ByControlType(ControlType.Edit));  
        //        //var test  = (TestStack.White.UIItems.TextBox)ab.Items[6];
        //        //test.Enter("Central, Lab Dn's");

        //        Atos.SyntBots.Common.Logger.Log("Debug Before Clicking search button", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");

        //        AutomationElement AEResponsibleBtn = driver.GetElmtByWlker(AEResponsibleGroup, "Begin Search", null, null, "button", null, -1, 15);
        //        driver.AEClick(AEResponsibleBtn);


        //        Atos.SyntBots.Common.Logger.Log("Debug After Clicking search button", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");

        //        //Thread.Sleep(1000);            

        //        // Operations in Personnel Search Dialog
        //        /*
        //        AutomationElement AEPersonnelSearchWindow = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Personnel Search", null, null, "Dialog", null);
        //        var child1 = GetRawChildrenByWlkerWithWait(AEPersonnelSearchWindow, 10);//driver.GetRawChildrenByWlker(AEPersonnelSearchWindow);
        //        var child2 = driver.GetRawChildrenByWlker(child1[0]);
        //        var child3 = driver.GetRawChildrenByWlker(child2[0]);
        //        var child4 = driver.GetRawChildrenByWlker(child3[0]);
        //        var child5 = driver.GetRawChildrenByWlker(child4[0]);
        //        var child6 = driver.GetRawChildrenByWlker(child5[0]);
        //        var fname = child6[4];
        //        var lname = child6[5];

        //        var splitName = personResponsible.Split(' ');
        //        driver.AESetText(fname, splitName[1]);
        //        driver.AESetText(lname, splitName[0]);

        //        Thread.Sleep(1000);            

        //        AutomationElement AEPersonnelSearchBtn = driver.GetElmtByWlker(AEPersonnelSearchWindow, "Search", null, null, "button", null, -1, 10);
        //        driver.AEClick(AEPersonnelSearchBtn);
        //        _accessHIMWindow.WaitWhileBusy();
        //        Thread.Sleep(1000);            

        //        AutomationElement AEPersonnelSearchdatagrid = driver.GetElmtByWlker(AEPersonnelSearchWindow, "", null, null, "data grid", null);
        //        var AEperson = driver.GetRawChildrenByWlker(AEPersonnelSearchdatagrid);
        //        driver.AEDoubleClick(AEperson[2]);
        //        */
        //        Thread.Sleep(1000);

        //        AutomationElement AEDescrDocuments = driver.GetElmtByWlker(AESaveCodesDialog, "Description", null, "Edit", "document", null, -1, 15);
        //        driver.AEClick(AEDescrDocuments);
        //        Keyboard.Instance.Enter(dnComment);

        //        //Thread.Sleep(1000);
        //        Atos.SyntBots.Common.Logger.Log("Debug Before Clicking Ok button", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");


        //        AutomationElement AESaveCodesOk = driver.GetElmtByWlker(AESaveCodesDialog, "Ok", null, null, "button", null, -1, 15);
        //        InvokePattern aeInvk = AESaveCodesOk.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
        //        aeInvk.Invoke();
        //        //driver.AEClick(AESaveCodesOk);
        //        Atos.SyntBots.Common.Logger.Log("Debug After Clicking Ok button", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
        //    }
        //    catch (Exception ex)
        //    {
        //        Atos.SyntBots.Common.Logger.Log("Debug Clicking Ok button" + ex.Message + "" + ex.StackTrace, @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
        //        throw new Exception("Error :" + ex.Message + "line :" + ex.StackTrace);
        //    }

        //    #endregion

        //}
        #endregion


        //Send DiscernNotification() function has been updated by Velan on 20-06-2022

        #region Commented by velan
        //public void SendDiscernNotification(string dnComment, string codingPassReason, string personResponsible)
        //{
        //    try
        //    {
        //        #region Discern Notification

        //        TestStack.White.UIItems.WindowItems.Window _accessHIMWindow;

        //        _accessHIMWindow = GetWindowWithWait("AccessHIM", 15); //Desktop.Instance.Windows().Find(K => K.Name.Contains("AccessHIM - "));
        //        _accessHIMWindow.SetForeground();


        //        Thread.Sleep(500);



        //        // Find Notes Pane
        //        #region Commented by Velan
        //        //AutomationElement notesPane3 = GetPopUpWindowWithWait(_accessHIMWindow.AutomationElement, 3, 10); //driver.GetRawChildrenByWlker(_accessHIMWindow.AutomationElement)[3];
        //        //AutomationElement notesPane30 = GetPopUpWindowWithWait(notesPane3, 0, 10); //driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //AutomationElement notesPane300 = GetPopUpWindowWithWait(notesPane30, 0, 10);//driver.GetRawChildrenByWlker(notesPane30)[0];
        //        //AutomationElement notesPane308 = GetPopUpWindowWithWait(notesPane30, 8, 10);//driver.GetRawChildrenByWlker(notesPane30)[8];
        //        #endregion

        //        #region commenting 20-06-2022
        //        // AutomationElement notesPane3 = GetPopUpWindowWithWait(_accessHIMWindow.AutomationElement, 3, 10); //driver.GetRawChildrenByWlker(_accessHIMWindow.AutomationElement)[3];
        //        // AutomationElement notesPane3_1 = GetPopUpWindowWithWait(notesPane3, 0, 10); //driver.GetRawChildrenByWlker(notesPane3)[0];
        //        // Atos.SyntBots.Common.Logger.Log("Check1", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //// AutomationElement notesPane3_2 = GetPopUpWindowWithWait(notesPane3_1, 0, 10); //driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //// Atos.SyntBots.Common.Logger.Log("Check2", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        // AutomationElement notesPane3_3 = GetPopUpWindowWithWait(notesPane3_1, 0, 10);
        //        // Atos.SyntBots.Common.Logger.Log("Check3", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        // AutomationElement notesPane3_4 = GetPopUpWindowWithWait(notesPane3_3, 1, 10);
        //        // Atos.SyntBots.Common.Logger.Log("Check4", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        // AutomationElement notesPane3_5 = GetPopUpWindowWithWait(notesPane3_4, 0, 10);
        //        // Atos.SyntBots.Common.Logger.Log("Check5", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        // AutomationElement notesPane3_6 = GetPopUpWindowWithWait(notesPane3_5, 0, 10);
        //        // Atos.SyntBots.Common.Logger.Log("Check6", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        // AutomationElement notesPane3_7 = GetPopUpWindowWithWait(notesPane3_6, 1, 10);
        //        // Atos.SyntBots.Common.Logger.Log("Check7", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        // AutomationElement notesPane3_8 = GetPopUpWindowWithWait(notesPane3_7, 5, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        // Atos.SyntBots.Common.Logger.Log("Check8", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        #endregion

        //        //Atos.SyntBots.Common.Logger.Log("Check1", @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");
        //        AutomationElement NotesTab = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Notes", null, null, "tab item", null);
        //        //Atos.SyntBots.Common.Logger.Log("NotesTab IsNull?" + (NotesTab == null), @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");
        //        driver.AEClick(NotesTab);
        //        //Atos.SyntBots.Common.Logger.Log("Check2", @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");
        //        Thread.Sleep(1000);

        //        // Click Notes Tab  
        //        //Thread.Sleep(1000);

        //        //AutomationElement AENotesTab  = GetPopUpWindowWithWait(notesPane308, 1 , 15);//driver.GetRawChildrenByWlker(notesPane308)[3];
        //        //driver.AEClick(AENotesTab);

        //        #region Commenting on 20-06-2022
        //        //Keyboard.Instance.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
        //        //Keyboard.Instance.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
        //        //Keyboard.Instance.Enter("v");
        //        //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
        //        //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RIGHT);
        //        //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
        //        //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
        //        //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
        //        //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
        //        #endregion
        //        //AutomationElement AE4 = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "View", null, null, null, null); 
        //        //driver.AEClick(AE4); 
        //        ////-----------------------------------------------------------------------------------
        //        // AutomationElement AE5 = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Open View", null, null, null, null); 
        //        //driver.AEClick(AE5); 
        //        ////-----------------------------------------------------------------------------------
        //        // AutomationElement AE6 = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Notes", null, null, null, null); 
        //        //driver.AEClick(AE6); 

        //        //Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.RIGHT);
        //        //Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.RIGHT);
        //        //System.Windows.Rect r  = AENotesTab.Current.BoundingRectangle;
        //        //Thread.Sleep(3000);

        //        //System.Windows.Point pp = r.TopRight; 
        //        //driver.AEClick(AENotesTab);

        //        //            if(AENotesTab.Current.Name=="Notes")
        //        //            {
        //        //             AENotesTab.SetFocus();
        //        //            
        //        //            System.Windows.Rect r  = AENotesTab.Current.BoundingRectangle;
        //        //            Thread.Sleep(3000);
        //        //             
        //        //            System.Windows.Point pp = r.TopRight; 
        //        //            driver.MouseClick(pp);
        //        //            //driver.AEClick(AENotesTab);
        //        //            }

        //        // Click Add Note Button
        //        AutomationElement AddNotesButton = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Add Note", null, null, "button", null);
        //        driver.AEClick(AddNotesButton);
        //       // Atos.SyntBots.Common.Logger.Log("Check3", @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");
        //        Thread.Sleep(1000);

        //        AutomationElement NotesComboedit = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, null, "1001", null, null, null);
        //       // Atos.SyntBots.Common.Logger.Log("NotesComboedit IsNull?" + (NotesComboedit == null), @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");
        //        //driver.AESetText(NotesComboedit, "Encounter");
        //        driver.AEClick(NotesComboedit);
        //        Keyboard.Instance.Enter("Encounter");
        //        //Atos.SyntBots.Common.Logger.Log("Check4", @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");
        //        Thread.Sleep(1000);

        //        AutomationElement NotesDocument = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "", null, null, "document", null);
        //       // Atos.SyntBots.Common.Logger.Log("NotesDocument IsNull?" + (NotesDocument == null), @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");
        //        //driver.AESetText(NotesDocument, dnComment);
        //        driver.AEClick(NotesDocument);
        //        Keyboard.Instance.Enter(dnComment);
        //        Atos.SyntBots.Common.Logger.Log("DN COmment : " + dnComment, @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");

        //        Thread.Sleep(1000);

        //        AutomationElement NotesSaveButton = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Save Note", null, null, "button", null);
        //        driver.AEClick(NotesSaveButton);
        //       // Atos.SyntBots.Common.Logger.Log("Check6", @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");
        //        Thread.Sleep(1000);
        //        //
        //        #region Commented by Velan
        //        //AutomationElement N2notesPane3 = GetPopUpWindowWithWait(_accessHIMWindow.AutomationElement, 3, 20); //driver.GetRawChildrenByWlker(_accessHIMWindow.AutomationElement)[3];
        //        //AutomationElement N2notesPane30 = GetPopUpWindowWithWait(notesPane3, 0, 20); //driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //AutomationElement N2notesPane300 = GetPopUpWindowWithWait(notesPane30, 0, 20);//driver.GetRawChildrenByWlker(notesPane30)[0];
        //        //AutomationElement AEnotesToolbar = driver.GetElmtByWlker(N2notesPane300, "", null, null, "tool bar", null, -1, 20);
        //        #endregion

        //        #region commenting on 20-06-2022
        //        //AutomationElement notesPane3_9 = GetPopUpWindowWithWait(notesPane3_7, 2, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check8", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //AutomationElement notesPane3_10 = GetPopUpWindowWithWait(notesPane3_9, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check9", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //AutomationElement notesPane3_11 = GetPopUpWindowWithWait(notesPane3_10, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check10", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //AutomationElement notesPane3_12 = GetPopUpWindowWithWait(notesPane3_11, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check11", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //AutomationElement notesPane3_13 = GetPopUpWindowWithWait(notesPane3_12, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check12", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //AutomationElement notesPane3_14 = GetPopUpWindowWithWait(notesPane3_13, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check13", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //AutomationElement notesPane3_15 = GetPopUpWindowWithWait(notesPane3_14, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check14", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        ////AutomationElement notesPane3_16 = GetPopUpWindowWithWait(notesPane3_7, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];

        //        //AutomationElement AEnotesToolbar = driver.GetElmtByWlker(notesPane3_15, "", null, null, "tool bar", null, -1, 20);
        //        //Atos.SyntBots.Common.Logger.Log("Check15", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //var AEAddNoteBtn = GetPopUpWindowWithWait(AEnotesToolbar, 0, 20); //driver.GetRawChildrenByWlker(AEnotesToolbar)[0];
        //        //Atos.SyntBots.Common.Logger.Log("AEAddNoteBtn is null? : "+(AEAddNoteBtn==null), @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //driver.AEClick(AEAddNoteBtn);
        //        //Atos.SyntBots.Common.Logger.Log("Check15", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //#endregion
        //        //Thread.Sleep(1000);
        //        // Get Note Type combo box
        //        // 

        //        #endregion

        //        #region Commented by Velan
        //        //AutomationElement AENoteTypeCB = driver.GetElmtByWlker(N2notesPane300, "Note Type: ", null, null, "combo box", null, -1, 15);

        //        //// Click Combo box drop down            
        //        //AutomationElement AENoteTypeDD = driver.GetElmtByWlker(AENoteTypeCB, "Drop Down Button", null, null, "button", null, -1, 15);
        //        //driver.AEClick(AENoteTypeDD);

        //        ////Thread.Sleep(1000);            
        //        //// Select Encounter from DropDown
        //        //AutomationElement AESelectEncounter = driver.GetElmtByWlker(AENoteTypeCB, "Encounter", null, null, "list item", null, -1, 15);
        //        //driver.AEClick(AESelectEncounter);

        //        #endregion


        //        #region commenting on 20-06-2022
        //        //AutomationElement notesPane3_16 = GetPopUpWindowWithWait(notesPane3_13, 1, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check17", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //AutomationElement notesPane3_17 = GetPopUpWindowWithWait(notesPane3_16, 1, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check18", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //AutomationElement notesPane3_18 = GetPopUpWindowWithWait(notesPane3_17, 1, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check19", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //AutomationElement notesPane3_19 = GetPopUpWindowWithWait(notesPane3_18, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check20", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //AutomationElement notesPane3_20 = GetPopUpWindowWithWait(notesPane3_19, 1, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check21", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //AutomationElement notesPane3_21 = GetPopUpWindowWithWait(notesPane3_20, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check22", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //Thread.Sleep(1000);
        //        //driver.AESetText(notesPane3_21, "Encounter");
        //        //Thread.Sleep(1000);

        //        #endregion
        //        // Choose the document to enter text

        //        #region Commented by Velan
        //        //AutomationElement AEDocuments = driver.GetElmtByWlker(N2notesPane300, "", null, null, "document", null, -1, 15);
        //        //driver.AEClick(AEDocuments);

        //        //Keyboard.Instance.Enter(dnComment);
        //        #endregion


        //        #region 20-06-2022
        //        //AutomationElement notesPane3_22 = GetPopUpWindowWithWait(notesPane3_18, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check23", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //Thread.Sleep(1000);
        //        //driver.AEClick(notesPane3_22);
        //        //Keyboard.Instance.Enter(dnComment);
        //        //Thread.Sleep(1000);

        //        #endregion
        //        // Click Notes Savebutton 
        //        /*
        //        AutomationElement AECbParent = TreeWalker.RawViewWalker.GetParent(AENoteTypeCB);
        //        AutomationElement AECbParentParent = TreeWalker.RawViewWalker.GetParent(AECbParent);
        //        var AECbParentParentChld = driver.GetRawChildrenByWlker(AECbParentParent);
        //        var notesToolbar = driver.GetRawChildrenByWlker(AECbParentParentChld[0]);
        //        var AENotesSaveBtn = notesToolbar[0];
        //        driver.AEClick(AENotesSaveBtn);
        //        */
        //        //AutomationElement saveNotesAE = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Save Note", null, null, "button", null, -1, 15);

        //        //driver.AEClick(saveNotesAE);

        //        //Thread.Sleep(1000);

        //        // Click Savebutton
        //        #region commented by velan
        //        //var toolbarSave = GetPopUpWindowWithWait(N2notesPane30, 1, 10);//driver.GetRawChildrenByWlker(N2notesPane30)[1];
        //        //var AESaveBtn = GetPopUpWindowWithWait(toolbarSave, 1, 10);// driver.GetRawChildrenByWlker(toolbarSave)[1];
        //        //driver.AEClick(AESaveBtn);
        //        #endregion


        //        #region commenting on 20-06-2022
        //        //AutomationElement notesPane3_23 = GetPopUpWindowWithWait(notesPane3_17, 1, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check24", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //AutomationElement notesPane3_24 = GetPopUpWindowWithWait(notesPane3_23, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check25", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //AutomationElement notesPane3_25 = GetPopUpWindowWithWait(notesPane3_24, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check26", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //AutomationElement notesPane3_26 = GetPopUpWindowWithWait(notesPane3_25, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
        //        //Atos.SyntBots.Common.Logger.Log("Check27", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
        //        //Thread.Sleep(1000);
        //        //driver.AEClick(notesPane3_26);
        //        //Thread.Sleep(1000);
        //        #endregion'
        //        // Navigate to Save Codes dialog
        //        // 


        //        AutomationElement SaveBtn = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Save", null, null, "button", null);
        //        driver.AEClick(SaveBtn);
        //        //Atos.SyntBots.Common.Logger.Log("Check7", @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");
        //        Thread.Sleep(1000);

        //        AutomationElement AESaveCodesDialog = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Save Codes", null, null, "dialog", null, -1, 15);

        //        var radiobuttons = _accessHIMWindow.GetMultiple<TestStack.White.UIItems.RadioButton>(TestStack.White.UIItems.Finders.SearchCriteria.ByControlType(ControlType.RadioButton));
        //        bool isFinalSelected = radiobuttons.FirstOrDefault(x => x.Name == "Final").IsSelected;
        //        if (isFinalSelected)
        //        {
        //            Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
        //        }

        //        //Atos.SyntBots.Common.Logger.Log("Check8", @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");
        //        Thread.Sleep(1000);


        //        //AutomationElement AEDraftDD = driver.GetElmtByWlker(AESaveCodesDialog, "Draft", "396598", null, "radio button", null);
        //        //driver.AEClick(AEDraftDD);
        //        AutomationElement AEPassReasonTxt = driver.GetElmtByWlker(AESaveCodesDialog, "Pass Reason:", null, null, "edit", null, -1, 15);
        //        if (AEPassReasonTxt == null)
        //        {
        //            AutomationElement aeSaveCodesOk = driver.GetElmtByWlker(AESaveCodesDialog, "Cancel", null, null, "button", null, -1, 15);
        //            //AutomationElement aeSaveCodesOk = driver.GetElmtByWlker(AESaveCodesDialog, "Ok", null, null, "button", null, -1, 15);
        //            driver.AEClick(aeSaveCodesOk);
        //            //Atos.SyntBots.Common.Logger.Log("Check9.1", @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");
        //            throw new Exception("Pass Reason dropdown not found");
        //        }
        //        //driver.AESetText(AEPassReasonTxt, codingPassReason);
        //        driver.AEClick(AEPassReasonTxt);
        //        Keyboard.Instance.Enter(codingPassReason);
        //        //Atos.SyntBots.Common.Logger.Log("Check9", @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");
        //        AutomationElement AEPassReasonDD = driver.GetElmtByWlker(AESaveCodesDialog, "OK", null, null, "button", null, -1, 15);
        //        driver.AEClick(AEPassReasonDD);


        //        //Thread.Sleep(1000);
        //        //Atos.SyntBots.Common.Logger.Log("Check10", @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");
        //        // Navigate to Responsible Pane
        //        AutomationElement AEResponsibleGroup = driver.GetElmtByWlker(AESaveCodesDialog, "Responsible", null, null, "group", null, -1, 15);

        //        //var popup = _accessHIMWindow.MdiChild(SearchCriteria.ByText("Save Codes"));
        //        //var ab = popup.Get<TestStack.White.UIItems.GroupBox>(SearchCriteria.ByControlType(ControlType.Group));
        //        //var abc = ab.Get<TestStack.White.UIItems.TextBox>(SearchCriteria.ByControlType(ControlType.Edit));  
        //        //var test  = (TestStack.White.UIItems.TextBox)ab.Items[6];
        //        //test.Enter("Central, Lab Dn's");

        //        //Atos.SyntBots.Common.Logger.Log("Debug Before Clicking search button", @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");

        //        AutomationElement AEResponsibleBtn = driver.GetElmtByWlker(AEResponsibleGroup, "Begin Search", null, null, "button", null, -1, 15);
        //        driver.AEClick(AEResponsibleBtn);


        //        //Atos.SyntBots.Common.Logger.Log("Debug After Clicking search button", @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");

        //        //Thread.Sleep(1000);            

        //        // Operations in Personnel Search Dialog
        //        /*
        //        AutomationElement AEPersonnelSearchWindow = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Personnel Search", null, null, "Dialog", null);
        //        var child1 = GetRawChildrenByWlkerWithWait(AEPersonnelSearchWindow, 10);//driver.GetRawChildrenByWlker(AEPersonnelSearchWindow);
        //        var child2 = driver.GetRawChildrenByWlker(child1[0]);
        //        var child3 = driver.GetRawChildrenByWlker(child2[0]);
        //        var child4 = driver.GetRawChildrenByWlker(child3[0]);
        //        var child5 = driver.GetRawChildrenByWlker(child4[0]);
        //        var child6 = driver.GetRawChildrenByWlker(child5[0]);
        //        var fname = child6[4];
        //        var lname = child6[5];

        //        var splitName = personResponsible.Split(' ');
        //        driver.AESetText(fname, splitName[1]);
        //        driver.AESetText(lname, splitName[0]);

        //        Thread.Sleep(1000);            

        //        AutomationElement AEPersonnelSearchBtn = driver.GetElmtByWlker(AEPersonnelSearchWindow, "Search", null, null, "button", null, -1, 10);
        //        driver.AEClick(AEPersonnelSearchBtn);
        //        _accessHIMWindow.WaitWhileBusy();
        //        Thread.Sleep(1000);            

        //        AutomationElement AEPersonnelSearchdatagrid = driver.GetElmtByWlker(AEPersonnelSearchWindow, "", null, null, "data grid", null);
        //        var AEperson = driver.GetRawChildrenByWlker(AEPersonnelSearchdatagrid);
        //        driver.AEDoubleClick(AEperson[2]);
        //        */
        //        Thread.Sleep(1000);

        //        AutomationElement AEDescrDocuments = driver.GetElmtByWlker(AESaveCodesDialog, "Description", null, "Edit", "document", null, -1, 15);
        //        driver.AEClick(AEDescrDocuments);
        //        Keyboard.Instance.Enter(dnComment);

        //        Thread.Sleep(1000);
        //        //Atos.SyntBots.Common.Logger.Log("Debug Before Clicking Ok button", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");


        //        AutomationElement AESaveCodesOk = driver.GetElmtByWlker(AESaveCodesDialog, "OK", null, null, "button", null, -1, 15);
        //        InvokePattern aeInvk = AESaveCodesOk.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
        //        aeInvk.Invoke();
        //        //driver.AEClick(AESaveCodesOk);
        //        //Atos.SyntBots.Common.Logger.Log("Debug After Clicking Ok button", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
        //    }
        //    catch (Exception ex)
        //    {
        //        //Atos.SyntBots.Common.Logger.Log("Debug Clicking Ok button" + ex.Message + "" + ex.StackTrace, @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
        //        throw new Exception("Error :" + ex.Message + "line :" + ex.StackTrace);
        //    }

        //    #endregion

        //}

        #endregion

        public void SendDiscernNotification(string dnComment, string codingPassReason, string personResponsible)
        {
            try
            {
                #region Discern Notification

                TestStack.White.UIItems.WindowItems.Window _accessHIMWindow;

                _accessHIMWindow = GetWindowWithWait("AccessHIM", 15); //Desktop.Instance.Windows().Find(K => K.Name.Contains("AccessHIM - "));
                _accessHIMWindow.SetForeground();


                Thread.Sleep(500);



                // Find Notes Pane
                #region Commented by Velan
                //AutomationElement notesPane3 = GetPopUpWindowWithWait(_accessHIMWindow.AutomationElement, 3, 10); //driver.GetRawChildrenByWlker(_accessHIMWindow.AutomationElement)[3];
                //AutomationElement notesPane30 = GetPopUpWindowWithWait(notesPane3, 0, 10); //driver.GetRawChildrenByWlker(notesPane3)[0];
                //AutomationElement notesPane300 = GetPopUpWindowWithWait(notesPane30, 0, 10);//driver.GetRawChildrenByWlker(notesPane30)[0];
                //AutomationElement notesPane308 = GetPopUpWindowWithWait(notesPane30, 8, 10);//driver.GetRawChildrenByWlker(notesPane30)[8];
                #endregion

                #region commenting 20-06-2022
                // AutomationElement notesPane3 = GetPopUpWindowWithWait(_accessHIMWindow.AutomationElement, 3, 10); //driver.GetRawChildrenByWlker(_accessHIMWindow.AutomationElement)[3];
                // AutomationElement notesPane3_1 = GetPopUpWindowWithWait(notesPane3, 0, 10); //driver.GetRawChildrenByWlker(notesPane3)[0];
                // Atos.SyntBots.Common.Logger.Log("Check1", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //// AutomationElement notesPane3_2 = GetPopUpWindowWithWait(notesPane3_1, 0, 10); //driver.GetRawChildrenByWlker(notesPane3)[0];
                //// Atos.SyntBots.Common.Logger.Log("Check2", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                // AutomationElement notesPane3_3 = GetPopUpWindowWithWait(notesPane3_1, 0, 10);
                // Atos.SyntBots.Common.Logger.Log("Check3", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                // AutomationElement notesPane3_4 = GetPopUpWindowWithWait(notesPane3_3, 1, 10);
                // Atos.SyntBots.Common.Logger.Log("Check4", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                // AutomationElement notesPane3_5 = GetPopUpWindowWithWait(notesPane3_4, 0, 10);
                // Atos.SyntBots.Common.Logger.Log("Check5", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                // AutomationElement notesPane3_6 = GetPopUpWindowWithWait(notesPane3_5, 0, 10);
                // Atos.SyntBots.Common.Logger.Log("Check6", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                // AutomationElement notesPane3_7 = GetPopUpWindowWithWait(notesPane3_6, 1, 10);
                // Atos.SyntBots.Common.Logger.Log("Check7", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                // AutomationElement notesPane3_8 = GetPopUpWindowWithWait(notesPane3_7, 5, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                // Atos.SyntBots.Common.Logger.Log("Check8", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                #endregion

               //Atos.SyntBots.Common.Logger.Log("Check1", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                AutomationElement NotesTab;
                try
                {
                    NotesTab = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Notes", null, null, "tab item", null);
                    //Atos.SyntBots.Common.Logger.Log("NotesTab IsNull?" + (NotesTab == null), @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");
                    driver.AEClick(NotesTab);
                }
                catch(Exception e)
                {
                    NotesTab = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Notes (1)", null, null, "tab item", null);
                    //Atos.SyntBots.Common.Logger.Log("NotesTab IsNull?" + (NotesTab == null), @"E:\DiscernNotification\Lansing\TaskQueue\Logs\");
                    driver.AEClick(NotesTab);
                }
                //Atos.SyntBots.Common.Logger.Log("Check2", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                Thread.Sleep(1000);

                // Click Notes Tab  
                //Thread.Sleep(1000);

                //AutomationElement AENotesTab  = GetPopUpWindowWithWait(notesPane308, 1 , 15);//driver.GetRawChildrenByWlker(notesPane308)[3];
                //driver.AEClick(AENotesTab);

                #region Commenting on 20-06-2022
                //Keyboard.Instance.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                //Keyboard.Instance.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.LEFT_ALT);
                //Keyboard.Instance.Enter("v");
                //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
                //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RIGHT);
                //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
                //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
                //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
                //Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
                #endregion
                //AutomationElement AE4 = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "View", null, null, null, null); 
                //driver.AEClick(AE4); 
                ////-----------------------------------------------------------------------------------
                // AutomationElement AE5 = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Open View", null, null, null, null); 
                //driver.AEClick(AE5); 
                ////-----------------------------------------------------------------------------------
                // AutomationElement AE6 = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Notes", null, null, null, null); 
                //driver.AEClick(AE6); 

                //Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.RIGHT);
                //Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.RIGHT);
                //System.Windows.Rect r  = AENotesTab.Current.BoundingRectangle;
                //Thread.Sleep(3000);

                //System.Windows.Point pp = r.TopRight; 
                //driver.AEClick(AENotesTab);

                //            if(AENotesTab.Current.Name=="Notes")
                //            {
                //             AENotesTab.SetFocus();
                //            
                //            System.Windows.Rect r  = AENotesTab.Current.BoundingRectangle;
                //            Thread.Sleep(3000);
                //             
                //            System.Windows.Point pp = r.TopRight; 
                //            driver.MouseClick(pp);
                //            //driver.AEClick(AENotesTab);
                //            }

                // Click Add Note Button
                AutomationElement AddNotesButton = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Add Note", null, null, "button", null);
                driver.AEClick(AddNotesButton);
                //Atos.SyntBots.Common.Logger.Log("Check3", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                Thread.Sleep(1000);

                AutomationElement NotesComboedit = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, null, "1001", null, null, null);
                //Atos.SyntBots.Common.Logger.Log("NotesComboedit IsNull?" + (NotesComboedit == null), @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                //driver.AESetText(NotesComboedit, "Encounter");
                driver.AEClick(NotesComboedit);
                Keyboard.Instance.Enter("Encounter");
                //Atos.SyntBots.Common.Logger.Log("Check4", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                Thread.Sleep(1000);

                AutomationElement NotesDocument = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "", null, null, "document", null);
                //Atos.SyntBots.Common.Logger.Log("NotesDocument IsNull?" + (NotesDocument == null), @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                //driver.AESetText(NotesDocument, dnComment);
                driver.AEClick(NotesDocument);
                Keyboard.Instance.Enter(dnComment);
                //Atos.SyntBots.Common.Logger.Log("Check5", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");

                Thread.Sleep(1000);

                AutomationElement NotesSaveButton = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Save Note", null, null, "button", null);
                driver.AEClick(NotesSaveButton);
                //Atos.SyntBots.Common.Logger.Log("Check6", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                Thread.Sleep(1000);
                //
                #region Commented by Velan
                //AutomationElement N2notesPane3 = GetPopUpWindowWithWait(_accessHIMWindow.AutomationElement, 3, 20); //driver.GetRawChildrenByWlker(_accessHIMWindow.AutomationElement)[3];
                //AutomationElement N2notesPane30 = GetPopUpWindowWithWait(notesPane3, 0, 20); //driver.GetRawChildrenByWlker(notesPane3)[0];
                //AutomationElement N2notesPane300 = GetPopUpWindowWithWait(notesPane30, 0, 20);//driver.GetRawChildrenByWlker(notesPane30)[0];
                //AutomationElement AEnotesToolbar = driver.GetElmtByWlker(N2notesPane300, "", null, null, "tool bar", null, -1, 20);
                #endregion

                #region commenting on 20-06-2022
                //AutomationElement notesPane3_9 = GetPopUpWindowWithWait(notesPane3_7, 2, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check8", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //AutomationElement notesPane3_10 = GetPopUpWindowWithWait(notesPane3_9, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check9", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //AutomationElement notesPane3_11 = GetPopUpWindowWithWait(notesPane3_10, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check10", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //AutomationElement notesPane3_12 = GetPopUpWindowWithWait(notesPane3_11, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check11", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //AutomationElement notesPane3_13 = GetPopUpWindowWithWait(notesPane3_12, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check12", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //AutomationElement notesPane3_14 = GetPopUpWindowWithWait(notesPane3_13, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check13", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //AutomationElement notesPane3_15 = GetPopUpWindowWithWait(notesPane3_14, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check14", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                ////AutomationElement notesPane3_16 = GetPopUpWindowWithWait(notesPane3_7, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];

                //AutomationElement AEnotesToolbar = driver.GetElmtByWlker(notesPane3_15, "", null, null, "tool bar", null, -1, 20);
                //Atos.SyntBots.Common.Logger.Log("Check15", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //var AEAddNoteBtn = GetPopUpWindowWithWait(AEnotesToolbar, 0, 20); //driver.GetRawChildrenByWlker(AEnotesToolbar)[0];
                //Atos.SyntBots.Common.Logger.Log("AEAddNoteBtn is null? : "+(AEAddNoteBtn==null), @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //driver.AEClick(AEAddNoteBtn);
                //Atos.SyntBots.Common.Logger.Log("Check15", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //#endregion
                //Thread.Sleep(1000);
                // Get Note Type combo box
                // 

                #endregion

                #region Commented by Velan
                //AutomationElement AENoteTypeCB = driver.GetElmtByWlker(N2notesPane300, "Note Type: ", null, null, "combo box", null, -1, 15);

                //// Click Combo box drop down            
                //AutomationElement AENoteTypeDD = driver.GetElmtByWlker(AENoteTypeCB, "Drop Down Button", null, null, "button", null, -1, 15);
                //driver.AEClick(AENoteTypeDD);

                ////Thread.Sleep(1000);            
                //// Select Encounter from DropDown
                //AutomationElement AESelectEncounter = driver.GetElmtByWlker(AENoteTypeCB, "Encounter", null, null, "list item", null, -1, 15);
                //driver.AEClick(AESelectEncounter);

                #endregion


                #region commenting on 20-06-2022
                //AutomationElement notesPane3_16 = GetPopUpWindowWithWait(notesPane3_13, 1, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check17", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //AutomationElement notesPane3_17 = GetPopUpWindowWithWait(notesPane3_16, 1, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check18", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //AutomationElement notesPane3_18 = GetPopUpWindowWithWait(notesPane3_17, 1, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check19", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //AutomationElement notesPane3_19 = GetPopUpWindowWithWait(notesPane3_18, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check20", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //AutomationElement notesPane3_20 = GetPopUpWindowWithWait(notesPane3_19, 1, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check21", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //AutomationElement notesPane3_21 = GetPopUpWindowWithWait(notesPane3_20, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check22", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //Thread.Sleep(1000);
                //driver.AESetText(notesPane3_21, "Encounter");
                //Thread.Sleep(1000);

                #endregion
                // Choose the document to enter text

                #region Commented by Velan
                //AutomationElement AEDocuments = driver.GetElmtByWlker(N2notesPane300, "", null, null, "document", null, -1, 15);
                //driver.AEClick(AEDocuments);

                //Keyboard.Instance.Enter(dnComment);
                #endregion


                #region 20-06-2022
                //AutomationElement notesPane3_22 = GetPopUpWindowWithWait(notesPane3_18, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check23", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //Thread.Sleep(1000);
                //driver.AEClick(notesPane3_22);
                //Keyboard.Instance.Enter(dnComment);
                //Thread.Sleep(1000);

                #endregion
                // Click Notes Savebutton 
                /*
                AutomationElement AECbParent = TreeWalker.RawViewWalker.GetParent(AENoteTypeCB);
                AutomationElement AECbParentParent = TreeWalker.RawViewWalker.GetParent(AECbParent);
                var AECbParentParentChld = driver.GetRawChildrenByWlker(AECbParentParent);
                var notesToolbar = driver.GetRawChildrenByWlker(AECbParentParentChld[0]);
                var AENotesSaveBtn = notesToolbar[0];
                driver.AEClick(AENotesSaveBtn);
                */
                //AutomationElement saveNotesAE = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Save Note", null, null, "button", null, -1, 15);

                //driver.AEClick(saveNotesAE);

                //Thread.Sleep(1000);

                // Click Savebutton
                #region commented by velan
                //var toolbarSave = GetPopUpWindowWithWait(N2notesPane30, 1, 10);//driver.GetRawChildrenByWlker(N2notesPane30)[1];
                //var AESaveBtn = GetPopUpWindowWithWait(toolbarSave, 1, 10);// driver.GetRawChildrenByWlker(toolbarSave)[1];
                //driver.AEClick(AESaveBtn);
                #endregion


                #region commenting on 20-06-2022
                //AutomationElement notesPane3_23 = GetPopUpWindowWithWait(notesPane3_17, 1, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check24", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //AutomationElement notesPane3_24 = GetPopUpWindowWithWait(notesPane3_23, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check25", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //AutomationElement notesPane3_25 = GetPopUpWindowWithWait(notesPane3_24, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check26", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //AutomationElement notesPane3_26 = GetPopUpWindowWithWait(notesPane3_25, 0, 10);//driver.GetRawChildrenByWlker(notesPane3)[0];
                //Atos.SyntBots.Common.Logger.Log("Check27", @"E:\DiscernNotification\Lansing\TaskQueue\Logs");
                //Thread.Sleep(1000);
                //driver.AEClick(notesPane3_26);
                //Thread.Sleep(1000);
                #endregion'
                // Navigate to Save Codes dialog
                // 


                AutomationElement SaveBtn = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Save", null, null, "button", null);
                driver.AEClick(SaveBtn);
                //Atos.SyntBots.Common.Logger.Log("Check7", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                Thread.Sleep(1000);

                AutomationElement AESaveCodesDialog = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Save Codes", null, null, "dialog", null, -1, 15);

                var radiobuttons = _accessHIMWindow.GetMultiple<TestStack.White.UIItems.RadioButton>(TestStack.White.UIItems.Finders.SearchCriteria.ByControlType(ControlType.RadioButton));
                bool isFinalSelected = radiobuttons.FirstOrDefault(x => x.Name == "Final").IsSelected;
                if (isFinalSelected)
                {
                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
                }

                //Atos.SyntBots.Common.Logger.Log("Check8, Pass reason :"+codingPassReason, @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                Thread.Sleep(1000);


                //AutomationElement AEDraftDD = driver.GetElmtByWlker(AESaveCodesDialog, "Draft", "396598", null, "radio button", null);
                //driver.AEClick(AEDraftDD);
                //AutomationElement AEPassReasonTxt = driver.GetElmtByWlker(AESaveCodesDialog, "Pass Reason:", null, null, "combo box", null, -1, 15);
                AutomationElement AEPassReasonTxt = driver.GetElmtByWlker(AESaveCodesDialog, "Pass Reason:", null, null, "combo box", null, -1, 15);
                Thread.Sleep(1000);
                if (AEPassReasonTxt == null)
                {
                    AutomationElement aeSaveCodesOk = driver.GetElmtByWlker(AESaveCodesDialog, "Cancel", null, null, "button", null, -1, 15);
                    //AutomationElement aeSaveCodesOk = driver.GetElmtByWlker(AESaveCodesDialog, "Ok", null, null, "button", null, -1, 15);
                    driver.AEClick(aeSaveCodesOk);
                    //Atos.SyntBots.Common.Logger.Log("Check9.1", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                    throw new Exception("Pass Reason dropdown not found");
                }

                //driver.AESetText(AEPassReasonTxt, codingPassReason);
                //Atos.SyntBots.Common.Logger.Log("Check  value pattern1", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                ValuePattern aeText = AEPassReasonTxt.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                //Atos.SyntBots.Common.Logger.Log("Check  value pattern2", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                Thread.Sleep(1000);
                try
                {
                    aeText.SetValue(codingPassReason);
                }
                catch(Exception ex)
                {
                    //Atos.SyntBots.Common.Logger.Log("Exception in Pass Reason Set Text :"+ex.Message, @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                }
                //Atos.SyntBots.Common.Logger.Log("Check  value pattern3", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");

                Thread.Sleep(1000);
                //driver.AEClick(AEPassReasonTxt);
                //Thread.Sleep(1000);
                //Keyboard.Instance.Enter(codingPassReason);
                //Thread.Sleep(3000);

                //AutomationElement AEPassreasonList = driver.GetElmtByWlker(AESaveCodesDialog, "Open", null, null, "button", null, -1, 15);
                //driver.AEClick(AEPassreasonList);
                //Thread.Sleep(1000);
                //AutomationElement AEPassreasonListItem = driver.GetElmtByWlker(AESaveCodesDialog, "Coding - Charge Correction Lab", null, null, "list item", null, -1, 15);
                //driver.AEClick(AEPassreasonListItem);

                //Atos.SyntBots.Common.Logger.Log("Check9", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                AutomationElement AEPassReasonDD = driver.GetElmtByWlker(AESaveCodesDialog, "OK", null, null, "button", null, -1, 15);
                driver.AEClick(AEPassReasonDD);


                //Thread.Sleep(1000);
                //Atos.SyntBots.Common.Logger.Log("Check10", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                // Navigate to Responsible Pane
                AutomationElement AEResponsibleGroup = driver.GetElmtByWlker(AESaveCodesDialog, "Responsible", null, null, "group", null, -1, 15);

                //var popup = _accessHIMWindow.MdiChild(SearchCriteria.ByText("Save Codes"));
                //var ab = popup.Get<TestStack.White.UIItems.GroupBox>(SearchCriteria.ByControlType(ControlType.Group));
                //var abc = ab.Get<TestStack.White.UIItems.TextBox>(SearchCriteria.ByControlType(ControlType.Edit));  
                //var test  = (TestStack.White.UIItems.TextBox)ab.Items[6];
                //test.Enter("Central, Lab Dn's");



                Thread.Sleep(1000);
                //Atos.SyntBots.Common.Logger.Log("Debug Before Clicking search button", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                Thread.Sleep(1000);
                AutomationElement AEResponsibleBtn = driver.GetElmtByWlker(AEResponsibleGroup, "Begin Search", null, null, "button", null, -1, 15);
                Thread.Sleep(1000);
                driver.AEClick(AEResponsibleBtn);


                //Atos.SyntBots.Common.Logger.Log("Debug After Clicking search button", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");

                //Thread.Sleep(1000);            

                // Operations in Personnel Search Dialog
                /*
                AutomationElement AEPersonnelSearchWindow = driver.GetElmtByWlker(_accessHIMWindow.AutomationElement, "Personnel Search", null, null, "Dialog", null);
                var child1 = GetRawChildrenByWlkerWithWait(AEPersonnelSearchWindow, 10);//driver.GetRawChildrenByWlker(AEPersonnelSearchWindow);
                var child2 = driver.GetRawChildrenByWlker(child1[0]);
                var child3 = driver.GetRawChildrenByWlker(child2[0]);
                var child4 = driver.GetRawChildrenByWlker(child3[0]);
                var child5 = driver.GetRawChildrenByWlker(child4[0]);
                var child6 = driver.GetRawChildrenByWlker(child5[0]);
                var fname = child6[4];
                var lname = child6[5];

                var splitName = personResponsible.Split(' ');
                driver.AESetText(fname, splitName[1]);
                driver.AESetText(lname, splitName[0]);

                Thread.Sleep(1000);            

                AutomationElement AEPersonnelSearchBtn = driver.GetElmtByWlker(AEPersonnelSearchWindow, "Search", null, null, "button", null, -1, 10);
                driver.AEClick(AEPersonnelSearchBtn);
                _accessHIMWindow.WaitWhileBusy();
                Thread.Sleep(1000);            

                AutomationElement AEPersonnelSearchdatagrid = driver.GetElmtByWlker(AEPersonnelSearchWindow, "", null, null, "data grid", null);
                var AEperson = driver.GetRawChildrenByWlker(AEPersonnelSearchdatagrid);
                driver.AEDoubleClick(AEperson[2]);
                */
                Thread.Sleep(1000);

                AutomationElement AEDescrDocuments = driver.GetElmtByWlker(AESaveCodesDialog, "Description", null, "Edit", "document", null, -1, 15);
                driver.AEClick(AEDescrDocuments);
                Keyboard.Instance.Enter(dnComment);

                Thread.Sleep(1000);
                //Atos.SyntBots.Common.Logger.Log("Debug Before Clicking Ok button", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");


                AutomationElement AESaveCodesOk = driver.GetElmtByWlker(AESaveCodesDialog, "OK", null, null, "button", null, -1, 15);
                InvokePattern aeInvk = AESaveCodesOk.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                aeInvk.Invoke();
                //driver.AEClick(AESaveCodesOk);
                //Atos.SyntBots.Common.Logger.Log("Debug After Clicking Ok button", @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
            }
            catch (Exception ex)
            {
                //Atos.SyntBots.Common.Logger.Log("Debug Clicking Ok button" + ex.Message + "" + ex.StackTrace, @"E:\DiscernNotification\CentralOakland\TaskQueue\Logs\");
                throw new Exception("Error :" + ex.Message + "line :" + ex.StackTrace);
            }

            #endregion

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
