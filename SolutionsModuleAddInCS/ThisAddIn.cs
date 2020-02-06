using System;
using System.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace SolutionsModuleAddInCS
{
    public partial class ThisAddIn
    {
        Outlook.SolutionsModule solutionsModule;
        Outlook.Explorer explorer;
        IntPtr hwndExplorer = IntPtr.Zero;
        Outlook.Folder switchedFolder;
        string solutionEntryId;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        private EmptyUserControl emptyUserControl;
        private MyUserControl myUserControl1;
        private Form1 form1;
        private MainForm mainForm;
        WinApiSubClass shellWinApiClass;
        WinApiSubClass leftPaneWinApiClass;
        Window1 window1;
        Window2 window2;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //InitTaskPane();

            explorer = Application.ActiveExplorer();
            explorer.BeforeFolderSwitch += Explorer_BeforeFolderSwitch;
            explorer.FolderSwitch += Explorer_FolderSwitch;
            hwndExplorer = WinApiProvider.GetExplorerWindowHandle(explorer);

            //var inspector = Application.ActiveInspector();
            //inspector.NewFormRegion();

            //explorer.ShowPane(Outlook.OlPane.olFolderList, false);
            //explorer.ShowPane(Outlook.OlPane.olNavigationPane, false);
            //explorer.ShowPane(Outlook.OlPane.olOutlookBar, false);
            //explorer.ShowPane(Outlook.OlPane.olPreview, false);
            //explorer.ShowPane(Outlook.OlPane.olToDoBar, false);

            //Call EnsureSolutionsModule to ensure that
            //Solutions module and custom folder icons
            //appear in Outlook Navigation Pane
            EnsureSolutionsModule();

            //Microsoft.Office.Tools.Outlook.FormRegionType
        }

        private static string outlookClassName = "rctrl_renwnd32";
        private static string internetExplorerClassName = "Internet Explorer_Server";
        private static string shellEmbeddingClassName = "Shell Embedding";
        private static string netUINativeClassName = "NetUINativeHWNDHost";

        private IntPtr GetHWNDInExplorer(string className, int? controlID = null)
        {
            IntPtr hWnd;
            bool flag = controlID != null;

            //hWnd = WinApiProvider.FindWindowEx(hwndExplorer, IntPtr.Zero, className, string.Empty);
            //if (hWnd != IntPtr.Zero)
            //    return hWnd;

            List<IntPtr> childWindows = WinApiProvider.EnumChildWindows(hwndExplorer);
            int targetIndex = WinApiProvider.FindChildByClassName(childWindows, className);
            IntPtr targetHWnd = childWindows[targetIndex];
            var parentHWnd = WinApiProvider.GetParent(targetHWnd);
            while (parentHWnd != hwndExplorer && flag || parentHWnd != hwndExplorer)
            {
                childWindows.RemoveAt(targetIndex);
                targetIndex = WinApiProvider.FindChildByClassName(childWindows, className);
                targetHWnd = childWindows[targetIndex];
                parentHWnd = WinApiProvider.GetParent(targetHWnd);
                if (controlID != null)
                {
                    var cID = WinApiProvider.GetDlgCtrlID(targetHWnd);
                    flag = cID == controlID;
                }
            }
            hWnd = targetHWnd;

            return hWnd;
        }

        private void SetChildWindowStyle(IntPtr windowHWND)
        {
            var style = WinApiProvider.GetWindowLong(windowHWND, WinApiProvider.GWL_STYLE);
            style = (style & ~WinApiProvider.WS_POPUP & ~WinApiProvider.WS_OVERLAPPEDWINDOW) | WinApiProvider.WS_CHILD;
            //var style = WinApiProvider.WS_CHILD | WinApiProvider.WS_CLIPSIBLINGS | WinApiProvider.WS_CLIPCHILDREN | WinApiProvider.WS_EX_CONTROLPARENT | WinApiProvider.WS_VISIBLE;
            WinApiProvider.SetWindowLong(windowHWND, WinApiProvider.GWL_STYLE, style);
        }

        private void mainForm_WndProc(ref Message m)
        {
            if (m.Msg == WinApiProvider.WM_SIZE)
            {
                int lParam = m.LParam.ToInt32();
                System.Drawing.Size newSize = new System.Drawing.Size(lParam & 0xFFFF, (int)(lParam & 0xFFFF0000) / 0x10000);
                if (window1 != null)
                {
                    window1.Width = newSize.Width;
                    window1.Height = newSize.Height;
                }
            }
        }

        private void leftPaneForm_WndProc(ref Message m)
        {
            if (m.Msg == WinApiProvider.WM_SIZE)
            {
                int lParam = m.LParam.ToInt32();
                System.Drawing.Size newSize = new System.Drawing.Size(lParam & 0xFFFF, (int)(lParam & 0xFFFF0000) / 0x10000);
                if (window2 != null)
                {
                    window2.Width = newSize.Width;
                    window2.Height = newSize.Height;
                }
            }
        }

        private void ReplaceIE()
        {

            SetThreadDPIContext(hwndExplorer);

            //var dpi = WinApiProvider.GetDpiForWindow(form1.Handle);
            //dpi = WinApiProvider.GetDpiForWindow(targetHWnd);

            /*if (window is null)
                window = new Window1();
            var wih = new System.Windows.Interop.WindowInteropHelper(window);
            IntPtr windowHWND = wih.EnsureHandle();*/

            //if (mainForm is null)
            //    mainForm = new MainForm();

            if (window1 is null)
                window1 = new Window1();
            var wih = new System.Windows.Interop.WindowInteropHelper(window1);
            IntPtr window1HWND = wih.EnsureHandle();
            var shellHWnd = GetHWNDInExplorer(shellEmbeddingClassName);
            var ph = WinApiProvider.SetParent(window1HWND, shellHWnd);
            if (shellWinApiClass is null)
            {
                shellWinApiClass = new WinApiSubClass(shellHWnd);
                shellWinApiClass.CallbackProc += mainForm_WndProc;
            }

            Rect tempRect = new Rect();
            WinApiProvider.GetWindowRect(shellHWnd, ref tempRect);
            window1.Width = tempRect.right - tempRect.left;
            window1.Height = tempRect.bottom - tempRect.top;
            //mainForm.Location = new System.Drawing.Point(0, 0);
            //mainForm.Size = new System.Drawing.Size(tempRect.right - tempRect.left, tempRect.bottom - tempRect.top);
            SetChildWindowStyle(window1HWND);
            window1.Show();


            if (window2 is null)
                window2 = new Window2();
            wih = new System.Windows.Interop.WindowInteropHelper(window2);
            IntPtr window2HWND = wih.EnsureHandle();
            var leftPaneHWND = GetHWNDInExplorer(netUINativeClassName, 0x67);
            ph = WinApiProvider.SetParent(window2HWND, leftPaneHWND);
            if (leftPaneWinApiClass is null)
            {
                leftPaneWinApiClass = new WinApiSubClass(leftPaneHWND);
                leftPaneWinApiClass.CallbackProc += leftPaneForm_WndProc;
            }

            WinApiProvider.GetWindowRect(leftPaneHWND, ref tempRect);
            window2.Width = tempRect.right - tempRect.left;
            window2.Height = tempRect.bottom - tempRect.top;
            //form1.Location = new System.Drawing.Point(0, 0);
            //form1.Size = new System.Drawing.Size(tempRect.right - tempRect.left, tempRect.bottom - tempRect.top);
            SetChildWindowStyle(window2HWND);
            window2.Show();

            WinApiProvider.SetFocus(hwndExplorer);

            var a = Marshal.GetLastWin32Error();

            /*var t = WinApiProvider.SetParent(emptyUserControl.Handle, parentHWnd);
            emptyUserControl.Visible = true;
            emptyUserControl.Show();
            var emptyUCParentHWnd = WinApiProvider.GetParent(emptyUserControl.Handle);

            if (myUserControl1 is null)
                myUserControl1 = new MyUserControl();

            WinApiProvider.SetParent(myUserControl1.Handle, emptyUserControl.Handle);
            myUserControl1.Visible = true;
            var myUCParentHWnd = WinApiProvider.GetParent(myUserControl1.Handle);

            myUserControl1.Show();*/

            //WinApiProvider.ShowWindow(hWnd, WinApiProvider.SW_HIDE);
            //myUserControl1.Show();
        }

        public int SetThreadDPIContext(IntPtr contextWindow)
        {
            int num = -1;
            num = WinApiProvider.GetThreadDpiAwarenessContext();
            int num2 = WinApiProvider.GetWindowDpiAwarenessContext(contextWindow);
            if (num != num2 && WinApiProvider.IsValidDpiAwarenessContext(num2))
            {
                WinApiProvider.SetThreadDpiAwarenessContext(num2);
            }
            return num;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //If needed, your cleanup code goes here
        }

        #region Testing
        private void InitTaskPane()
        {
            myUserControl1 = new MyUserControl();
            myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "My Task Pane");
            myCustomTaskPane.Visible = !myCustomTaskPane.Visible;
        }

        private void Explorer_FolderSwitch()
        {
            if (switchedFolder != null && switchedFolder.EntryID == solutionEntryId && switchedFolder.WebViewOn)
            {
                ReplaceIE();
            }
            else
            {
                if (window1 != null && window1.IsVisible)
                    window1.Hide();
                if (window2 != null && window2.IsVisible)
                    window2.Hide();
            }

            /*if (switchedFolder != null)
            {
                Outlook.MailItem newMailItem = Application.CreateItem(Outlook.OlItemType.olMailItem);
                newMailItem.Body = "Hello fvckbldfbv dsf";
                newMailItem.Recipients.Add("Outlook");
                newMailItem.Subject = "Following";
                switchedFolder.Items.Add(newMailItem);
            }*/
        }

        private void Explorer_BeforeFolderSwitch(object NewFolder, ref bool Cancel)
        {
            switchedFolder = NewFolder as Outlook.Folder;

            //if (switchedFolder != null && (switchedFolder.Parent as Outlook.Folder).EntryID == solutionEntryId)
            //{
            //    switchedFolder.WebViewURL = "https://www.microsoft.com";
            //    switchedFolder.WebViewOn = true;
            //}

            /*switchedFolder = NewFolder as Outlook.Folder;
            if (switchedFolder != null && (switchedFolder.Parent as Outlook.Folder).EntryID == solutionEntryId)
            {
                Process[] processes = Process.GetProcessesByName("OUTLOOK");
                var proc = processes[0];
                IntPtr pointer = proc.MainWindowHandle;
                Rect1 rect = new Rect1();
                GetWindowRect(pointer, ref rect);
                Form newForm = new Form();
                newForm.Show();
                newForm.Top = rect.Top;
                newForm.Left = rect.Left;
            }*/

            /*switchedFolder = NewFolder as Outlook.Folder;
            if (switchedFolder != null && (switchedFolder.Parent as Outlook.Folder).EntryID == solutionEntryId)
            {
                System.Windows.Forms.MessageBox.Show($"You switch this folder: {switchedFolder.Name}");
            }*/
        }

        /* switch (i)
            {
                case 0:
                    folderType = Outlook.OlDefaultFolders.olFolderCalendar;
                    break;
                case 1:
                    folderType = Outlook.OlDefaultFolders.olFolderConflicts;
                    break;
                case 2:
                    folderType = Outlook.OlDefaultFolders.olFolderContacts;
                    break;
                case 3:
                    folderType = Outlook.OlDefaultFolders.olFolderDeletedItems;
                    break;
                case 4:
                    folderType = Outlook.OlDefaultFolders.olFolderDrafts;
                    break;
                case 5:
                    folderType = Outlook.OlDefaultFolders.olFolderInbox;
                    break;
                case 6:
                    folderType = Outlook.OlDefaultFolders.olFolderJournal;
                    break;
                case 7:
                    folderType = Outlook.OlDefaultFolders.olFolderJunk;
                    break;
                case 8:
                    folderType = Outlook.OlDefaultFolders.olFolderLocalFailures;
                    break;
                case 9:
                    folderType = Outlook.OlDefaultFolders.olFolderManagedEmail;
                    break;
                case 10:
                    folderType = Outlook.OlDefaultFolders.olFolderNotes;
                    break;
                case 11:
                    folderType = Outlook.OlDefaultFolders.olFolderOutbox;
                    break;
                case 12:
                    folderType = Outlook.OlDefaultFolders.olFolderRssFeeds;
                    break;
                case 13:
                    folderType = Outlook.OlDefaultFolders.olFolderSentMail;
                    break;
                case 14:
                    folderType = Outlook.OlDefaultFolders.olFolderServerFailures;
                    break;
                case 15:
                    folderType = Outlook.OlDefaultFolders.olFolderSuggestedContacts;
                    break;
                case 16:
                    folderType = Outlook.OlDefaultFolders.olFolderSyncIssues;
                    break;
                case 17:
                    folderType = Outlook.OlDefaultFolders.olFolderTasks;
                    break;
                case 18:
                    folderType = Outlook.OlDefaultFolders.olFolderToDo;
                    break;
                case 19:
                    folderType = Outlook.OlDefaultFolders.olPublicFoldersAllPublicFolders;
                    break;
            }
             */
        #endregion

        private void EnsureSolutionsModule()
        {
            try
            {
                //Declarations
                //int foldersCount = 10;
                //List<Outlook.Folder> subFoldersList = new List<Outlook.Folder>();
                Outlook.Folder solutionRoot;
                bool firstRun = false;
                Outlook.Folder rootStoreFolder =
                    Application.Session.DefaultStore.GetRootFolder()
                    as Outlook.Folder;
                //If solution root folder does not exist, create it
                //Note that solution root 
                //could also be in PST or custom store
                try
                {
                    solutionRoot =
                        rootStoreFolder.Folders["Search (test)"]
                        as Outlook.Folder;
                }
                catch
                {
                    firstRun = true;
                }

                if (firstRun == true)
                {
                    solutionRoot =
                        rootStoreFolder.Folders.Add("Search (test)",
                        Outlook.OlDefaultFolders.olFolderInbox)
                        as Outlook.Folder;

                    /*Outlook.OlDefaultFolders folderType = Outlook.OlDefaultFolders.olFolderInbox;

                    for (int i = 0; i < foldersCount; i++)
                    {
                        try
                        {
                            subFoldersList.Add(solutionRoot.Folders.Add(
                            $"Location {i}",
                            folderType)
                            as Outlook.Folder);
                        }
                        catch (Exception ex)
                        {
                            Debug.Write(ex.Message);
                        }
                    }*/
                }
                else
                {
                    solutionRoot =
                        rootStoreFolder.Folders["Search (test)"]
                        as Outlook.Folder;

                    /*for (int i = 0; i < foldersCount; i++)
                        try
                        {
                            subFoldersList.Add(solutionRoot.Folders[$"Location {i}"] as Outlook.Folder);
                        }
                        catch (Exception ex)
                        {
                            Debug.Write(ex.Message);
                        }*/
                }

                solutionEntryId = solutionRoot.EntryID;
                if (!solutionRoot.WebViewOn || solutionRoot.WebViewURL == string.Empty)
                {
                    solutionRoot.WebViewURL = @"file:\\\C:\Users\Andrey\AppData\Local\Temp\AddinExpress\ADXOlFormGeneral.html";
                    solutionRoot.WebViewOn = true;
                }

                //Get the icons for the solution
                //stdole.StdPicture rootPict =
                //    PictureDispConverter.ToIPictureDisp(
                //    Properties.Resources.folder)
                //    as stdole.StdPicture;
                //Set the icons for solution folders
                //solutionRoot.SetCustomIcon(rootPict);
                //subFoldersList.ForEach(f => f.SetCustomIcon(rootPict));

                //Obtain a reference to the SolutionsModule
                solutionsModule =
                    explorer.NavigationPane.Modules.GetNavigationModule(
                    Outlook.OlNavigationModuleType.olModuleSolutions)
                    as Outlook.SolutionsModule;
                //Add the solution and hide folders in default modules
                solutionsModule.AddSolution(solutionRoot,
                    Outlook.OlSolutionScope.olHideInDefaultModules);
                //The following code sets the position and visibility
                //of the SolutionsModule
                if (solutionsModule.Visible == false)
                {
                    //Set Visibile to true
                    solutionsModule.Visible = true;
                }
                if (solutionsModule.Position != 1)
                {
                    //Move SolutionsModule to Position = 5
                    solutionsModule.Position = 1;
                }
                //Create instance variable for Outlook.NavigationPane
                Outlook.NavigationPane navPane = explorer.NavigationPane;
                if (navPane.DisplayedModuleCount != 5)
                {
                    //Ensure that Solutions Module button is large
                    navPane.DisplayedModuleCount = 5;
                }
            }
            catch (Exception ex)
            {
                Debug.Write(ex.Message);
            }
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
