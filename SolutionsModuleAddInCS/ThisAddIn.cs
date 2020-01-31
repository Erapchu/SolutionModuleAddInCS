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
        Outlook.Folder switchedFolder;
        string solutionEntryId;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        private MyUserControl myUserControl1;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //InitTaskPane();

            explorer = Application.ActiveExplorer();
            explorer.BeforeFolderSwitch += Explorer_BeforeFolderSwitch;
            explorer.FolderSwitch += Explorer_FolderSwitch;

            var inspector = Application.ActiveInspector();
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
            //ReplaceIE();
        }

        private static string windowClassName = "rctrl_renwnd32";
        private static string placeholderClassName = "Internet Explorer_Server";

        private void ReplaceIE()
        {
            IntPtr hBuiltInWindow = WinApiProvider.FindWindow(windowClassName, null);
            if (hBuiltInWindow != IntPtr.Zero)
            {
                List<IntPtr> childWindows = WinApiProvider.EnumChildWindows(hBuiltInWindow);
                int childIndex = WinApiProvider.FindChildByClassName(childWindows, placeholderClassName);
                myUserControl1 = new MyUserControl();
                IntPtr hWnd = childWindows[childIndex];
                //WinApiProvider.ShowWindow(hWnd, WinApiProvider.SW_HIDE);
                myUserControl1.Show();
            }
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
            if (switchedFolder != null && (switchedFolder.Parent as Outlook.Folder).EntryID == solutionEntryId)
            {
                ReplaceIE();
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
                int foldersCount = 10;
                List<Outlook.Folder> subFoldersList = new List<Outlook.Folder>();
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
                        rootStoreFolder.Folders["All locations (test)"]
                        as Outlook.Folder;
                }
                catch
                {
                    firstRun = true;
                }

                if (firstRun == true)
                {
                    solutionRoot =
                        rootStoreFolder.Folders.Add("All locations (test)",
                        Outlook.OlDefaultFolders.olFolderInbox)
                        as Outlook.Folder;

                    Outlook.OlDefaultFolders folderType = Outlook.OlDefaultFolders.olFolderInbox;

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
                    }
                }
                else
                {
                    solutionRoot =
                        rootStoreFolder.Folders["All locations (test)"]
                        as Outlook.Folder;

                    for (int i = 0; i < foldersCount; i++)
                        try
                        {
                            subFoldersList.Add(solutionRoot.Folders[$"Location {i}"] as Outlook.Folder);
                        }
                        catch (Exception ex)
                        {
                            Debug.Write(ex.Message);
                        }
                }

                solutionEntryId = solutionRoot.EntryID;

                //Get the icons for the solution
                stdole.StdPicture rootPict =
                    PictureDispConverter.ToIPictureDisp(
                    Properties.Resources.folder)
                    as stdole.StdPicture;
                //Set the icons for solution folders
                solutionRoot.SetCustomIcon(rootPict);
                subFoldersList.ForEach(f => f.SetCustomIcon(rootPict));

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
