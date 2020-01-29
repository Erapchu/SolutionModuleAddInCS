using System;
using System.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Threading;

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
            myUserControl1 = new MyUserControl();
            myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "My Task Pane");
            myCustomTaskPane.Visible = !myCustomTaskPane.Visible;

            explorer = Application.ActiveExplorer();
            explorer.BeforeFolderSwitch += Explorer_BeforeFolderSwitch;
            explorer.FolderSwitch += Explorer_FolderSwitch;

            //Call EnsureSolutionsModule to ensure that
            //Solutions module and custom folder icons
            //appear in Outlook Navigation Pane
            EnsureSolutionsModule();
        }

        private void Explorer_FolderSwitch()
        {
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
            if (switchedFolder != null && (switchedFolder.Parent as Outlook.Folder).EntryID == solutionEntryId)
            {
                System.Windows.Forms.MessageBox.Show($"You switch this folder: {switchedFolder.Name}");
            }
                
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //If needed, your cleanup code goes here
        }

        private void EnsureSolutionsModule()
        {
            try
            {
                //Declarations
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

                    for (int i = 0; i < 10; i++)
                    {
                        subFoldersList.Add(solutionRoot.Folders.Add(
                        $"Location {i}",
                        Outlook.OlDefaultFolders.olFolderInbox)
                        as Outlook.Folder);
                    }
                }
                else
                {
                    solutionRoot =
                        rootStoreFolder.Folders["All locations (test)"]
                        as Outlook.Folder;
                    for (int i = 0; i < 10; i++)
                    {
                        subFoldersList.Add(solutionRoot.Folders[$"Location {i}"] as Outlook.Folder);
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
                if (solutionsModule.Position != 5)
                {
                    //Move SolutionsModule to Position = 5
                    solutionsModule.Position = 5;
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
