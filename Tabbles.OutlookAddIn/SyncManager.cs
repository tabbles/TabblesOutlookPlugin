using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.Office.Interop.Outlook;
using Win = System.Windows;
using Tabbles.OutlookAddIn.Controls;
using Res = Tabbles.OutlookAddIn.Properties.Resources;

namespace Tabbles.OutlookAddIn
{
    public class SyncManager
    {
        private const string CategorizedMailItemFilter = @"@SQL=""urn:schemas-microsoft-com:office:office#Keywords"" like '%'";
        private const string MessageClassIpmNote = "IPM.Note";


        public MenuManager mMenuManager;

        // SUJAYXML
       // private XMLFileManager xmlFileManager;

        //private Folders rootFolders;

        //private bool cancel;

        //public event Action<List<string>> SendEmailCategories;

        public bool InProcess
        {
            get;
            private set;
        }

        public SyncManager()
        {
            //this.rootFolders = rootFolders;

            // SUJAYXML
           // xmlFileManager = new XMLFileManager();
        }

        //public System.Action GetSyncAction()
        //{
        //    System.Action action = new System.Action(() =>
        //    {
        //        if (InProcess)
        //        {
        //            System.Windows.MessageBox.Show(Res.MsgSyncIsRunning);
        //            return;
        //        }

        //        InProcess = true;

        //        PromptDialog confirmationWindow = new PromptDialog()
        //        {
        //            OkText = Res.LabelOk,
        //            CancelText = Res.LabelCancel,
        //            Message = Res.MsgDoYouWantToSync,
        //            DontShowAgainMessage = Res.MsgDontAskAgain,

        //             //SUJAYXML
        //          //  WasDontAskAgain = xmlFileManager.IsDontAskForSync()
        //           //  WasDontAskAgain = RegistryManager.IsDontAskForSync()
        //        };

        //        bool? answer = confirmationWindow.ShowDialog();

        //        RegistryManager.SetDontAskForSync(confirmationWindow.IsDontAskAgain);

        //        //SUJAYXML
        //        //xmlFileManager.SetDontAskForSync(confirmationWindow.IsDontAskAgain);

        //        if (answer.HasValue && answer.Value)
        //        {
        //            SyncProgress progressWindow = new SyncProgress(Run);
        //            //progressWindow.Cancel += OnCancel;
        //            progressWindow.ShowDialog();
        //        }
        //        else
        //        {
        //            InProcess = false;
        //        }
        //    });

        //    return action;
        //}

        //private void Run()
        //{
        //    try
        //    {
        //        foreach (Folder folder in this.rootFolders)
        //        {
        //            //if (this.cancel)
        //            //{
        //            //    break;
        //            //}

        //            int totalCount = SyncCategorizedItems(folder);
        //            Log.log(string.Format("{0} items were synced in '{1}' folder.", totalCount, folder.Name));

        //            Thread.Sleep(20); //Take a little break to make Outlook responsive
        //        }

        //        InProcess = false;
        //        RegistryManager.SetSyncPerformed(true);

        //        // SUJAYXML
        //       // xmlFileManager.SetSyncPerformed(true);

        //    }
        //    catch (System.Exception)
        //    {
        //        //do nothing
        //    }
        //}

        /// <summary>
        /// Synchonized categorized items with Tabbles and returns synced items total count.
        /// </summary>
        /// <param name="folder"></param>
        /// <returns></returns>
        //private int SyncCategorizedItems(Folder folder = null)
        //{
        //    int totalCount = 0;

        //    List<string> entryIds = null;
        //    Folders subFolders;
        //    if (folder != null)
        //    {
        //        Table table = folder.GetTable(CategorizedMailItemFilter);
        //        Row row;
        //        while ((row = table.GetNextRow()) != null)
        //        {
        //            //if (this.cancel)
        //            //{
        //            //    break;
        //            //}

        //            object[] values = (object[])row.GetValues();
        //            if (values.Length == 5 && values[0] is string &&
        //                MessageClassIpmNote.Equals(values[4]))
        //            {
        //                if (entryIds == null)
        //                {
        //                    entryIds = new List<string>();
        //                }

        //                entryIds.Add((string)values[0]);
        //            }

        //            Thread.Sleep(20); //Take a little break to make Outlook responsive
        //        }

        //        subFolders = folder.Folders;
        //    }
        //    else
        //    {
        //        subFolders = this.rootFolders;
        //    }

        //    if (entryIds != null && entryIds.Count > 0 )
        //    {
        //        //mMenuManager.SendEmailCategories(entryIds);
        //        //SendEmailCategories(entryIds);

        //        totalCount += entryIds.Count;
        //    }

        //    foreach (Folder subFolder in subFolders)
        //    {
        //        //if (this.cancel)
        //        //{
        //        //    break;
        //        //}

        //        totalCount += SyncCategorizedItems(subFolder);

        //        Thread.Sleep(20); //Take a little break to make Outlook responsive
        //    }

        //    return totalCount;
        //}

        //private void OnCancel(object sender, EventArgs e)
        //{
        //    this.cancel = true;
        //}
    }
}
