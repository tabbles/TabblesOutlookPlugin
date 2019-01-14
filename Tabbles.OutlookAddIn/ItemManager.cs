using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;

namespace Tabbles.OutlookAddIn
{
    class ItemManager
    {
        struct ItemsActionStruct
        {
            public Items Items { get; set; }
            public System.Action Action { get; set; }
        }

        Dictionary<string, ItemsActionStruct> itemsActionsDict = new Dictionary<string, ItemsActionStruct>();
        Dictionary<string, List<string>> mailsRemaining = new Dictionary<string, List<string>>();

        public static readonly object LockObj = new object();

        ///// <param name="callbackAction"></param>
        //public void SaveMailWithCallback(MailItem mail, System.Action callbackAction)
        //{
        //    lock (LockObj)
        //    {
        //        Folder folder = (Folder)mail.Parent;
        //        if (folder != null && folder.Items != null)
        //        {
        //            string folderId = folder.EntryID;
        //            if (this.itemsActionsDict.ContainsKey(folderId))
        //            {
        //                if (this.mailsRemaining.ContainsKey(folderId)) //100% should be true
        //                {
        //                    this.mailsRemaining[folderId].Add(mail.EntryID);
        //                }
        //            }
        //            else
        //            {
        //                ItemsActionStruct itemsAction = new ItemsActionStruct()
        //                {
        //                    Items = folder.Items,
        //                    Action = callbackAction
        //                };
        //                this.itemsActionsDict.Add(folderId, itemsAction);
        //                itemsAction.Items.ItemChange += NotifyItemChanged;

        //                this.mailsRemaining[folderId] = new List<string>() { mail.EntryID };
        //            }
        //        }
        //    }

        //    mail.Save();
        //}

        //private void NotifyItemChanged(object item)
        //{
        //    lock (LockObj)
        //    {
        //        MailItem mail = item as MailItem;
        //        if (mail != null)
        //        {
        //            Folder folder = (Folder)mail.Parent;
        //            if (folder != null)
        //            {
        //                string folderId = folder.EntryID;
        //                if (this.mailsRemaining.ContainsKey(folderId))
        //                {
        //                    List<string> mailIds = this.mailsRemaining[folderId];
        //                    mailIds.Remove(mail.EntryID);
        //                    if (mailIds.Count == 0)
        //                    {
        //                        if (this.itemsActionsDict.ContainsKey(folderId))
        //                        {
        //                            ItemsActionStruct itemsAction = this.itemsActionsDict[folderId];
        //                            itemsAction.Items.ItemChange -= NotifyItemChanged;
        //                            itemsAction.Action();
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //}
    }
}
