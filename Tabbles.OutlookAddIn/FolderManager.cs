using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;

namespace Tabbles.OutlookAddIn
{
    class FolderManager
    {
        //private Folders tmpFolders;
        //private System.Action folderRemoveCallback;
        //private object folderRemoveLock = new object();

        //public void RemoveFolderByName(Folders folders, string name, System.Action callback = null)
        //{
        //    lock (this.folderRemoveLock)
        //    {
        //            if (folders != null)
        //            {
        //                Folder folder = null;
        //                try
        //                {
        //                    folder = (Folder)folders[name];
        //                }
        //                catch (System.Exception)
        //                {
        //                    //this means that the folder wasn't found
        //                }

        //                if (folder != null)
        //                {
        //                    if (callback == null)
        //                    {
        //                        folder.Delete();
        //                    }
        //                    else
        //                    {
        //                        this.tmpFolders = folders;
        //                        this.folderRemoveCallback = callback;
        //                        this.tmpFolders.FolderRemove += FolderRemoved;
        //                        folder.Delete();
        //                    }
        //                }
        //                else if (callback != null)
        //                {
        //                    callback();
        //                }
        //            }
        //    }
        //}

        //private void FolderRemoved()
        //{
        //    lock (this.folderRemoveLock)
        //    {
        //        if (this.tmpFolders != null)
        //        {
        //            this.tmpFolders.FolderRemove -= FolderRemoved;

        //            if (this.folderRemoveCallback != null)
        //            {
        //                this.folderRemoveCallback();
        //            }
        //        }
        //    }
        //}
    }
}
