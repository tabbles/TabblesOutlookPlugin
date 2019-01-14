using System;
using System.Collections.Generic;
using System.IO.Pipes;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.FSharp.Core;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Action = System.Action;
using Exception = System.Exception;
using Res = Tabbles.OutlookAddIn.Properties.Resources;
// ReSharper disable CollectionNeverQueried.Local


namespace Tabbles.OutlookAddIn
{
    public partial class ThisAddIn
    {

#if TABBLES
        public static bool isConfidential = false;
#else
        public static bool isConfidential = true;
#endif

        private static string emptyStringOfNull(string s)
        {
            return (s ?? "-null-");
        }

        private static readonly string nl = Environment.NewLine + Environment.NewLine;

        public static string AssemblyVer_safe()
        {
            string assemblyVer;
            try
            {
                assemblyVer = "outlook: " + Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
            catch (Exception e)
            {
                assemblyVer = "outlook: Error in assembly ver: " + e.GetType().ToString() + " --- " + e.Message + nl;

            }
            return assemblyVer;
        }

        //public static string preamboloEmail_safe()
        //{
        //    try
        //    {

        //        return "outlook plugin: username = " + emptyStringOfNull(Environment.UserName) + ", machine = " + emptyStringOfNull(Environment.MachineName) + ", ver = " + AssemblyVer_safe() + nl;
        //    }
        //    catch
        //    {
        //        return "errore in preambolo outlook plugin" + nl;
        //    }
        //}

        private static readonly string[] OutlookCmdSeparator = new string[] { @"/select outlook:" };



        public static FSharpFunc<string, Unit> logError = FuncConvert.ToFSharpFunc<string>(a => Log.log(a));
        private MenuManager menuManager;
        private ItemManager itemManager;
        //private SyncManager syncManager;
        private TabblesRibbon ribbon;


        private BinaryFormatter formatter = new BinaryFormatter();
        private Thread listenerThread;

        private Inspectors inspectors;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                //string redemptionDllPath = @"D:\Projects\Tabbles\TabblesOutlookAddIn\TabblesLibrary\";
                //RedemptionLoader.DllLocation32Bit = redemptionDllPath + "Redemption.dll";
                //RedemptionLoader.DllLocation64Bit = redemptionDllPath + "Redemption64.dll";

                Log.log("Outlook plugin dev initializing. debug sub version: a.");


                //inspectors = Application.Inspectors;

                //inspectors.NewInspector += Inspectors_NewInspector;



                this.menuManager = new MenuManager(this.Application);

                //this.menuManager.Ribbon = this.ribbon;
                ribbon.mMenuManager = menuManager;

                this.itemManager = new ItemManager();


                //var lSession = Application.Session;
                //var lSessionFolders = lSession.Folders;


                //this.syncManager = new SyncManager();
                //syncManager.mMenuManager = menuManager;


                this.listenerThread = new Thread(ListenTabblesEvents);
                this.listenerThread.Start();


                // thread which deletes the log when it is too big

                CrashReportFsharp.execInThreadForceNewThreadDur(true, logError, FuncConvert.ToFSharpFunc<Unit>(aa =>
                {
                    try
                    {
                        Log.corpoThreadLog();
                    }
                    catch
                    {
                        // non posso scrivere nel log ovviamente
                    }
                }));


                var mustAttachHandlers = VediSeDevoFareAttachHandlers();

                if (mustAttachHandlers)
                    ApriThreadCheAttaccaHandlersInBg();
                else
                {
                    Log.log("I am not attaching handlers.");
                }


                var strFinished = Log.stringAppConfidOrTabbles() + " initialization (outside of thread which attaches handlers) finished. ";
                
                Log.log(strFinished);
            }
            catch (Exception ex)
            {
                try
                {
                    var crashId = Log.stringAppConfidOrTabbles() + " outlook-addin: error startup outer ";
                    var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(ex);
                    var str = crashId + stackTrace;
                    Log.log(str);
                    
                }
                catch
                {
                }


            }
        }

        //private void Inspectors_NewInspector(Inspector Inspector)
        //{
        //    // l'utente ha aperto un inspector, quindi probabilmente ha aperto una mail.

        //    Log.log(" aperto inspector. currentitem = " + (Inspector.CurrentItem.GetType().Name));

        //    if (Inspector.CurrentItem is MailItem)
        //    {
        //        Log.log(" is mail item. attaching handler");

        //        var mi = Inspector.CurrentItem as MailItem;
        //        mi.BeforeAttachmentAdd += (Attachment at, ref bool xx) =>
        //        {

        //            var pn = at.PathName;

        //            var path = at.GetTemporaryFilePath();

        //            Log.log($"attachment detected. path name = {pn},  path = {path}" );


        //            // send message to tabbles asking if it is confidential


        //            CrashReportFsharp.execInThreadForceNewThreadDur(false, logError, FuncConvert.ToFSharpFunc<Unit>(aa =>
        //            {
        //                try
        //                {
        //                    //var emails = (from m in pairs
        //                    //              let atSubj = new XAttribute("subject", m.Subject ?? "")
        //                    //              let atOldId = new XAttribute("old_cmd_line", outlookPrefix + m.OldId)
        //                    //              let atNewId = new XAttribute("new_cmd_line", outlookPrefix + m.NewId)
        //                    //              let ats = new[] { atSubj, atOldId, atNewId }
        //                    //              select new XElement("id_change", ats)).ToArray();
        //                    //var xelRoot = new XElement("tell_me_if_attachment_is_confidential");
        //                    //xelRoot.Add();
        //                    //var xdoc = new XDocument(xelRoot);
        //                    //var tabblesWasRunning = MenuManager.sendXmlToTabbles(xdoc);
        //                    //if (!tabblesWasRunning) {
        //                    //    // non posso, darebbe troppo fastidio
        //                    //    //showMessageTabblesIsNotRunning();
        //                    //}
        //                }
        //                catch (Exception ecc)
        //                {
        //                    try
        //                    {
        //                        var crashId = "outlook-addin: error in explorer before item paste subthread";
        //                        var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(ecc);
        //                        var str = crashId + stackTrace;
        //                        Log.log(str);
        //                        CrashReportFsharp.sendSilentCrashIfEnoughTimePassed3(ThisAddIn.logError, stackTrace, crashId, Environment.UserName ?? "", Environment.MachineName ?? "");
        //                    }
        //                    catch
        //                    {
        //                    }

        //                }
        //            }));
        //        };


        //    }
        //}


        private static bool VediSeDevoFareAttachHandlers()
        {
            try
            {
                var doc = GetTabblesConfigXdoc();

                var root = doc.Element("config");

                var outlook_attach_handlers = root.Attribute("outlook_attach_handlers");
                if (outlook_attach_handlers == null)
                {
                    return false;
                }
                else
                {
                    var mustAttachHandlers = (Boolean.Parse(outlook_attach_handlers.Value));
                    return mustAttachHandlers;
                }
            }
            catch (Exception ex)
            {
                try
                {
                    var crashId =
                        Log.stringAppConfidOrTabbles() + " outlook-addin: error vedi se devo fare attach handlers - non riesco a leggere dal config di tabbles ";
                    var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(ex);
                    var str = crashId + stackTrace;
                    Log.log(str);
                    
                }
                catch
                {
                }
                return false;
            }
        }


        private void ApriThreadCheAttaccaHandlersInBg()
        {
            Log.log("initialization: starting thread that scans folders.");

            // do this in a thread, because if exchange is not available, curfolder.items throws exceptions and each time you
            // wait a lot of time, minutes, causing outlook to shut down the plugin!
            CrashReportFsharp.execInThreadForceNewThreadDur(false, logError, FuncConvert.ToFSharpFunc<Unit>(a =>
            {
                // recursively set itemchange handlers for all folders


                try
                {
                    Log.log("initialization thread: started. 1 - attaching itemChange handlers.");

                    var startTime = DateTime.Now;
                    var lSessionFolders = Application.Session.Folders;
                    AttachHandlersItemChange(lSessionFolders);


                    AttachHandlersFolderAdd(lSessionFolders);

                    var strRep = AssemblyVer_safe() + "--" + CalcolaStringaTimeReport(startTime);
                    
                    Log.log(strRep);
                }
                catch (Exception ecc)
                {

                    try
                    {

                        var crashId = Log.stringAppConfidOrTabbles() + " outlook-addin: error in init thread";
                        var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(ecc);
                        var str = crashId + stackTrace;
                        Log.log(str);
                        
                    }
                    catch
                    {
                    }

                }
            }));
        }

        //private static void GestisciEccezione(string preambolo, Exception ecc)
        //{
        //    var str = preambolo + CrashReportFsharp.stringOfException(ecc);
        //    Log.log(str);
        //    CrashReportFsharp.sendSilentCrashIfEnoughTimePassed2(logError, preamboloEmail_safe() + str);
        //}

        private static string CalcolaStringaTimeReport(DateTime startTime)
        {
            var endTime = DateTime.Now;
            var durataSec = endTime.Subtract(startTime).TotalSeconds;
            var strRep = "initialization thread: finished. Duration = " + durataSec;
            return strRep;
        }

        public static string getTabblesFolder()
        {
            var folderDocs = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData);

            var foldName = ThisAddIn.isConfidential ? "Confidential" : "Tabbles";
            var tabblesFolder = System.IO.Path.Combine(folderDocs, foldName);
            return tabblesFolder;
        }


        private static XDocument GetTabblesConfigXdoc()
        {
            XDocument doc;
            var configDir = System.IO.Path.Combine(getTabblesFolder(), "Config");
            var fname_with_path = System.IO.Path.Combine(configDir, "config.xml");
            try
            {
                doc = XDocument.Load(fname_with_path, LoadOptions.PreserveWhitespace);
            }
            catch (Exception e)
            {
                if (e is UriFormatException
                    || e is System.Xml.XmlException
                    || e is System.IO.FileNotFoundException)
                {
                    var xelConf = new XElement("config");
                    doc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), new object[] { xelConf });
                }
                else
                    throw;
            }
            return doc;
        }

        private void AttachHandlersItemChange(Folders lSessionFolders)
        {
            var frontier = new Queue<Folder>();
            foreach (Folder f in lSessionFolders)
            {
                Log.log("initialization thread: enqueuing folder:" + f.Name);
                frontier.Enqueue(f);
            }


            while (frontier.Any())
            {
                var curFolder = frontier.Dequeue();

                Log.log("initialization thread: processing folder:" + curFolder.Name);
                try
                {
                    var itemsOfCurFolder = curFolder.Items; // attento: throws

                    Log.log("initialization thread: retrieved items of folder:" + curFolder.Name);

                    mItemsGC.Add(itemsOfCurFolder); // see comment below, bm_75h57fh57
                    itemsOfCurFolder.ItemChange += Items_ItemChange;
                }
                catch (COMException e2)
                {
                    try
                    {
                        var crashId = Log.stringAppConfidOrTabbles() + " outlook-addin: initialization thread: could not retrieve items of folder. continuing.  " +
                                      curFolder.Name;
                        var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(e2);
                        var str = crashId + stackTrace;
                        Log.log(str);
                        
                    }
                    catch // in case curfolder.name crashes
                    {
                    }

                }


                try
                {
                    var foldersOfCurFolder = curFolder.Folders; // lo fisso per sicurezza
                    Log.log("initialization thread: retrieved subfolders of folder:" + curFolder.Name);
                    foreach (Folder ch in foldersOfCurFolder)
                    {
                        frontier.Enqueue(ch);
                    }
                }
                catch (Exception e3)
                {
                    try
                    {
                        var crashId =
                            "outlook-addin: error in initialization thread: could not retrieve subfolders of folder. continuing. " +
                            curFolder.Name;
                        var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(e3);
                        var str = crashId + stackTrace;
                        Log.log(str);
                        
                    }
                    catch // in case .Name should crash...
                    {

                    }


                }
            }
        }

        private void AttachHandlersFolderAdd(Folders lSessionFolders)
        {
            Log.log("initialization thread: 2 - attaching folderAdd handlers.");

            {
                // set folderadd handlers. altrimenti se aggiungo una cartella nel corso di outlook, 
                // e poi ci sposto e ci modifico una mail, non scatta itemchange ecc.
                var fr = new Queue<Folders>();
                fr.Enqueue(lSessionFolders);

                while (fr.Any())
                {
                    var curFolders = fr.Dequeue();
                    curFolders.FolderAdd += lSessionFolders_FolderAdd;
                    mFoldersGC.Add(curFolders);

                    foreach (Folder ch in curFolders)
                    {
                        try
                        {
                            fr.Enqueue(ch.Folders);
                        }
                        catch (Exception e4)
                        {

                            var crashId =
                               Log.stringAppConfidOrTabbles() + " outlook-addin: error in initialization thread 2: could not retrieve subfolders of folder.  foldername: " +
                                ch.Name;
                            var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(e4);
                            var str = crashId + stackTrace;
                            Log.log(str);
                            


                        }
                    }
                }
            }
        }


        /// <summary>
        /// Questo metodo funziona ma è inutile, perché viene chiamato solo quando  elimino l'email dall'inspector window, cioè NON dall'elenco. devo fare doppio click su una mail, aprirla, e premere delete.
        /// </summary>
        /// <param name="Item"></param>
        /// <param name="Cancel"></param>
        void mi_BeforeDelete2(object Item, ref bool Cancel)
        {
            var item = Item as MailItem;
            if (item != null)
            {
                var m = item;
                Log.log("beforedelete called on email " + (m.Subject ?? "---"));



                CrashReportFsharp.execInThreadForceNewThreadDur(false, logError, FuncConvert.ToFSharpFunc<Unit>(a =>
                {
                    try
                    {
                        var cmdLine = new XAttribute("command_line", menuManager.outlookPrefix + getEntryIdDebug(m, "gknhj5h945t9845uytr4"));
                        var subject = new XAttribute("subject", m.Subject);
                        var ats = new object[] { cmdLine, subject };

                        var xelRoot = new XElement("email_was_deleted", ats);


                        var xdoc = new XDocument(xelRoot);
                        var tabblesWasRunning = MenuManager.sendXmlToTabbles(xdoc);
                        if (!tabblesWasRunning)
                        {
                            Utils.appendToXml(xelRoot);
                        }
                    }
                    catch (Exception e)
                    {

                        var crashId = Log.stringAppConfidOrTabbles() + " outlook-addin: error mi-before-delete inner ";
                        var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(e);
                        var str = crashId + stackTrace;
                        Log.log(str);
                        

                    }

                }));
            }
        }

        //void itemsOfCurFolder_ItemRemove()
        //{
        //    Log.log("itemremove called ");
        //    //if (Item is MailItem)
        //    //{
        //    //    ThreadUtils.execInThreadForceNewThread(() =>
        //    //    {

        //    //        var emails = new MailItem[] { (MailItem)Item };
        //    //        var cont = new Action<bool, XElement>((tabblesWasRunning, xelRoot) =>
        //    //        {
        //    //            if (!tabblesWasRunning)
        //    //            {
        //    //                Utils.appendToXml(xelRoot);
        //    //            }
        //    //        });
        //    //        sendMessageToTabblesUpdateTagsForEmails(emails, cont, null, null);
        //    //    });
        //    //}
        //}

        void lSessionFolders_FolderAdd(MAPIFolder Folder)
        {
            var curFolder = Folder;
            //
            try
            {
                var itemsOfCurFolder = curFolder.Items;
                mItemsGC.Add(itemsOfCurFolder); // see comment below, bm_75h57fh57
                itemsOfCurFolder.ItemChange += Items_ItemChange;
                //itemsOfCurFolder.ItemRemove += itemsOfCurFolder_ItemRemove; // inutile, non riceve come argomento l'email eliminata. devo usare Item.beforedelete.

                //
                var subFolders = curFolder.Folders;
                subFolders.FolderAdd += lSessionFolders_FolderAdd;
                mFoldersGC.Add(subFolders);

            }
            catch (Exception e)
            {
                try
                {
                    var crashId = Log.stringAppConfidOrTabbles() + " outlook-addin: error lSessionFolders_FolderAdd failed";
                    var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(e);
                    var str = crashId + stackTrace;
                    Log.log(str);
                    
                }
                catch
                {
                }
            }

        }




        static List<Items> mItemsGC = new List<Items>(); // prevents garbage collection. otherwise itemchange stops being fired the next time I iterate folders recursively, e.g. when doing recursive sync. bm_75h57fh57
        static List<Folders> mFoldersGC = new List<Folders>(); // prevents garbage collection. otherwise FolderAdd stops being fired the next time I iterate folders recursively,.

        private bool checkNotNull(string debugStr, object o)
        {
            if (o == null)
            {
                Log.log("skipped because null: " + debugStr);
                return false;
            }
            else
                return true;
        }


        /// <summary>
        /// Needed because this call often fails and I don't know why; if it's because I need to retry, or because the email does not exist any longer.
        /// </summary>
        /// <param name="m">the mailitem to retrieve the entryid for</param>
        /// <returns>null if I could not retrieve the entry id due to some exception</returns>
        public static string getEntryIdDebug(MailItem m, string callerId)
        {


            var exceptionsReceived = new List<Exception>();
            string result = null;
            while (result == null && exceptionsReceived.Count < 1) // prima era < 2 ma inutile : non ha mai avuto successo il retry.
            {
                try
                {
                    result = m.EntryID;
                }
                catch (Exception e)
                {
                    exceptionsReceived.Add(e);
                    Thread.Sleep(1); // I don't know if it's needed.
                }
            }

            if (exceptionsReceived.Count > 0)
            {
                var strEventualSuccess = (result == null ? "failed" : "success");
                var str0 = "getEntryIdDebug: retried at least once. Did I succeed after retrying? " + strEventualSuccess + " ; code place id = " + callerId + " ;  here are the exceptions: ";

                var strExceptions = Utils.throwsWrapper(
                    f: () =>
                    {
                        return exceptionsReceived.ConvertAll(a => CrashReportFsharp.stringOfException(a)).Aggregate((a, b) => a + nl + "-------------------" + nl + b);
                    },
                    contIfThrows: e => "error in creating exception list as string");

                var stackTrace = AssemblyVer_safe() + str0 + strExceptions;

                var crashId = "outlook-addin: error get entry id debug ";

                var str = crashId + stackTrace;
                Log.log(str);
                


            }

            return result;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="mails"></param>
        /// <param name="cont">an action which is told if tabbles-was-running</param>
        /// <param name="checkIfUserAborted"></param>
        /// <param name="beforeSendingMessageToPipe"></param>
        void sendMessageToTabblesUpdateTagsForEmails(IEnumerable<MailItem> mails, Action<bool, XElement> cont,
                Action checkIfUserAborted,
            Action beforeSendingMessageToPipe)
        {

            //var atSubj = new XAttribute("subject", m.Subject);
            //var atCmdLine = new XAttribute("command_line", outlookPrefix + m.EntryID);
            //var ats = new[] { atSubj, atCmdLine };

            var emails = new List<XElement>();

            foreach (var m in mails) // fill emails
            {
                if (checkIfUserAborted != null)
                    checkIfUserAborted();

                if (getEntryIdDebug(m, "sendMessageToTabblesUpdateTagsForEmails") != null) // the call to .entryId was failing often!
                {

                    var catNames = Utils.GetCategories(m);


                    var els = new List<XElement>();
                    foreach (var c in catNames) // fill els
                    {
                        var nameAt = new XAttribute("name", c);

                        var category = Application.Session.Categories[c];
                        if (category != null) // superfluo, è struct  && category.Color != null)
                        {
                            var col = Utils.GetRgbFromOutlookColor(category.Color);
                            var colAt = new XAttribute("color", col);

                            var atsTag = new object[] { colAt, nameAt };
                            els.Add(new XElement("tag", atsTag));
                        }
                        else
                        {
                            var atsTag = new object[] { nameAt };
                            els.Add(new XElement("tag", atsTag));
                        }

                    }
                    var subj = (m.Subject ?? "");
                    var cmdLine = new XAttribute("command_line", menuManager.outlookPrefix + getEntryIdDebug(m, "hlkgtkjr0459o54àà"));
                    var subject = new XAttribute("subject", subj);
                    var ats = new object[] { cmdLine, subject };
                    var atsAndEls = els.Concat(ats).ToList();
                    emails.Add(new XElement("email", atsAndEls));
                }
            }



            //var emails = (from m in mails

            //              let catNames = Utils.GetCategories(m)
            //              //where cats.Any() // non posso, altrimenti se tolgo l'ultima categoria non aggiorna in tabbles.
            //              let els = (from c in catNames
            //                         let category = Application.Session.Categories[c]
            //                         where checkNotNull("3 - cat name =" + c, category)
            //                         where checkNotNull("4 - cat name =" + c, category.Color)
            //                         let col = Utils.GetRgbFromOutlookColor(category.Color)
            //                         let colAt = new XAttribute("color", col)
            //                         let nameAt = new XAttribute("name", c)
            //                         //let colorNameAt = new XAttribute("color_name", category.Name)
            //                         let ats = new object[] { colAt, nameAt , null} //, colorNameAt }
            //                         select new XElement("tag", ats)).ToList()
            //              let subj = (m.Subject == null? "" : m.Subject)
            //              where checkNotNull ("6 " , m.EntryID)

            //              let cmdLine = new XAttribute("command_line", menuManager.outlookPrefix + m.EntryID)
            //              let subject = new XAttribute("subject", subj)
            //              let ats = new object[] { cmdLine, subject }
            //              let atsAndEls = els.Concat(ats).ToList()
            //              select new XElement("email", atsAndEls)).ToArray();

            if (emails.Any())
            {
                if (beforeSendingMessageToPipe != null)
                {
                    beforeSendingMessageToPipe();
                }
                var debugStr = (from e in emails
                                select e.ToString()).Aggregate((a, b) => a + Environment.NewLine + b);
                Log.log("Sending these emails to tabbles: " + debugStr);

                var xelRoot = new XElement("update_tags_for_these_emails", emails);
                var xdoc = new XDocument(xelRoot);
                //var text = xdoc.ToString();
                var tabblesWasRunning = MenuManager.sendXmlToTabbles(xdoc);
                cont(tabblesWasRunning, xelRoot);
            }
        }

        class AbortingOperation : Exception { }

        public void importOutlookTaggingIntoTabbles()
        {

            var frontier = new Queue<Folder>();


            // prendo la cartella attiva e metto nella frontiera solo quella. // prima prendevo tutte le cartelle, ma ci metteva troppo tempo
            Folder fold = (Folder)Application.ActiveExplorer().CurrentFolder;
            frontier.Enqueue(fold);


            // prima prendevo tutte le cartelle, ma ci metteva troppo tempo
            //foreach (Folder f in Application.Session.Folders)
            //{
            //    var fname = f.Name;
            //    var itemsCount = f.Items.Count;
            //    frontier.Enqueue(f);
            //}



            var msg = Res.thisWillImport2.sd().Replace("{FOLDER}", fold.Name);

            var res = MessageBox.Show(msg, Res.info, MessageBoxButtons.YesNo);
            if (res == DialogResult.Yes)
            {

                var wndPr = new progress
                {
                    pb1 = { Maximum = 100.0, IsIndeterminate = true },
                    lbl1 = { Text = Res.gatheringEmails.Replace("{N}", "0") }
                };

                wndPr.Show();

                var userWantsToAbort = false;
                wndPr.btn_cancel.Click += ((a, b) =>
                {
                    userWantsToAbort = true;
                });

                wndPr.Closing += ((a, b) => { userWantsToAbort = true; });

                var numEmailsFound = 0;




                CrashReportFsharp.execInThreadForceNewThreadDur(true, logError, FuncConvert.ToFSharpFunc<Unit>(a =>
                {
                    try
                    {
                        
                               //Application.Session.GetDefaultFolder(
                               //OlDefaultFolders.olFolderInbox)
                               //as Folder,
                               //OlFolderDisplayMode.olFolderDisplayNormal);

                        

                      

                        var emails = new Queue<MailItem>();
                        while (frontier.Any())
                        {
                            var curFolder = frontier.Dequeue();
                            var emailsInFolder =
                                Utils.throwsWrapper(
                                    () => curFolder.Items,

                                    ec =>
                                    {

                                        try
                                        {
                                            var crashId = "outlook-addin: error getting folder items. foldername = " + curFolder.Name;
                                            var stackTrace = ThisAddIn.AssemblyVer_safe() +
                                                             CrashReportFsharp.stringOfException(ec);

                                            


                                        }
                                        catch // in case curfolder.name crashes
                                        {
                                        }
                                        return null;
                                    });

                            if (emailsInFolder != null)
                            {
                                foreach (var m in emailsInFolder)
                                {



                                    var ma = m as MailItem;
                                    if (ma != null)
                                    {
                                        emails.Enqueue(ma);

                                        { // update message
                                            numEmailsFound++;
                                            if (numEmailsFound % 16 == 0)
                                            {
                                                ThreadUtils.gui(wndPr, () =>
                                                {
                                                    wndPr.lbl1.Text = Res.gatheringEmails.Replace("{N}", numEmailsFound.ToString());
                                                });
                                            }
                                        }
                                    }


                                    if (userWantsToAbort)
                                    {
                                        throw new AbortingOperation();
                                    }
                                }

                                foreach (Folder ch in curFolder.Folders)
                                {
                                    frontier.Enqueue(ch);
                                }
                            }
                        }

                        ThreadUtils.gui(wndPr, () =>
                        {
                            wndPr.pb1.IsIndeterminate = false;

                        });



                        var curEmailInXml = 0;

                        // ora mando le email a Tabbles
                        sendMessageToTabblesUpdateTagsForEmails(emails,

                                            (tabblesWasRunning, _) =>
                                            {
                                                if (!tabblesWasRunning)
                                                {
                                                    ThreadUtils.gui(wndPr, MenuManager.showMessageTabblesIsNotRunning);
                                                }
                                                else
                                                {
                                                    ThreadUtils.gui(wndPr, MenuManager.showMessageSyncSentToTabbles);
                                                    //MessageBox.Show(Res.messageSyncSentToTabbles.sd(), Res.info, MessageBoxButtons.OK, owner: wndPr);
                                                }

                                            },

                                            (() =>
                                            {
                                                { // update message
                                                    curEmailInXml++;
                                                    if (curEmailInXml % 16 == 0)
                                                    {
                                                        ThreadUtils.gui(wndPr, () =>
                                                        {
                                                            var msg2 = ThisAddIn.isConfidential ? Res.buildingMessageToSendToConfid : Res.buildingMessageToSendToTabbles2;
                                                            wndPr.lbl1.Text = msg2.Replace("{A}", curEmailInXml.ToString()).Replace("{B}", numEmailsFound.ToString());
                                                            // x / curEm = 100 / numEmails
                                                            // x = 100 * curEm / numEmails

                                                            if (numEmailsFound > 0)
                                                            {
                                                                var pbValue = 100.0 * (float)curEmailInXml / (float)numEmailsFound;
                                                                wndPr.pb1.Value = pbValue;
                                                            }
                                                        });
                                                    }
                                                }

                                                if (userWantsToAbort)
                                                {
                                                    throw new AbortingOperation();
                                                }
                                            }),

                                             (() =>
                                            {
                                                ThreadUtils.gui(wndPr, () =>
                                                {
                                                    
                                                    wndPr.lbl1.Text = Res.sendingMessageToConfidential4.sd();
                                                    
                                                    wndPr.btn_cancel.IsEnabled = false;

                                                });

                                            })

                                            );

                        ThreadUtils.gui(wndPr, () => { wndPr.Close(); });

                    }
                    catch (AbortingOperation)
                    {
                        ThreadUtils.gui(wndPr, () => { wndPr.Close(); });
                    }
                    catch (Exception otherE)
                    {
                        try
                        {
                            ThreadUtils.gui(wndPr, () => { wndPr.Close(); });

                            var crashId = "outlook-addin: error import outlook tagging into tabbles";
                            var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(otherE);
                            var str = crashId + stackTrace;
                            Log.log(str);
                            
                        }
                        catch
                        {
                        }

                    }

                }));

            }
        }


        void Items_ItemChange(object Item)
        {
            var item = Item as MailItem;
            if (item != null)
            {


                CrashReportFsharp.execInThreadForceNewThreadDur(false, logError, FuncConvert.ToFSharpFunc<Unit>(a =>
                {
                    try
                    {
                        var emails = new MailItem[] { item };
                        var cont = new Action<bool, XElement>((tabblesWasRunning, xelRoot) =>
                        {
                            if (!tabblesWasRunning)
                            {
                                Utils.appendToXml(xelRoot);
                            }
                        });
                        sendMessageToTabblesUpdateTagsForEmails(emails, cont, null, null);
                    }
                    catch (Exception ecc)
                    {
                        try
                        {
                            var crashId = "outlook-addin: error item change";
                            var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(ecc);
                            var str = crashId + stackTrace;
                            Log.log(str);
                            
                        }
                        catch
                        {
                        }
                    }
                }));
            }
        }

        //private void StartSyncThread()
        //{
        //    System.Action syncAction = this.syncManager.GetSyncAction();
        //    syncAction();
        //}

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            this.ribbon = new TabblesRibbon();
            ribbon.mAddin = this;



            return this.ribbon;
        }

        //private bool OnSendMessageToTabbles(object message)
        //{
        //    return SendMessageToTabblesBlocking(message);
        //}

        //private bool SendMessageToTabblesBlocking(object msg, bool retry = false)
        //{
        //    try
        //    {
        //        // I commented this block because this function should should never fail without showing an error message box.
        //        //if (msg.GetType().GetCustomAttributes(typeof(SerializableAttribute), false).Length == 0)
        //        //{
        //        //    return false;
        //        //}

        //        if (this.outlookToTabblesClientPipe == null || retry)
        //        {
        //            this.outlookToTabblesClientPipe = new NamedPipeClientStream(".", "OutlookToTabblesPipe",
        //                PipeDirection.Out, PipeOptions.Asynchronous);
        //            Logger.Log("connecting to Tabbles pipe server...");
        //            this.outlookToTabblesClientPipe.Connect(200); // blocks the thread
        //            Logger.Log("connected.");
        //        }

        //        Logger.Log("sendMessageToTabblesBlocking: serialize: " + msg.GetType().ToString());
        //        this.formatter.Serialize(this.outlookToTabblesClientPipe, msg);
        //        this.outlookToTabblesClientPipe.Flush();

        //        return true;
        //        //logFile.Print("sendMessageToTabblesBlocking: sent");
        //    }
        //    catch (TimeoutException)
        //    {
        //        string str = "Tabbles plugin not active. Cannot send message to Tabbles: " + msg.GetType().ToString();
        //        Logger.Log(str);

        //        try
        //        {
        //            this.outlookToTabblesClientPipe.Dispose();
        //        }
        //        catch (System.Exception)
        //        { }
        //        finally
        //        {
        //            this.outlookToTabblesClientPipe = null;
        //        }

        //        return false;
        //    }
        //    catch (System.Exception)
        //    {
        //        if (!retry)
        //        {
        //            try
        //            {
        //                this.outlookToTabblesClientPipe.Dispose();
        //            }
        //            catch (System.Exception)
        //            { }
        //            finally
        //            {
        //                this.outlookToTabblesClientPipe = null;
        //            }

        //            //try once more to re-connect the pipe
        //            if (SendMessageToTabblesBlocking(msg, true))
        //            {
        //                return true;
        //            }
        //            else
        //            {
        //                Logger.Log("The Tabbles plugin for Outlook is not running.");
        //            }
        //        }

        //        return false;
        //    }
        //}

        private void handleMessageFromTabbles(XDocument xdoc)
        {
            var root = xdoc.Root;
            if (root.Name.LocalName == "emails_tagged")
            {
                var emails = root.Elements("email");
                var tags = root.Elements("tag").ToList();

                foreach (var email in emails)
                {
                    var cmdLine = email.Attribute("command_line").Value;
                    // I have to tag the same email with categories corresponding to the tags
                    var arguments = cmdLine.Split(OutlookCmdSeparator, StringSplitOptions.None);

                    var entryId = arguments[1];

                    var mail = (MailItem)Application.Session.GetItemFromID(entryId);

                    var currentCategories = mail.Categories != null ? Utils.GetCategories(mail) : new string[] { };

                    var tagsToAddWithColors = (from tag in tags
                                               let tagName = tag.Attribute("name").Value
                                               let tagColor = tag.Attribute("color").Value
                                               where currentCategories.All(cat => cat != tagName)
                                               select new { name = tagName, color = tagColor }).ToList();

                    if (!tagsToAddWithColors.Any())
                    {
                        continue;
                    }

                    foreach (var tag in tagsToAddWithColors)
                    {
                        Category cat = !CategoryExists(tag.name) ? this.Application.Session.Categories.Add(tag.name) : this.Application.Session.Categories[tag.name];

                        //change colors for all categories, in case if they were changed in Tabbles
                        cat.Color = Utils.GetOutlookColorFromRgb(tag.color);
                    }

                    var tagsToAdd = (from x in tagsToAddWithColors
                                     select x.name).ToList();
                    var finalTags = tagsToAdd.Union(currentCategories).ToList();

                    //IEnumerable<string> newCats = tagsToAdd.Concat<string>(currentCategories);
                    // todo newcats is empty: ???? check, are they

                    if (finalTags.Any())
                    {

                        mail.Categories = finalTags.Aggregate((a, b) => a + ";" + b);
                    }

                    this.menuManager.InternallyChangedMailIds.Add(entryId);

                    mail.Save();

                }

            }
            else if (root.Name.LocalName == "emails_untagged")
            {
                var emails = root.Elements("email");
                var tags = root.Elements("tag");
                foreach (var email in emails)
                {
                    var cmdLine = email.Attribute("command_line").Value;

                    // I have to tag the same email with categories corresponding to the tags
                    var arguments = cmdLine.Split(OutlookCmdSeparator, StringSplitOptions.None);

                    var entryId = arguments[1];

                    var mail = (MailItem)Application.Session.GetItemFromID(entryId);

                    string[] currentCategories;
                    if (mail.Categories != null)
                    {
                        currentCategories = Utils.GetCategories(mail);
                    }
                    else
                    {
                        continue;
                    }


                    var tagnames = (from tag in tags
                                    select tag.Attribute("name").Value);
                    var newCats = currentCategories.Except<string>(tagnames).ToList();

                    if (newCats.Any())
                    {
                        mail.Categories = newCats.Aggregate((a, b) => a + ";" + b);
                    }
                    else
                    {

                        mail.Categories = null;
                    }


                    this.menuManager.InternallyChangedMailIds.Add(entryId);

                    mail.Save();
                }

            }
            else if (root.Name.LocalName == "find_emails_which_have_these_tags")
            {
                var tags = root.Elements("tag");

                var tagnames = (from tag in tags
                                select tag.Attribute("name").Value);

                //MsgOpenMailsWithTags msgOpenMailsWithTags = (MsgOpenMailsWithTags)messageObj;
                //if (msgOpenMailsWithTags.tags != null)
                //{
                SearchByCategories(tagnames);
                //}
            }
            //else if (root.Name.LocalName == "tag_created")
            //{
            //    //MsgAtomKeyCreated msgAtomKeyCreated = (MsgAtomKeyCreated)messageObj;
            //    //string categoryName = msgAtomKeyCreated.AtomKeyName;

            //    //Category category;
            //    //if (!CategoryExists(categoryName))
            //    //{
            //    //    category = this.Application.Session.Categories.Add(categoryName);
            //    //}
            //    //else
            //    //{
            //    //    category = this.Application.Session.Categories[categoryName];
            //    //}

            //    //category.Color = Utils.GetOutlookColorFromRgb(msgAtomKeyCreated.AtomKeyColor);

            //    //Logger.Log("detected ak created: " + msgAtomKeyCreated.AtomKeyName);
            //}
            else if (root.Name.LocalName == "tag_deleted")
            {
                Log.log("detected ak deleted");
            }
            else
            {
                Log.log("message from Tabbles not recognized: " + root.ToString());
            }
        }

        //int mCountFailurePipeCreation = 0;
        private void ListenTabblesEvents()
        {

            while (true)
            {
                try
                {


                    var cont = new Action<NamedPipeServerStream>((pipeServer) =>
                                  {

                                      try
                                      {
                                          Log.log("Waiting for Tabbles to connect to outlook pipe...");
                                          pipeServer.WaitForConnection(); //blocking

                                          Log.log("Connection established.");

                                          var xdoc = XDocument.Load(pipeServer);
                                          //pipeServer.Dispose();


                                          CrashReportFsharp.execInThreadForceNewThreadDur(false, logError,
                                              FuncConvert.ToFSharpFunc<Unit>(a =>
                                              {
                                                  try
                                                  {
                                                      handleMessageFromTabbles(xdoc);
                                                  }
                                                  catch (Exception ec)
                                                  {
                                                      try
                                                      {

                                                          var crashId =
                                                              "outlook-addin: listen tabbles event - crash subthread";
                                                          var stackTrace = ThisAddIn.AssemblyVer_safe() +
                                                                           CrashReportFsharp.stringOfException(ec);
                                                          var str = crashId + stackTrace;
                                                          Log.log(str);
                                                          
                                                      }
                                                      catch
                                                      {
                                                      }

                                                  }
                                              }));
                                      }
                                      finally
                                      {
                                          pipeServer.Dispose();
                                      }

                                  });


                    var res = CrashReportFsharp.createPipeServerSafeForExternalProcesses(AssemblyVer_safe(), logError, @"TABBLES_PIPE_TO_OUTLOOK");
                    if (res.IsCpsrFailed)
                    {
                        var res2 = (CrashReportFsharp.createPipeServerResult.CpsrFailed)res;
                        Log.log("ERROR: failed creating pipe: " + res2.Item);
                        Thread.Sleep(3000); // per non prendere il 50% cpu, che succederebbe riprovando subito.
                    }
                    else if (res.IsCpsrOkSecondTry)
                    {
                        var res2 = (CrashReportFsharp.createPipeServerResult.CpsrOkSecondTry)res;
                        Log.log("ERROR: failed creating pipe 1st try, succeded 2nd try: " + res2.Item2);
                        cont(res2.Item1);
                    }
                    else if (res.IsCpsrOkThirdTry)
                    {
                        var res2 = (CrashReportFsharp.createPipeServerResult.CpsrOkThirdTry)res;
                        Log.log("ERROR: failed creating pipe 2nd try, succeded 3rd try: " + res2.Item2);
                        cont(res2.Item1);
                    }
                    else if (res.IsCpsrOkFirstTry)
                    {
                        var res2 = (CrashReportFsharp.createPipeServerResult.CpsrOkFirstTry)res;
                        cont(res2.Item);
                    }
                    else
                    {
                        try
                        {
                            Log.log("ERROR: failed creating pipe: unhandled case cvru84u8r4jife4 ");
                        }
                        catch
                        {
                        }
                        finally
                        {
                            Thread.Sleep(2000);
                        }
                    }

                }
                catch (Exception e)
                {
                    try
                    {
                        // ad esempio ha crashato waitForConnection
                        var crashId = "outlook-addin: error in cont after creating pipe server";
                        var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(e);
                        var str = crashId + stackTrace;
                        Log.log(str);
                        
                    }
                    catch
                    {
                    }
                    finally
                    {


                        Thread.Sleep(2000);
                        // per sicurezza. non prendere il 50% cpu, che succederebbe riprovando subito.
                    }
                }
            }

        }

        private bool CategoryExists(string categoryName)
        {
            try
            {
                var category =
                    this.Application.Session.Categories[categoryName];

                return category != null;
            }
            catch
            {
                return false;
            }
        }

        public void SearchByCategories(IEnumerable<string> categories)
        {

            var explorer = Application.Explorers.Add(
                   Application.Session.GetDefaultFolder(
                   OlDefaultFolders.olFolderInbox)
                   as Folder,
                   OlFolderDisplayMode.olFolderDisplayNormal);

            var cats = (from c in categories
                            //select "category:\"" + c + "\"").Aggregate((a, b) => a + " AND " + b);

                        select "System.Category: = \"" + c + "\"").Aggregate((a, b) => a + " AND " + b);

            explorer.Search(cats, OlSearchScope.olSearchScopeAllFolders);
            explorer.Display();


            //Folder currentFolder = (Folder)Application.ActiveExplorer().CurrentFolder;



            //Folder rootFolder;


            //if (currentFolder != null)
            //{
            //    rootFolder = (Folder)currentFolder.Store.GetRootFolder();
            //}
            //else
            //{
            //    rootFolder = (Folder)Application.Session.Folders[1];
            //}




            ////example: ("urn:schemas-microsoft-com:office:office#Keywords" = 'aa' OR "urn:schemas-microsoft-com:office:office#Keywords" = 'bb')
            //int count = categories.Count<string>();
            //StringBuilder filterSql = new StringBuilder("(");
            //if (count > 0)
            //{
            //    filterSql.AppendFormat("\"urn:schemas-microsoft-com:office:office#Keywords\" = '{0}'", categories.First<string>());
            //}
            //else
            //{
            //    return;
            //}

            //for (int i = 1; i < count; i++)
            //{
            //    // andrej aveva messo OR. a me sembra senza senso.
            //    filterSql.Append(" AND ").AppendFormat("\"urn:schemas-microsoft-com:office:office#Keywords\" = '{0}'", categories.ElementAt<string>(i));
            //}
            //filterSql.Append(")");

            //#region old comment by andrej
            ////-- We use Redemption instead of these code (together with AdvancedSearchComplete event, see another Commented out section)
            ////-- Currently there is a problem with calling Results.Save() for search on a folder of non-default store
            ////See: http://social.msdn.microsoft.com/Forums/en-US/outlookdev/thread/7d1d3494-988f-4c42-a391-e732b5dfb2c6

            ////string folderStr = string.Format("'{0}'", rootFolder.FolderPath);

            ////string logMessage = string.Format("Started search with filter {0} in folder {1} ...", filter.ToString(), folderStr);
            ////this.logger.Log(logMessage);

            ////Application.AdvancedSearch(folderStr, filter.ToString(), true, "Tabbles categories");
            ////--------------------------------------------------------------------------------------
            //#endregion

            //var performSearch = new System.Action(() =>
            //    {
            //            #region Sujay
            //            //if (this.rdoSession == null)
            //            //{
            //            //    this.rdoSession = RedemptionLoader.new_RDOSession();
            //            //}
            //            //if (!this.rdoSession.LoggedOn)
            //            //{
            //            //    this.rdoSession.Logon();
            //            //}

            //            //RDOStore2 store = (RDOStore2)this.rdoSession.GetStoreFromID(rootFolder.StoreID);

            //            //oInbox = oApp.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);


            //            //NameSpace olNS = this.Application.GetNamespace("MAPI");
            //            //Store olStore = olNS.GetStoreFromID(rootFolder.StoreID);

            //            //MAPIFolder olSearchFolder;

            //            //  olStore.

            //            //  Application.AdvancedSearchComplete -= new ApplicationEvents_11_AdvancedSearchCompleteEventHandler(Application_AdvancedSearchComplete);

            //            string folderStr = string.Format("'{0}'", rootFolder.FolderPath);
            //            var olSearch = Application.AdvancedSearch(Scope: folderStr, Filter: filterSql.ToString(), SearchSubFolders: true, Tag: SearchResultsFolderName);
            //            olSearch.Save(SearchResultsFolderName);

            //            //     olSearchFolder = olSearch.Save("Sujay Search");

            //            //Application.AdvancedSearchComplete -= new ApplicationEvents_11_AdvancedSearchCompleteEventHandler(Application_AdvancedSearchComplete);

            //            //store.OnSearchComplete += store_OnSearchComplete;



            //            //MAPIFolder olFolderFromID = olNS.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID);


            //            //RDOFolder folder = this.rdoSession.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID);

            //            // Sujay code

            //            //store.Searches.AddCustom(SearchResultsFolderName, filterSql.ToString(), folder, true); 


            //            #endregion

            //    });

            //Folders searchFolders = rootFolder.Store.GetSearchFolders();
            ////if (searchFolders != null)
            ////{
            ////    if (this.folderManager == null)
            ////    {
            ////        this.folderManager = new FolderManager();
            ////    }

            ////    //in case if there is a search folder
            ////    this.folderManager.RemoveFolderByName(folders: searchFolders, name: SearchResultsFolderName, callback: performSearch);
            ////}
            ////else
            ////{
            //    //in case if there is no any search folder
            //    performSearch();
            ////}

            return;
        }

        //private void store_OnSearchComplete(string searchFolderID)
        //{
        //    #region Sujay
        //    //Folder searchFolder = (Folder)Application.Session.GetFolderFromID(searchFolderID);
        //    //if (this.rdoSession != null && this.rdoSession.LoggedOn)
        //    //{
        //    //    RDOStore2 store = (RDOStore2)this.rdoSession.GetStoreFromID(searchFolder.StoreID);
        //    //    store.OnSearchComplete -= store_OnSearchComplete;
        //    //}

        //    //Application.ActiveExplorer().CurrentFolder = searchFolder; 
        //    #endregion
        //}

        #region Commented out
        //see comment in SearchByCategories() for the explanation

        //private void Application_AdvancedSearchComplete(Search search)
        //{
        //    #region Sujay Comments
        //    //string logMessage = string.Format("Search is completed with {0} results.", search.Results.Count.ToString());
        //    ////this.logger.Log(logMessage);

        //    //if (search.Results.Count != 0)
        //    //{
        //    //    search.Save("Sujay Search");
        //    //    return;
        //    //}

        //    if (search.Results.Count == 0)
        //    {
        //        MessageBox.Show(Res.MsgNoResultsFound);
        //    }
        //    else
        //    {
        //        //search.Save("Sujay Search");

        //        Folders searchFolders = null;
        //        MailItem aMail = search.Results[1] as MailItem;
        //        Folder aFolder = aMail.Parent as Folder;
        //        searchFolders = aFolder.Store.GetSearchFolders();

        //        var showResultsAction = new System.Action(() =>
        //            {
        //                Folder searchResultsFolder = (Folder)search.Save(SearchResultsFolderName);
        //                Application.ActiveExplorer().CurrentFolder = searchResultsFolder;
        //            });

        //        //if (searchFolders != null)
        //        //{
        //        //    if (this.folderManager == null)
        //        //    {
        //        //        this.folderManager = new FolderManager();
        //        //    }

        //        //    //in case if there is a search folder
        //        //    this.folderManager.RemoveFolderByName(searchFolders, SearchResultsFolderName, showResultsAction);
        //        //}
        //        //else
        //        //{
        //        //in case if there is no any search folder
        //        showResultsAction();
        //        //}

        //        return;

        //        //give some response in any case
        //        //MessageBox.Show(Res.MsgNoResultsFound);
        //    }
        //    #endregion

        //    //MessageBox.Show(" In advanced search");

        //    //Application.AdvancedSearchComplete -= new ApplicationEvents_11_AdvancedSearchCompleteEventHandler(Application_AdvancedSearchComplete);

        //    //  Application.ActiveExplorer().CurrentView = searchFolder.Application.ActiveExplorer().CurrentView;//  = searchFolder.f;
        //}
        #endregion

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {

            //Application.AdvancedSearchComplete -= new ApplicationEvents_11_AdvancedSearchCompleteEventHandler(Application_AdvancedSearchComplete);
            //Logger.Dispose();
        }



        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
