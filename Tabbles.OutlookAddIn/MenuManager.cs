using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Pipes;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Xml.Linq;
using Microsoft.FSharp.Core;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Exception = System.Exception;
using WinForms = System.Windows.Forms;
using Res = Tabbles.OutlookAddIn.Properties.Resources;
using u = Tabbles.OutlookAddIn.Utils;

namespace Tabbles.OutlookAddIn
{


    //public delegate bool IsAnyEmailSelectedHandler();

    public class MenuManager
    {


        private readonly FSharpFunc<string, Unit> logError = FuncConvert.ToFSharpFunc<string>(Log.log);
        // SUJAYXML
        //   private XMLFileManager xmlFileManager;


        //private const string CommandBarName = "Tabbles Toolbar";
        private const string ButtonIdTagUsingTabbles = "tagUsingTabbles";
        private const string ButtonIdOpenInTabbles = "openInTabbles";
        private const string ButtonIdTabblesSearch = "tabblesSearch";
        private const string ButtonIdSyncWithTabbles = "syncWithTabbles";
        private const string PropertyNameCategories = "Categories";
        private const string PropertyNameFlagRequest = "FlagRequest";


        //public event System.Action StartSync;

        //public readonly object syncObj = new object();

        private OutlookVersion outlookVersion;
        public string outlookPrefix;
        private CultureInfo outlookCulture;

        private Application outlookApp;
        private Explorers explorers;

        //keep the list members to avoid VSTO garbage collection problem
        private List<Explorer> explorerList;
        //private List<CommandBarButton> buttonList;

        private List<MailItem> selectedMails;

        //private Items currentFolderItems;

        //private ISet<string> onceItemChanged;

        public ISet<string> InternallyChangedMailIds
        {
            get;
            private set;
        }

        private bool trackItemMove = true;

        //public TabblesRibbon Ribbon
        //{
        //    set
        //    {
        //        //value.TagEmailsWithTabbles += (sender, args) =>
        //        //{
        //        //    TagSelectedEmailsWithTabbles();
        //        //};
        //        //value.OpenEmailInTabbles += (sender, args) =>
        //        //{
        //        //    OpenSelectedEmailInTabbles();
        //        //};
        //        //value.TabblesSearch += (sender, args) =>
        //        //{
        //        //    TabblesSearch();
        //        //};
        //        //value.SyncWithTabbles += (sender, args) =>
        //        //{
        //        //    //RegistryManager.SetSyncPerformed(false);

        //        //    // SUJAYXML
        //        //    //xmlFileManager.SetSyncPerformed(false);

        //        //    if (StartSync != null)
        //        //    {
        //        //        StartSync();
        //        //    }
        //        //};
        //        //value.IsAnyEmailSelected += () =>
        //        //{
        //        //    return IsAnyEmailSelected(true);
        //        //};
        //    }
        //}

        public MenuManager(Application outlookApp)
        {

            this.outlookApp = outlookApp;
            this.outlookVersion = Utils.ParseMajorVersion(outlookApp);
            this.outlookPrefix = Utils.GetOutlookPrefix();

            this.explorerList = new List<Explorer>();
            //this.buttonList = new List<CommandBarButton>();

            //this.onceItemChanged = new HashSet<string>();

            InternallyChangedMailIds = new HashSet<string>();

            //culture info for localization
            int languageId = outlookApp.LanguageSettings.LanguageID[MsoAppLanguageID.msoLanguageIDUI];
            this.outlookCulture = new CultureInfo(languageId);
            Thread.CurrentThread.CurrentUICulture = this.outlookCulture;

            //CheckMenus();


            this.explorers = this.outlookApp.Explorers;
            this.explorers.NewExplorer += OnNewExplorer;

            //FillItemsToListen();

            foreach (Explorer explorer in this.explorers)
            {

                AddExplorerListeners(explorer);
            }

        }


        private void OnNewExplorer(Explorer explorer)
        {
            try
            {
                AddExplorerListeners(explorer);
            }
            catch (Exception e)
            {

                try
                {
                    var crashId = "outlook-addin: error onNewExplorer ";
                    var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(e);
                    var str = crashId + stackTrace;
                    Log.log(str);
                    
                }
                catch
                {
                }
            }
        }

        private void AddExplorerListeners(Explorer explorer)
        {
            this.explorerList.Add(explorer);

            explorer.SelectionChange += UpdateSelectedEmails_safe;
            //explorer.BeforeItemCopy += explorer_BeforeItemCopy;
            //explorer.BeforeItemCut += explorer_BeforeItemCut;
            explorer.BeforeItemPaste += explorer_BeforeItemPaste_safe;

            


            //explorer.FolderSwitch += () =>
            //    {
            //        FillItemsToListen();
            //    };
        }

        class EntryIdChange
        {
            public string NewId { get; set; }
            public string OldId { get; set; }

            public string Subject { get; set; }

        }

        void explorer_BeforeItemPaste_safe(ref object clipboardContent, MAPIFolder Target, ref bool Cancel)
        {
            try
            {
                if (!this.trackItemMove) //prevent infinite loop
                {
                    return;
                }

                if (clipboardContent is Selection)
                {
                    var mailsToMove = new List<MailItem>();

                    var selection = (Selection)clipboardContent;
                    foreach (object itemObj in selection)
                    {
                        var obj = itemObj as MailItem;
                        if (obj != null)
                        {
                            mailsToMove.Add(obj);
                        }
                    }

                    if (mailsToMove.Count == 0)
                    {
                        return;
                    }


                    try
                    {
                        bool mailMovedToDifferentStore = u.c(() =>
                        {
                            foreach (MailItem mail in mailsToMove)
                            {
                                if (string.IsNullOrEmpty(mail.Categories))
                                {
                                    continue;
                                }

                                if (mail.Parent is Folder)
                                {
                                    var parent = (Folder)mail.Parent;
                                    if (parent.StoreID != Target.StoreID)
                                    {
                                        return true;
                                    }
                                }
                            }
                            return false;

                        });

                        if (!mailMovedToDifferentStore)
                        {
                            return;
                        }


                        Cancel = true; // because I am doing the move myself with mail.Move()
                        this.trackItemMove = false;

                        var pairs = new List<EntryIdChange>();
                        foreach (MailItem mail in mailsToMove)
                        {
                            var mailAfterMove = (MailItem)mail.Move(Target);
                            Log.log("moved mail. old id = " + ThisAddIn.getEntryIdDebug(mail, "bljbkghjrhje") + " ---- new id = " + mailAfterMove.EntryID);
                            pairs.Add(new EntryIdChange { OldId = ThisAddIn.getEntryIdDebug(mail, "gflibfkhjdsbnmdbfjdhjg"), NewId = mailAfterMove.EntryID, Subject = mail.Subject ?? "" });
                            Utils.ReleaseComObject(mailAfterMove);
                        }
                        this.trackItemMove = true;



                        CrashReportFsharp.execInThreadForceNewThreadDur(false, logError, FuncConvert.ToFSharpFunc<Unit>(aa =>
                        {
                            try
                            {
                                var emails = (from m in pairs
                                              let atSubj = new XAttribute("subject", m.Subject ?? "")
                                              let atOldId = new XAttribute("old_cmd_line", outlookPrefix + m.OldId)
                                              let atNewId = new XAttribute("new_cmd_line", outlookPrefix + m.NewId)
                                              let ats = new[] { atSubj, atOldId, atNewId }
                                              select new XElement("id_change", ats)).ToArray();
                                var xelRoot = new XElement("update_email_ids", emails);
                                var xdoc = new XDocument(xelRoot);
                                var tabblesWasRunning = sendXmlToTabbles(xdoc);
                                if (!tabblesWasRunning)
                                {
                                    Utils.appendToXml(xelRoot);
                                }
                            }
                            catch (Exception ecc)
                            {
                                try
                                {
                                    var crashId = "outlook-addin: error in explorer before item paste subthread";
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
                    finally
                    {
                        foreach (MailItem mail in mailsToMove)
                        {
                            Utils.ReleaseComObject(mail);
                        }
                    }
                }
            }
            catch (Exception eOuter)
            {
                try
                {
                    var crashId = "outlook-addin: error before item paste ";
                    var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(eOuter);
                    var str = crashId + stackTrace;
                    Log.log(str);
                    
                }
                catch
                {
                }
            }
        }

        //void explorer_BeforeItemCut(ref bool Cancel)
        //{
        //    var y = 5;
        //}

        //void explorer_BeforeItemCopy(ref bool Cancel)
        //{
        //    var y = 5;
        //}

        //// era chiamata in explorer_BeforeItemPaste
        //private void TrackEmailMove(ref object clipboardContent, MAPIFolder target, ref bool cancel)
        //{
        //    if (!this.trackItemMove) //prevent infinite loop
        //    {
        //        return;
        //    }

        //    if (clipboardContent is Selection)
        //    {
        //        List<MailItem> mails = new List<MailItem>();

        //        Selection selection = (Selection)clipboardContent;
        //        foreach (object itemObj in selection)
        //        {
        //            if (itemObj is MailItem)
        //            {
        //                mails.Add((MailItem)itemObj);
        //            }
        //        }

        //        if (mails.Count == 0)
        //        {
        //            return;
        //        }

        //        bool movedFromStore = false;
        //        try
        //        {
        //            foreach (MailItem mail in mails)
        //            {
        //                if (string.IsNullOrEmpty(mail.Categories))
        //                {
        //                    continue;
        //                }

        //                if (mail.Parent is Folder)
        //                {
        //                    Folder parent = (Folder)mail.Parent;
        //                    if (parent.StoreID != target.StoreID)
        //                    {
        //                        movedFromStore = true;
        //                        break;
        //                    }
        //                }
        //            }

        //            if (!movedFromStore)
        //            {
        //                return;
        //            }

        //            // todo maurizio
        //            //if (!CheckTabblesRunning())
        //            //{
        //            //    cancel = true;
        //            //    WinForms.MessageBox.Show(Res.MsgTabblesIsNotRunning, Res.MsgCaptionTabblesAddIn);
        //            //    return;
        //            //}

        //            cancel = true;
        //            this.trackItemMove = false;

        //            foreach (MailItem mail in mails)
        //            {
        //                MailItem mailAfterMove = (MailItem)mail.Move(target);
        //                Utils.ReleaseComObject(mailAfterMove);
        //                //WinForms.MessageBox.Show(mail.EntryID + "\n\n" + mailAfterMove.EntryID);
        //                //TODO Maurizio: call Tabbles API at this point
        //            }
        //            this.trackItemMove = true;

        //        }
        //        finally
        //        {
        //            foreach (MailItem mail in mails)
        //            {
        //                Utils.ReleaseComObject(mail);
        //            }
        //        }
        //    }
        //}

        private void UpdateSelectedEmails_safe()
        {
            try
            {
                var selection = this.outlookApp.ActiveExplorer().Selection;
                FillSelectedMails(selection);
            }
            catch (Exception ec)
            {
                try
                {
                    var crashId = "outlook-addin: error in update-selected-emails";
                    var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(ec);
                    var str = crashId + stackTrace;
                    Log.log(str);
                    
                }
                catch
                {
                }
            }
        }

        //public bool CheckTabblesRunning()
        //{
        //    if (SendMessageToTabbles == null)
        //    {
        //        return false;
        //    }

        //    return SendMessageToTabbles(new INeedToPingTabbles());
        //}

        //public void SendEmailCategories(List<string> entryIds)
        //{

        //    // todo
        //    //if (SendMessageToTabbles == null)
        //    //{
        //    //    return;
        //    //}

        //    //foreach (string entryId in entryIds)
        //    //{
        //    //    try
        //    //    {
        //    //        MailItem mail = this.outlookApp.Session.GetItemFromID(entryId) as MailItem;
        //    //        if (mail != null)
        //    //        {
        //    //            SendCategoriesToTabbles(mail);
        //    //        }
        //    //    }
        //    //    catch (System.Exception)
        //    //    {
        //    //    }
        //    //}
        //}


        //private void CheckMenus()
        //{
        //    if (this.outlookVersion == OutlookVersion.OUTLOOK_2003 ||
        //        this.outlookVersion == OutlookVersion.OUTLOOK_2007)
        //    {
        //        CommandBar commandBar = null;
        //        try
        //        {
        //            commandBar = this.outlookApp.ActiveExplorer().CommandBars[CommandBarName];
        //            if (commandBar != null)
        //            {
        //                commandBar.Delete();
        //            }
        //        }
        //        catch (System.Exception)
        //        {
        //        }

        //        commandBar = this.outlookApp.ActiveExplorer().CommandBars.Add(CommandBarName, MsoBarPosition.msoBarTop, Temporary: true);

        //        CommandBarButton tagUsingTabbles = CreateCommandBarButton(commandBar, Res.MenuTagUsingTabbles, ButtonIdTagUsingTabbles, "tag_using_tabbles");
        //        tagUsingTabbles.Click += tagUsingTabblesMenuButton_Click;
        //        this.buttonList.Add(tagUsingTabbles);

        //        CommandBarButton openEmailInTabbles = CreateCommandBarButton(commandBar, Res.MenuOpenInTabbles, ButtonIdOpenInTabbles, "open_in_tabbles");
        //        openEmailInTabbles.Click += openInTabblesMenuButton_Click;
        //        this.buttonList.Add(openEmailInTabbles);

        //        //CommandBarButton tabblesSearch = CreateCommandBarButton(commandBar, Res.MenuTabblesSearch, ButtonIdTabblesSearch, "search");
        //        //tabblesSearch.Click += tabblesSearch_Click;
        //        //this.buttonList.Add(tabblesSearch);

        //        //CommandBarButton syncWithTabbles = CreateCommandBarButton(commandBar, Res.MenuSyncWithTabbles, ButtonIdSyncWithTabbles, "sync_with_tabbles");
        //        //syncWithTabbles.Click += syncWithTabbles_Click;
        //        //this.buttonList.Add(syncWithTabbles);

        //        commandBar.Protection = MsoBarProtection.msoBarNoCustomize;
        //        commandBar.Visible = true;

        //        this.outlookApp.ItemContextMenuDisplay += outlookApp_ItemContextMenuDisplay;
        //    }
        //}

        //private CommandBarButton CreateCommandBarButton(CommandBar commandBar, string caption, string tag, string pictureAlias)
        //{
        //    CommandBarButton button = (CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton);
        //    button.Caption = caption;
        //    button.Tag = tag;
        //    SetButtonPicture(button, pictureAlias + "_16_bmp", pictureAlias + "_16_mask");

        //    return button;
        //}

        //private void outlookApp_ItemContextMenuDisplay(CommandBar commandBar, Selection selection)
        //{
        //    if (IsAnyEmailSelected(true))
        //    {
        //        CommandBarButton tagUsingTabbles = (CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, Temporary: true);
        //        tagUsingTabbles.Caption = Res.MenuTagUsingTabbles;
        //        tagUsingTabbles.Click += tagUsingTabblesContextMenu_Click;
        //        SetButtonPicture(tagUsingTabbles, "tag_using_tabbles_16_bmp", "tag_using_tabbles_16_mask");
        //        this.buttonList.Add(tagUsingTabbles);

        //        if (this.selectedMails != null && this.selectedMails.Count == 1)
        //        {
        //            CommandBarButton openInTabbles = (CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, Temporary: true);
        //            openInTabbles.Caption = Res.MenuOpenInTabbles;
        //            openInTabbles.Click += openInTabblesContextMenu_Click;
        //            SetButtonPicture(openInTabbles, "open_in_tabbles_16_bmp", "open_in_tabbles_16_mask");
        //            this.buttonList.Add(openInTabbles);
        //        }
        //    }
        //}

        //private void SetButtonPicture(CommandBarButton button, string imageName, string maskName)
        //{
        //    IPictureDisp picture = GetPictureDispFromResource(imageName);
        //    if (picture != null)
        //    {
        //        button.Style = MsoButtonStyle.msoButtonIconAndCaption;
        //        button.Picture = picture;
        //        IPictureDisp mask = GetPictureDispFromResource(maskName);
        //        if (mask != null)
        //        {
        //            button.Mask = mask;
        //        }
        //    }
        //    else
        //    {
        //        button.Style = MsoButtonStyle.msoButtonCaption;
        //    }
        //}

        //private void tagUsingTabblesMenuButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        //{
        //    TagSelectedEmailsWithTabbles();
        //}

        public void TagSelectedEmailsWithTabbles()
        {
            if (IsAnyEmailSelected(true))
            {
                TagEmailsWithTabbles_safe(this.selectedMails);
            }
            else
            {
                MessageBox.Show(Res.noEmailSelected);
            }
        }

        public void TagOpenEmailWithTabbles()
        {
            var it = outlookApp.ActiveInspector().CurrentItem;
            var item = it as MailItem;
            if (item != null)
            {

                var ar = new[] { item }.ToList();
                TagEmailsWithTabbles_safe(ar);
            }
            else
            {
                MessageBox.Show(Res.notAnEmail);
            }
        }

        //private void tagUsingTabblesContextMenu_Click(CommandBarButton ctrl, ref bool cancelDefault)
        //{
        //    if (IsAnyEmailSelected(false))
        //    {
        //        TagEmailsWithTabbles(this.selectedMails);
        //    }
        //}

        // Non posso usare il mio lock! CrashReportFsharp ne usa già uno. se lo faccio, deadlock.
        //private static object mLock = new object();

        public static bool sendXmlToTabbles(XDocument xdoc)
        {
            try
            {
                //Log.log("before trying to lock to send this message: " + xdoc.ToString());
                lock (CrashReportFsharp.g_lock) // only one thread at a time must attempt this. Otherwise pipe crashes.
                {

                    using (var pc = new NamedPipeClientStream("TABBLES_PIPE_FROM_OUTLOOK"))
                    {
                        pc.Connect(500);
                        xdoc.Save(pc);
                        try
                        {
                            pc.WaitForPipeDrain(); // senza questo arriva ioexception pipe is being closed
                        }
                        catch
                        {
                        }
                    }
                }
                Log.log("Message sent to Tabbles successfully: " + xdoc.ToString());
                return true;
            }
            catch (TimeoutException)
            {
                try
                {
                    Log.log("Tabbles is not running. Message lost: " + xdoc.ToString());
                }
                catch
                {
                }
                return false;
            }
            //catch(UnauthorizedAccessException)
            //{
            //    WinForms.MessageBox.Show("No permission to send message to tabbles' pipe.");
            //}

        }

        public static void showMessageTabblesIsNotRunning()
        {
            if (ThisAddIn.isConfidential)
            {
                MessageBox.Show(Res.MsgConfidIsNotRunning3, Res.info);
            }
            else
            {
                MessageBox.Show(Res.MsgTabblesIsNotRunning3, Res.info);
            }
        }

        public static void showMessageSyncSentToTabbles()
        {
            MessageBox.Show(Res.messageSyncSentToTabbles.sd());
            
        }

        public void openQuickTagAndShowResultInOutlook_safe()
        {

            try
            {
                var xelRoot = new XElement("quick_open_tags_in_outlook");
                var xdoc = new XDocument(xelRoot);
                var tabblesWasRunning = sendXmlToTabbles(xdoc);
                if (!tabblesWasRunning)
                    showMessageTabblesIsNotRunning();
            }
            catch (Exception ec)
            {
                try
                {
                    var crashId = "outlook-addin: error openQuickTagAndShowResultInOutlook ";
                    var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(ec);
                    var str = crashId + stackTrace;
                    Log.log(str);
                    
                }
                catch
                {
                }


            }
        }


        public void TagEmailsWithTabbles_safe(List<MailItem> mails)
        {
            try
            {
                var emails = (from m in mails
                              let atSubj = new XAttribute("subject", m.Subject ?? "")
                              let atCmdLine = new XAttribute("command_line", outlookPrefix + ThisAddIn.getEntryIdDebug(m, "kflbjfghwjkkfre"))
                              let ats = new[] { atSubj, atCmdLine }
                              select new XElement("email", ats)).ToArray();
                var xelRoot = new XElement("i_need_to_tag_emails", emails);
                var xdoc = new XDocument(xelRoot);
                var tabblesWasRunning = sendXmlToTabbles(xdoc);
                if (!tabblesWasRunning)
                    showMessageTabblesIsNotRunning();

            }
            catch (Exception ec)
            {
                try
                {
                    var crashId = "outlook-addin: error TagEmailsWithTabbles ";
                    var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(ec);
                    var str = crashId + stackTrace;
                    Log.log(str);
                    
                }
                catch
                {
                }
            }

            // todo 
            //if (SendMessageToTabbles == null)
            //{
            //    return;
            //}

            //var emails = (from MailItem mi in this.selectedMails
            //              select new Generic
            //              {
            //                  name = mi.Subject,
            //                  commandLine = this.outlookPrefix + mi.EntryID,
            //                  icon = new IconOther(),
            //                  showCommandLine = false
            //              }).ToList();

            //SendMessageToTabbles(new INeedToTagGenericsWithTabblesQuickTagDialog()
            //{
            //    gens = emails
            //});
        }

        //private void openInTabblesMenuButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        //{
        //    OpenSelectedEmailInTabbles();
        //}

        public void OpenTheOpenEmailInTabbles_safe()
        {
            try
            {

                var item = outlookApp.ActiveInspector().CurrentItem;
                var mailItem = item as MailItem;
                if (mailItem != null)
                {
                    OpenEmailInTabbles_safe(mailItem);
                }
                else
                {
                    MessageBox.Show(Res.notAnEmail);
                }
            }
            catch (Exception e)
            {
                try
                {
                    var crashId = "outlook-addin: error in open the open email";
                    var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(e);
                    var str = crashId + stackTrace;
                    Log.log(str);
                    
                }
                catch
                {
                }


            }
        }

        public void OpenSelectedEmailInTabbles_safe()
        {
            try
            {
                if (IsAnyEmailSelected(true))
                {
                    OpenEmailInTabbles_safe(this.selectedMails[0]);
                }
                else
                {
                    MessageBox.Show(Res.noEmailSelected);
                }
            }
            catch (Exception e)
            {
                try
                {
                    var crashId = "outlook-addin: error in open selected email in tabbles";
                    var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(e);
                    var str = crashId + stackTrace;
                    Log.log(str);
                    
                }
                catch
                {
                }

            }
        }

        //private void openInTabblesContextMenu_Click(CommandBarButton ctrl, ref bool cancelDefault)
        //{
        //    if (IsAnyEmailSelected(false))
        //    {
        //        OpenEmailInTabbles(this.selectedMails[0]);
        //    }
        //}

        public void OpenEmailInTabbles_safe(MailItem m)
        {
            try
            {
                var atSubj = new XAttribute("subject", m.Subject ?? "");
                var atCmdLine = new XAttribute("command_line", outlookPrefix + ThisAddIn.getEntryIdDebug(m, "nh,klhtuy748"));
                var ats = new[] { atSubj, atCmdLine };
                var xelRoot = new XElement("locate_email", ats);
                var xdoc = new XDocument(xelRoot);
                var tabblesWasRunning = sendXmlToTabbles(xdoc);
                if (!tabblesWasRunning)
                    showMessageTabblesIsNotRunning();
            }
            catch (Exception ec)
            {
                try
                {
                    var crashId = "outlook-addin: error open email in tabbles";
                    var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(ec);
                    var str = crashId + stackTrace;
                    Log.log(str);
                    
                }
                catch
                {
                }

            }
        }


        //private void tabblesSearch_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        //{
        //    // 
        //}

        //private void TabblesSearch()
        //{
        //    // todo
        //    //if (SendMessageToTabbles == null)
        //    //{
        //    //    return;
        //    //}

        //    //SendMessageToTabbles(new INeedToOpenSearch());
        //}

        //private void syncWithTabbles_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        //{
        //    // todo implem
        //    //if (StartSync != null)
        //    //{
        //    //    StartSync();
        //    //}
        //}

        private bool IsAnyEmailSelected(bool fillAtFirst)
        {
            if (fillAtFirst)
            {
                try
                {
                    FillSelectedMails(this.outlookApp.ActiveExplorer().Selection);
                }
                catch
                {
                    return false;
                }
            }

            return (this.selectedMails != null && this.selectedMails.Count > 0);
        }

        private void FillSelectedMails(Selection selection)
        {
            if (selection.Count > 0 && selection[1] is MailItem)
            {
                if (this.selectedMails == null)
                {
                    this.selectedMails = new List<MailItem>();
                }
                else
                {
                    this.selectedMails.Clear();
                }

                foreach (var sel in selection)
                {
                    var item = sel as MailItem;
                    if (item != null)
                    {
                        MailItem mail = item;
                        this.selectedMails.Add(mail);
                    }
                }
            }
            else if (this.selectedMails != null)
            {
                this.selectedMails.Clear();
            }
        }

        //private void FillItemsToListen()
        //{
        //    if (this.currentFolderItems != null)
        //    {
        //        try
        //        {
        //            this.currentFolderItems.ItemChange -= Items_ItemChange;
        //        }
        //        catch (System.Exception)
        //        {
        //        }
        //    }

        //    Folder currentFolder = (Folder)this.outlookApp.ActiveExplorer().CurrentFolder;

        //    if (currentFolder != null)
        //    {
        //        this.currentFolderItems = currentFolder.Items;

        //        //avoid double adding
        //        this.currentFolderItems.ItemChange -= Items_ItemChange;
        //        this.currentFolderItems.ItemChange += Items_ItemChange;
        //    }
        //}

        //private void Items_ItemChange(object item)
        //{
        //    if (item is MailItem)
        //    {
        //        MailItem mail = (MailItem)item;
        //        string mailId = mail.EntryID;
        //        //lock (this.syncObj)
        //        {
        //            //if (this.onceItemChanged.Contains(mailId))
        //            //{
        //            //    this.onceItemChanged.Remove(mailId);
        //            //}
        //            //else
        //            if (InternallyChangedMailIds.Contains(mailId))
        //            {
        //                InternallyChangedMailIds.Remove(mailId);
        //                //this.onceItemChanged.Add(mailId);
        //            }
        //            else
        //            {
        //                SendCategoriesToTabbles(mail);
        //                {
        //                    //this.onceItemChanged.Add(mailId);
        //                }
        //            }
        //        }
        //    }
        //}

        //private void SendCategoriesToTabbles(MailItem mail)
        //{

        //    var categoriesWithColors = new Dictionary<string, string>();
        //    string[] categories = Utils.GetCategories(mail);
        //    foreach (string categoryName in categories)
        //    {
        //        try
        //        {
        //            Category category = this.outlookApp.Session.Categories[categoryName];
        //            if (category != null)
        //            {
        //                string categoryRgb = Utils.GetRgbFromOutlookColor(category.Color);
        //                categoriesWithColors[categoryName] = categoryRgb;
        //            }
        //        }
        //        catch (System.Exception)
        //        {
        //            //ignore the category
        //        }
        //    }

        //    if (!string.IsNullOrEmpty(mail.FlagRequest))
        //    {
        //        categoriesWithColors[mail.FlagRequest] = Utils.GetRgbForFlagRequest(mail.FlagRequest);
        //    }

        //}

        //public void AddEntryIdToSkip(string entryId)
        //{
        //    this.itemsToSkipChanges.Add(entryId);
        //}

        //private IPictureDisp GetPictureDispFromResource(string resourceName)
        //{
        //    object resource = Res.ResourceManager.GetObject(resourceName);
        //    if (resource is Image)
        //    {
        //        return ImageConverter.GetPictureDisp((Image)resource);
        //    }

        //    return null;
        //}
    }
}
