using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Res = Tabbles.OutlookAddIn.Properties.Resources;
using System.Drawing;

using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using IWshRuntimeLibrary;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace Tabbles.OutlookAddIn
{
    [ComVisible(true)]
    public class TabblesRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        //public event EventHandler TagEmailsWithTabbles;
        //public event EventHandler OpenEmailInTabbles;
        //public event EventHandler TabblesSearch;
        //public event EventHandler SyncWithTabbles;

        //public event IsAnyEmailSelectedHandler IsAnyEmailSelected;

        public MenuManager mMenuManager;
        public ThisAddIn mAddin;
        public TabblesRibbon()
        {
        }


        private Dictionary<string, Tuple<string, Bitmap>> Files { get; set; }

        public void OnAddAttahment(Office.IRibbonControl control)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                if (MessageBox.Show("Is the file " + ofd.FileName + " OK to attach?", "Good to attach",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    //MessageBox.Show("We add a file as attachment to curremt mailitem");
                    var item = control.Context as Inspector;
                    if (item != null)
                    {
                        var mailItem = item.CurrentItem as MailItem;
                        if (mailItem != null)
                        {
                            mailItem.Attachments.Add(ofd.FileName,
                                OlAttachmentType.olByValue,
                                1,
                                Path.GetFileName(ofd.FileName));
                        }

                    }
                }
            }
        }

        public void OnAddAttahmentClicked(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            var file = Files.ToList()[selectedIndex];
            MessageBox.Show(@"Do something with file " + file.Key);
            if (MessageBox.Show("Is the file " + file.Key + " OK to attach?", "Good to attach",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                //MessageBox.Show("We add a file as attachment to curremt mailitem");
                var item = control.Context as Inspector;
                if (item != null)
                {
                    var mailItem = item.CurrentItem as MailItem;
                    if (mailItem != null)
                    {
                        mailItem.Attachments.Add(file.Key,
                            OlAttachmentType.olByValue,
                            1,
                            Path.GetFileName(file.Key));
                    }

                }
            }
        }

        public Bitmap GetImage(Office.IRibbonControl control, int itemIndex)
        {
            return Files.ToList()[itemIndex].Value.Item2;
            //try
            //{
            //    if (itemIndex < Files.Length)
            //    {
            //        var filename = Files[itemIndex];
            //        var fi = new FileInfo(filename);
            //        if (fi.FullName.ToLower().EndsWith("lnk"))
            //        {
            //            IWshShell shell = new WshShell();
            //            var lnk = shell.CreateShortcut(fi.FullName) as IWshShortcut;
            //            if (lnk != null)
            //            {
            //                return Icon.ExtractAssociatedIcon(lnk.TargetPath).ToBitmap(); ;
            //            }
            //            else
            //            {
            //                return Icon.ExtractAssociatedIcon(filename).ToBitmap(); ;
            //            }

            //        }
            //        else
            //        {
            //            return Icon.ExtractAssociatedIcon(filename).ToBitmap(); ;
            //        }

            //    }

            //}
            //catch
            //{
            //}


            //return null;
        }


        public int GetItemCount(Office.IRibbonControl control)
        {
            //return Files.Length;
            return Files.Count;
        }

        /// <summary>
        ///  attenzione, cambiare nome, se no si confonde con l'altro menu e non mostra i titoli sotto i 3 pulsanti nella ribbon
        /// </summary>
        /// <param name="ribbonID"></param>
        /// <returns></returns>
        //public string GetLabel(Office.IRibbonControl control, int index)
        //{
        //    return Files.ToList()[index].Value.Item1;

        //    //try
        //    //{
        //    //    if (index < Files.Length)
        //    //    {
        //    //        var filename = Files[index];
        //    //        var fi = new FileInfo(filename);
        //    //        if (fi.FullName.ToLower().EndsWith("lnk"))
        //    //        {
        //    //            IWshShell shell = new WshShell();
        //    //            var lnk = shell.CreateShortcut(fi.FullName) as IWshShortcut;
        //    //            if (lnk != null)
        //    //            {
        //    //                return Path.GetFileName(lnk.TargetPath);
        //    //            }
        //    //            else
        //    //            {
        //    //                return Path.GetFileName(filename);
        //    //            }
        //    //        }
        //    //        else
        //    //        {
        //    //            return Path.GetFileName(filename);
        //    //        }
        //    //    }
        //    //}
        //    //catch
        //    //{
        //    //}

        //    //return null;
        //}

        


        public string GetCustomUI(string ribbonID)
        {
            try
            {
                if (ribbonID == "Microsoft.Outlook.Explorer")
                {
                    return GetResourceText("Tabbles.OutlookAddIn.RibbonExplorer.xml");
                }
                else if (ribbonID == "Microsoft.Outlook.Mail.Compose" ||
                    ribbonID == "Microsoft.Outlook.Mail.Read")
                {
                    return GetResourceText("Tabbles.OutlookAddIn.RibbonInspector.xml");
                }

                return null;
            }
            catch (System.Exception e)
            {
                var crashId = "outlook-addin: error get custom ui ";
                var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(e);
                var str = crashId + stackTrace;
                Log.log(str);
                


                return null;
            }
        }


        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;








            //Files = new Dictionary<string, Tuple<string, Bitmap>>();


            //var directory = Environment.GetFolderPath(Environment.SpecialFolder.Recent);
            //var files = Directory.EnumerateFiles(directory).ToArray();

            //foreach (string file in files)
            //{
            //    try
            //    {
            //        var fi = new FileInfo(file);
            //        if (fi.FullName.ToLower().EndsWith("lnk"))
            //        {
            //            IWshShell shell = new WshShell();
            //            var lnk = shell.CreateShortcut(fi.FullName) as IWshShortcut;
            //            if (lnk != null)
            //            {
            //                try
            //                {
            //                    var bitmap = Icon.ExtractAssociatedIcon(lnk.TargetPath).ToBitmap();
            //                    var label = Path.GetFileName(lnk.TargetPath);

            //                    Files.Add(lnk.TargetPath, new Tuple<string, Bitmap>(label, bitmap));
            //                }
            //                catch (System.IO.FileNotFoundException)
            //                {

            //                }
            //            }
            //        }
            //        else
            //        {
            //            var bitmap = Icon.ExtractAssociatedIcon(file).ToBitmap();
            //            var label = Path.GetFileName(file);

            //            Files.Add(file, new Tuple<string, Bitmap>(label, bitmap));
            //        }
            //    }
            //    catch (System.Exception e)
            //    {
            //    }
            //}



        }

        public void OnAction(Office.IRibbonControl control)
        {

            try
            {
                switch (control.Id)
                {
                    case "tagUsingTabblesButton":
                    case "tagUsingTabblesMenuSingle":
                    case "tagUsingTabblesMenuMultiple":
                        mMenuManager.TagSelectedEmailsWithTabbles();
                        break;

                    case "tagUsingTabblesButtonSingleMail":
                        mMenuManager.TagOpenEmailWithTabbles();
                        break;
                    case "openInTabblesButton":
                    case "openInTabblesMenu":
                        mMenuManager.OpenSelectedEmailInTabbles_safe();
                        break;

                    case "openInTabblesButtonSingleMail":
                        mMenuManager.OpenTheOpenEmailInTabbles_safe();
                        break;
                    case "tabblesSearchButton":
                        mMenuManager.openQuickTagAndShowResultInOutlook_safe();
                        break;
                    case "syncWithTabblesButton":
                        mAddin.importOutlookTaggingIntoTabbles();
                        //if (SyncWithTabbles != null)
                        //{
                        //    SyncWithTabbles(control, EventArgs.Empty);
                        //}
                        break;
                    default:
                        break;
                }
            }
            catch (System.Exception e)
            {

                var crashId = "outlook-addin: error on-action";
                var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(e);
                var str = crashId + stackTrace;
                Log.log(str);
                


            }
        }

        public string getLabelTabTitle(Office.IRibbonControl control)
        {
            if (ThisAddIn.isConfidential)
            {
                return "Confidential";
            }
            else
            {
                return "Tabbles";
            }
        }

        public string GetLabel(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "tagUsingTabblesButton":
                    case "tagUsingTabblesButtonSingleMail":
                    case "tagUsingTabblesMenuSingle":
                    case "tagUsingTabblesMenuMultiple":
                        return Res.MenuTagUsingTabbles2;
                    case "openInTabblesButton":
                    case "openInTabblesButtonSingleMail":
                    case "openInTabblesMenu":
                        return Res.MenuOpenInTabbles2;
                    case "tabblesSearchButton":
                        return Res.MenuTabblesSearch3;
                    case "syncWithTabblesButton":
                        return Res.MenuSyncWithTabbles3;
                    default:
                        return string.Empty;
                }
            }
            catch (System.Exception e)
            {
                var crashId = "outlook-addin: error get label ";
                var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(e);
                var str = crashId + stackTrace;
                Log.log(str);
                


                return string.Empty;
            }
        }

        public Bitmap OnLoadImage(string imageName)
        {
            try
            {
                Bitmap image = null;

                switch (imageName)
                {
                    case "tag_using_tabbles":
                        image = Res.Outlook_32x32_tag;
                        break;
                    case "open_in_tabbles":
                        image = Res.Outlook_32x32_open;
                        break;
                    case "search":
                        image = Res.Outlook_32x32_search;
                        break;
                    case "sync_with_tabbles":
                        image = Res.Outlook_32x32_sync;
                        break;
                    case "tag_using_tabbles_small":
                        image = Res.Outlook_16x16_tag;
                        break;
                    case "open_in_tabbles_small":
                        image = Res.Outlook_16x16_open;
                        break;
                    default:
                        break;
                }

                return image;
            }
            catch (System.Exception e)
            {
                var crashId = "outlook-addin: error onloadimage ";
                var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(e);
                var str = crashId + stackTrace;
                Log.log(str);
                


                return null;
            }
        }

        //public bool IsAnythingSelected(Office.IRibbonControl control)
        //{
        //    return IsAnyEmailSelected != null && IsAnyEmailSelected();
        //}

        #endregion

        #region Helpers

        /// <summary>
        /// this function returns the content of a resource file as a string. 
        /// </summary>
        /// <param name="resourceName">the file name. this file must be an embedded resource.</param>
        /// <returns></returns>
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
