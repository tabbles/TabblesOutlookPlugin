using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using io = System.IO;


namespace Tabbles.OutlookAddIn
{
    public static class Utils
    {


        public static string sd(this string s)
        {
            var prodName = ThisAddIn.isConfidential ? "Confidential" : "Tabbles";
            return s.Replace("{PRODUCT}", prodName);
        }

        public static T throwsWrapper<T>(Func<T> f, Func<System.Exception, T> contIfThrows)
        {
            try
            {

                return f();
            }
            catch (System.Exception e)
            {
                return contIfThrows(e);
            }
        }

        //public static A trySafe<A>(A valueIfThrows,  Func<A> f )
        //{
        //    try
        //    {
        //        return f.Invoke();
        //    }
        //    catch (Exception e)
        //    {
        //        CrashReportFsharp.sendSilentCrashIfEnoughTimePassed2(
        //    }
        //}

        public static void appendToXmlCore(int maxNumOperationsInXml, string xmlLogPath, XElement xel, string rootElementName)
        {
            var parentDir = System.IO.Path.GetDirectoryName(xmlLogPath);
            if (!io.Directory.Exists(parentDir))
            {
                io.Directory.CreateDirectory(parentDir);
            }

            io.FileStream fs;
            bool xmlWasCreated;
            {
                try
                {
                    fs = new io.FileStream(xmlLogPath, io.FileMode.Open, io.FileAccess.ReadWrite);
                    xmlWasCreated = false;
                }
                catch (io.FileNotFoundException)
                {
                    fs = new io.FileStream(xmlLogPath, io.FileMode.Create, io.FileAccess.ReadWrite);
                    xmlWasCreated = true;
                }
            }



            var xdoc = new Func<XDocument>(() =>
            {
                var createNew = new Func<XDocument>(() =>
                {
                    var root = new XElement(rootElementName, xel);
                    return new XDocument(root);
                });

                if (xmlWasCreated)
                {
                    return createNew();
                }
                else
                {
                    try
                    {
                        var xdoc2 = XDocument.Load(fs);
                        xdoc2.Root.Add(xel);
                        return xdoc2;

                    }
                    catch (System.Xml.XmlException)
                    {
                        return createNew();
                    }
                }
            })();

            {

                var count = xdoc.Root.Elements().Count();
                if (count > maxNumOperationsInXml)
                {
                    var howManyToRemove = count - maxNumOperationsInXml;
                    var elementsToRemove = xdoc.Root.Elements().Take(howManyToRemove).ToList();
                    foreach (var el in elementsToRemove)
                    {
                        el.Remove();
                    }
                }
            }

            {
                fs.SetLength(0);
                xdoc.Save(fs);
                fs.Dispose();

            }

        }

        public static string getTabblesRootDirUser()
        {
            string folderName = ThisAddIn.isConfidential ? "Confidential" : "Tabbles";
            return (System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData), folderName));
        }
        public static string xmlLogPath()
        {
            return System.IO.Path.Combine(getTabblesRootDirUser(), "outlook-operations-log.xml");
        }
        public static string errorLogPath()
        {
            return System.IO.Path.Combine(getTabblesRootDirUser(), "outlook-operations-errors.xml");
        }


        public static void appendToXml(XElement xel)
        {
            appendToXmlCore(2000, xmlLogPath(), xel, "outlook_operations_log");
        }

        public static void appendToXmlError(XElement xel)
        {
            appendToXmlCore(10, xmlLogPath(), xel, "outlook_operations_errors");
        }
        /*
         * 

let rec create_message_of_exception (ex: Exception) msg  level = 
        let hres = Marshal.GetHRForException(ex).ToString()

        let nl = Environment.NewLine + Environment.NewLine
        let sqlDetail = 
                match ex with
                | :? System.Data.SqlClient.SqlException as e ->
                        "sql exception. Number = " + e.Number.ToString() + nl
                        + " - LineNumber = " + e.LineNumber.ToString() + nl
                        

                |  _ -> ""
        if ex = null then
                msg
                //"-------------------------------------------------------\n\nException level " + level + ":\n\n" +  msg
        else
                //"-------------------------------------------------------\n\nException level " + level + ":\n\n" + 
                let msg' = msg + nl +  " -------------------- Level " + string level + " -------------" + nl +  "Exception type = " + ex.GetType().ToString() + nl 
                                + "hresult = " + hres  + nl
                                + sqlDetail +  "Message = " + ex.Message + nl + "StackTrace = " + ex.StackTrace
                create_message_of_exception  ex.InnerException   msg'  (level + 1)


let stringOfException (e: Exception) = 
        create_message_of_exception e "" 0
         * 
         * */

        public static string emptyStringOfNull(string s)
        {
            if (s == null)
                return "";
            else
                return s;

        }
        //private static string createMessageOfException(System.Exception ex, int level, string initialMessage)
        //{

        //    var hres = Marshal.GetHRForException(ex).ToString();

        //    var nl = Environment.NewLine + Environment.NewLine;
        //    var sqlDetail = "";
        //    if (ex == null)
        //    {
        //        return initialMessage;
        //    }
        //    else
        //    {
        //        var msg = initialMessage + nl + " -------------------- Level " + level.ToString() + " -------------" + nl + "Exception type = " + ex.GetType().ToString() + nl
        //                        + "hresult = " + hres + nl
        //                        + sqlDetail + "Message = " + ex.Message + nl + "StackTrace = " + emptyStringOfNull( ex.StackTrace);
        //        return createMessageOfException(ex.InnerException, level + 1, msg);
        //    }


        //}

        //public static string stringOfException(System.Exception e)
        //{
        //    return createMessageOfException(e, 0, "");
        //}

        public const int MAPI_E_COLLISION = -2147219964;

        private static readonly Dictionary<Outlook.OlCategoryColor, string> OutlookColorsStr = new Dictionary<Outlook.OlCategoryColor, string>()
        {
            {Outlook.OlCategoryColor.olCategoryColorBlack, "4F4F4F"},
            {Outlook.OlCategoryColor.olCategoryColorBlue, "9DB7E8"},
            {Outlook.OlCategoryColor.olCategoryColorDarkBlue, "2858A5"},
            {Outlook.OlCategoryColor.olCategoryColorDarkGray, "6F6F6F"},
            {Outlook.OlCategoryColor.olCategoryColorDarkGreen, "3F8F2B"},
            {Outlook.OlCategoryColor.olCategoryColorDarkMaroon, "93446B"},
            {Outlook.OlCategoryColor.olCategoryColorDarkOlive, "778B45"},
            {Outlook.OlCategoryColor.olCategoryColorDarkOrange, "E2620D"},
            {Outlook.OlCategoryColor.olCategoryColorDarkPeach, "C79930"},
            {Outlook.OlCategoryColor.olCategoryColorDarkPurple, "5C3FA3"},
            {Outlook.OlCategoryColor.olCategoryColorDarkRed, "C11A25"},
            {Outlook.OlCategoryColor.olCategoryColorDarkSteel, "6B7994"},
            {Outlook.OlCategoryColor.olCategoryColorDarkTeal, "329B7A"},
            {Outlook.OlCategoryColor.olCategoryColorDarkYellow, "B9B300"},
            {Outlook.OlCategoryColor.olCategoryColorGray, "BFBFBF"},
            {Outlook.OlCategoryColor.olCategoryColorGreen, "78D168"},
            {Outlook.OlCategoryColor.olCategoryColorMaroon, "DAAEC2"},
            {Outlook.OlCategoryColor.olCategoryColorNone, "FFFFFF"},
            {Outlook.OlCategoryColor.olCategoryColorOlive, "C6D2B0"},
            {Outlook.OlCategoryColor.olCategoryColorOrange, "F9BA89"},
            {Outlook.OlCategoryColor.olCategoryColorPeach, "F7DD8F"},
            {Outlook.OlCategoryColor.olCategoryColorPurple, "B5A1E2"},
            {Outlook.OlCategoryColor.olCategoryColorRed, "E7A1A2"},
            {Outlook.OlCategoryColor.olCategoryColorSteel, "DAD9DC"},
            {Outlook.OlCategoryColor.olCategoryColorTeal, "9FDCC9"},
            {Outlook.OlCategoryColor.olCategoryColorYellow, "FCFA90"}
        };

        private const string OutlookColorStrFollowUp = "F6532F";

        private static readonly string[] CategorySeparator = new string[] { ";" , ","}; // entrambi, perché sul mio sistema è ; , ma per questo utente era ,
                                                                                            //  https://mail.google.com/mail/u/0/#inbox/1464268c706c39c3

        private const string RegKeyTabbles = @"SOFTWARE\Yellow Blue Soft\Tabbles";
        private const string RegValueTabblesInstallDir = "installation_dir";

        private static readonly Dictionary<Outlook.OlCategoryColor, System.Drawing.Color> OutlookColorsRgb
            = new Dictionary<OlCategoryColor, System.Drawing.Color>();

        static Utils()
        {
            foreach (var outlookColorStr in OutlookColorsStr)
            {
                System.Drawing.Color color = System.Drawing.ColorTranslator.FromHtml("#" + outlookColorStr.Value);
                OutlookColorsRgb.Add(outlookColorStr.Key, color);
            }
        }

        public static OutlookVersion ParseMajorVersion(Outlook.Application outlookApplication)
        {
            string majorVersionString = outlookApplication.Version.Split(new char[] { '.' })[0];
            switch (majorVersionString)
            {
                case "11":
                    return OutlookVersion.OUTLOOK_2003;
                case "12":
                    return OutlookVersion.OUTLOOK_2007;
                case "14":
                    return OutlookVersion.OUTLOOK_2010;
                default:
                    return OutlookVersion.UNKNOWN;
            }
        }

        public static string GetOutlookPrefix()
        {
            string path = GetOutlookPath();
            return "\"" + path + @""" /select outlook:";
        }

        private static string GetOutlookPath()
        {
            // Fetch the Outlook Class ID
            var key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Classes\\Outlook.Application\\CLSID");
            var objOutlookClassID = key.GetValue("");
            var outlookClassId = ((string)objOutlookClassID).Trim();
            key.Dispose();

            // Using the class ID from above pull up the path
            var key2 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\CLSID\" + outlookClassId + @"\LocalServer32");
            var outlookPath = ((string)key2.GetValue("")).Trim();
            key2.Dispose();

            return outlookPath;
        }

        public static string GetTabblesInstallDir()
        {
            RegistryKey key = Registry.LocalMachine.OpenSubKey(RegKeyTabbles) ??
                              Registry.CurrentUser.OpenSubKey(RegKeyTabbles);

            if (key != null)
            {
                return key.GetValue(RegValueTabblesInstallDir) as string;
            }

            return null;
        }

        public static Outlook.OlCategoryColor GetOutlookColorFromRgb(string rgb)
        {
            if (!string.IsNullOrEmpty(rgb) && (rgb = rgb.Trim()).Length != 9)
            {
                return OlCategoryColor.olCategoryColorNone;
            }

            string rStr = rgb.Substring(3, 2);
            string gStr = rgb.Substring(5, 2);
            string bStr = rgb.Substring(7, 2);

            int r;
            int g;
            int b;
            try
            {
                r = Int32.Parse(rStr, System.Globalization.NumberStyles.HexNumber);
                g = Int32.Parse(gStr, System.Globalization.NumberStyles.HexNumber);
                b = Int32.Parse(bStr, System.Globalization.NumberStyles.HexNumber);

                OlCategoryColor olColor = OlCategoryColor.olCategoryColorNone;

                //calculate nearest color with two algorithms
                double minDistance = double.MaxValue;
                foreach (var olColorRgb in OutlookColorsRgb)
                {
                    int rDiff = r - olColorRgb.Value.R;
                    int gDiff = g - olColorRgb.Value.G;
                    int bDiff = b - olColorRgb.Value.B;

                    //1. consider RGB as a three-dimensional space and get the nearest color
                    double curDistance = Math.Sqrt(rDiff * rDiff + gDiff * gDiff + bDiff * bDiff);
                    //2. count the total difference of color and choose the nearest one
                    curDistance += Math.Abs(rDiff) + Math.Abs(gDiff) + Math.Abs(bDiff);
                    if (curDistance < minDistance)
                    {
                        minDistance = curDistance;
                        olColor = olColorRgb.Key;
                    }
                }

                return olColor;
            }
            catch (System.Exception)
            {
                return OlCategoryColor.olCategoryColorNone;
            }
        }

        public static string GetRgbFromOutlookColor(Outlook.OlCategoryColor color)
        {
            //Tabbles needs # and alpha value in addition to RGB
            return "#FF" + OutlookColorsStr[color];
        }

        public static string GetRgbForFlagRequest(string flagRequest)
        {
            return "#FF" + OutlookColorStrFollowUp;
        }

        /// <summary>
        /// Returns array of categories of given mail item.
        /// </summary>
        /// <param name="mail"></param>
        /// <returns></returns>
        public static string[] GetCategories(MailItem mail)
        {
            var fl = mail.FlagRequest;
            if (mail.Categories != null)
            {
                var spl = mail.Categories.Split(CategorySeparator, StringSplitOptions.None);
                var x = (from q in spl
                         select q.Trim());
                return x.ToArray();
            }
            else
            {
                return new string[] { };
            }


        }

        /// <summary>
        /// Removes given folder from its parent.
        /// </summary>
        /// <param name="folder"></param>
        public static void RemoveFolder(Folder folder)
        {
            try
            {
                if (folder != null)
                {
                    object parentObj = folder.Parent;
                    var obj = parentObj as Folder;
                    if (obj != null)
                    {
                        Folders subfolders = obj.Folders;
                        for (int i = 1, length = subfolders.Count; i <= length; i++) //folder index is 1-based
                        {
                            if (subfolders[i].Name == folder.Name)
                            {
                                subfolders.Remove(i);
                            }
                        }
                    }
                }
            }
            catch (System.Exception)
            {
                //ignore this exception
            }
        }

        public static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
                }
                catch (System.Exception)
                {
                    //do nothing
                }
            }
        }

        public static T c<T>(Func<T> f)
        {
            return f();
        }
    }

    public enum OutlookVersion
    {
        OUTLOOK_2003,
        OUTLOOK_2007,
        OUTLOOK_2010,
        UNKNOWN
    }
}
