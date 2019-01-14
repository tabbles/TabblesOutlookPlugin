using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace Tabbles.OutlookAddIn
{
    public static class Log
    {
        private static string getLogFilePath()
        {
            string folderDocs;
            
            // application data. non va bene per tabbles vecchio forked, ma tanto non va bene nemmeno il nome plugin tagger.vsto.
            folderDocs = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData);
            

            string folderName;
            if (ThisAddIn.isConfidential)
                folderName = "Confidential";
            else
                folderName = "Tabbles";

            var tabblesFolder = System.IO.Path.Combine(folderDocs, folderName);

            System.IO.Directory.CreateDirectory(tabblesFolder);
            return (System.IO.Path.Combine(tabblesFolder, "log_outlook_addin.txt"));
        }

        /// <summary>
        /// Meant to be called every 10 minutes or so.
        /// </summary>
        public static void deleteLogIfTooLong()
        {

            var logFilePath = getLogFilePath();

            var fi = new System.IO.FileInfo(logFilePath);
            if (fi.Exists && fi.Length > 1000000)
            {
                System.IO.File.Delete(logFilePath);
            }
            


        }

        public  static void corpoThreadLog()
        {
            var count = 0;
            var logFilePath = getLogFilePath();
            while (true)
            {
                try
                {
                    Thread.Sleep(2400);

                    if (count % 1000 == 0)
                        deleteLogIfTooLong();


                    var righe = new List<string>();
                    lock (CrashReportFsharp.g_lock)
                    {
                        righe = mLog.ToList();
                        mLog.Clear();
                    }

                    System.IO.File.AppendAllLines(path: logFilePath, contents: righe);
                }
                catch (Exception ee)
                {
                    try
                    {
                        // probably an access problem (some other thread might be writing to the log).

                        //var crashId =
                        //    "outlook-addin: error: thread which deletes the log when it is too big. it will probably work next time.  ";
                        //var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(ee);

                        //CrashReportFsharp.sendSilentCrashIfEnoughTimePassed3(ThisAddIn.logError, stackTrace, crashId,
                        //    Environment.UserName ?? "", Environment.MachineName ?? "");
                    }
                    catch
                    {
                    }
                    finally
                    {
                        // non scrivere nel log qui... non mi sento sicuro
                        Thread.Sleep(2000); // retry in 2 seconds
                    }
                }
            }
        }

        private static Queue<string> mLog = new Queue<string>(); 
        public static void log(string txt)
        {
            try
            {
                string app = stringAppConfidOrTabbles();
                //var logFilePath = getLogFilePath();
                var riga = DateTime.Now + ": " + app + ": " + txt + System.Environment.NewLine + System.Environment.NewLine;

                lock (CrashReportFsharp.g_lock)
                {
                    mLog.Enqueue(riga);
                }

                //using (var sw = System.IO.File.AppendText(logFilePath))
                //{

                //    sw.WriteLine(riga);
                //}
            }
            catch { }
        }

        public static string stringAppConfidOrTabbles()
        {
            return ThisAddIn.isConfidential ? "confid" : "tabbles";
        }
    }
}
