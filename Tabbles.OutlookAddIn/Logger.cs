using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Tabbles.OutlookAddIn
{
    /// <summary>
    /// Logs messages in "outlook-plugin-log.txt" file of My Documents folder. If an exception occurred
    /// during opening a stream nothing happens, but messages will not be logged.
    /// </summary>
    class Logger
    {
        private const string LogFileName = "outlook-plugin-log.txt";
        private static Logger instance;

        private StreamWriter writer;

        private bool IsInited
        {
            get;
            set;
        }

        private Logger()
        {
            string myDocsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string loggerPath = Path.Combine(myDocsFolder, LogFileName);
            try
            {
                writer = File.CreateText(loggerPath);
                IsInited = true;
            }
            catch (Exception)
            {
                IsInited = false;
            }
        }

        /// <summary>
        /// Logs message with current timestamp into the log file.
        /// </summary>
        /// <param name="message"></param>
        public static void Log(string message)
        {
            if (instance == null)
            {
                instance = new Logger();
            }

            if (instance.IsInited)
            {
                instance.LogInternal(message);
            }
        }

        public static void Dispose()
        {
            if (instance != null && instance.IsInited)
            {
                instance.DisposeInternal();
            }
        }

        private void LogInternal(string message)
        {
            lock (this)
            {
                message = DateTime.Now.ToString() + ": " + message;
                try
                {
                    writer.WriteLine(message);
                    writer.Flush();
                }
                catch (Exception)
                {
                    //do nothing
                }
            }
        }

        /// <summary>
        /// Closes the writer stream.
        /// </summary>
        private void DisposeInternal()
        {
            if (writer != null)
            {
                try
                {
                    writer.Close();
                }
                catch (Exception)
                { }
            }
        }
    }
}
