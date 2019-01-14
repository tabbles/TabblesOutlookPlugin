using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.FSharp.Core;
namespace Tabbles.OutlookAddIn
{
    public static class ThreadUtils
    {

        /*
        let execInThread f =
        if Threading.Thread.CurrentThread.IsBackground then // modo rapido per capire se siamo già in un thread
                f ()
        else
                let body () = 
                        try
                                f()
                        with
                        | ecc  ->
                             showCrashDialog ecc None "execInThread - "
                (  // faccio partire il thread
                let th = new System.Threading.Thread(body)
                th.CurrentUICulture <-  new Globalization.CultureInfo(g_lang) 
                th.Priority <- Threading.ThreadPriority.Normal
                th.IsBackground <- true
                th.Start ()                
                )

         * 
         * */


        /// <summary>
        /// Execute a given piece of code in background, in a thread, in order to return quickly.
        /// </summary>
        /// <param name="dobj"></param>
        /// <param name="a"></param>
        //public static void execInThread(Action a)
        //{
        //    if (System.Threading.Thread.CurrentThread.IsBackground)
        //    {
        //        // we are already in a thread, no need to start another.

        //        a.Invoke();
        //    }
        //    else
        //    {

        //        var th = new Task((() =>
        //        {
        //            try
        //            {
        //                a.Invoke();
        //            }
        //            catch (Exception e)
        //            {
        //                Log.log(">>> execInThread: exception:" + Utils.stringOfException(e));
        //            }

        //        }));

        //        th.Start();
        //    }

        //}

        //public static void execInThreadForceNewThread(Action a)
        //{

        //    var th = new Task((() =>
        //    {
        //        try
        //        {
        //            a.Invoke();
        //        }
        //        catch (Exception e)
        //        {
        //            Log.log(">>> execInThreadForceNewThread: exception:" + Utils.stringOfException(e));
        //        }

        //    }));

        //    // th.CurrentUICulture = ...
        //    //th.Priority = ThreadPriority.Normal;
        //    //th.IsBackground = true;
        //    th.Start();

        //}
        public static void gui(System.Windows.Threading.DispatcherObject dobj, Action a)
        {

            var aWrapped = new Action (() =>
            {
                try
                {
                    a();
                }
                catch (Exception e)
                {
                    var crashId = "outlook-addin: crash in gui fun";
                    var stackTrace = ThisAddIn.AssemblyVer_safe() + CrashReportFsharp.stringOfException(e);
                    var str = crashId + stackTrace;
                    Log.log(str);
                    



                }
            });

            dobj.Dispatcher.Invoke(aWrapped, null);

        }
    }
}
